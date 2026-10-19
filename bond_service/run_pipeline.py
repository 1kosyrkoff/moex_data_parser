"""One bounded update run. Use a fresh interpreter for every run (upstream cache)."""
import argparse
from datetime import date, datetime, timedelta, timezone
import hashlib
import logging
import os
from pathlib import Path
import time
from uuid import uuid4

import moex_bonds
import moex_cls

from .storage import Store, connect, frame_payload, technical_errors, writer_lock

LOG = logging.getLogger(__name__)
MSK = timezone(timedelta(hours=3))


def source_hash():
    digest = hashlib.sha256()
    for module in (moex_bonds, moex_cls):
        digest.update(Path(module.__file__).read_bytes())
    return digest.hexdigest()


def execute(store, args, run_id):
    counts = {'instruments': 0, 'cards': 0, 'yields': 0, 'sectors': 0, 'candles': 0, 'errors': 0}
    params = {**vars(args), 'date': args.date.isoformat(), 'source_sha256': source_hash(),
              'revision': os.environ.get('SOURCE_REVISION', 'unknown')}
    store.start(run_id, params)

    def error(stage, exc, isin=None):
        counts['errors'] += 1
        LOG.warning('run=%s stage=%s isin=%s: %s', run_id, stage, isin, exc)
        store.error(run_id, stage, exc, isin)

    try:
        market = moex_bonds.market_snapshot()
        if market.empty or 'ISIN' not in market:
            raise ValueError('MOEX returned an empty or invalid market list')
        # Preserve source ordering: find_bond() selects the first matching row too.
        valid = market['ISIN'].notna() & market['ISIN'].fillna('').str.strip().ne('')
        market = market.loc[valid].copy()
        market['ISIN'] = market['ISIN'].str.strip().str.upper()
        isins = list(dict.fromkeys(market['ISIN']))
        if args.isins:
            requested = list(dict.fromkeys(v.strip().upper() for v in args.isins.split(',') if v.strip()))
            missing = set(requested) - set(isins)
            if missing:
                raise ValueError('Not in current market list: ' + ', '.join(sorted(missing)))
            isins = requested
        if args.limit:
            isins = isins[:args.limit]
        if not isins:
            raise ValueError('No ISINs selected')
        for isin in isins:
            store.put('instruments', run_id, isin, frame_payload(market[market['ISIN'].eq(isin)]))
        counts['instruments'] = len(isins)

        for offset in range(0, len(isins), args.batch_size):
            batch = isins[offset:offset + args.batch_size]
            for isin in batch:
                try:
                    card = moex_bonds.get_bond_card(isin)
                    store.card(run_id, isin, card)
                    issues = technical_errors(card['schedule'])
                    if issues:
                        error('cards', '; '.join(issues), isin)
                    else:
                        counts['cards'] += 1
                except Exception as exc:
                    error('cards', exc, isin)
                if args.pause:
                    time.sleep(args.pause)

            for stage, function in (
                ('yields', lambda: moex_bonds.get_bond_yields(batch, date=args.date, price=args.price,
                    lookback=args.lookback, assumed_rate=args.assumed_rate, details=True)),
                ('sectors', lambda: moex_bonds.get_issuer_sector(batch)),
            ):
                try:
                    frame = function()
                    for isin in batch:
                        part = frame[frame['ISIN'].eq(isin)]
                        if len(part) != 1:
                            error(stage, 'Expected one result row', isin)
                            continue
                        issues = technical_errors(part)
                        store.put(stage, run_id, isin, frame_payload(part), usable=not issues)
                        if issues:
                            error(stage, '; '.join(issues), isin)
                        else:
                            counts[stage] += 1
                except Exception as exc:
                    error(stage, exc)

            if args.history_days:
                for isin in batch:
                    try:
                        instrument = market[market['ISIN'].eq(isin)].iloc[0]
                        board = args.candle_board or instrument['BOARDID']
                        if not isinstance(board, str) or not board:
                            raise ValueError('No trading board for candles')
                        start = args.date - timedelta(days=args.history_days)
                        previous = store.candle_start(isin, board)
                        if previous:
                            start = max(start, min(args.date, previous) - timedelta(days=7))
                        candles = moex_cls.get_candles(instrument['SECID'], start, args.date, '1d', board=board)
                        store.candles(run_id, isin, board, candles)
                        counts['candles'] += len(candles)
                    except Exception as exc:
                        error('candles', exc, isin)
                    if args.pause:
                        time.sleep(args.pause)
            LOG.info('run=%s processed=%s/%s counts=%s', run_id, min(offset + len(batch), len(isins)), len(isins), counts)
            if args.pause:
                time.sleep(args.pause)
        status = 'partial' if counts['errors'] else 'success'
        store.finish(run_id, status, counts)
        return 2 if counts['errors'] else 0
    except Exception as exc:
        error('run', exc)
        store.finish(run_id, 'failed', counts)
        return 1


def parser():
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument('--date', type=date.fromisoformat, default=datetime.now(MSK).date())
    p.add_argument('--isins', help='Comma-separated ISINs from the current market list')
    p.add_argument('--limit', type=int, default=0, help='0 = all market ISINs')
    p.add_argument('--batch-size', type=int, default=25)
    p.add_argument('--pause', type=float, default=0.2, help='Pause between card/candle calls and batches')
    p.add_argument('--price', choices=['close', 'legal', 'mp3', 'wap'], default='close')
    p.add_argument('--lookback', type=int, default=moex_bonds.LOOKBACK)
    p.add_argument('--assumed-rate', type=float)
    p.add_argument('--history-days', type=int, default=365, help='0 disables candles')
    p.add_argument('--candle-board', help='Explicit override for candles only; otherwise use the source board')
    return p


def main(argv=None):
    logging.basicConfig(level=logging.INFO, format='%(asctime)s %(levelname)s %(message)s')
    p = parser()
    args = p.parse_args(argv)
    if args.batch_size < 1 or min(args.limit, args.pause, args.lookback, args.history_days) < 0:
        p.error('batch-size must be positive; limits, periods and pause must be nonnegative')
    with connect() as conn:
        with writer_lock(conn):
            return execute(Store(conn), args, uuid4())


if __name__ == '__main__':
    raise SystemExit(main())
