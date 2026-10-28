"""Offline checks; integration tests use an isolated PostgreSQL from TEST_PGPORT."""
from datetime import date
import json
import os
from uuid import uuid4

import pandas as pd
import psycopg
import pytest
import requests

from bond_service import run_pipeline as pipeline
from bond_service.storage import Store, connect, frame_payload, technical_errors, writer_lock

ISIN = 'RU000A106E90'
SECOND = 'RU000A0JS3W6'


@pytest.fixture(autouse=True)
def no_network(monkeypatch):
    def blocked(*args, **kwargs):
        raise AssertionError('External requests are not allowed in these tests')
    monkeypatch.setattr(requests.sessions.Session, 'request', blocked)


def schedule():
    return pd.DataFrame([{'ISIN': ISIN, 'date': pd.Timestamp('2026-11-01'),
        'STARTDATE': pd.Timestamp('2026-05-01'), 'coupon_amount': float('nan'),
        'amort_amount': 100.0, 'total_amount': float('nan'), 'STATUS': None}])


def card():
    return {'passport': pd.DataFrame([{'Поле': 'ISIN', 'Значение': ISIN}]),
            'quotes': pd.DataFrame([{'Поле': 'PRICE', 'Значение': 80.12567891234567}]),
            'schedule': schedule(), 'boards': pd.DataFrame(),
            'indices': pd.DataFrame(), 'aggregates': pd.DataFrame()}


def test_serialization_keeps_unknowns_dates_precision_and_input():
    original = schedule()
    before = original.copy(deep=True)
    payload = frame_payload(original)
    assert payload['rows'][0]['coupon_amount'] is None
    assert payload['rows'][0]['total_amount'] is None
    assert payload['rows'][0]['amort_amount'] == 100.0
    assert payload['rows'][0]['date'] == '2026-11-01T00:00:00'
    assert payload['columns'] == list(original.columns)
    json.dumps(payload, allow_nan=False)
    pd.testing.assert_frame_equal(original, before)
    amount = 80.12567891234567
    assert frame_payload(pd.DataFrame([{'x': amount}]))['rows'][0]['x'] == amount


def test_missing_sector_and_unknown_yield_are_not_technical_errors():
    frame = pd.DataFrame({'STATUS': [None, 'ok', 'нет цены за 20 дней до даты оценки',
        'бессрочная бумага: доходность к погашению не определена',
        'эмитент не представлен в отраслевых индексах MOEX']})
    assert technical_errors(frame) == []
    assert technical_errors(pd.DataFrame({'STATUS': ['ошибка запроса: timeout']}))


@pytest.fixture
def db(monkeypatch):
    if not os.environ.get('TEST_PGPORT'):
        pytest.skip('Set TEST_PGPORT for isolated PostgreSQL integration tests')
    monkeypatch.setenv('PGHOST', '127.0.0.1')
    monkeypatch.setenv('PGPORT', os.environ['TEST_PGPORT'])
    database = os.environ.get('TEST_PGDATABASE', 'bond_service_test')
    assert database.endswith('_test'), 'Integration tests require an isolated database ending in _test'
    monkeypatch.setenv('PGDATABASE', database)
    monkeypatch.setenv('PGUSER', 'bond_writer')
    monkeypatch.setenv('PGPASSWORD', os.environ['TEST_WRITER_PASSWORD'])
    with connect() as conn:
        # DELETE privileges intentionally only on this isolated test database.
        for name in ('candles', 'cards', 'schedules', 'yields', 'sectors', 'instruments', 'runs'):
            conn.execute(f'DELETE FROM bond.{name}')
        yield Store(conn)


@pytest.fixture
def upstream(monkeypatch):
    market = pd.DataFrame([
        {'ISIN': ISIN, 'SECID': 'SU26243RMFS4', 'BOARDID': 'SPOB'},
        {'ISIN': ISIN, 'SECID': 'SU26243RMFS4', 'BOARDID': 'TQOB'},
        {'ISIN': SECOND, 'SECID': 'SU26207RMFS9', 'BOARDID': 'TQOB'},
        {'ISIN': None, 'SECID': 'NO_ISIN', 'BOARDID': 'TQOB'},
    ])
    calls = {'cards': [], 'yields': [], 'sectors': [], 'candles': []}
    monkeypatch.setattr(pipeline.moex_bonds, 'market_snapshot', lambda: market)

    def get_card(isin):
        calls['cards'].append(isin)
        return card()

    def get_yields(isins, **kwargs):
        calls['yields'].append((isins, kwargs))
        return pd.DataFrame([{'ISIN': isin, 'YTM': 13.6509, 'YTP': None, 'STATUS': 'ok'} for isin in isins])

    def sectors(isins):
        calls['sectors'].append(isins)
        return pd.DataFrame([{'ISIN': isin, 'SECTOR': None,
            'STATUS': 'эмитент не представлен в отраслевых индексах MOEX'} for isin in isins])

    def candles(ticker, start, end, period, board):
        calls['candles'].append((ticker, start, end, period, board))
        return pd.DataFrame([{'ticker': ticker, 'close': 80.12567891234567}],
                            index=pd.DatetimeIndex(['2026-10-15'], name='begin'))

    monkeypatch.setattr(pipeline.moex_bonds, 'get_bond_card', get_card)
    monkeypatch.setattr(pipeline.moex_bonds, 'get_bond_yields', get_yields)
    monkeypatch.setattr(pipeline.moex_bonds, 'get_issuer_sector', sectors)
    monkeypatch.setattr(pipeline.moex_cls, 'get_candles', candles)
    return calls


def args(*extra):
    return pipeline.parser().parse_args(['--date', '2026-10-15', '--pause', '0',
        '--batch-size', '1', '--history-days', '30', *extra])


def test_full_run_roundtrip_and_original_call_parameters(db, upstream):
    run_id = uuid4()
    assert pipeline.execute(db, args(), run_id) == 0
    assert upstream['cards'] == [ISIN, SECOND]  # Duplicate ISIN processed once.
    assert upstream['yields'][0][1] == {'date': date(2026, 10, 15), 'price': 'close',
        'lookback': 20, 'assumed_rate': None, 'details': True}
    assert upstream['candles'][0][-1] == 'SPOB'  # No implicit board correction.
    row = db.conn.execute('SELECT * FROM bond.cards WHERE run_id=%s AND isin=%s', (run_id, ISIN)).fetchone()
    assert row['payload'] == {k: frame_payload(v) for k, v in card().items() if k != 'schedule'}
    saved = db.conn.execute('SELECT payload FROM bond.schedules WHERE run_id=%s AND isin=%s', (run_id, ISIN)).fetchone()
    assert saved['payload'] == frame_payload(schedule())
    event = db.conn.execute('SELECT * FROM bond.coupon_events WHERE isin=%s', (ISIN,)).fetchone()
    assert event['coupon_amount'] is None and event['total_amount'] is None
    assert event['amort_amount'] == 100
    run = db.conn.execute('SELECT * FROM bond.runs WHERE id=%s', (run_id,)).fetchone()
    assert run['status'] == 'success'
    assert run['counts']['instruments'] == 2 and run['counts']['errors'] == 0
    assert len(run['params']['source_sha256']) == 64


def test_snapshot_idempotency_history_and_incremental_candles(db, upstream):
    first, second = uuid4(), uuid4()
    assert pipeline.execute(db, args('--limit', '1'), first) == 0
    db.card(first, ISIN, card())
    assert db.conn.execute('SELECT count(*) AS n FROM bond.cards').fetchone()['n'] == 1
    assert pipeline.execute(db, args('--limit', '1'), second) == 0
    assert db.conn.execute('SELECT count(*) AS n FROM bond.cards').fetchone()['n'] == 2
    assert db.conn.execute('SELECT run_id FROM bond.latest_cards').fetchone()['run_id'] == second
    assert db.conn.execute('SELECT count(*) AS n FROM bond.latest_candles').fetchone()['n'] == 1
    assert upstream['candles'][-1][1] == date(2026, 10, 8)  # Seven-day overlap.


def test_failed_stage_retains_previous_success_and_processes_other_stages(db, upstream, monkeypatch):
    first, second = uuid4(), uuid4()
    assert pipeline.execute(db, args('--limit', '1'), first) == 0

    def broken(isin):
        raise requests.Timeout('offline test failure')

    monkeypatch.setattr(pipeline.moex_bonds, 'get_bond_card', broken)
    assert pipeline.execute(db, args('--limit', '1'), second) == 2
    assert db.conn.execute('SELECT run_id FROM bond.latest_cards').fetchone()['run_id'] == first
    assert db.conn.execute('SELECT run_id FROM bond.latest_yields').fetchone()['run_id'] == second
    assert db.conn.execute('SELECT status FROM bond.runs WHERE id=%s', (second,)).fetchone()['status'] == 'partial'


def test_internal_status_failure_is_recorded_but_not_published_as_latest(db, upstream, monkeypatch):
    first, second = uuid4(), uuid4()
    pipeline.execute(db, args('--limit', '1'), first)
    monkeypatch.setattr(pipeline.moex_bonds, 'get_bond_yields', lambda isins, **kw:
        pd.DataFrame([{'ISIN': ISIN, 'YTM': None, 'STATUS': 'ошибка запроса: timeout'}]))
    assert pipeline.execute(db, args('--limit', '1'), second) == 2
    assert db.conn.execute('SELECT usable FROM bond.yields WHERE run_id=%s', (second,)).fetchone()['usable'] is False
    assert db.conn.execute('SELECT run_id FROM bond.latest_yields').fetchone()['run_id'] == first


def test_card_and_schedule_are_atomic(db, upstream):
    run_id = uuid4()
    pipeline.execute(db, args('--limit', '1'), run_id)
    old = db.conn.execute('SELECT payload FROM bond.cards').fetchone()['payload']
    broken = card()
    broken['passport'] = pd.DataFrame([{'new': 'must roll back'}])
    broken['schedule'] = pd.DataFrame([[1, 2]], columns=['duplicate', 'duplicate'])
    with pytest.raises(ValueError):
        db.card(run_id, ISIN, broken)
    assert db.conn.execute('SELECT payload FROM bond.cards').fetchone()['payload'] == old


def test_lock_and_interrupted_run(db, upstream):
    with writer_lock(db.conn):
        with connect() as other:
            with pytest.raises(RuntimeError, match='already holds'):
                with writer_lock(other):
                    pass
    old = uuid4()
    db.start(old, {})
    pipeline.execute(db, args('--limit', '1'), uuid4())
    assert db.conn.execute('SELECT status FROM bond.runs WHERE id=%s', (old,)).fetchone()['status'] == 'interrupted'


def test_read_api_and_database_privileges(db, upstream, monkeypatch):
    from fastapi.testclient import TestClient
    from bond_service.api import app
    run_id = uuid4()
    pipeline.execute(db, args('--limit', '1'), run_id)
    monkeypatch.setenv('PGUSER', 'bond_reader')
    monkeypatch.setenv('PGPASSWORD', os.environ['TEST_READER_PASSWORD'])
    monkeypatch.delenv('BOND_API_TOKEN', raising=False)
    with connect() as reader:
        with pytest.raises(psycopg.Error):
            reader.execute('DELETE FROM bond.runs')
    with TestClient(app) as client:
        assert client.get('/health').status_code == 200
        assert client.get('/bonds').status_code == 200
        assert not client.get('/openapi.json').json().get('components', {}).get('securitySchemes')
        response = client.get(f'/bonds/{ISIN}')
        assert response.status_code == 200
        assert response.json()['schedules']['payload'] == frame_payload(schedule())
        assert response.json()['cards']['run_id'] == str(run_id)
        assert client.get('/calendar?start=2026-11-01&end=2026-11-02').json()[0]['coupon_amount'] is None
        assert client.get(f'/bonds/{ISIN}/candles').status_code == 200
        assert client.get(f'/bonds/{ISIN}?run_id={run_id}').status_code == 200
        assert client.get('/bonds/NOT_FOUND').status_code == 404
        assert client.get('/bonds?limit=1000000').status_code == 422
        assert client.get('/calendar?start=2026-11-02&end=2026-11-01').status_code == 422


def test_empty_market_fails_without_deleting_history(db, upstream, monkeypatch):
    first = uuid4()
    pipeline.execute(db, args('--limit', '1'), first)
    monkeypatch.setattr(pipeline.moex_bonds, 'market_snapshot', lambda: pd.DataFrame())
    failed = uuid4()
    assert pipeline.execute(db, args(), failed) == 1
    assert db.conn.execute('SELECT run_id FROM bond.latest_cards').fetchone()['run_id'] == first
    assert db.conn.execute('SELECT status FROM bond.runs WHERE id=%s', (failed,)).fetchone()['status'] == 'failed'
