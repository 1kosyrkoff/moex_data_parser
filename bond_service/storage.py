"""Persist results, without interpreting or recalculating financial values."""
from contextlib import contextmanager
from datetime import date, datetime
import math
import os

import numpy as np
import pandas as pd
import psycopg
from psycopg import sql
from psycopg.rows import dict_row
from psycopg.types.json import Jsonb

LOCK_ID = 109824014
SNAPSHOTS = frozenset({'instruments', 'cards', 'schedules', 'yields', 'sectors'})


def json_value(value):
    if value is None or value is pd.NA or value is pd.NaT:
        return None
    if isinstance(value, (datetime, date)):
        return value.isoformat()
    if isinstance(value, np.generic):
        return json_value(value.item())
    if isinstance(value, float) and not math.isfinite(value):
        return None
    if isinstance(value, dict):
        return {str(k): json_value(v) for k, v in value.items()}
    if isinstance(value, (list, tuple)):
        return [json_value(v) for v in value]
    return value


def frame_payload(frame):
    """Keep columns, original row order, index and pandas type descriptions."""
    if not frame.columns.is_unique:
        raise ValueError('Cannot store duplicate column names as JSON records')
    return json_value({
        'columns': list(frame.columns),
        'dtypes': [str(t) for t in frame.dtypes],
        'index': list(frame.index),
        'index_name': frame.index.name,
        'index_dtype': str(frame.index.dtype),
        'rows': frame.to_dict(orient='records'),
    })


def connect(*, readonly=False):
    options = '-c statement_timeout=30000'
    if readonly:
        options += ' -c default_transaction_read_only=on'
    return psycopg.connect(
        host=os.environ.get('PGHOST', 'localhost'),
        port=os.environ.get('PGPORT', '5432'),
        dbname=os.environ.get('PGDATABASE', 'bonds'),
        user=os.environ['PGUSER'], password=os.environ['PGPASSWORD'],
        connect_timeout=5, autocommit=True, row_factory=dict_row,
        options=options,
    )


@contextmanager
def writer_lock(conn):
    acquired = conn.execute('SELECT pg_try_advisory_lock(%s) AS acquired', (LOCK_ID,)).fetchone()['acquired']
    if not acquired:
        raise RuntimeError('Another update process already holds the database lock')
    try:
        yield
    finally:
        conn.execute('SELECT pg_advisory_unlock(%s)', (LOCK_ID,))


def technical_errors(frame):
    """Upstream STATUS distinguishes technical failures from valid missing data."""
    if 'STATUS' not in frame.columns:
        return []
    return [str(v) for v in frame['STATUS'].dropna()
            if str(v).lower().startswith(('ошибка запроса', 'ошибка разбора', 'ошибка расчета'))]


class Store:
    def __init__(self, conn):
        self.conn = conn

    def start(self, run_id, params):
        # Only call while holding the lock: a prior running process has exited.
        with self.conn.transaction():
            self.conn.execute("UPDATE bond.runs SET status='interrupted', finished_at=now() WHERE status='running'")
            self.conn.execute('INSERT INTO bond.runs (id, params) VALUES (%s, %s)', (run_id, Jsonb(params)))

    def error(self, run_id, stage, message, isin=None):
        item = {'stage': stage, 'isin': isin, 'message': str(message)[:1000]}
        self.conn.execute('UPDATE bond.runs SET errors=errors || %s WHERE id=%s', (Jsonb([item]), run_id))

    def put(self, table, run_id, isin, payload, *, usable=True):
        if table not in SNAPSHOTS:
            raise ValueError('Unsupported snapshot table')
        statement = sql.SQL('''INSERT INTO bond.{} (run_id, isin, payload, usable)
            VALUES (%s, %s, %s, %s)
            ON CONFLICT (run_id, isin) DO UPDATE SET
                payload=EXCLUDED.payload, usable=EXCLUDED.usable, collected_at=clock_timestamp()''')
        self.conn.execute(statement.format(sql.Identifier(table)), (run_id, isin, Jsonb(payload), usable))

    def card(self, run_id, isin, card):
        # Do not mutate the caller's dict or request the schedule a second time.
        schedule = card['schedule']
        usable = not technical_errors(schedule)
        with self.conn.transaction():
            self.put('cards', run_id, isin, {k: frame_payload(v) for k, v in card.items() if k != 'schedule'}, usable=usable)
            self.put('schedules', run_id, isin, frame_payload(schedule), usable=usable)

    def candle_start(self, isin, board):
        row = self.conn.execute('SELECT max(begin) AS last FROM bond.candles WHERE isin=%s AND board=%s', (isin, board)).fetchone()
        return row['last'].date() if row['last'] else None

    def candles(self, run_id, isin, board, frame):
        with self.conn.transaction():
            for begin, row in frame.iterrows():
                self.conn.execute('''INSERT INTO bond.candles (run_id, isin, board, begin, payload)
                    VALUES (%s, %s, %s, %s, %s)
                    ON CONFLICT (run_id, isin, board, begin) DO UPDATE SET payload=EXCLUDED.payload''',
                    (run_id, isin, board, begin.to_pydatetime(), Jsonb(json_value(row.to_dict()))))

    def finish(self, run_id, status, counts):
        self.conn.execute('UPDATE bond.runs SET status=%s, finished_at=now(), counts=%s WHERE id=%s',
                          (status, Jsonb(counts), run_id))
