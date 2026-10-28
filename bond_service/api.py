"""Public read API. Does not import or call any MOEX function."""
from contextlib import asynccontextmanager
from datetime import date
from uuid import UUID

from fastapi import FastAPI, HTTPException, Query

from .storage import connect

@asynccontextmanager
async def lifespan(app):
    with connect(readonly=True) as conn:
        if conn.execute("SELECT has_table_privilege(current_user, 'bond.cards', 'INSERT') AS writable").fetchone()['writable']:
            raise RuntimeError('The API must use a database role without write privileges')
    yield


app = FastAPI(title='MOEX bonds — saved results', lifespan=lifespan)


@app.get('/health')
def health():
    try:
        with connect(readonly=True) as conn:
            conn.execute('SELECT 1 FROM bond.runs LIMIT 1')
        return {'ok': True}
    except Exception:
        raise HTTPException(503, 'Database is unavailable') from None


@app.get('/runs')
def runs(limit: int = Query(20, ge=1, le=100)):
    with connect(readonly=True) as conn:
        return conn.execute('SELECT * FROM bond.runs ORDER BY started_at DESC LIMIT %s', (limit,)).fetchall()


@app.get('/bonds')
def bonds(limit: int = Query(100, ge=1, le=500), offset: int = Query(0, ge=0)):
    with connect(readonly=True) as conn:
        return conn.execute('SELECT * FROM bond.latest_instruments ORDER BY isin LIMIT %s OFFSET %s', (limit, offset)).fetchall()


@app.get('/bonds/{isin}')
def bond(isin: str, run_id: UUID | None = None):
    isin = isin.strip().upper()
    with connect(readonly=True) as conn, conn.transaction():
        conn.execute('SET TRANSACTION ISOLATION LEVEL REPEATABLE READ')
        instrument = conn.execute('''SELECT * FROM bond.instruments
            WHERE isin=%s AND (%s::uuid IS NULL OR run_id=%s)
            ORDER BY collected_at DESC LIMIT 1''', (isin, run_id, run_id)).fetchone()
        if not instrument:
            raise HTTPException(404, 'ISIN not found')
        result = {'instrument': instrument}
        # Fixed table names only. Each stage includes its own run_id and timestamp.
        for name in ('cards', 'schedules', 'yields', 'sectors'):
            result[name] = conn.execute(f'''SELECT * FROM bond.{name}
                WHERE isin=%s AND (%s::uuid IS NULL OR run_id=%s)
                  AND (%s::uuid IS NOT NULL OR usable)
                ORDER BY collected_at DESC LIMIT 1''', (isin, run_id, run_id, run_id)).fetchone()
        return result


@app.get('/bonds/{isin}/candles')
def candles(isin: str, start: date | None = None, end: date | None = None,
            board: str | None = None, limit: int = Query(1000, ge=1, le=10000)):
    if start and end and start > end:
        raise HTTPException(422, 'start must not be later than end')
    with connect(readonly=True) as conn:
        return conn.execute('''SELECT * FROM bond.latest_candles
            WHERE isin=%s AND (%s::date IS NULL OR begin >= %s::date)
              AND (%s::date IS NULL OR begin < %s::date + interval '1 day')
              AND (%s::text IS NULL OR board=%s)
            ORDER BY begin, board LIMIT %s''',
            (isin.strip().upper(), start, start, end, end, board, board, limit)).fetchall()


@app.get('/calendar')
def calendar(start: date, end: date, limit: int = Query(1000, ge=1, le=10000), offset: int = Query(0, ge=0)):
    if start > end:
        raise HTTPException(422, 'start must not be later than end')
    with connect(readonly=True) as conn:
        return conn.execute('''SELECT * FROM bond.coupon_events
            WHERE event_date BETWEEN %s AND %s
            ORDER BY event_date, isin LIMIT %s OFFSET %s''', (start, end, limit, offset)).fetchall()
