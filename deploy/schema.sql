\set ON_ERROR_STOP on
\getenv writer_password BOND_WRITER_PASSWORD
\getenv reader_password BOND_READER_PASSWORD

SELECT format('CREATE ROLE bond_writer LOGIN PASSWORD %L', :'writer_password')
WHERE NOT EXISTS (SELECT FROM pg_roles WHERE rolname='bond_writer') \gexec
SELECT format('CREATE ROLE bond_reader LOGIN PASSWORD %L', :'reader_password')
WHERE NOT EXISTS (SELECT FROM pg_roles WHERE rolname='bond_reader') \gexec
ALTER ROLE bond_reader SET default_transaction_read_only = on;

BEGIN;
CREATE SCHEMA IF NOT EXISTS bond;
CREATE TABLE IF NOT EXISTS bond.runs (
    id uuid PRIMARY KEY,
    started_at timestamptz NOT NULL DEFAULT clock_timestamp(),
    finished_at timestamptz,
    status text NOT NULL DEFAULT 'running'
        CHECK (status IN ('running','success','partial','failed','interrupted')),
    params jsonb NOT NULL,
    counts jsonb NOT NULL DEFAULT '{}',
    errors jsonb NOT NULL DEFAULT '[]'
);
CREATE TABLE IF NOT EXISTS bond.instruments (
    run_id uuid NOT NULL REFERENCES bond.runs(id),
    isin text NOT NULL,
    collected_at timestamptz NOT NULL DEFAULT clock_timestamp(),
    payload jsonb NOT NULL,
    usable boolean NOT NULL DEFAULT true,
    PRIMARY KEY (run_id, isin)
);
CREATE INDEX IF NOT EXISTS instruments_latest ON bond.instruments (isin, collected_at DESC);
CREATE OR REPLACE VIEW bond.latest_instruments AS
    SELECT DISTINCT ON (isin) * FROM bond.instruments WHERE usable
    ORDER BY isin, collected_at DESC, run_id DESC;

DO $$
DECLARE name text;
BEGIN
    FOREACH name IN ARRAY ARRAY['cards','schedules','yields','sectors'] LOOP
        EXECUTE format('CREATE TABLE IF NOT EXISTS bond.%I (
            run_id uuid NOT NULL,
            isin text NOT NULL,
            collected_at timestamptz NOT NULL DEFAULT clock_timestamp(),
            payload jsonb NOT NULL,
            usable boolean NOT NULL DEFAULT true,
            PRIMARY KEY (run_id, isin),
            FOREIGN KEY (run_id, isin) REFERENCES bond.instruments(run_id, isin)
        )', name);
        EXECUTE format('CREATE INDEX IF NOT EXISTS %I ON bond.%I (isin, collected_at DESC)', name || '_latest', name);
        EXECUTE format('CREATE OR REPLACE VIEW bond.%I AS
            SELECT DISTINCT ON (isin) * FROM bond.%I WHERE usable
            ORDER BY isin, collected_at DESC, run_id DESC', 'latest_' || name, name);
    END LOOP;
END $$;

CREATE TABLE IF NOT EXISTS bond.candles (
    run_id uuid NOT NULL,
    isin text NOT NULL,
    board text NOT NULL,
    begin timestamp NOT NULL,
    collected_at timestamptz NOT NULL DEFAULT clock_timestamp(),
    payload jsonb NOT NULL,
    PRIMARY KEY (run_id, isin, board, begin),
    FOREIGN KEY (run_id, isin) REFERENCES bond.instruments(run_id, isin)
);
CREATE INDEX IF NOT EXISTS candles_latest ON bond.candles (isin, board, begin, collected_at DESC);
CREATE OR REPLACE VIEW bond.latest_candles AS
    SELECT DISTINCT ON (isin, board, begin) * FROM bond.candles
    ORDER BY isin, board, begin, collected_at DESC, run_id DESC;

-- SQL calendar for all instruments; unknown amounts stay NULL.
CREATE OR REPLACE VIEW bond.coupon_events AS
    SELECT s.isin, s.run_id, s.collected_at, (event->>'date')::date AS event_date,
        (event->>'coupon_amount')::numeric AS coupon_amount,
        (event->>'amort_amount')::numeric AS amort_amount,
        (event->>'total_amount')::numeric AS total_amount, event AS payload
    FROM bond.latest_schedules s
    CROSS JOIN LATERAL jsonb_array_elements(s.payload->'rows') AS event
    WHERE event->>'date' IS NOT NULL;

REVOKE ALL ON SCHEMA bond FROM PUBLIC;
GRANT USAGE ON SCHEMA bond TO bond_writer, bond_reader;
GRANT SELECT, INSERT, UPDATE, DELETE ON ALL TABLES IN SCHEMA bond TO bond_writer;
GRANT SELECT ON ALL TABLES IN SCHEMA bond TO bond_reader;
COMMIT;
