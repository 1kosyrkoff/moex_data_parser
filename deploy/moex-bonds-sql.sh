# Default connection for libpq clients in login shells; preserve explicit choices.
export PGSERVICEFILE="${PGSERVICEFILE:-/etc/moex-bonds/pg_service.conf}"
export PGSERVICE="${PGSERVICE:-bonds}"
