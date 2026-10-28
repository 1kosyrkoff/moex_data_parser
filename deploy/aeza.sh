#!/usr/bin/env bash
# User-run deployment from the Git checkout; no remote execution by the agent.
set -euo pipefail
cd "$(dirname "$0")/.."

if [ "$(id -u)" -ne 0 ]; then
  echo 'Run as root on the VM.' >&2
  exit 1
fi
if [ "${BOND_DEPLOY_LOCK_HELD:-0}" != 1 ]; then
  exec 9>.git/moex-deploy.lock
  if ! flock -n 9; then
    echo 'Another deployment is already running.' >&2
    exit 75
  fi
fi
if [ "$(git branch --show-current)" != new_arch ]; then
  echo 'Expected branch new_arch.' >&2
  exit 1
fi
if ! git diff --quiet || ! git diff --cached --quiet; then
  echo 'Tracked local changes found. Preserve/review them before deployment.' >&2
  exit 1
fi

if ! command -v docker >/dev/null; then
  . /etc/os-release
  if [ "$ID" != ubuntu ]; then
    echo 'Automatic Docker installation is limited to Ubuntu.' >&2
    exit 1
  fi
  apt-get update
  # Use Ubuntu's packages; do not mix installations or restart unrelated services.
  NEEDRESTART_MODE=l DEBIAN_FRONTEND=noninteractive \
    apt-get install -y --no-install-recommends docker.io docker-compose-v2
fi
docker compose version
systemctl enable --now docker

if [ ! -e .env ]; then
  python3 - <<'PY'
import os
from pathlib import Path
import secrets
text = Path('.env.example').read_text()
for name in ['BOND_DB_PASSWORD', 'BOND_WRITER_PASSWORD', 'BOND_READER_PASSWORD']:
    text = text.replace(name + '=\n', name + '=' + secrets.token_hex(32) + '\n')
fd = os.open('.env', os.O_WRONLY | os.O_CREAT | os.O_EXCL, 0o600)
with os.fdopen(fd, 'w') as stream:
    stream.write(text)
PY
fi
chmod 600 .env

# Install the host SQL client and defaults for new login shells.
if ! command -v psql >/dev/null; then
  apt-get update
  NEEDRESTART_MODE=l DEBIAN_FRONTEND=noninteractive \
    apt-get install -y --no-install-recommends postgresql-client
fi

export SOURCE_REVISION
SOURCE_REVISION=$(git rev-parse HEAD)
export BOND_IMAGE_TAG
BOND_IMAGE_TAG=$(git rev-parse --short=12 HEAD)
export COMPOSE_PARALLEL_LIMIT=1
compose=(docker compose -f compose.yaml)
"${compose[@]}" config --quiet
"${compose[@]}" build
if [ -n "$("${compose[@]}" ps -q worker)" ]; then
  "${compose[@]}" stop worker
fi
"${compose[@]}" up -d --wait --wait-timeout 120 postgres api
"${compose[@]}" up -d worker
# Use Compose's resolved port, including an existing .env override.
db_endpoint=$("${compose[@]}" port postgres 5432)
install -d -m 755 /etc/moex-bonds
printf '[bonds]\nhost=127.0.0.1\nport=%s\ndbname=bonds\nuser=bond_reader\nconnect_timeout=5\n' \
  "${db_endpoint##*:}" > /etc/moex-bonds/pg_service.conf
chmod 644 /etc/moex-bonds/pg_service.conf
install -m 644 deploy/moex-bonds-sql.sh /etc/profile.d/moex-bonds-sql.sh
PGSERVICEFILE=/etc/moex-bonds/pg_service.conf PGSERVICE=bonds \
  psql -X -w -v ON_ERROR_STOP=1 -c 'SELECT current_user, current_database();'
"${compose[@]}" ps
endpoint=$("${compose[@]}" port api 8000)
curl --fail --silent --show-error --max-time 10 "http://127.0.0.1:${endpoint##*:}/health"
printf '%s\n' "$SOURCE_REVISION" > .git/moex-deployed-revision
echo
echo "Deployed ${SOURCE_REVISION}; API bound to ${endpoint}."
echo 'Run scope is configured in .env (default: complete current MOEX list). Secrets stay in .env.'
