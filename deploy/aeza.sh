#!/usr/bin/env bash
# User-run deployment from the Git checkout; no remote execution by the agent.
set -euo pipefail
cd "$(dirname "$0")/.."

if [ "$(id -u)" -ne 0 ]; then
  echo 'Run as root on the VM.' >&2
  exit 1
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
for name in ['BOND_DB_PASSWORD', 'BOND_WRITER_PASSWORD', 'BOND_READER_PASSWORD', 'BOND_API_TOKEN']:
    text = text.replace(name + '=\n', name + '=' + secrets.token_hex(32) + '\n')
fd = os.open('.env', os.O_WRONLY | os.O_CREAT | os.O_EXCL, 0o600)
with os.fdopen(fd, 'w') as stream:
    stream.write(text)
PY
fi
chmod 600 .env

export SOURCE_REVISION
SOURCE_REVISION=$(git rev-parse HEAD)
export BOND_IMAGE_TAG
BOND_IMAGE_TAG=$(git rev-parse --short=12 HEAD)
export COMPOSE_PARALLEL_LIMIT=1
compose=(docker compose -f compose.yaml -f deploy/compose.aeza.yaml)
"${compose[@]}" config --quiet
"${compose[@]}" build
if [ -n "$("${compose[@]}" ps -q worker)" ]; then
  "${compose[@]}" stop worker
fi
"${compose[@]}" up -d --wait --wait-timeout 120 postgres api
"${compose[@]}" up -d worker
"${compose[@]}" ps
endpoint=$("${compose[@]}" port api 8000)
curl --fail --silent --show-error --max-time 10 "http://${endpoint}/health"
echo
echo "Deployed ${SOURCE_REVISION}; API bound to ${endpoint}."
echo 'Initial run uses .env settings (default: 5 issues). Secrets stay in .env.'
