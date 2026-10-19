#!/usr/bin/env bash
# Fast-forward only; never reset, clean, force-pull or deploy a different branch.
set -euo pipefail
export GIT_TERMINAL_PROMPT=0

main() {
  local root current remote deployed running busy
  root=$(cd "$(dirname "$0")/.." && pwd)
  cd "$root"
  exec 9>.git/moex-deploy.lock
  flock -n 9 || exit 0
  if [ "$(git branch --show-current)" != new_arch ]; then
    echo 'Refusing deployment: expected branch new_arch.' >&2
    exit 1
  fi
  if ! git diff --quiet || ! git diff --cached --quiet; then
    echo 'Refusing deployment: preserve/review tracked local changes first.' >&2
    exit 1
  fi
  git fetch --no-tags origin refs/heads/new_arch:refs/remotes/origin/new_arch
  current=$(git rev-parse HEAD)
  remote=$(git rev-parse origin/new_arch)
  if ! git merge-base --is-ancestor "$current" "$remote"; then
    echo 'Refusing deployment: branch history diverged or was rewritten.' >&2
    exit 1
  fi
  deployed=$(cat .git/moex-deployed-revision 2>/dev/null || true)
  if [ "$current" = "$remote" ] && [ "$deployed" = "$current" ]; then
    exit 0
  fi

  # Do not interrupt an active data collection. Recheck on the next timer tick.
  if command -v docker >/dev/null; then
    export SOURCE_REVISION="$current" BOND_IMAGE_TAG="${current:0:12}"
    running=$(docker compose -f compose.yaml ps --status running -q worker)
    if [ -n "$running" ]; then
      busy=$(docker compose -f compose.yaml exec -T worker python - <<'PY'
from pathlib import Path
busy = False
for p in Path('/proc').glob('[0-9]*/cmdline'):
    try:
        if b'bond_service.run_pipeline' in p.read_bytes().split(b'\0'):
            busy = True
            break
    except (FileNotFoundError, ProcessLookupError, PermissionError):
        continue
print('busy' if busy else 'idle')
PY
      )
      if [ "$busy" = busy ]; then
        echo 'Collection is active; deployment deferred.'
        exit 0
      fi
    fi
  fi
  git merge --ff-only origin/new_arch
  # Exec the new version instead of continuing a shell file just changed by Git.
  export BOND_DEPLOY_LOCK_HELD=1
  exec /bin/bash "$root/deploy/aeza.sh"
}

main "$@"
