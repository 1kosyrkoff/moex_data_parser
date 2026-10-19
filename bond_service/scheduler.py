"""Schedule fresh child interpreters; no scheduler dependency or shell commands."""
import logging
import os
import shlex
import signal
import subprocess
import sys
import threading

LOG = logging.getLogger(__name__)


def main():
    logging.basicConfig(level=logging.INFO, format='%(asctime)s %(levelname)s %(message)s')
    interval = int(os.environ.get('BOND_UPDATE_INTERVAL', '86400'))
    timeout = int(os.environ.get('BOND_RUN_TIMEOUT', '43200'))
    if min(interval, timeout) < 1:
        raise ValueError('BOND_UPDATE_INTERVAL and BOND_RUN_TIMEOUT must be positive')
    stop = threading.Event()
    child = None

    def terminate(signum, frame):
        stop.set()
        if child is not None and child.poll() is None:
            child.terminate()

    signal.signal(signal.SIGTERM, terminate)
    signal.signal(signal.SIGINT, terminate)
    command = [sys.executable, '-m', 'bond_service.run_pipeline',
               *shlex.split(os.environ.get('BOND_RUN_ARGS', ''))]
    while not stop.is_set():
        child = subprocess.Popen(command)
        try:
            code = child.wait(timeout=timeout)
            LOG.info('Update process exited: %s', code)
        except subprocess.TimeoutExpired:
            LOG.error('Update exceeded BOND_RUN_TIMEOUT; stopping the process')
            child.terminate()
            try:
                child.wait(timeout=30)
            except subprocess.TimeoutExpired:
                child.kill()
                child.wait()
        if not stop.is_set():
            stop.wait(interval)


if __name__ == '__main__':
    main()
