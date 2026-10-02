"""표준 라이브러리만으로 전용 실행환경 설치 및 시작."""
import os
from pathlib import Path
import subprocess
import sys
import time
from runtime import exclusive_lock, AlreadyRunning

BASE = Path(__file__).resolve().parent
VENV = BASE / '.venv'
PYTHON = VENV / 'Scripts' / 'python.exe'

def usable():
    if not PYTHON.is_file():
        return False
    try:
        code = "import pymupdf,win32com.client,pythoncom,win32file; assert tuple(map(int,pymupdf.__version__.split('.'))) >= (1,26,5)"
        return subprocess.run([str(PYTHON), '-c', code], timeout=30,
                              stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                              creationflags=0x08000000 if os.name == 'nt' else 0).returncode == 0
    except (OSError, subprocess.TimeoutExpired):
        return False

def ensure_environment():
    if usable():
        return
    print('최초 설치 또는 실행환경 복구 중입니다. 인터넷 연결이 필요합니다.', flush=True)
    # 기존 venv가 살아 있으면 의존성만 보충한다.
    healthy = False
    if PYTHON.exists():
        try:
            healthy = subprocess.run([str(PYTHON), '-c', 'import sys; assert sys.version_info >= (3,10)'],
                timeout=15, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                              creationflags=0x08000000 if os.name == 'nt' else 0).returncode == 0
        except (OSError, subprocess.TimeoutExpired):
            pass
    if not healthy:
        if VENV.exists():
            backup = BASE / ('.venv-backup-' + str(time.time_ns()))
            VENV.rename(backup)
        subprocess.run([sys.executable, '-m', 'venv', str(VENV)], check=True, timeout=180)
    subprocess.run([str(PYTHON), '-m', 'pip', 'install', '--disable-pip-version-check',
                    '--only-binary=:all:', '-r', str(BASE / 'requirements.txt')],
                   check=True, timeout=600)
    if not usable():
        raise RuntimeError('설치 후 검사에 실패했습니다. 설치기록.log를 확인하세요.')
    print('설치 및 의존성 검사 완료.', flush=True)

def main():
    if os.name != 'nt':
        raise RuntimeError('실행.bat은 Windows용입니다.')
    if sys.version_info < (3,10):
        raise RuntimeError('Python 3.10 이상이 필요합니다. Python 3.12 또는 3.13을 설치하세요.')
    env = dict(os.environ, PYTHONUTF8='1', PYTHONIOENCODING='utf-8')
    os.environ.update(env)
    # 설치가 완료되기 전 실행, 중복 설치, 동일 런처 중복 실행 모두 차단.
    with exclusive_lock(BASE / '.launcher.lock'):
        # 직접 실행 중인 split.py가 있으면 환경 수정을 시도하지 않는다.
        with exclusive_lock(BASE / '.running.lock'):
            ensure_environment()
            (BASE / '.stop-request').unlink(missing_ok=True)
        if '--install-only' in sys.argv:
            return 0
        return subprocess.call([str(PYTHON), '-u', str(BASE / 'split.py'), *sys.argv[1:]], env=env)

if __name__ == '__main__':
    try:
        raise SystemExit(main())
    except KeyboardInterrupt:
        print('종료했습니다.')
        raise SystemExit(130)
    except Exception as exc:
        message = time.strftime('%Y-%m-%d %H:%M:%S ') + str(exc)
        print('시작 실패: ' + message, flush=True)
        with (BASE / '설치기록.log').open('a', encoding='utf-8') as stream:
            stream.write(message + '\n')
        raise SystemExit(1)
