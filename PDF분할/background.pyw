"""Windows pythonw 전용 시작/종료/상태/로그인 자동시작 관리."""
import ctypes
import os
from pathlib import Path
import subprocess
import sys
import time
from runtime import exclusive_lock, AlreadyRunning

BASE = Path(__file__).resolve().parent
STOP = BASE / '.stop-request'
STARTUP_NAME = 'PDFSplitBackground.lnk'
IS_WORKER = '--worker' in sys.argv


def message(text, error=False, silent=False):
    if not silent:
        ctypes.windll.user32.MessageBoxW(None, text, 'PDF분할', 0x10 if error else 0x40)


def running():
    try:
        with exclusive_lock(BASE / '.running.lock'):
            return False
    except AlreadyRunning:
        return True


def startup_link():
    import win32com.client
    shell = win32com.client.Dispatch('WScript.Shell')
    return shell, Path(shell.SpecialFolders('Startup')) / STARTUP_NAME


def main():
    mode = sys.argv[1] if len(sys.argv)>1 else 'start'
    silent = '--silent' in sys.argv
    if mode == '--worker':
        # pythonw는 stdout/stderr가 없으므로 split은 회전 로그만 사용한다.
        import split
        sys.argv = ['split.py', '--background-worker']
        return split.main()
    if mode == 'status':
        state = '종료 요청 대기 중 (현재 파일 처리 후 종료)' if STOP.exists() else '감시 실행 중'
        message((state if running() else '현재 실행 중이 아닙니다.') + '\n\n처리 내역: 실행기록.log', silent=silent)
        return 0
    with exclusive_lock(BASE / '.control.lock'):
        if mode == 'stop':
            if running():
                STOP.touch()
                message('종료를 요청했습니다. 현재 파일 처리를 마친 뒤 감시를 종료합니다.\n한글이 응답하지 않으면 최대 5분 걸릴 수 있습니다.', silent=silent)
            else:
                message('현재 실행 중이 아닙니다.', silent=silent)
            return 0
        if mode in ('enable','disable'):
            shell, link_path = startup_link()
            if mode == 'enable':
                shortcut = shell.CreateShortcut(str(link_path))
                shortcut.TargetPath = sys.executable
                shortcut.Arguments = subprocess.list2cmdline([str(Path(__file__).resolve()), 'start', '--silent'])
                shortcut.WorkingDirectory = str(BASE)
                shortcut.Description = 'PDF분할 폴더 자동 감시'
                shortcut.WindowStyle = 7
                shortcut.Save()
                message('Windows 로그인 시 자동 시작하도록 등록했습니다.\n지금 감시를 시작하려면 백그라운드 시작.vbs를 실행하세요.\n이 폴더를 옮기면 다시 등록해야 합니다.', silent=silent)
            else:
                link_path.unlink(missing_ok=True)
                message('로그인 자동 시작을 해제했습니다.\n현재 감시도 끝내려면 백그라운드 종료.vbs를 실행하세요.', silent=silent)
            return 0
        if mode != 'start':
            raise ValueError('지원하지 않는 실행 명령입니다.')
        if running():
            message('이미 감시 중입니다.' if not STOP.exists() else '종료 중입니다. 잠시 후 다시 시작하세요.', silent=silent)
            return 0
        # 설치/복구 중에는 실행하지 않는다.
        with exclusive_lock(BASE / '.launcher.lock'):
            import bootstrap
            if not bootstrap.usable():
                raise RuntimeError('먼저 설치.bat을 한 번 실행하여 설치를 완료하세요.')
            STOP.unlink(missing_ok=True)
            worker = subprocess.Popen([sys.executable, str(Path(__file__).resolve()), '--worker'],
                cwd=BASE, stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                creationflags=0x08000000, env=dict(os.environ, PYTHONUTF8='1'))
            for _ in range(80):
                if running():
                    message('백그라운드 폴더 감시를 시작했습니다.\n변환할 파일 폴더에 파일을 넣으세요.\n종료: 백그라운드 종료.vbs', silent=silent)
                    return 0
                if worker.poll() is not None:
                    raise RuntimeError('감시 시작에 실패했습니다. 백그라운드오류.log를 확인하세요.')
                time.sleep(0.1)
            message('감시 시작 요청을 보냈습니다. 잠시 후 상태 확인.vbs로 확인하세요.', silent=silent)
    return 0


if __name__ == '__main__':
    try:
        raise SystemExit(main())
    except Exception as exc:
        with (BASE / '백그라운드오류.log').open('a', encoding='utf-8') as log:
            log.write(time.strftime('%Y-%m-%d %H:%M:%S ') + str(exc) + '\n')
        if not IS_WORKER:
            message(str(exc), error=True, silent='--silent' in sys.argv)
        raise SystemExit(1)
