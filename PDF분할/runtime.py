"""공유 실행 잠금. 잠긴 파일을 쓰지 않고 첫 바이트만 잠근다."""
import os
from contextlib import contextmanager

class AlreadyRunning(RuntimeError):
    pass

@contextmanager
def exclusive_lock(path):
    handle = open(path, 'a+b')
    try:
        handle.seek(0)
        try:
            if os.name == 'nt':
                import msvcrt
                msvcrt.locking(handle.fileno(), msvcrt.LK_NBLCK, 1)
            else:
                import fcntl
                fcntl.flock(handle, fcntl.LOCK_EX | fcntl.LOCK_NB)
        except OSError as exc:
            raise AlreadyRunning('이미 실행 중입니다. 기존 실행 창을 확인하세요.') from exc
        yield
    finally:
        handle.close()


@contextmanager
def hwp_session():
    """두 프로그램의 한글 COM 작업이 겹치지 않도록 사용자 임시 폴더에서 공유 잠금."""
    import tempfile
    import time
    from pathlib import Path
    path=Path(tempfile.gettempdir())/'pdf_tools_hwp_session.lock'
    deadline=time.monotonic()+240
    while True:
        ctx=exclusive_lock(path)
        try:
            ctx.__enter__()
            break
        except AlreadyRunning:
            if time.monotonic()>deadline:
                raise RuntimeError('다른 PDF 변환의 한글 작업을 기다리다 시간이 초과되었습니다.')
            time.sleep(0.5)
    try:
        yield
    finally:
        ctx.__exit__(None,None,None)
