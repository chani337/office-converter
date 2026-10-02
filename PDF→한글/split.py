"""PDF→한글 폴더 감시. 원본 보존."""
import argparse
import hashlib
import json
import os
from pathlib import Path
import tempfile
import time
import converter
import logging
from logging.handlers import RotatingFileHandler
from runtime import exclusive_lock, AlreadyRunning

BASE = Path(__file__).resolve().parent

def log(message):
    logging.getLogger('pdfsplit').info(message)

def digest(path):
    h = hashlib.sha256()
    with path.open('rb') as f:
        for block in iter(lambda: f.read(1024 * 1024), b''):
            h.update(block)
    return h.hexdigest()

OCR_MODE = 'auto'

def process(src, watch, output, engine, dpi=300):
    from reverse import process as convert_pdf
    return convert_pdf(src, watch, output, engine, dpi, OCR_MODE)

def copy_snapshot(src, dst):
    """Windows에서는 쓰기 중인 입력을 열지 않는다."""
    if os.name != 'nt':
        import shutil
        shutil.copy2(src, dst)
        return
    import win32file
    import win32con
    handle = win32file.CreateFile(str(src), win32con.GENERIC_READ,
        win32con.FILE_SHARE_READ, None, win32con.OPEN_EXISTING,
        win32con.FILE_ATTRIBUTE_NORMAL, None)
    try:
        with dst.open('wb') as stream:
            while True:
                _, block = win32file.ReadFile(handle, 1024 * 1024)
                if not block:
                    break
                stream.write(block)
    finally:
        handle.Close()


def setup_logging():
    import sys
    logger = logging.getLogger('pdfsplit')
    logger.setLevel(logging.INFO)
    for handler in logger.handlers:
        handler.close()
    logger.handlers.clear()
    logger.propagate = False
    stream = logging.StreamHandler(sys.stdout) if sys.stdout is not None else None
    logfile = RotatingFileHandler(BASE / '실행기록.log', maxBytes=2_000_000,
                                  backupCount=2, encoding='utf-8')
    for handler in ([stream] if stream is not None else []) + [logfile]:
        handler.setFormatter(logging.Formatter('[%(asctime)s] %(message)s', '%H:%M:%S'))
        logger.addHandler(handler)
    # pdf2docx의 경고도 pythonw 실행 중 파일에 남긴다.
    root = logging.getLogger()
    root.handlers = [logfile]
    root.setLevel(logging.WARNING)


def signature(path):
    st = path.stat()
    return (st.st_size, st.st_mtime_ns, st.st_ctime_ns)


def scan_files(watch):
    # 링크를 통한 폴더 외부 입력이나 순환 경로는 따라가지 않는다.
    for root, dirs, files in os.walk(watch):
        dirs[:] = sorted(d for d in dirs if not d.startswith('.')
                         and not (Path(root) / d).is_symlink())
        for name in sorted(files):
            path = Path(root) / name
            if (path.suffix.lower() in converter.SUPPORTED_EXTS
                    and not name.startswith(('~', '.')) and not path.is_symlink()):
                yield path


def run(watch, output, args):
    observations = {}
    completed = {}
    retries = {}
    failed = set()
    first_seen = {}
    initial = set(scan_files(watch)) if args.once else None
    cfg = {'convert_timeout_sec': args.timeout}
    with converter.Converter(cfg, log=log) as engine:
        while True:
            if (BASE / '.stop-request').exists():
                log('종료 요청을 받아 감시를 종료했습니다.')
                return 0
            pending = False
            live = set()
            for src in scan_files(watch):
                if (BASE / '.stop-request').exists():
                    log('현재 파일 처리를 마치고 종료했습니다.')
                    return 0
                if initial is not None and src not in initial:
                    continue
                live.add(src)
                try:
                    sig = signature(src)
                    now = time.monotonic()
                    first_seen.setdefault(src, now)
                    old_sig, since = observations.get(src, (None, now))
                    if sig != old_sig:
                        observations[src] = (sig, now)
                        retries.pop(src, None)
                        completed.pop(src, None)
                        failed.discard(src)
                        since = now
                    # 3초 동안 크기와 수정 시간이 모두 그대로인 파일만 처리.
                    if now - since < 3:
                        if args.once and now - first_seen[src] > 60:
                            failed.add(src)
                            log(f'건너뜀: {src.name} — 60초 동안 복사가 완료되지 않았습니다.')
                            continue
                        pending = True
                        continue
                    attempts, next_try = retries.get(src, (0, 0))
                    if attempts >= 3:
                        continue
                    if now < next_try:
                        pending = True
                        continue
                    if src in completed and now < completed[src]:
                        continue
                    if not sig[0]:
                        raise ValueError('비어 있는 파일입니다.')
                    process(src, watch, output, engine, args.dpi)
                    # 기존 파일을 매 2초마다 다시 해시/대기하지 않는다.
                    # 결과 삭제 여부는 30초마다 재확인한다.
                    completed[src] = time.monotonic() + 30
                    retries.pop(src, None)
                    failed.discard(src)
                except converter.ConversionTimeout as exc:
                    retries[src] = (3, 0)
                    failed.add(src)
                    log(f'변환 중지: {src.name} — {exc}')
                except Exception as exc:
                    count = retries.get(src, (0, 0))[0] + 1
                    retries[src] = (3 if args.once else count, time.monotonic() + 30)
                    failed.add(src)
                    action = '파일 수정 또는 프로그램 재실행 필요' if args.once or count >= 3 else '30초 후 재시도'
                    log(f'실패 ({count}/3): {src.name} — {exc} / {action}')
                    if not args.once and count < 3:
                        pending = True
            for cache in (observations, completed, retries, first_seen):
                for missing in set(cache) - live:
                    cache.pop(missing, None)
            failed.intersection_update(live)
            if args.once and not pending:
                log(f'일괄 처리 종료. 실패 {len(failed)}건. 자세한 내용: 실행기록.log')
                return 2 if failed else 0
            time.sleep(1 if args.once else 2)


def main():
    parser = argparse.ArgumentParser(description='PDF → 편집 가능한 한글 문서 (2018/2022 대상)')
    parser.add_argument('--background-worker', action='store_true', help=argparse.SUPPRESS)
    parser.add_argument('--once', action='store_true', help='현재 파일만 처리 후 종료')
    parser.set_defaults(dpi=300)
    parser.add_argument('--ocr', choices=('auto','always'), default='auto', help='자동 판단 또는 전 페이지 OCR')
    parser.add_argument('--timeout', type=int, default=600, help='문서 변환 대기 한도(초)')
    parser.add_argument('--language', choices=('kor','kor+eng','eng'), default='kor', help='OCR 언어 (기본 한국어 우선)')
    args = parser.parse_args()
    os.environ['PDF_OCR_LANGUAGE'] = args.language
    global OCR_MODE
    OCR_MODE = args.ocr
    if not 72 <= args.dpi <= 600 or not 10 <= args.timeout <= 3600:
        parser.error('DPI는 72~600, 대기 한도는 10~3600초로 지정하세요.')
    import pymupdf
    watch, output = BASE / '변환할 PDF', BASE / '한글 변환 완료'
    watch.mkdir(exist_ok=True)
    output.mkdir(exist_ok=True)
    (BASE / '확인 필요').mkdir(exist_ok=True)
    with exclusive_lock(BASE / '.running.lock'):
        if not args.background_worker:
            (BASE / '.stop-request').unlink(missing_ok=True)
        setup_logging()
        log('자동 추출 시작. 종료: Ctrl+C. 문서 원본은 유지됩니다.')
        log('입력: ' + str(watch))
        try:
            return run(watch, output, args)
        except KeyboardInterrupt:
            log('종료했습니다.')
            return 0


if __name__ == '__main__':
    try:
        raise SystemExit(main())
    except Exception as exc:
        print('실행 오류: ' + str(exc), flush=True)
        raise SystemExit(1)
