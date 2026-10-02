"""Windows 한글 문서 → 임시 PDF → 전체 PDF 및 페이지별 JPG/PDF. 원본 보존."""
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

def render(pdf, folder, stem, dpi=200, pdf_folder=None):
    import pymupdf
    names = []
    with pymupdf.open(pdf) as doc:
        if not doc.is_pdf or doc.needs_pass or not len(doc):
            raise ValueError('암호가 있거나 페이지가 없는 PDF입니다.')
        for number, page in enumerate(doc, 1):
            name = f'{stem}_{number:03d}.jpg'
            pixels = page.rect.width * page.rect.height * (dpi / 72) ** 2
            if pixels > 80_000_000:
                raise ValueError('페이지가 너무 큽니다. --dpi 값을 낮춰 다시 실행하세요.')
            page.get_pixmap(dpi=dpi, colorspace=pymupdf.csRGB, alpha=False).save(
                folder / name, jpg_quality=95)
            if pdf_folder is not None:
                with pymupdf.open() as single:
                    single.insert_pdf(doc, from_page=number - 1, to_page=number - 1)
                    single.save(pdf_folder / f'{stem}_{number:03d}.pdf', garbage=4, deflate=True)
            names.append(name)
    return names

def process(src, watch, output, engine, dpi=200):
    relative = src.relative_to(watch)
    # 확장자를 포함하여 같은 이름의 HWP/HWPX/PDF도 충돌하지 않음.
    parent = output / "페이지별 JPG" / relative.parent
    pdf_parent = output / "페이지별 PDF" / relative.parent
    full_parent = output / "전체 PDF" / relative.parent
    full_parent.mkdir(parents=True, exist_ok=True)
    pdf_parent.mkdir(parents=True, exist_ok=True)
    parent.mkdir(parents=True, exist_ok=True)
    fingerprint = digest(src)
    # 완료된 동일 버전의 모든 JPG가 남아 있을 때만 건너뜀.
    for candidate in parent.glob('*'):
        manifest = candidate / '.complete.json'
        if not manifest.is_file():
            continue
        try:
            data = json.loads(manifest.read_text(encoding='utf-8'))
            if (data['source'] == relative.as_posix() and data['hash'] == fingerprint
                    and data.get('version') == 4
                    and (full_parent / candidate.name / (src.stem + '.pdf')).is_file()
                    and (full_parent / candidate.name / (src.stem + '.pdf')).stat().st_size and data['dpi'] == dpi
                    and isinstance(data['pages'], list) and data['pages']
                    and all(isinstance(n, str) and Path(n).name == n and n.endswith('.jpg')
                            for n in data['pages'])
                    and all((pdf_parent / candidate.name / Path(n).with_suffix('.pdf')).is_file()
                            and (pdf_parent / candidate.name / Path(n).with_suffix('.pdf')).stat().st_size
                            for n in data['pages'])
                    and all((candidate / n).is_file() and (candidate / n).stat().st_size
                            for n in data['pages'])):
                return False
        except (OSError, ValueError, KeyError, TypeError, AttributeError):
            continue
    log('변환 중: ' + str(relative))
    with tempfile.TemporaryDirectory(prefix='.작업-', dir=parent) as tmp:
        tmp = Path(tmp)
        source_dir = tmp / 'source'
        source_dir.mkdir()
        snapshot = source_dir / ('input' + src.suffix.lower())
        import shutil
        copy_snapshot(src, snapshot)
        if digest(snapshot) != fingerprint:
            raise RuntimeError('복사 중 원본이 변경되었습니다. 다시 시도합니다.')
        pdf = tmp / 'intermediate.pdf'
        engine.convert(str(snapshot), str(pdf))
        pages_dir = tmp / 'pages'
        pages_dir.mkdir()
        pdf_dir = tmp / 'pdf_pages'
        pdf_dir.mkdir()
        names = render(pdf, pages_dir, src.stem, dpi, pdf_dir)
        full_dir = tmp / 'full_pdf'
        full_dir.mkdir()
        pdf.rename(full_dir / (src.stem + '.pdf'))
        if digest(src) != fingerprint:
            raise RuntimeError('변환 중 원본이 변경되었습니다. 다시 시도합니다.')
        (pages_dir / '.complete.json').write_text(json.dumps({
            'source': relative.as_posix(), 'hash': fingerprint,
            'version': 4, 'dpi': dpi, 'pages': names}, ensure_ascii=False), encoding='utf-8')
        destination = parent / src.name
        index = 1
        while (destination.exists() or (pdf_parent / destination.name).exists()
               or (full_parent / destination.name).exists()):
            destination = parent / f'{src.name} ({index})'
            index += 1
        pdf_destination = pdf_parent / destination.name
        full_destination = full_parent / destination.name
        published = []
        try:
            for staged, final in ((full_dir, full_destination), (pdf_dir, pdf_destination),
                                  (pages_dir, destination)):
                staged.rename(final)
                published.append(final)
        except BaseException:
            # 이번 실행에서 게시한 폴더만 되돌린다. 완료 기록은 마지막에 게시한다.
            for final in reversed(published):
                shutil.rmtree(final)
            raise
    log(f'완료: {relative} → 전체 PDF 1개 + JPG {len(names)}장 + 페이지별 PDF {len(names)}개 ({destination.name})')
    return True

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
    logger.handlers.clear()
    stream = logging.StreamHandler(sys.stdout) if sys.stdout is not None else None
    logfile = RotatingFileHandler(BASE / '실행기록.log', maxBytes=2_000_000,
                                  backupCount=2, encoding='utf-8')
    for handler in ([stream] if stream is not None else []) + [logfile]:
        handler.setFormatter(logging.Formatter('[%(asctime)s] %(message)s', '%H:%M:%S'))
        logger.addHandler(handler)


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
    parser = argparse.ArgumentParser(description='한글/PDF → 페이지별 JPG/PDF 자동 추출')
    parser.add_argument('--background-worker', action='store_true', help=argparse.SUPPRESS)
    parser.add_argument('--once', action='store_true', help='현재 파일만 처리 후 종료')
    parser.add_argument('--dpi', type=int, default=200)
    parser.add_argument('--timeout', type=int, default=300, help='문서 변환 대기 한도(초)')
    args = parser.parse_args()
    if not 72 <= args.dpi <= 600 or not 10 <= args.timeout <= 3600:
        parser.error('DPI는 72~600, 대기 한도는 10~3600초로 지정하세요.')
    import pymupdf
    watch, output = BASE / '변환할 파일', BASE / 'PDF 변환 완료'
    watch.mkdir(exist_ok=True)
    output.mkdir(exist_ok=True)
    for name in ('전체 PDF', '페이지별 PDF', '페이지별 JPG'):
        (output / name).mkdir(exist_ok=True)
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
