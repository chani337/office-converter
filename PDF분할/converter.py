"""한글/Word/PowerPoint/Excel → PDF. 형식별 작업을 별도 프로세스로 격리."""
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys

IS_WINDOWS = os.name == 'nt'
KINDS = {
    '.hwp': 'hwp', '.hwpx': 'hwp',
    '.doc': 'word', '.docx': 'word', '.rtf': 'word', '.odt': 'word',
    '.ppt': 'powerpoint', '.pptx': 'powerpoint',
    '.xls': 'excel', '.xlsx': 'excel', '.pdf': 'pdf',
}
SUPPORTED_EXTS = frozenset(KINDS)
LABELS = {'hwp': '한글', 'word': 'Word', 'powerpoint': 'PowerPoint', 'excel': 'Excel'}

class ConversionError(RuntimeError):
    pass

class ConversionTimeout(ConversionError):
    pass

# 이전 버전 테스트/API 호환
HwpTimeout = ConversionTimeout


def hwp_worker(src, dst):
    from runtime import hwp_session
    with hwp_session():
        return _hwp_worker(src, dst)


def _hwp_worker(src, dst):
    import pythoncom
    import win32com.client
    import winreg
    dll = Path(__file__).with_name('FilePathCheckerModuleExample.dll')
    if not dll.is_file():
        raise ConversionError('보안 DLL이 없습니다. ZIP을 전체 압축 해제하세요.')
    # 기존 PDFLocal의 레지스트리 등록 경로와 충돌하지 않는 전용 이름.
    module = 'PDFSplitSecurityModule'
    for key in (r'SOFTWARE\HNC\HwpAutomation\Modules', r'SOFTWARE\HNC\HwpCtrl\Modules'):
        with winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, key, 0, winreg.KEY_SET_VALUE) as k:
            winreg.SetValueEx(k, module, 0, winreg.REG_SZ, str(dll.resolve()))
    pythoncom.CoInitialize()
    app = None
    try:
        try:
            app = win32com.client.DispatchEx('HWPFrame.HwpObject')
        except Exception as exc:
            raise ConversionError('한글 자동화를 시작하지 못했습니다. 데스크톱 한글을 설치하고 한 번 직접 실행하세요. ' + str(exc)) from exc
        # 기존 PDFLocal처럼 변환용 한글 창을 숨긴다.
        try:
            app.XHwpWindows.Item(0).Visible = False
        except Exception:
            pass
        if not app.RegisterModule('FilePathCheckDLL', module):
            raise ConversionError('한글 보안 모듈을 불러오지 못했습니다. 한글 버전과 DLL 호환성을 확인하세요.')
        if not app.Open(str(src), '', 'forceopen:true'):
            raise ConversionError('한글 파일을 열지 못했습니다. 암호/손상 여부를 확인하세요.')
        saved = False
        try:
            saved = bool(app.SaveAs(str(dst), 'PDF', ''))
        except Exception:
            pass
        if not saved or not dst.is_file() or not dst.stat().st_size:
            pset = app.HParameterSet.HFileOpenSave
            app.HAction.GetDefault('FileSaveAsPdf', pset.HSet)
            pset.filename = str(dst)
            pset.Format = 'PDF'
            if not app.HAction.Execute('FileSaveAsPdf', pset.HSet):
                raise ConversionError('한글 PDF 저장에 실패했습니다.')
        if not dst.is_file() or not dst.stat().st_size:
            raise ConversionError('한글이 PDF를 만들지 못했습니다.')
    finally:
        if app is not None:
            try:
                app.Clear(1)
            except Exception:
                pass
            try:
                app.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


def office_worker(src, dst, kind):
    import pythoncom
    import win32com.client
    ids = {'word': 'Word.Application', 'powerpoint': 'PowerPoint.Application', 'excel': 'Excel.Application'}
    pythoncom.CoInitialize()
    app = document = None
    try:
        try:
            app = win32com.client.DispatchEx(ids[kind])
        except Exception as exc:
            raise ConversionError(f'{LABELS[kind]}를 시작하지 못했습니다. Windows에 데스크톱 {LABELS[kind]}가 설치되어 있는지 확인하세요. {exc}') from exc
        # 입력 문서의 매크로는 실행하지 않는다.
        app.AutomationSecurity = 3
        try:
            app.DisplayAlerts = 1 if kind == 'powerpoint' else 0
        except Exception:
            pass
        if kind == 'word':
            document = app.Documents.Open(str(src), ReadOnly=True, AddToRecentFiles=False,
                                          ConfirmConversions=False)
            document.ExportAsFixedFormat(str(dst), 17)
        elif kind == 'powerpoint':
            document = app.Presentations.Open(str(src), WithWindow=False, ReadOnly=True)
            document.SaveAs(str(dst), 32)
        elif kind == 'excel':
            document = app.Workbooks.Open(str(src), ReadOnly=True, UpdateLinks=0)
            document.ExportAsFixedFormat(0, str(dst))
        if not dst.is_file() or not dst.stat().st_size:
            raise ConversionError(f'{LABELS[kind]}가 PDF를 만들지 못했습니다.')
    finally:
        if document is not None:
            try:
                if kind == 'powerpoint':
                    document.Close()
                else:
                    document.Close(False)
            except Exception:
                pass
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


def convert_worker(src, dst):
    kind = KINDS.get(src.suffix.lower())
    if kind == 'hwp':
        hwp_worker(src, dst)
    elif kind in ('word', 'powerpoint', 'excel'):
        office_worker(src, dst, kind)
    else:
        raise ConversionError('지원하지 않는 작업 형식입니다.')


class Converter:
    def __init__(self, cfg=None, log=None):
        self.timeout = (cfg or {}).get('convert_timeout_sec', 300) or 300
        self.log = log or print
        self.blocked_kinds = set()

    def __enter__(self):
        return self

    def __exit__(self, *args):
        return False

    def convert(self, src, dst):
        src, dst = Path(src).resolve(), Path(dst).resolve()
        dst.parent.mkdir(parents=True, exist_ok=True)
        if src == dst:
            raise ConversionError('입력과 출력 경로가 같습니다.')
        if src.suffix.lower() == '.pdf':
            shutil.copy2(src, dst)
            return str(dst)
        kind = KINDS.get(src.suffix.lower())
        if kind is None:
            raise ConversionError('지원하지 않는 형식입니다.')
        label = LABELS[kind]
        if not IS_WINDOWS:
            raise ConversionError(f'{label} 변환은 Windows + 데스크톱 {label} 설치가 필요합니다.')
        if kind in self.blocked_kinds:
            raise ConversionTimeout(f'{label} 응답 대기 초과로 해당 형식의 변환을 중지했습니다. 확인창을 정리하고 다시 실행하세요.')
        report = dst.with_suffix('.result.json')
        env = dict(os.environ, PYTHONUTF8='1', PYTHONIOENCODING='utf-8')
        try:
            result = subprocess.run(
                [sys.executable, str(Path(__file__).resolve()), str(src), str(dst), str(report)],
                timeout=self.timeout, capture_output=True, env=env)
        except (subprocess.TimeoutExpired, KeyboardInterrupt) as exc:
            self.blocked_kinds.add(kind)
            if isinstance(exc, KeyboardInterrupt):
                raise
            raise ConversionTimeout(f'{label}가 {self.timeout}초 동안 응답하지 않았습니다. '
                             '확인창을 정리하고 다시 실행하세요. 다른 형식은 계속 처리합니다.') from exc
        try:
            payload = json.loads(report.read_text(encoding='utf-8'))
        except (OSError, ValueError):
            detail = result.stderr.decode('utf-8', errors='replace')[-1000:]
            raise ConversionError(f'{label} 작업 프로세스 오류: ' + detail)
        finally:
            report.unlink(missing_ok=True)
        if result.returncode or not payload.get('ok'):
            raise ConversionError(payload.get('error', f'{label} 변환 오류'))
        if not dst.is_file() or not dst.stat().st_size:
            raise ConversionError('변환 결과 PDF가 없습니다.')
        return str(dst)


if __name__ == '__main__':
    src, dst, report = map(Path, sys.argv[1:4])
    try:
        convert_worker(src, dst)
        payload = {'ok': True}
    except Exception as exc:
        payload = {'ok': False, 'error': str(exc)}
    report.write_text(json.dumps(payload, ensure_ascii=False), encoding='utf-8')
    sys.exit(0 if payload['ok'] else 1)
