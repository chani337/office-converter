"""한글 2018/2022 공통 COM API로 DOCX를 HWP 5.x로 저장."""
import json
import os
from pathlib import Path
import subprocess
import sys
from runtime import hwp_session

SUPPORTED_EXTS = {'.pdf'}
class ConversionTimeout(RuntimeError):
    pass


def open_docx_in_hwp(app, path):
    """DOCX를 한워드로 전달하지 않고 현재 한글 문서로 삽입한다.

    한컴 공식 안내: https://forum.developer.hancom.com/t/topic/1639
    형식명: https://forum.developer.hancom.com/t/saveas/660
    """
    errors = []
    for file_format in ('OOXML', ''):
        # 실패한 불러오기에서 일부 내용이 남아도 다음 시도에 중복 삽입하지 않는다.
        app.Clear(1)
        try:
            if app.Open(str(path), file_format, 'forceopen:true;insert:true'):
                return 'OOXML 삽입' if file_format else '자동 인식 삽입'
            errors.append((file_format or '자동 인식') + ': Open 반환값 False')
        except Exception as exc:
            errors.append((file_format or '자동 인식') + ': ' + str(exc))
    raise RuntimeError('DOCX를 한글 안으로 불러오지 못했습니다. ' + ' / '.join(errors))


def save_hwp(docx, dst):
    import pythoncom
    import win32com.client
    import winreg
    from prepare import docx_text, text_coverage
    dll = Path(__file__).with_name('FilePathCheckerModuleExample.dll')
    module = 'PDFToHwpSecurityModule'
    if not dll.exists():
        raise RuntimeError('한글 보안 DLL이 없습니다. ZIP을 전체 압축 해제하세요.')
    for key in (r'SOFTWARE\HNC\HwpAutomation\Modules', r'SOFTWARE\HNC\HwpCtrl\Modules'):
        with winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER,key,0,winreg.KEY_SET_VALUE) as reg:
            winreg.SetValueEx(reg,module,0,winreg.REG_SZ,str(dll.resolve()))
    pythoncom.CoInitialize()
    app=None
    try:
        app=win32com.client.DispatchEx('HWPFrame.HwpObject')
        try: app.XHwpWindows.Item(0).Visible=False
        except Exception: pass
        if not app.RegisterModule('FilePathCheckDLL',module):
            raise RuntimeError('한글 보안 모듈을 읽지 못했습니다. 한글 버전과 DLL 호환성을 확인하세요.')
        import_method = open_docx_in_hwp(app, docx)
        try: version=str(app.Version)
        except Exception: version='확인 불가'
        if not app.SaveAs(str(dst),'HWP',''):
            raise RuntimeError('한글 HWP 저장에 실패했습니다.')
        if not dst.exists() or dst.read_bytes()[:8] != bytes.fromhex('d0cf11e0a1b11ae1'):
            raise RuntimeError('HWP 파일 형식 검사에 실패했습니다.')
        app.Clear(1)
        if not app.Open(str(dst),'HWP','forceopen:true'):
            raise RuntimeError('저장한 HWP를 다시 열지 못했습니다.')
        actual=app.GetTextFile('TEXT','')
        expected=docx_text(docx)
        if expected.strip() and not str(actual).strip():
            raise RuntimeError('저장한 HWP에서 편집 가능한 글자를 확인하지 못했습니다.')
        return {'hancom_version':version,'import_method':import_method,'hwp_text_retention':round(text_coverage(expected,str(actual)),4)}
    finally:
        if app is not None:
            try: app.Clear(1)
            except Exception: pass
            try: app.Quit()
            except Exception: pass
        pythoncom.CoUninitialize()


class Converter:
    def __init__(self,cfg=None,log=None):
        self.timeout=(cfg or {}).get('convert_timeout_sec',300)
        self.log=log or print
        self.blocked=False
    def __enter__(self):return self
    def __exit__(self,*args):return False
    def convert(self,src,dst):
        if os.name!='nt':
            raise RuntimeError('HWP 저장은 Windows + 설치된 한글 2018/2022가 필요합니다.')
        if self.blocked:
            raise ConversionTimeout('한글 응답 대기 초과 상태입니다. 한글을 확인하고 다시 실행하세요.')
        src,dst=Path(src).resolve(),Path(dst).resolve()
        report=dst.with_suffix('.worker.json')
        try:
            result=subprocess.run([sys.executable,str(Path(__file__).resolve()),str(src),str(dst),str(report)],
                timeout=self.timeout,capture_output=True,creationflags=0x08000000,
                env=dict(os.environ,PYTHONUTF8='1'))
        except subprocess.TimeoutExpired as exc:
            self.blocked=True
            raise ConversionTimeout('한글이 응답하지 않아 HWP 저장을 중지했습니다. 확인창을 정리하고 다시 실행하세요.') from exc
        if not report.exists():
            raise RuntimeError('한글 작업 실행 실패: '+result.stderr.decode('utf-8',errors='replace')[-800:])
        payload=json.loads(report.read_text(encoding='utf-8'))
        report.unlink(missing_ok=True)
        if result.returncode or not payload.get('ok'):
            raise RuntimeError(payload.get('error','HWP 저장 실패'))
        return payload['validation']


if __name__=='__main__':
    src,dst,report=map(Path,sys.argv[1:4])
    try:
        with hwp_session():
            validation=save_hwp(src,dst)
        payload={'ok':True,'validation':validation}
    except Exception as exc:
        payload={'ok':False,'error':str(exc)}
    report.write_text(json.dumps(payload,ensure_ascii=False),encoding='utf-8')
    sys.exit(0 if payload['ok'] else 1)
