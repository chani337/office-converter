"""파일 하나의 전체 변환/완료 기록. 실패 시 편집용 결과는 확인 필요에 보관."""
import hashlib
import json
import os
from pathlib import Path
import shutil
import tempfile
from prepare import prepare


def unique(parent,name):
    candidate=parent/name
    number=1
    while candidate.exists():
        candidate=parent/f'{name} ({number})'
        number+=1
    return candidate


def process(src,watch,output,engine,dpi=300,mode='auto'):
    from split import digest,copy_snapshot,log
    src,watch,output=map(Path,(src,watch,output))
    relative=src.relative_to(watch)
    parent=output/relative.parent
    parent.mkdir(parents=True,exist_ok=True)
    fingerprint=digest(src)
    for folder in parent.iterdir():
        manifest=folder/'.complete.json'
        try:
            data=json.loads(manifest.read_text(encoding='utf-8'))
            if (data['version']==2 and data['source']==relative.as_posix()
                    and data['hash']==fingerprint and data['mode']==mode
                    and data.get('language','kor')==os.environ.get('PDF_OCR_LANGUAGE','kor')
                    and isinstance(data.get('outputs'),dict)
                    and src.stem+'.hwp' in data['outputs']
                    and all(isinstance(name,str) and Path(name).name==name and isinstance(size,int) and size>0 for name,size in data['outputs'].items())
                    and all((folder/name).is_file() and (folder/name).stat().st_size==size
                            for name,size in data['outputs'].items())):
                return False
        except (OSError,ValueError,KeyError,TypeError,AttributeError):pass
    with tempfile.TemporaryDirectory(prefix='.작업-',dir=parent) as temp:
        temp=Path(temp);snapshot=temp/'source.pdf';work=temp/'result';work.mkdir()
        copy_snapshot(src,snapshot)
        if digest(snapshot)!=fingerprint:raise RuntimeError('입력 파일이 변경되었습니다. 다시 시도합니다.')
        log('PDF→한글 변환: '+str(relative))
        docx,report=prepare(snapshot,work,mode,log)
        shutil.copy2(snapshot,work/'원본비교.pdf')
        hwp=work/(src.stem+'.hwp')
        try:
            validation=engine.convert(docx,hwp)
            report.update(validation)
            if not hwp.is_file() or hwp.stat().st_size==0:
                raise RuntimeError('HWP 파일이 만들어지지 않았습니다.')
        except Exception as exc:
            report['hwp_error']=str(exc)
            (work/'변환보고서.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
            failures=output.parent/'확인 필요'/relative.parent
            failures.mkdir(parents=True,exist_ok=True)
            partial=unique(failures,src.name)
            work.rename(partial)
            log('HWP 미완료. 편집용 DOCX/추출문자 보관: '+str(partial))
            raise
        if digest(src)!=fingerprint:raise RuntimeError('변환 중 입력 파일이 변경되었습니다. 다시 시도합니다.')
        if report.get('hwp_text_retention',1)<.97:
            report['warnings'].append('한글로 저장하는 과정에서 문자가 누락되었을 수 있습니다. DOCX 및 추출문자와 비교하세요.')
        (work/'변환보고서.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
        lines=['PDF→한글 변환 결과',f'원본: {src.name}',f'원본 페이지 수: {len(report["pages"])}',
               '저장 형식: HWP (한글 2018/2022 공통 형식)',
               '실행한 한글 버전: '+str(report.get('hancom_version','확인 불가')),
               '원본 PDF와 비교하여 확인해 주세요. 자동 검사는 OCR 정답률이나 서식 일치를 보증하지 않습니다.','']
        lines.extend(report['warnings'])
        for page in report['pages']:
            lines.append(f'\n{page["page"]}쪽: '+('OCR' if page['ocr'] else '직접 추출'))
            lines.extend(page['warnings'])
        (work/'확인사항.txt').write_text('\n'.join(lines),encoding='utf-8-sig')
        files={p.name:p.stat().st_size for p in work.iterdir() if p.is_file()}
        (work/'.complete.json').write_text(json.dumps({'version':2,'source':relative.as_posix(),'hash':fingerprint,
                    'mode':mode,'language':os.environ.get('PDF_OCR_LANGUAGE','kor'),'outputs':files},ensure_ascii=False),encoding='utf-8')
        final=unique(parent,src.name)
        work.rename(final)
    log('HWP 저장 완료 (원본 대조 필요): '+str(final))
    return True
