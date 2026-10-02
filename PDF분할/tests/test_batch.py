"""여러 입력을 한 번에 넣는 경우의 통합 테스트."""
import argparse
from pathlib import Path
import shutil
import sys
import tempfile
import unittest
from unittest.mock import patch
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
import pymupdf
import split
import converter

class BatchTests(unittest.TestCase):
    def test_mixed_batch_and_failure_isolation(self):
        with tempfile.TemporaryDirectory() as tmp:
            base=Path(tmp);watch=base/'in';watch.mkdir();out=base/'out'
            counts={'보고서.hwp':3,'보고서.pptx':5,'표.xlsx':2,'문서.docx':4,'첨부.pdf':1}
            fixtures={}
            for name,count in counts.items():
                fixture=base/(name+'.pdf')
                with pymupdf.open() as doc:
                    for i in range(count):doc.new_page().insert_text((40,40),str(i+1))
                    doc.save(fixture)
                fixtures[Path(name).suffix]=fixture
                (watch/name).write_bytes(fixture.read_bytes() if name.endswith('.pdf') else b'office input')
            (watch/'bad.pdf').write_bytes(b'broken')
            original=converter.Converter.convert
            def convert(engine,src,dst):
                if Path(src).suffix=='.pdf':return original(engine,src,dst)
                shutil.copyfile(fixtures[Path(src).suffix],dst)
                return dst
            clock=[0]
            def sleep(n):clock[0]+=n
            with patch.object(split,'BASE',base),patch.object(split.time,'monotonic',side_effect=lambda:clock[0]),patch.object(split.time,'sleep',side_effect=sleep),patch.object(converter.Converter,'convert',convert):
                result=split.run(watch,out,argparse.Namespace(once=True,dpi=200,timeout=300))
            self.assertEqual(result,2)
            for name,count in counts.items():
                with pymupdf.open(out/'전체 PDF'/name/(Path(name).stem+'.pdf')) as doc:self.assertEqual(len(doc),count)
                self.assertEqual(len(list((out/'페이지별 PDF'/name).glob('*.pdf'))),count)
                self.assertEqual(len(list((out/'페이지별 JPG'/name).glob('*.jpg'))),count)
            self.assertEqual(len(list((out/'전체 PDF').iterdir())),len(counts))
            self.assertFalse((out/'전체 PDF'/'bad.pdf').exists())

if __name__=='__main__':unittest.main()
