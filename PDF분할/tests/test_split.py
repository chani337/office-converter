"""Run: python -m unittest discover -s tests -v"""
import argparse
import json
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest
from unittest.mock import patch
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import pymupdf
import split
import converter
from runtime import exclusive_lock, AlreadyRunning

class PipelineTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.watch = self.root / '변환할 파일'; self.watch.mkdir()
        self.out = self.root / 'PDF 변환 완료'
        self.engine = converter.Converter()

    def source(self, name='한글 문서.pdf', count=3):
        path = self.watch / name
        path.parent.mkdir(parents=True, exist_ok=True)
        with pymupdf.open() as doc:
            for i in range(count):
                page = doc.new_page(width=300+i*30,height=400)
                page.insert_text((30,50), f'Page {i+1}',fontsize=20)
                if i==1: page.set_rotation(90)
            doc.save(path)
        return path

    def process(self, src):
        return split.process(src,self.watch,self.out,self.engine)

    def files(self, kind):
        return sorted(p for p in (self.out/('페이지별 '+kind)).rglob('*') if p.suffix.lower()=='.'+kind.lower() and p.is_file())

    def test_pages_text_rotation_and_jpeg(self):
        src=self.source(); before=src.read_bytes();self.process(src)
        self.assertEqual(len(self.files('JPG')),3)
        self.assertEqual(len(self.files('PDF')),3)
        for i,p in enumerate(self.files('PDF')):
            with pymupdf.open(p) as doc:
                self.assertEqual(len(doc),1)
                self.assertIn(f'Page {i+1}',doc[0].get_text())
                self.assertEqual(doc[0].rotation,90 if i==1 else 0)
                image=pymupdf.Pixmap(self.files('JPG')[i])
                self.assertGreater(image.width,0)
        self.assertEqual(src.read_bytes(),before)
        full = self.out/'전체 PDF'/src.name/(src.stem+'.pdf')
        self.assertEqual(full.read_bytes(), before)
        with pymupdf.open(full) as doc: self.assertEqual(len(doc),3)

    def test_intermediate_filename(self):
        self.assertTrue(self.process(self.source('intermediate.pdf')))

    def test_skip_completed(self):
        src=self.source();self.process(src)
        self.assertFalse(self.process(src));self.assertEqual(len(self.files('PDF')),3)

    def test_missing_pdf_recovers(self):
        src=self.source();self.process(src);self.files('PDF')[0].unlink()
        self.assertTrue(self.process(src));self.assertEqual(len(self.files('PDF')),5)
        self.assertFalse(self.process(src))

    def test_empty_jpg_recovers(self):
        src=self.source();self.process(src);self.files('JPG')[0].write_bytes(b'')
        self.assertTrue(self.process(src))

    def test_changed_source_preserves_old(self):
        src=self.source();self.process(src);old=self.files('PDF')[0].read_bytes()
        src.unlink();self.source(count=1);self.process(src)
        self.assertEqual(len(self.files('PDF')),4)
        self.assertIn(old,[p.read_bytes() for p in self.files('PDF')])

    def test_nested_same_names(self):
        for name in ['A/같은 이름.pdf','B/같은 이름.pdf']:
            self.process(self.source(name,1))
        self.assertEqual(len(self.files('PDF')),2)

    def test_invalid_pdf_no_outputs(self):
        src=self.watch/'broken.pdf';src.write_bytes(b'bad')
        with self.assertRaises(Exception): self.process(src)
        self.assertFalse(self.files('PDF'));self.assertFalse(self.files('JPG'))
        self.assertFalse(list(self.out.rglob('.작업-*')))

    def test_encrypted_pdf(self):
        src=self.watch/'locked.pdf'
        with pymupdf.open() as d:
            d.new_page();d.save(src,encryption=pymupdf.PDF_ENCRYPT_AES_256,user_pw='secret',owner_pw='owner')
        with self.assertRaises(ValueError):self.process(src)
        self.assertFalse(self.files('PDF'))

    def test_malformed_manifest_recovery(self):
        src=self.source();self.process(src)
        next(self.out.rglob('.complete.json')).write_text('[]')
        self.assertTrue(self.process(src))

    def test_change_during_conversion(self):
        src=self.source()
        real=self.engine.convert
        def convert(a,b):
            real(a,b);src.write_bytes(b'changed')
        with patch.object(self.engine,'convert',side_effect=convert):
            with self.assertRaises(RuntimeError):self.process(src)
        self.assertFalse(self.files('PDF'))

    def test_publish_failure_rolls_back(self):
        src=self.source();real=Path.rename
        def rename(path,target):
            if path.name=='pages':raise OSError('simulated permission error')
            return real(path,target)
        with patch.object(Path,'rename',rename):
            with self.assertRaises(OSError):self.process(src)
        self.assertFalse(self.files('PDF'));self.assertFalse(self.files('JPG'))
        self.assertFalse(list((self.out/'전체 PDF').iterdir()))

    def test_missing_full_pdf_recovers(self):
        src=self.source();self.process(src)
        (self.out/'전체 PDF'/src.name/(src.stem+'.pdf')).unlink()
        self.assertTrue(self.process(src))
        self.assertFalse(self.process(src))

    def test_single_lock(self):
        lock=self.root/'lock'
        with exclusive_lock(lock):
            with self.assertRaises(AlreadyRunning):
                with exclusive_lock(lock):pass
        with exclusive_lock(lock):pass

    def test_once_exit_codes_and_continues(self):
        self.source('good.pdf',1)
        bad=self.watch/'bad.pdf';bad.write_bytes(b'invalid')
        clock=[0]
        def sleep(n):clock[0]+=n
        with patch.object(split.time,'monotonic',side_effect=lambda:clock[0]),patch.object(split.time,'sleep',side_effect=sleep):
            result=split.run(self.watch,self.out,argparse.Namespace(once=True,dpi=200,timeout=300))
        self.assertEqual(result,2);self.assertEqual(len(self.files('PDF')),1)
        self.assertLess(clock[0],10)

    def test_hwp_path_mocked(self):
        src=self.watch/'한글 문서.hwpx';src.write_bytes(b'fake hwp input')
        pdf=self.source('reference.pdf',2)
        import shutil
        with patch.object(self.engine,'convert',side_effect=lambda a,b:shutil.copyfile(pdf,b)):
            self.process(src)
        self.assertEqual(len(self.files('PDF')),2)


class WorkerTests(unittest.TestCase):
    def test_timeout_blocks_only_hwp(self):
        with tempfile.TemporaryDirectory() as tmp:
            root=Path(tmp);src=root/'a.hwp';src.write_bytes(b'x');dst=root/'out.pdf'
            engine=converter.Converter({'convert_timeout_sec':10})
            with patch.object(converter,'IS_WINDOWS',True),patch.object(converter.subprocess,'run',side_effect=subprocess.TimeoutExpired('worker',10)):
                with self.assertRaises(converter.HwpTimeout):engine.convert(src,dst)
                with self.assertRaises(converter.HwpTimeout):engine.convert(src,dst)
            pdf=root/'input.pdf';pdf.write_bytes(b'pdf')
            engine.convert(pdf,dst)
            self.assertEqual(dst.read_bytes(),b'pdf')

    def test_worker_error_report(self):
        with tempfile.TemporaryDirectory() as tmp:
            root=Path(tmp);src=root/'a.hwp';src.write_bytes(b'x');dst=root/'out.pdf'
            def result(*args,**kwargs):
                dst.with_suffix('.result.json').write_text(json.dumps({'ok':False,'error':'test error'}))
                return subprocess.CompletedProcess([],1,b'',b'')
            with patch.object(converter,'IS_WINDOWS',True),patch.object(converter.subprocess,'run',side_effect=result):
                with self.assertRaisesRegex(converter.ConversionError,'test error'):
                    converter.Converter().convert(src,dst)
            self.assertFalse(dst.with_suffix('.result.json').exists())

    def test_windows_lock_does_not_write(self):
        import runtime
        import types
        from unittest.mock import MagicMock
        handle=MagicMock();ms=types.SimpleNamespace(LK_NBLCK=2,locking=MagicMock())
        with patch.object(runtime,'os',types.SimpleNamespace(name='nt')),patch('builtins.open',return_value=handle),patch.dict(sys.modules,{'msvcrt':ms}):
            with exclusive_lock('unused'):pass
        handle.write.assert_not_called();ms.locking.assert_called_once();handle.close.assert_called_once()

    def test_bootstrap_reuses_healthy_environment(self):
        import bootstrap
        with patch.object(bootstrap,'usable',return_value=True),patch.object(bootstrap.subprocess,'run') as runner:
            bootstrap.ensure_environment()
        runner.assert_not_called()



class BackgroundTests(unittest.TestCase):
    def test_stop_before_scan(self):
        with tempfile.TemporaryDirectory() as tmp:
            base=Path(tmp);(base/'.stop-request').touch()
            with patch.object(split,'BASE',base),patch.object(split,'process') as process:
                result=split.run(base,base,argparse.Namespace(once=False,dpi=200,timeout=300))
            self.assertEqual(result,0);process.assert_not_called()

    def test_stop_after_current_file(self):
        with tempfile.TemporaryDirectory() as tmp:
            base=Path(tmp);watch=base/'in';watch.mkdir()
            (watch/'a.pdf').write_bytes(b'a');(watch/'b.pdf').write_bytes(b'b')
            clock=[0]
            def sleep(n):clock[0]+=n
            def process(*args): (base/'.stop-request').touch()
            with patch.object(split,'BASE',base),patch.object(split.time,'sleep',side_effect=sleep),patch.object(split.time,'monotonic',side_effect=lambda:clock[0]),patch.object(split,'process',side_effect=process) as worker:
                result=split.run(watch,base/'out',argparse.Namespace(once=False,dpi=200,timeout=300))
            self.assertEqual(result,0);self.assertEqual(worker.call_count,1)

    def test_pythonw_logging(self):
        import logging
        with tempfile.TemporaryDirectory() as tmp:
            with patch.object(split,'BASE',Path(tmp)),patch.object(sys,'stdout',None):
                split.setup_logging();split.log('background test')
            logger=logging.getLogger('pdfsplit')
            for h in logger.handlers:h.close()
            logger.handlers.clear()
            self.assertIn('background test',(Path(tmp)/'실행기록.log').read_text())



class OfficeTests(unittest.TestCase):
    def test_all_original_formats_detected(self):
        formats={'.hwp','.hwpx','.doc','.docx','.rtf','.odt','.ppt','.pptx','.xls','.xlsx','.pdf'}
        self.assertEqual(converter.SUPPORTED_EXTS,formats)
        with tempfile.TemporaryDirectory() as tmp:
            root=Path(tmp)
            for ext in formats:(root/('문서'+ext.upper())).write_bytes(b'x')
            (root/'~$lock.docx').write_bytes(b'x')
            self.assertEqual(len(list(split.scan_files(root))),len(formats))

    def test_office_native_pdf_calls_and_cleanup(self):
        from unittest.mock import MagicMock
        import types
        for kind in ('word','powerpoint','excel'):
            with self.subTest(kind=kind),tempfile.TemporaryDirectory() as tmp:
                root=Path(tmp);src=root/'input';dst=root/'output.pdf'
                app=MagicMock();doc=MagicMock();pc=MagicMock()
                app.Documents.Open.return_value=doc
                app.Presentations.Open.return_value=doc
                app.Workbooks.Open.return_value=doc
                def write(*args):dst.write_bytes(b'%PDF-test')
                doc.ExportAsFixedFormat.side_effect=write
                doc.SaveAs.side_effect=write
                client=types.ModuleType('win32com.client');client.DispatchEx=MagicMock(return_value=app)
                package=types.ModuleType('win32com');package.client=client
                with patch.dict(sys.modules,{'win32com':package,'win32com.client':client,'pythoncom':pc}):
                    converter.office_worker(src,dst,kind)
                if kind=='word':doc.ExportAsFixedFormat.assert_called_once_with(str(dst),17)
                elif kind=='powerpoint':doc.SaveAs.assert_called_once_with(str(dst),32)
                else:doc.ExportAsFixedFormat.assert_called_once_with(0,str(dst))
                self.assertEqual(app.AutomationSecurity,3)
                doc.Close.assert_called_once();app.Quit.assert_called_once();pc.CoUninitialize.assert_called_once()

    def test_office_exports_reach_whole_pdf_and_splits(self):
        import shutil
        with tempfile.TemporaryDirectory() as tmp:
            root=Path(tmp);watch=root/'in';watch.mkdir();output=root/'out';fixture=root/'fixture.pdf'
            with pymupdf.open() as doc:
                doc.new_page();doc.new_page();doc.save(fixture)
            engine=converter.Converter()
            for ext in ('.doc','.docx','.rtf','.odt','.ppt','.pptx','.xls','.xlsx'):
                src=watch/('보고서'+ext);src.write_bytes(b'office placeholder')
                with patch.object(engine,'convert',side_effect=lambda a,b:shutil.copyfile(fixture,b)) as export:
                    self.assertTrue(split.process(src,watch,output,engine))
                    export.assert_called_once()
                with pymupdf.open(output/'전체 PDF'/src.name/'보고서.pdf') as doc:self.assertEqual(len(doc),2)
                self.assertEqual(len(list((output/'페이지별 PDF'/src.name).glob('*.pdf'))),2)
                self.assertEqual(len(list((output/'페이지별 JPG'/src.name).glob('*.jpg'))),2)

    def test_word_timeout_does_not_block_powerpoint(self):
        with tempfile.TemporaryDirectory() as tmp:
            root=Path(tmp);src=root/'a.docx';src.write_bytes(b'x');dst=root/'out.pdf'
            engine=converter.Converter({'convert_timeout_sec':10})
            with patch.object(converter,'IS_WINDOWS',True),patch.object(converter.subprocess,'run',side_effect=subprocess.TimeoutExpired('worker',10)):
                with self.assertRaises(converter.ConversionTimeout):engine.convert(src,dst)
            self.assertEqual(engine.blocked_kinds,{'word'})
            ppt=root/'a.pptx';ppt.write_bytes(b'x')
            def success(*a,**kw):
                dst.write_bytes(b'%PDF');dst.with_suffix('.result.json').write_text('{"ok":true}')
                return subprocess.CompletedProcess([],0,b'',b'')
            with patch.object(converter,'IS_WINDOWS',True),patch.object(converter.subprocess,'run',side_effect=success):
                self.assertEqual(engine.convert(ppt,dst),str(dst.resolve()))

if __name__=='__main__':unittest.main()
