from __future__ import annotations
import io,json,sys,unittest,zipfile
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
from fastapi.testclient import TestClient
from quote_engine import preview,generate,source,QuoteError,sheet_and_shared,value
from app import app
ROOT=Path(__file__).resolve().parents[1]
T01=(ROOT/'private/T01_MASTER_CLEAN.xlsx').read_bytes() if (ROOT/'private/T01_MASTER_CLEAN.xlsx').exists() else None
SRC1=(ROOT/'tests/fixtures/SMARTOPTICS_SYNTHETIC_SOURCE.xlsx').read_bytes()

class PortalTests(unittest.TestCase):
 def setUp(self):self.c=TestClient(app)
 def test_status(self):
  r=self.c.get('/api/status');self.assertEqual(r.status_code,200);self.assertTrue(r.json()['demo_ready']);self.assertEqual(r.json()['t01_master_ready'],T01 is not None)
 def test_home(self):
  r=self.c.get('/');self.assertEqual(r.status_code,200);self.assertIn('Buat Quotation',r.text)
 def test_real_requires_auth(self):
  r=self.c.post('/api/preview',files={'quote':('supplier.xlsx',SRC1)},data={'margin':'15','fx':'18300'});self.assertIn(r.status_code,(401,503))
  r=self.c.post('/api/generate',files={'quote':('supplier.xlsx',SRC1)},data={'margin':'15','fx':'18300'});self.assertIn(r.status_code,(401,503))
 def test_demo_preview(self):
  p=self.c.get('/api/demo',params={'margin':'15','fx':'18300'});self.assertEqual(p.status_code,200)
  self.assertEqual(len(p.json()['items']),3)
  p2=self.c.get('/api/demo',params={'margin':'20','fx':'18300'});self.assertGreater(p2.json()['total_usd'],p.json()['total_usd'])
 def test_demo_xlsx(self):
  r=self.c.get('/api/demo-export',params={'margin':'20','fx':'18300','format':'T01','customer':'Demo Test','project':'New Test'})
  self.assertEqual(r.status_code,200)
  z=zipfile.ZipFile(io.BytesIO(r.content));s,ss=sheet_and_shared(z)
  self.assertEqual(value(s,ss,'B4'),'T01-DEMO-SAMPLE');self.assertEqual(value(s,ss,'A18'),'Demo Test');self.assertEqual(value(s,ss,'C18'),'New Test')
  self.assertEqual(float(value(s,ss,'E15')),self.c.get('/api/demo',params={'margin':'20','fx':'18300'}).json()['total_usd'])
 def test_demo_zip(self):
  r=self.c.get('/api/demo-export',params={'margin':'15','fx':'18300','format':'BOTH'})
  self.assertEqual(r.status_code,200);z=zipfile.ZipFile(io.BytesIO(r.content));self.assertEqual(len(z.namelist()),2)
  for filename in z.namelist():
   b=z.read(filename);self.assertTrue(zipfile.is_zipfile(io.BytesIO(b)))
 def test_demo_invalid(self):
  self.assertEqual(self.c.get('/api/demo',params={'margin':'99'}).status_code,422)
  self.assertEqual(self.c.get('/api/demo-export',params={'format':'T09'}).status_code,422)
 def test_vendor_option1(self):
  p=preview(SRC1,15,18300);self.assertEqual(p['item_count'],3);self.assertEqual(p['total_usd'],3631.80)
 def test_vendor_option2(self):
  p=preview(SRC1,15,18300);self.assertEqual(p['item_count'],3);self.assertEqual(p['total_usd'],3631.80)
 def test_real_output_t01(self):
  if T01 is None:self.skipTest('Branded T01 template is intentionally excluded from public repository')
  name,raw,mime=generate(SRC1,'T01',15,18300,'{"customer":"PT TEST ONLY","attention":"Recipient A + Recipient B","date":"2026-11-26"}',T01)
  self.assertTrue(name.endswith('.xlsx'))
  z=zipfile.ZipFile(io.BytesIO(raw));s,ss=sheet_and_shared(z)
  self.assertEqual(value(s,ss,'C10'),'PT TEST ONLY');self.assertEqual(value(s,ss,'C11'),'Recipient A + Recipient B')
  self.assertEqual(float(value(s,ss,'H45')),3631.80)
  self.assertEqual(len([x for x in z.namelist() if x.startswith('xl/media/')]),3)
  self.assertEqual(len([x for x in z.namelist() if x.startswith('xl/worksheets/sheet') and x.endswith('.xml')]),1)
 def test_real_output_t02(self):
  name,raw,mime=generate(SRC1,'T02',15,18300,'{"customer":"PT TEST ONLY","attention":"Recipient A + Recipient B"}')
  z=zipfile.ZipFile(io.BytesIO(raw));s,ss=sheet_and_shared(z)
  self.assertEqual(value(s,ss,'E9'),'PT TEST ONLY');self.assertEqual(value(s,ss,'E8'),'Recipient A + Recipient B')
  self.assertEqual(float(value(s,ss,'M89')),3631.80)
  self.assertIn('DRAFT',value(s,ss,'E12'))
 def test_real_output_zip(self):
  if T01 is None:self.skipTest('Branded T01 template is intentionally excluded from public repository')
  name,raw,mime=generate(SRC1,'BOTH',15,18300,'{}',T01)
  z=zipfile.ZipFile(io.BytesIO(raw));self.assertEqual(len(z.namelist()),2)
 def test_no_double_margin(self):
  _,repriced,_=generate(SRC1,'T02',15,18300)
  with self.assertRaises(QuoteError):source(repriced)
 def test_no_metadata_item_override(self):
  with self.assertRaises(QuoteError):generate(SRC1,'T01',15,18300,'{"sku":"SOMETHING"}',T01)
 def test_no_markup_confusion(self):
  p=preview(SRC1,20,18300);self.assertGreater(p['total_usd'],3631.80)
 def test_wrong_format(self):
  if T01 is None:self.skipTest('Branded T01 template is intentionally excluded from public repository')
  with self.assertRaises(QuoteError):source(T01)
 def test_no_vendor_source_files_in_package(self):
  self.assertFalse(any(x.suffix=='.xlsx' for x in (ROOT/'static').glob('**/*')))

if __name__=='__main__':unittest.main()
