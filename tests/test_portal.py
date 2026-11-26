from __future__ import annotations
import io,json,re,sys,unittest,zipfile
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
from unittest.mock import patch
from fastapi.testclient import TestClient
from quote_engine import preview,generate,source,QuoteError,sheet_and_shared,value,workbook,setval,pack
from pdf_export import generate_pdf
from preview_compare import compare_workbooks
from app import app,permitted
ROOT=Path(__file__).resolve().parents[1]
T01=(ROOT/'private/T01_MASTER_CLEAN.xlsx').read_bytes() if (ROOT/'private/T01_MASTER_CLEAN.xlsx').exists() else None
SRC1=(ROOT/'tests/fixtures/SMARTOPTICS_SYNTHETIC_SOURCE.xlsx').read_bytes()

def synthetic_commverge_priced():
    """Synthetic single-sheet T01-like input: already sold, one broken subtotal."""
    z=workbook(SRC1);sheet,strings=sheet_and_shared(z)
    for address,val,kind in [
        ('D6','PT COMMVERGE SOLUTIONS INDONESIA','text'),
        ('D16','CUSTOMER QUOTE','text'),
        ('C21','Product Name','text'),('G21','Unit Price (USD)','text'),
        ('C23','SYNTHETIC-SKU','text'),('D23','Synthetic demo device','text'),
        ('F23',2,'number'),('G23',120,'number'),('H23','#VALUE!','text')
    ]:setval(sheet,address,val,kind)
    stamped=pack(z,sheet);output=io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(stamped)) as src,zipfile.ZipFile(output,'w') as dst:
        for entry in src.infolist():
            content=src.read(entry.filename)
            if entry.filename=='xl/workbook.xml':
                content=re.sub(rb'(\bsheet name=")[^"]+',rb'\g<1>Quote to DE',content)
            dst.writestr(entry,content)
    return output.getvalue()

class PortalTests(unittest.TestCase):
 def setUp(self):self.c=TestClient(app)
 def test_status(self):
  r=self.c.get('/api/status');self.assertEqual(r.status_code,200);self.assertTrue(r.json()['demo_ready']);self.assertIsNone(r.json()['t01_master_ready']);self.assertEqual(r.json()['t01_master_location'],'private_supabase_storage')
 def test_home(self):
  r=self.c.get('/');self.assertEqual(r.status_code,200)
  self.assertIn('<select id="rEmail"',r.text)
  self.assertIn('id="rPassword"',r.text)
  self.assertNotIn('rPinSetup',r.text)
  self.assertIn('id="rMasterStatus"',r.text)
  self.assertNotIn('id="rT01"',r.text)
  self.assertIn('id="rDraftDocument"',r.text)
  self.assertNotIn('id="rSourceDocument"',r.text)
  self.assertIn('<option value="T01">CommVerge</option>',r.text)
  self.assertIn('<option value="T02">Smartoptics</option>',r.text)
  self.assertNotIn('CommVerge (T01)',r.text)
  self.assertNotIn('Smartoptics (T02)',r.text)
  self.assertNotIn('id="demoPanel"',r.text)
  self.assertNotIn('id="rResultInfo"',r.text)
  self.assertNotIn('id="rPricingNote"',r.text)
  self.assertNotIn('id="rCompareNote"',r.text)
  self.assertIn('class="review-grid"',r.text)
  self.assertIn('<ul id="rCompareChecks" class="qa-list"',r.text)
  self.assertLess(r.text.index('class="cr-logo"'),r.text.index('class="brand-title"'))
  self.assertNotIn('qa-item-index',self.c.get('/static/app.js').text)
  self.assertIn('contain:size',self.c.get('/static/style.css').text)
  self.assertIn('id="rCompareChecks"',r.text)
  self.assertIn('id="rDraftDocument"',r.text)
  self.assertLess(r.text.index('id="rCompareChecks"'),r.text.index('id="rDraftDocument"'))
  self.assertLess(r.text.index('id="rExportType"'),r.text.index('id="rUsd"'))
  self.assertLess(r.text.index('id="rAck"'),r.text.index('id="rGenerate"'))
  self.assertIn('class="download-summary"',r.text)
  self.assertLess(r.text.index('id="rUsd"'),r.text.index('id="rGenerate"'))
  self.assertNotIn('id="rRefreshPdf"',r.text)
  self.assertNotIn('rRefreshPdf',self.c.get('/static/app.js').text)
  self.assertIn('id="rSendEmail"',r.text)
  self.assertIn('id="rPreviewFormat"',r.text)
  self.assertIn('>Michael</option>',r.text)
  self.assertIn('>Daniel</option>',r.text)
  self.assertNotIn('Cuplikan hasil',self.c.get('/static/app.js').text)
  self.assertIn('Collective Revivers</span>',r.text)
 def test_real_requires_auth(self):
  r=self.c.post('/api/preview',files={'quote':('supplier.xlsx',SRC1)},data={'margin':'15','fx':'18300'});self.assertIn(r.status_code,(401,503))
  r=self.c.post('/api/generate',files={'quote':('supplier.xlsx',SRC1)},data={'margin':'15','fx':'18300'});self.assertIn(r.status_code,(401,503))
 def test_master_status_requires_auth(self):
  r=self.c.get('/api/master-status')
  self.assertIn(r.status_code,(401,503))
 def test_prepriced_commverge_no_double_margin(self):
  raw=synthetic_commverge_priced()
  p15=preview(raw,15,18300);p20=preview(raw,20,18300)
  self.assertEqual(p15['source_type'],'commverge_prepriced')
  self.assertIsNone(p15['actual_gm'])
  self.assertEqual(p15['total_usd'],240.0)
  self.assertEqual(p15['total_usd'],p20['total_usd'])
  self.assertEqual(p15['recomputed_subtotals'],1)
  name,rendered,mime=generate(raw,'T01',15,18300)
  self.assertTrue(name.endswith('.xlsx'))
  zipfile.ZipFile(io.BytesIO(rendered)).testzip()
  sheet,strings=sheet_and_shared(zipfile.ZipFile(io.BytesIO(rendered)))
  self.assertEqual(float(value(sheet,strings,'H45')),240.0)
  pdf_name,pdf,mime=generate_pdf(raw,'T01',15,18300)
  self.assertEqual(mime,'application/pdf')
  self.assertTrue(pdf.startswith(b'%PDF-'))
  # Reverse conversion is valid: original selling prices must be retained.
  t02_name,t02_bytes,t02_mime=generate(raw,'T02',15,18300)
  self.assertTrue(t02_name.endswith('.xlsx'))
  self.assertTrue(zipfile.is_zipfile(io.BytesIO(t02_bytes)))
  t02sheet,t02strings=sheet_and_shared(zipfile.ZipFile(io.BytesIO(t02_bytes)))
  self.assertEqual(value(t02sheet,t02strings,'D18'),'SYNTHETIC-SKU')
  self.assertEqual(float(value(t02sheet,t02strings,'G18')),120.0)
  self.assertEqual(float(value(t02sheet,t02strings,'M89')),240.0)
  self.assertEqual(value(t02sheet,t02strings,'F18'),'')
  self.assertIn('NOT SUPPLIER-ISSUED',value(t02sheet,t02strings,'E12'))
  othername,otherbytes,othermime=generate(raw,'BOTH',15,18300)
  with zipfile.ZipFile(io.BytesIO(otherbytes)) as both:
   self.assertEqual(len(both.namelist()),2)
 def test_prepriced_authenticated_api_bypasses_storage(self):
  raw=synthetic_commverge_priced()
  fake={'SUPABASE_URL':'https://test.supabase.co','SUPABASE_PUBLISHABLE_KEY':'sb_publishable_test','ALLOWED_EMAILS':'test@example.com'}
  app.dependency_overrides[permitted]=lambda:{'id':'test','email':'test@example.com','access_token':'test'}
  try:
   with patch.dict('os.environ',fake):
    files={'quote':('synthetic_prepriced.xlsx',raw)}
    p=self.c.post('/api/preview',files=files,data={'margin':'15','fx':'18300'})
    self.assertEqual(p.status_code,200,p.text[:200])
    self.assertEqual(p.json()['source_type'],'commverge_prepriced')
    # Format T02 is an independent, deliberately selected output.
    alternate=self.c.post('/api/preview',files=files,data={'format':'T02','margin':'15','fx':'18300'})
    self.assertEqual(alternate.status_code,200,alternate.text[:250])
    self.assertFalse(alternate.json()['format_auto_selected'])
    self.assertEqual(alternate.json()['format_resolved'],'T02')
    self.assertEqual(alternate.json()['comparison']['format_viewed'],'T02')
    self.assertEqual(alternate.json()['comparison']['summary']['unexpected_differences'],0)
    p2=self.c.post('/api/pdf-preview',files=files,data={'format':'T01','margin':'15','fx':'18300'})
    self.assertEqual(p2.status_code,200,p2.text[:200])
    self.assertTrue(p2.content.startswith(b'%PDF-'))
    p3=self.c.post('/api/generate',files=files,data={'format':'T01','margin':'15','fx':'18300','file_type':'xlsx'})
    self.assertEqual(p3.status_code,200,p3.text[:200])
    self.assertTrue(zipfile.is_zipfile(io.BytesIO(p3.content)))
  finally:app.dependency_overrides.clear()
 def test_smartoptics_source_to_commverge_t01_synthetic_master(self):
  # Synthetic brand-shape fixture: original supplier costs receive margin once.
  master=synthetic_commverge_priced()
  name,data,mime=generate(SRC1,'T01',15,18300,'{}',master)
  self.assertTrue(name.endswith('.xlsx'))
  c=compare_workbooks(SRC1,'T01',15,18300,'{}',master)
  self.assertEqual(c['summary']['sku_matches'],3)
  self.assertEqual(c['summary']['qty_matches'],3)
  self.assertEqual(c['summary']['unexpected_differences'],0)
  self.assertEqual(c['summary']['margin_adjustments'],3)
 def test_commverge_source_to_smartoptics_t02_comparison(self):
  raw=synthetic_commverge_priced()
  c=compare_workbooks(raw,'T02',15,18300)
  self.assertEqual(c['source_type'],'commverge_prepriced')
  self.assertEqual(c['format_viewed'],'T02')
  self.assertEqual(c['summary']['sku_matches'],1)
  self.assertEqual(c['summary']['qty_matches'],1)
  self.assertEqual(c['summary']['subtotals_repaired'],1)
  self.assertEqual(c['summary']['unexpected_differences'],0)
  self.assertEqual(float(c['draft_total']),240.0)
  c20=compare_workbooks(raw,'T02',20,18300)
  self.assertEqual(c['draft_total'],c20['draft_total'])
 def test_compare_original_smartoptics_vs_draft_t02(self):
  c=compare_workbooks(SRC1,'T02',15,18300)
  summary=c['summary']
  self.assertEqual(c['source_type'],'smartoptics_original')
  self.assertEqual(summary['rows'],3)
  self.assertEqual(summary['sku_matches'],3)
  self.assertEqual(summary['qty_matches'],3)
  self.assertEqual(summary['description_matches'],3)
  self.assertEqual(summary['price_validated'],3)
  self.assertEqual(summary['total_validated'],3)
  self.assertGreater(summary['margin_adjustments'],0)
  self.assertEqual(summary['unexpected_differences'],0)
  self.assertEqual(c['source_columns']['sku'],'D')
  self.assertEqual(c['draft_columns']['sku'],'D')
 def test_compare_existing_commverge_priced_vs_draft_t01(self):
  raw=synthetic_commverge_priced()
  c=compare_workbooks(raw,'T01',15,18300,master=raw)
  self.assertEqual(c['source_type'],'commverge_prepriced')
  self.assertEqual(c['summary']['rows'],1)
  self.assertEqual(c['summary']['subtotals_repaired'],1)
  self.assertEqual(c['summary']['unexpected_differences'],0)
  self.assertEqual(c['source_rows'][0]['total'],'#VALUE!')
  self.assertEqual(float(c['draft_rows'][0]['total']),240.0)
  c20=compare_workbooks(raw,'T01',20,18300,master=raw)
  self.assertEqual(c['draft_total'],c20['draft_total'])
 def test_preview_returns_only_generated_document(self):
  c=compare_workbooks(SRC1,'T02',15,18300)
  self.assertNotIn('source_document',c)
  self.assertIn('draft_document',c)
  self.assertEqual(c['draft_label'],'Smartoptics')
  self.assertEqual(c['draft_document']['template'],'Smartoptics')
  self.assertEqual(c['draft_document']['item_count'],3)
  self.assertEqual(len(c['draft_document']['items']),3)
  self.assertEqual(float(c['draft_document']['total']),3631.80)
  self.assertEqual(c['draft_document']['meta']['company'],'Smartoptics')
 def test_commverge_preview_displays_only_result_and_repaired_total(self):
  raw=synthetic_commverge_priced()
  c=compare_workbooks(raw,'T01',15,18300,master=raw)
  self.assertNotIn('source_document',c)
  self.assertEqual(c['draft_label'],'CommVerge')
  self.assertNotIn('T01',c['draft_label'])
  self.assertEqual(c['draft_document']['template'],'CommVerge')
  self.assertEqual(float(c['draft_document']['total']),240)
  self.assertEqual(c['source_rows'][0]['total'],'#VALUE!')
  self.assertEqual(float(c['draft_document']['items'][0]['total']),240)
 def test_preview_includes_all_generated_items_for_scrolling(self):
  from preview_compare import _document_snapshot
  raw=synthetic_commverge_priced()
  rows=[{'row':23+i,'sku':'TEST-'+str(i),'description':'Safe test line '+str(i),'qty':'1','unit':'10','total':'10','status':'COCOK'} for i in range(16)]
  doc=_document_snapshot(raw,'T01',rows)
  self.assertEqual(doc['item_count'],16)
  self.assertEqual(len(doc['items']),16)
  self.assertEqual(doc['omitted'],0)
 def test_authenticated_compare_api(self):
  fake={'SUPABASE_URL':'https://test.supabase.co','SUPABASE_PUBLISHABLE_KEY':'sb_publishable_test','ALLOWED_EMAILS':'test@example.com'}
  app.dependency_overrides[permitted]=lambda:{'id':'test','email':'test@example.com','access_token':'test'}
  try:
   with patch.dict('os.environ',fake):
    r=self.c.post('/api/preview',files={'quote':('supplier.xlsx',SRC1)},data={'margin':'15','fx':'18300','format':'T02','details':'{}'})
    self.assertEqual(r.status_code,200,r.text[:200])
    self.assertEqual(r.json()['comparison']['summary']['sku_matches'],3)
    self.assertEqual(r.json()['comparison']['summary']['unexpected_differences'],0)
  finally:app.dependency_overrides.clear()
 def test_pdf_export_t02_a4(self):
  name,data,mime=generate_pdf(SRC1,'T02',15,18300,'{"customer":"TEST"}')
  self.assertTrue(name.endswith('.pdf'))
  self.assertEqual(mime,'application/pdf')
  self.assertTrue(data.startswith(b'%PDF-'))
  self.assertIn(b'/MediaBox',data)
  self.assertGreater(len(data),2000)
 def test_pdf_requires_auth(self):
  r=self.c.post('/api/pdf-preview',files={'quote':('supplier.xlsx',SRC1)},data={'format':'T02'})
  self.assertIn(r.status_code,(401,503))
 def test_authenticated_pdf_endpoint(self):
  fake={'SUPABASE_URL':'https://my-test.supabase.co','SUPABASE_PUBLISHABLE_KEY':'sb_publishable_test','ALLOWED_EMAILS':'test@example.com'}
  app.dependency_overrides[permitted]=lambda:{'id':'example','email':'test@example.com','access_token':'test'}
  try:
   with patch.dict('os.environ',fake):
    r=self.c.post('/api/pdf-preview',files={'quote':('supplier.xlsx',SRC1)},data={'format':'T02','margin':'15','fx':'18300'})
    self.assertEqual(r.status_code,200,r.text[:200])
    self.assertEqual(r.headers['content-type'],'application/pdf')
    self.assertTrue(r.content.startswith(b'%PDF-'))
    r2=self.c.post('/api/generate',files={'quote':('supplier.xlsx',SRC1)},data={'format':'T02','margin':'15','fx':'18300','file_type':'pdf'})
    self.assertEqual(r2.status_code,200,r2.text[:200])
    self.assertTrue(r2.content.startswith(b'%PDF-'))
  finally:app.dependency_overrides.clear()
 def test_demo_preview(self):
  p=self.c.get('/api/demo',params={'margin':'15','fx':'18300'});self.assertEqual(p.status_code,200)
  self.assertEqual(len(p.json()['items']),3)
  p2=self.c.get('/api/demo',params={'margin':'20','fx':'18300'});self.assertGreater(p2.json()['total_usd'],p.json()['total_usd'])
 def test_demo_xlsx(self):
  r=self.c.get('/api/demo-export',params={'margin':'20','fx':'18300','format':'T01','customer':'Demo Test','project':'New Test'})
  self.assertEqual(r.status_code,200)
  z=zipfile.ZipFile(io.BytesIO(r.content));s,ss=sheet_and_shared(z)
  self.assertEqual(value(s,ss,'B4'),'CommVerge-DEMO-SAMPLE');self.assertEqual(value(s,ss,'A18'),'Demo Test');self.assertEqual(value(s,ss,'C18'),'New Test')
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
