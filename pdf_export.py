"""Authenticated A4 PDF quotation renderer using the same validated source as Excel."""
from __future__ import annotations
import datetime as dt
import html
import io
import zipfile
from decimal import Decimal, ROUND_HALF_UP
from reportlab.lib import colors
from reportlab.lib.enums import TA_RIGHT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import BaseDocTemplate,Frame,KeepTogether,PageTemplate,Paragraph,Spacer,Table,TableStyle,Image
from quote_engine import QuoteError,source,metadata,generate

NAVY=colors.HexColor('#0D1E3F')
MUTED=colors.HexColor('#647084')
CREAM=colors.HexColor('#F4F1EC')
BORDER=colors.HexColor('#D8D1C7')
WHITE=colors.white
GREY=colors.HexColor('#F7F7F6')
S={
 'title':ParagraphStyle('qt',fontName='Helvetica-Bold',fontSize=17,leading=20,textColor=NAVY),
 'sub':ParagraphStyle('qs',fontName='Helvetica',fontSize=8,leading=11,textColor=MUTED),
 'metahead':ParagraphStyle('qh',fontName='Helvetica-Bold',fontSize=7,leading=10,textColor=MUTED),
 'metavalue':ParagraphStyle('qv',fontName='Helvetica',fontSize=8.5,leading=11.5,textColor=NAVY),
 'th':ParagraphStyle('qth',fontName='Helvetica-Bold',fontSize=7.5,leading=10,textColor=WHITE),
 'cell':ParagraphStyle('qc',fontName='Helvetica',fontSize=7.5,leading=10,textColor=NAVY),
 'cellsmall':ParagraphStyle('qcs',fontName='Helvetica',fontSize=7.1,leading=9.5,textColor=NAVY),
 'num':ParagraphStyle('qn',fontName='Helvetica',fontSize=7.5,leading=10,textColor=NAVY,alignment=TA_RIGHT),
 'small':ParagraphStyle('qsm',fontName='Helvetica',fontSize=7.5,leading=10.5,textColor=MUTED),
 'total':ParagraphStyle('qtotal',fontName='Helvetica-Bold',fontSize=10,leading=13,textColor=NAVY,alignment=TA_RIGHT),
}

def esc(value):
 return html.escape(str(value if value is not None else '')).replace('\u2013','-').replace('\u2014','-').replace('\u2011','-').replace('\u00a0',' ')
def money(value):return f'{Decimal(str(value)):,.2f}'
def idr(value):return f'{int(Decimal(str(value)).quantize(Decimal("1"),rounding=ROUND_HALF_UP)):,}'
def para(value,kind='cell'):return Paragraph(esc(value),S[kind])

def pdf_for_one(q,code,meta,date,master_bytes):
 buf=io.BytesIO()
 page_width,page_height=A4
 ml=37;mr=37;top=48;bottom=51
 def decorate(canv,doc):
  canv.saveState()
  canv.setFillColor(NAVY)
  canv.rect(0,page_height-7,page_width,7,fill=1,stroke=0)
  canv.setStrokeColor(BORDER);canv.setLineWidth(.5)
  canv.line(ml,34,page_width-mr,34)
  canv.setFillColor(MUTED);canv.setFont('Helvetica',7)
  canv.drawString(ml,23,'DRAFT - Commercial review required before customer submission')
  canv.drawRightString(page_width-mr,23,f'Page {doc.page} | A4')
  canv.restoreState()
 frame=Frame(ml,bottom,page_width-ml-mr,page_height-top-bottom,
             leftPadding=0,rightPadding=0,topPadding=0,bottomPadding=0,id='body')
 doc=BaseDocTemplate(buf,pagesize=A4,pageTemplates=[PageTemplate(id='qt',frames=[frame],onPage=decorate)],
                     title=f'Quote {"CommVerge" if code=="T01" else "Smartoptics"} DRAFT',author='Quote Generator')
 story=[]
 if code=='T01' and master_bytes:
  try:
   with zipfile.ZipFile(io.BytesIO(master_bytes)) as z:
    candidates=['xl/media/image.jpg','xl/media/image1.jpeg','xl/media/image.jpeg']
    name=next((p for p in candidates if p in z.namelist()),None)
    if not name:raise KeyError('CommVerge logo tidak ada')
    logo=z.read(name)
   story.append(Image(io.BytesIO(logo),width=154,height=46,kind='proportional'))
   story.append(Spacer(1,7))
  except (KeyError,zipfile.BadZipFile):story.append(para('COMMVERGE SOLUTIONS','metahead'))
 else:story.append(para('SMARTOPTICS SOURCE FORMAT - REPRICED COPY','metahead'))
 story.append(Paragraph('QUOTATION <font color="#AA4B38">| DRAFT</font>',S['title']))
 story.append(Paragraph(
  'Already-priced CommVerge quotation. Existing unit selling prices retained; subtotals recomputed for review.'
  if q.get('source_type')=='commverge_prepriced'
  else 'Review copy only. Pricing is based on source vendor costs and selected gross margin.',
  S['sub']))
 story.append(Spacer(1,12))
 customer=meta.get('customer','CUSTOMER - TBC')
 project=meta.get('project') or q['quote_name'] or 'Quotation draft'
 quote_ref=meta.get('quote_ref') or q['quote_ref'] or 'DRAFT'
 hdr=Table([
  [para('CUSTOMER','metahead'),para('PROJECT','metahead'),para('DATE','metahead')],
  [para(customer,'metavalue'),para(project,'metavalue'),para(date.isoformat(),'metavalue')],
  [para('ATTENTION','metahead'),para('QUOTE REFERENCE','metahead'),para('FX USD/IDR','metahead')],
  [para(meta.get('attention') or 'TBC','metavalue'),para(str(quote_ref)+' - DRAFT','metavalue'),para(idr(q['fx']),'metavalue')]
 ],colWidths=[176,211,134],hAlign='LEFT')
 hdr.setStyle(TableStyle([
  ('BACKGROUND',(0,0),(-1,-1),CREAM),('VALIGN',(0,0),(-1,-1),'TOP'),
  ('LEFTPADDING',(0,0),(-1,-1),10),('RIGHTPADDING',(0,0),(-1,-1),9),
  ('TOPPADDING',(0,0),(-1,-1),6),('BOTTOMPADDING',(0,0),(-1,-1),5),
  ('LINEBELOW',(0,1),(-1,1),.4,BORDER),('BOX',(0,0),(-1,-1),.5,BORDER),
 ]))
 story.append(hdr);story.append(Spacer(1,16))
 story.append(Paragraph('ITEMS & PRICING',S['metahead']));story.append(Spacer(1,5))
 lines=[[para('#','th'),para('SKU','th'),para('DESCRIPTION','th'),para('QTY','th'),para('UNIT USD','th'),para('TOTAL USD','th')]]
 for idx,item in enumerate(q['items'],1):
  lines.append([
   para(idx,'cellsmall'),para(item['sku'],'cellsmall'),para(item['description'],'cellsmall'),
   para(item['qty'],'num'),para(money(item['sell']),'num'),para(money(item['total']),'num'),
  ])
 tbl=Table(lines,colWidths=[23,103,224,37,66,68],repeatRows=1,hAlign='LEFT',splitByRow=1)
 tbl.setStyle(TableStyle([
  ('BACKGROUND',(0,0),(-1,0),NAVY),('VALIGN',(0,0),(-1,-1),'TOP'),
  ('LEFTPADDING',(0,0),(-1,-1),5),('RIGHTPADDING',(0,0),(-1,-1),5),
  ('TOPPADDING',(0,0),(-1,-1),6),('BOTTOMPADDING',(0,0),(-1,-1),6),
  ('LINEBELOW',(0,1),(-1,-1),.32,BORDER),
  ('ROWBACKGROUNDS',(0,1),(-1,-1),[WHITE,GREY]),
 ]))
 story.append(tbl);story.append(Spacer(1,15))
 total_usd=q['revenue']
 total_idr=(total_usd*q['fx']).quantize(Decimal('1'),rounding=ROUND_HALF_UP)
 sums=Table([
  ([para('Pricing basis','small'),para('Existing selling prices','num')]
   if q.get('source_type')=='commverge_prepriced'
   else [para('Gross margin target','small'),para(str(q['margin'])+'%','num')]),
  [para('TOTAL USD (excl. tax)','metahead'),Paragraph('USD '+money(total_usd),S['total'])],
  [para('Reference total IDR','small'),Paragraph('Rp '+idr(total_idr),S['total'])],
 ],colWidths=[351,170],hAlign='LEFT')
 sums.setStyle(TableStyle([
  ('BACKGROUND',(0,0),(-1,-1),CREAM),('TOPPADDING',(0,0),(-1,-1),7),
  ('BOTTOMPADDING',(0,0),(-1,-1),7),('LEFTPADDING',(0,0),(-1,-1),10),
  ('RIGHTPADDING',(0,0),(-1,-1),10),('LINEBELOW',(0,0),(-1,-2),.4,BORDER),
  ('BOX',(0,0),(-1,-1),.5,BORDER)
 ]))
 story.append(KeepTogether([sums,Spacer(1,12)]))
 story.append(Paragraph('COMMERCIAL TERMS (DRAFT)',S['metahead']))
 story.append(Spacer(1,5))
 terms=Table([[para(k,'metahead'),para(v,'cell')] for k,v in [
  ('Delivery',meta.get('shipping') or 'To be confirmed'),
  ('Payment',meta.get('payment') or 'To be confirmed'),
  ('Lead time',meta.get('lead_time') or 'To be confirmed'),
  ('Validity',meta.get('validity') or 'To be confirmed')
 ]],colWidths=[87,434],hAlign='LEFT')
 terms.setStyle(TableStyle([
  ('VALIGN',(0,0),(-1,-1),'TOP'),('BOTTOMPADDING',(0,0),(-1,-1),5),
  ('TOPPADDING',(0,0),(-1,-1),4),('LINEBELOW',(0,0),(-1,-1),.3,BORDER)
 ]))
 story.append(terms);story.append(Spacer(1,7))
 note=meta.get('remarks') or 'Tax, freight, delivery and commercial terms must be confirmed before issue.'
 story.append(Paragraph('<b>Remarks:</b> '+esc(note),S['small']))
 doc.build(story)
 return buf.getvalue()

def generate_pdf(raw,format_id='T01',margin=15,fx=18300,details='',master_bytes=None):
 """Return PDF for a single template or ZIP containing 2 PDFs for BOTH."""
 if format_id not in ('T01','T02','BOTH'):raise QuoteError('Format tidak valid.')
 # Run the same QA checks as the Excel download.
 generate(raw,format_id,margin,fx,details,master_bytes)
 q=source(raw,margin,fx);meta=metadata(details)
 if q.get('source_type')=='commverge_prepriced':
  if master_bytes is None:master_bytes=raw
  meta={**{key:q.get(key) for key in ('customer','attention','delivery','billing','email') if q.get(key)},**meta}
 date=dt.date.fromisoformat(meta.get('date') or dt.datetime.now(dt.timezone(dt.timedelta(hours=7))).date().isoformat())
 codes=['T01','T02'] if format_id=='BOTH' else [format_id]
 exports=[]
 for code in codes:
  pdf=pdf_for_one(q,code,meta,date,master_bytes if code=='T01' else None)
  if not pdf.startswith(b'%PDF-'):raise QuoteError('PDF output tidak valid.')
  exports.append((f'DANIEL_QUOTE_{"CommVerge" if code=="T01" else "Smartoptics"}_{date.isoformat()}_DRAFT.pdf',pdf))
 if len(exports)==1:return exports[0][0],exports[0][1],'application/pdf'
 out=io.BytesIO()
 with zipfile.ZipFile(out,'w',zipfile.ZIP_DEFLATED) as z:
  for name,pdf in exports:z.writestr(name,pdf)
 return f'DANIEL_QUOTE_Keduanya_{date.isoformat()}_PDF_DRAFT.zip',out.getvalue(),'application/zip'
