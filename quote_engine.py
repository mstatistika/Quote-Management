"""Server-side, portable XLSX (OOXML ZIP) quotation adapter.
No artifact_tool runtime dependency; original styling and logo are copied untouched.
Supported vendor source: Smartoptics Detailed Quote with verified D16/G14/G89 anchors.
"""
from __future__ import annotations
import datetime as dt
import hashlib
import io
import json
import re
import zipfile
from decimal import Decimal, ROUND_CEILING, ROUND_HALF_UP, InvalidOperation
from pathlib import Path
from xml.etree import ElementTree as ET

M='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
R='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
Mtag=lambda name: f'{{{M}}}{name}'
ET.register_namespace('',M)
ET.register_namespace('r',R)
ET.register_namespace('mc','http://schemas.openxmlformats.org/markup-compatibility/2006')
ET.register_namespace('x14ac','http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac')
CENT=Decimal('.01')
SLOTS=[23,24,25,26,28,29,30,31,32,34,35,37,38,39,40,41]
ERR_VALUES={'#REF!','#DIV/0!','#VALUE!','#NAME?','#N/A','#NUM!','#NULL!'}
ALLOWED_META={'customer','attention','email','project','quote_ref','date','delivery','billing','shipping','payment','lead_time','validity','remarks'}
SOURCE_UNCOMPRESSED_LIMIT=48*1024*1024
SOURCE_PARTS_LIMIT=512
MAX_UPLOAD=8*1024*1024

class QuoteError(ValueError): pass

def decimal(value):
    try:
        d=Decimal(str(value))
        if not d.is_finite():raise ValueError()
        return d
    except (InvalidOperation,ValueError):raise QuoteError('Angka quotation tidak valid.')

def amount(n):return n.quantize(CENT,rounding=ROUND_CEILING)
def round2(n):return n.quantize(CENT,rounding=ROUND_HALF_UP)
def fmt(n):return format(round2(n),'f')

def workbook(raw):
    if len(raw)>MAX_UPLOAD or len(raw)<200:raise QuoteError('XLSX tidak boleh melebihi 8 MB.')
    try:z=zipfile.ZipFile(io.BytesIO(raw))
    except zipfile.BadZipFile:raise QuoteError('File bukan XLSX yang valid.')
    inf=z.infolist()
    if len(inf)>SOURCE_PARTS_LIMIT or sum(x.file_size for x in inf)>SOURCE_UNCOMPRESSED_LIMIT:raise QuoteError('Workbook melebihi batas keamanan.')
    names=set(z.namelist())
    if not {'xl/workbook.xml','xl/worksheets/sheet1.xml','[Content_Types].xml'}.issubset(names):raise QuoteError('Workbook tidak lengkap.')
    if any('vbaproject' in n.lower() or 'externallink' in n.lower() for n in names):raise QuoteError('Macro dan external links tidak didukung.')
    if any(n.endswith('..') or n.startswith('/') or '../' in n for n in names):raise QuoteError('Nama part XLSX tidak valid.')
    return z

def sheet_and_shared(z):
    sheet=ET.fromstring(z.read('xl/worksheets/sheet1.xml'))
    if 'xl/sharedStrings.xml' in z.namelist():
        st=ET.fromstring(z.read('xl/sharedStrings.xml'))
        shared=[''.join(t.text or '' for t in si.iter(Mtag('t'))) for si in st.findall(Mtag('si'))]
    else:shared=[]
    return sheet,shared

def cells(sheet):return {c.get('r'):c for c in sheet.iter(Mtag('c'))}
def cellval(c,shared):
    if c is None:return ''
    t=c.get('t')
    if t=='inlineStr':return ''.join(x.text or '' for x in c.iter(Mtag('t')))
    v=c.find(Mtag('v'))
    if v is None:return ''
    if t=='s':
        idx=int(v.text or '0')
        if idx>=len(shared):raise QuoteError('Shared strings XLSX rusak.')
        return shared[idx]
    return v.text or ''

def value(sheet,shared,address):return cellval(cells(sheet).get(address),shared)

def cell(sheet,addr,create=False):
    cs=cells(sheet)
    if addr in cs:return cs[addr]
    if not create:return None
    rn=int(re.search(r'\d+',addr).group())
    data=sheet.find(Mtag('sheetData'))
    if data is None:raise QuoteError('SheetData tidak ditemukan.')
    row=next((r for r in data.findall(Mtag('row')) if int(r.get('r'))==rn),None)
    if row is None:
        row=ET.Element(Mtag('row'),{'r':str(rn)})
        i=next((i for i,r in enumerate(data) if r.tag==Mtag('row') and int(r.get('r'))>rn),len(data))
        data.insert(i,row)
    c=ET.Element(Mtag('c'),{'r':addr})
    alpha=lambda a: sum((ord(x)-64)*26**i for i,x in enumerate(re.search('[A-Z]+',a).group()[::-1]))
    i=next((i for i,existing in enumerate(row) if existing.tag==Mtag('c') and alpha(existing.get('r'))>alpha(addr)),len(row))
    row.insert(i,c)
    return c

def setval(sheet,addr,val,kind='text'):
    c=cell(sheet,addr,True)
    for ch in list(c):c.remove(ch)
    if kind=='text':
        c.set('t','inlineStr'); el=ET.SubElement(c,Mtag('is')); t=ET.SubElement(el,Mtag('t')); t.text=str(val)
    elif kind=='number':
        c.attrib.pop('t',None);v=ET.SubElement(c,Mtag('v'));v.text=str(val)
    elif kind=='formula':
        c.attrib.pop('t',None);f=ET.SubElement(c,Mtag('f'));f.text=val[0].lstrip('=');v=ET.SubElement(c,Mtag('v'));v.text=str(val[1])
    else:raise AssertionError(kind)

def clear(sheet,addr):
    c=cell(sheet,addr)
    if c is not None:
        for ch in list(c):c.remove(ch)
        c.attrib.pop('t',None)

def metadata(raw):
    if not raw:return {}
    if len(raw)>5000:raise QuoteError('Metadata terlalu panjang.')
    try:data=json.loads(raw)
    except json.JSONDecodeError:raise QuoteError('Metadata JSON tidak valid.')
    if not isinstance(data,dict) or any(k not in ALLOWED_META for k in data):raise QuoteError('Hanya metadata yang boleh diubah.')
    for k,v in data.items():
        if not isinstance(v,str) or len(v)>200 or '\n' in v or '\r' in v or (v and v[0] in '=+@'):
            raise QuoteError('Metadata berisi nilai tidak aman.')
    if data.get('date'):
        try:dt.date.fromisoformat(data['date'])
        except ValueError:raise QuoteError('Tanggal quotation tidak valid.')
    return data

def settings(margin,fx):
    m=decimal(margin);f=decimal(fx)
    if not Decimal('0')<m<=Decimal('80'):raise QuoteError('Target gross margin harus 0–80%.')
    if not Decimal(1000)<=f<=Decimal(1000000) or f!=f.to_integral_value():raise QuoteError('Kurs harus bilangan bulat Rp1.000–Rp1.000.000.')
    return m,f

def prepriced_commverge(z,s,ss,fx):
    """Already customer-priced CommVerge T01; never apply a second margin."""
    if (value(s,ss,'C21').strip().lower()!='product name'
        or 'unit price' not in value(s,ss,'G21').lower()
        or 'commverge' not in (value(s,ss,'D6')+' '+value(s,ss,'J3')).lower()
        or 'Quote to DE' not in z.read('xl/workbook.xml').decode('utf-8')
        or sum(bool(re.fullmatch(r'xl/worksheets/sheet\d+\.xml',n)) for n in z.namelist())!=1):
        raise QuoteError('Gunakan quotation Smartoptics atau CommVerge yang valid.')
    items=[];adjustments=0
    for r in SLOTS:
        sku=value(s,ss,f'C{r}').strip()
        if not sku:continue
        qty=decimal(value(s,ss,f'F{r}'))
        sell=decimal(value(s,ss,f'G{r}'))
        if qty<=0 or qty!=qty.to_integral_value() or sell<=0:
            raise QuoteError(f'Quotation CommVerge baris {r}: qty/harga jual tidak valid.')
        if sell!=round2(sell):raise QuoteError(f'Baris {r}: harga jual lebih dari 2 desimal.')
        rev=round2(sell*qty)
        cached=value(s,ss,f'H{r}')
        try:err=abs(decimal(cached)-rev)>Decimal('.02')
        except QuoteError:err=True
        if err:adjustments+=1
        items.append({'row':r,'sku':sku,'description':value(s,ss,f'D{r}'),
          'qty':int(qty),'cost':None,'listcost':None,'sell':sell,'total':rev,
          'site':{},'source_type':'commverge_prepriced'})
    if not items:raise QuoteError('Tidak ada item aktif pada quotation CommVerge.')
    if len({item['sku'] for item in items})!=len(items):raise QuoteError('SKU duplikat pada quotation CommVerge.')
    revenue=sum((i['total'] for i in items),Decimal(0))
    return {'source_type':'commverge_prepriced','items':items,'margin':None,'fx':fx,
      'cost':None,'revenue':round2(revenue),'gm_actual':None,'recomputed_subtotals':adjustments,
      'customer':value(s,ss,'C10'),'attention':value(s,ss,'C11'),
      'delivery':value(s,ss,'D12'),'billing':value(s,ss,'D14'),'email':value(s,ss,'D18'),
      'quote_name':value(s,ss,'F12'),'quote_ref':value(s,ss,'G11') or value(s,ss,'P10'),
      'original_zip':z}

def source(raw,margin=15,fx=18300):
    m,rate=settings(margin,fx)
    z=workbook(raw);s,ss=sheet_and_shared(z)
    if value(s,ss,'D16').strip().lower()!='part #' or value(s,ss,'G14').strip()!='Quote Number:' or value(s,ss,'G89').strip()!='Total Price:' or 'Smartoptics' not in value(s,ss,'L5'):
        return prepriced_commverge(z,s,ss,rate)
    if 'USD/IDR' in value(s,ss,'D90') or 'repric' in value(s,ss,'E12').lower():
        raise QuoteError('File sudah pernah dinaikkan margin. Upload vendor quote asli.')
    cs=cells(s);items=[];cost_total=Decimal('0')
    for r in range(18,89):
        sku=str(cellval(cs.get(f'D{r}'),ss)).strip()
        if not sku:continue
        qtxt=cellval(cs.get(f'L{r}'),ss)
        if qtxt in ('',None):continue
        qty=decimal(qtxt)
        if qty<=0:continue
        if qty!=qty.to_integral_value():raise QuoteError(f'QTY baris {r} harus bulat.')
        cost=decimal(cellval(cs.get(f'G{r}'),ss))
        total=decimal(cellval(cs.get(f'M{r}'),ss))
        if cost<=0 or abs(cost*qty-total)>Decimal('.02'):raise QuoteError(f'Baris {r}: harga supplier × QTY tidak cocok dengan total.')
        sell=amount(cost/(Decimal(1)-m/100));rev=round2(sell*qty)
        if (rev-cost*qty)/rev<m/100:raise QuoteError(f'Baris {r}: margin jual di bawah target.')
        listval=cellval(cs.get(f'F{r}'),ss)
        listcost=decimal(listval) if listval else cost
        desc=str(cellval(cs.get(f'E{r}'),ss))
        site={col:cellval(cs.get(f'{col}{r}'),ss) for col in 'HIJK'}
        items.append({'row':r,'sku':sku,'description':desc,'qty':int(qty),'cost':cost,'listcost':listcost,'sell':sell,'total':rev,'site':site})
        cost_total+=cost*qty
    if not items:raise QuoteError('Tidak ada item aktif pada sumber.')
    if len({i['sku'] for i in items})<len(items):raise QuoteError('SKU duplikat; butuh review supplier.')
    vtotal=decimal(value(s,ss,'M89'))
    if abs(cost_total-vtotal)>Decimal('.02'):raise QuoteError('Grand total supplier tidak cocok dengan item aktif.')
    revenue=sum((i['total'] for i in items),Decimal(0))
    return {'source_type':'smartoptics_original','recomputed_subtotals':0,'items':items,'margin':m,'fx':rate,'cost':round2(cost_total),'revenue':round2(revenue),
        'gm_actual':round2((revenue-cost_total)/revenue*100),
        'quote_name':value(s,ss,'E11'),'quote_ref':value(s,ss,'L14'), 'original_zip':z}

def preview(raw,margin=15,fx=18300):
    q=source(raw,margin,fx)
    prepriced=q['source_type']=='commverge_prepriced'
    return {'status':'PASS_REVIEW_REQUIRED','source_type':q['source_type'],
      'price_note':('Harga jual sudah ada dalam quotation CommVerge. Margin tidak dihitung ulang; upload Smartoptics asli untuk mengubah margin.' if prepriced else 'Margin dihitung dari harga supplier asli.'),
      'recomputed_subtotals':q['recomputed_subtotals'],
      'target_gm':None if prepriced else float(q['margin']),
      'actual_gm':None if prepriced else float(q['gm_actual']),
      'rate':int(q['fx']),'item_count':len(q['items']),'total_usd':float(q['revenue']),
      'total_idr':int((q['revenue']*q['fx']).quantize(Decimal('1'),rounding=ROUND_HALF_UP)),
      'items':[{'sku':i['sku'],'description':i['description'],'qty':i['qty'],'unit_usd':float(i['sell']),
      'total_usd':float(i['total']),'margin_pct':None if prepriced else round(float((i['total']-i['cost']*i['qty'])/i['total']*100),3)} for i in q['items']]}

def pack(z,sheet):
    dest=io.BytesIO()
    with zipfile.ZipFile(dest,'w') as o:
        for entry in z.infolist():
            if entry.filename=='xl/worksheets/sheet1.xml':o.writestr(entry,ET.tostring(sheet,encoding='utf-8',xml_declaration=True))
            else:o.writestr(entry,z.read(entry.filename))
    return dest.getvalue()

def date_serial(date):return (date-dt.date(1899,12,30)).days

def output_t02(q,date,meta):
    from_commverge=q['source_type']=='commverge_prepriced'
    if from_commverge:
        # A CommVerge workbook has no Smartoptics layout. Use the sanitized,
        # PUBLIC synthetic T02 worksheet layout, never an unrelated customer file.
        # It carries no real customer/pricing data and is clearly marked DRAFT.
        template=Path(__file__).resolve().parent/'tests/fixtures/SMARTOPTICS_SYNTHETIC_SOURCE.xlsx'
        if not template.is_file():raise QuoteError('Layout Smartoptics belum terpasang.')
        z=workbook(template.read_bytes())
    else:
        z=q['original_zip']
    s,ss=sheet_and_shared(z)
    if value(s,ss,'D16').strip().lower()!='part #' or 'Smartoptics' not in value(s,ss,'L5'):
        raise QuoteError('Template output Smartoptics tidak valid.')
    if from_commverge:
        if len(q['items'])>71:raise QuoteError('Template Smartoptics hanya mendukung 71 baris item.')
        # Remove every synthetic/sample item and price before copying the source.
        for r in range(18,89):
            for col in 'DEFGHIJKLM':clear(s,f'{col}{r}')
        for addr in ('E8','E9','E10','E11','L14','E93','E94','E95','E97'):
            clear(s,addr)
    active=set()
    for index,i in enumerate(q['items']):
        r=18+index if from_commverge else i['row']
        active.add(r)
        if from_commverge:
            setval(s,f'D{r}',i['sku'])
            setval(s,f'E{r}',i['description'])
            setval(s,f'L{r}',i['qty'],'number')
            # Original vendor list price/cost cannot be recovered from a selling-price quote.
            clear(s,f'F{r}')
        else:
            setval(s,f'F{r}',amount(i['listcost']/(1-q['margin']/100)),'number')
        setval(s,f'G{r}',i['sell'],'number')
        setval(s,f'M{r}',(f'ROUND(G{r}*L{r},2)',i['total']),'formula')
    for r in range(18,89):
        if r not in active:
            for col in 'FGM':clear(s,f'{col}{r}')
    setval(s,'M89',('ROUND(SUM(M18:M88),2)',q['revenue']),'formula')
    setval(s,'D90','USD/IDR Rate:');setval(s,'E90',q['fx'],'number')
    setval(s,'G90','Reference Total (IDR):')
    ref_idr=int((q['revenue']*q['fx']).quantize(Decimal('1'),rounding=ROUND_HALF_UP))
    setval(s,'M90',('ROUND(M89*$E$90,0)',ref_idr),'formula')
    setval(s,'E92',date_serial(date),'number')
    if from_commverge:
        setval(s,'E12','DRAFT – FORMATTED FROM COMMVERGE SELLING PRICES; NO ORIGINAL SMARTOPTICS COST DATA; NOT SUPPLIER-ISSUED')
        setval(s,'E8',meta.get('attention') or 'TBC')
        setval(s,'E9',meta.get('customer') or 'CUSTOMER NAME - TBC')
        setval(s,'E10',meta.get('email') or 'TBC')
        setval(s,'E11',meta.get('project') or q.get('quote_name') or 'Quotation draft')
        setval(s,'L14',meta.get('quote_ref') or q.get('quote_ref') or 'DRAFT')
    else:
        setval(s,'E12','DRAFT – REPRICED COPY FOR REVIEW; NOT SUPPLIER-ISSUED')
    mapping={'customer':'E9','attention':'E8','email':'E10','project':'E11','quote_ref':'L14','shipping':'E93','payment':'E94','validity':'E95','remarks':'E97'}
    for field,addr in mapping.items():
        if meta.get(field):setval(s,addr,meta[field])
    return pack(z,s)

def output_t01(q,date,meta,master_bytes):
    if len(q['items'])>len(SLOTS):raise QuoteError('Template CommVerge maksimal 16 item; template perlu diperluas manual.')
    z=workbook(master_bytes);s,ss=sheet_and_shared(z)
    if value(s,ss,'C21')!='Product Name' or 'Quote to DE' not in z.read('xl/workbook.xml').decode('utf-8'):
        raise QuoteError('Master CommVerge tidak valid.')
    if len([n for n in z.namelist() if re.fullmatch(r'xl/worksheets/sheet\d+\.xml',n)])!=1:raise QuoteError('Master CommVerge memiliki sheet tambahan yang belum diamankan.')
    # Clear old input footprints while leaving sheet layout and three original logo assets intact.
    for r in list(range(22,42))+list(range(44,48))+list(range(49,62)):
        for col in 'BCDEFGHIJL':clear(s,f'{col}{r}')
    for r in range(22,82):
        for n in range(11,35):
            letters='';x=n
            while x: letters=chr(65+(x-1)%26)+letters;x=(x-1)//26
            clear(s,f'{letters}{r}')
    groups={22:'Optical Platform & Licensing',27:'Optical Transponders & Modules',33:'Passive Optical Components',36:'Protection & Services'}
    for index,(r,label) in enumerate(groups.items(),1):setval(s,f'B{r}',index,'number');setval(s,f'C{r}',label)
    for i,r in zip(q['items'],SLOTS):
        group=1 if r<=26 else 2 if r<=32 else 3 if r<=35 else 4
        start={1:23,2:28,3:34,4:37}[group]
        nonphysical=bool(re.search(r'^(Smart Care|PRE-STAGE)|LICEN[SC]E|_SW$',i['sku'],re.I) or re.search(r'^(service|support|software)\b',i['description'],re.I))
        for col,val in {'B':f'{group}.{r-start+1}','C':i['sku'],'D':i['description'],'E':'Non-physical' if nonphysical else 'Physical','F':i['qty'],'G':i['sell'],'L':'S' if nonphysical else 'H'}.items():setval(s,f'{col}{r}',val,'number' if col in ('F','G') else 'text')
        setval(s,f'H{r}',(f'ROUND(F{r}*G{r},2)',i['total']),'formula')
        setval(s,f'I{r}',(f'ROUND(G{r}*$P$19,0)',int((i['sell']*q['fx']).quantize(Decimal('1'),rounding=ROUND_HALF_UP))),'formula')
        setval(s,f'J{r}',(f'ROUND(H{r}*$P$19,0)',int((i['total']*q['fx']).quantize(Decimal('1'),rounding=ROUND_HALF_UP))),'formula')
    name=meta.get('customer') or 'CUSTOMER NAME — TBC'
    setval(s,'C10',name);setval(s,'C11',meta.get('attention') or 'TBC')
    setval(s,'C13',meta.get('delivery') or 'Delivery address: TBC');setval(s,'C15',meta.get('billing') or 'Billing address: TBC')
    setval(s,'C18',meta.get('email') or 'TBC');setval(s,'F12',meta.get('project') or q['quote_name'] or 'Quotation draft')
    setval(s,'P10',(meta.get('quote_ref') or q['quote_ref'] or 'DRAFT')+' — DRAFT')
    setval(s,'P19',q['fx'],'number');setval(s,'I10',date_serial(date),'number')
    setval(s,'H42',('ROUND(SUM(H22:H41),2)',q['revenue']),'formula')
    idr_total=sum((i['total']*q['fx']).quantize(Decimal('1'),rounding=ROUND_HALF_UP) for i in q['items'])
    setval(s,'J42',('ROUND(SUM(J22:J41),0)',idr_total),'formula')
    setval(s,'B44','5.1');setval(s,'C44','Freight / shipping — to be confirmed')
    setval(s,'D44','Excluded from draft');setval(s,'F44',1,'number');setval(s,'G44',0,'number')
    setval(s,'H44',('F44*G44',0),'formula');setval(s,'I44',('ROUND(G44*$P$19,0)',0),'formula');setval(s,'J44',('ROUND(H44*$P$19,0)',0),'formula')
    setval(s,'H45',('ROUND(SUM(H42:H44),2)',q['revenue']),'formula');setval(s,'J45',('ROUND(SUM(J42:J44),0)',idr_total),'formula')
    setval(s,'B46','5.2');setval(s,'C46','VAT/PPN — not included (confirm applicable tax)')
    setval(s,'D46','Confirm tax treatment with the issuer before customer submission.')
    setval(s,'F46',1,'number');setval(s,'G46',0,'number')
    setval(s,'H46',('F46*G46',0),'formula');setval(s,'I46',('ROUND(G46*$P$19,0)',0),'formula');setval(s,'J46',('ROUND(H46*$P$19,0)',0),'formula')
    setval(s,'H47',('ROUND(SUM(H45:H46),2)',q['revenue']),'formula');setval(s,'J47',('ROUND(SUM(J45:J46),0)',idr_total),'formula')
    setval(s,'B63','***Remarks');setval(s,'B64','DRAFT FOR REVIEW — '+(meta.get('remarks') or 'Verify items, prices, terms and customer details.'))
    setval(s,'B66','Conditions: SUBJECT TO COMMERCIAL REVIEW')
    for row,key,default in [(67,'shipping','Delivery terms: TBC'),(68,'payment','Payment terms: TBC'),(69,'lead_time','Lead time: TBC'),(70,'validity','Quotation validity: TBC')]:
        setval(s,f'D{row}',meta.get(key) or default)
    return pack(z,s)

def qa_export(raw,q,format_id):
    z=workbook(raw);s,ss=sheet_and_shared(z)
    if format_id=='T02':
        if decimal(value(s,ss,'M89'))!=q['revenue']:raise QuoteError('Subtotal output Smartoptics tidak sesuai.')
        if decimal(value(s,ss,'E90'))!=q['fx']:raise QuoteError('Kurs output Smartoptics tidak sesuai.')
        converted=q['source_type']=='commverge_prepriced'
        active={18+idx if converted else i['row'] for idx,i in enumerate(q['items'])}
        for idx,i in enumerate(q['items']):
            r=18+idx if converted else i['row']
            if value(s,ss,f'D{r}')!=i['sku'] or value(s,ss,f'E{r}')!=i['description']:
                raise QuoteError('SKU/deskripsi output Smartoptics tidak sesuai.')
            if decimal(value(s,ss,f'L{r}'))!=Decimal(i['qty']):
                raise QuoteError('QTY output Smartoptics tidak sesuai.')
            if decimal(value(s,ss,f'G{r}'))!=i['sell'] or decimal(value(s,ss,f'M{r}'))!=i['total']:
                raise QuoteError('Harga/subtotal output Smartoptics tidak sesuai.')
            if converted and value(s,ss,f'F{r}'):
                raise QuoteError('Biaya supplier tidak tersedia dan tidak boleh dikarang.')
        for r in range(18,89):
            check_cols=('D','E','F','G','L','M') if converted else ('F','G','M')
            if r not in active and any(value(s,ss,f'{col}{r}') for col in check_cols):
                raise QuoteError('Data lama tersisa pada baris output T02.')
    else:
        if decimal(value(s,ss,'H45'))!=q['revenue']:raise QuoteError('Subtotal output CommVerge tidak sesuai.')
        if decimal(value(s,ss,'P19'))!=q['fx']:raise QuoteError('Kurs output CommVerge tidak sesuai.')
        for i,r in zip(q['items'],SLOTS):
            if value(s,ss,f'C{r}')!=i['sku'] or decimal(value(s,ss,f'G{r}'))!=i['sell']:
                raise QuoteError('SKU/harga output CommVerge tidak sesuai.')
        if len([n for n in z.namelist() if re.fullmatch(r'xl/worksheets/sheet\d+\.xml',n)])!=1:raise QuoteError('Sheet biaya lama masih terdapat di output CommVerge.')
        assert all(not value(s,ss,addr) for row in range(22,82) for addr in [f'N{row}',f'AH{row}'])
    for c in s.iter(Mtag('c')):
        if cellval(c,ss) in ERR_VALUES:raise QuoteError('Formula error pada output XLSX.')
    return True

def generate(raw,format_id='T01',margin=15,fx=18300,details='',master_bytes=None):
    if format_id not in ('T01','T02','BOTH'):raise QuoteError('Format tidak valid.')
    meta=metadata(details);q=source(raw,margin,fx)
    if q['source_type']=='commverge_prepriced':
        # Format selection is independent from the file uploaded. Cross-format
        # conversion preserves existing selling prices (no second margin).
        if master_bytes is None and format_id in ('T01','BOTH'):master_bytes=raw
        meta={**{key:q.get(key) for key in ('customer','attention','delivery','billing','email') if q.get(key)},**meta}
    date=dt.date.fromisoformat(meta.get('date') or dt.datetime.now(dt.timezone(dt.timedelta(hours=7))).date().isoformat())
    out=[]
    for code in (['T01','T02'] if format_id=='BOTH' else [format_id]):
        if code=='T01':
            if master_bytes is None:raise QuoteError('Master CommVerge tidak tersedia di server.')
            data=output_t01(q,date,meta,master_bytes)
        else:data=output_t02(q,date,meta)
        qa_export(data,q,code)
        out.append((f'DANIEL_QUOTE_{"CommVerge" if code=="T01" else "Smartoptics"}_{date.isoformat()}_DRAFT.xlsx',data))
    if len(out)==1:return out[0][0],out[0][1],'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    buf=io.BytesIO()
    with zipfile.ZipFile(buf,'w',zipfile.ZIP_DEFLATED) as z:
        for name,data in out:z.writestr(name,data)
    return f'DANIEL_QUOTE_Keduanya_{date.isoformat()}_DRAFT.zip',buf.getvalue(),'application/zip'
