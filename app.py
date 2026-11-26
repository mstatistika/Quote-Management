"""Quote Generator — Vercel FastAPI entrypoint, quote engine v2 portable.
Demo is public synthetic data. Real supplier XLSX is AUTH-GATED, never persisted.
"""
from __future__ import annotations
import datetime as dt
import io
import json
import os
import re
import zipfile
from pathlib import Path
from urllib.parse import urlparse
from decimal import Decimal, ROUND_HALF_UP

import httpx
from fastapi import FastAPI, Depends, Header, File, UploadFile, HTTPException, Form, Query
from fastapi.responses import FileResponse, Response
from fastapi.staticfiles import StaticFiles
from pdf_export import generate_pdf
from preview_compare import compare_workbooks
from quote_engine import (QuoteError, preview, generate, settings, metadata, source,
    sheet_and_shared, setval, pack, workbook, decimal, amount, round2, qa_export, MAX_UPLOAD)

ROOT=Path(__file__).resolve().parent
T01_PRIVATE_BUCKET='qms-templates'
T01_PRIVATE_PATH='T01_MASTER_CLEAN.xlsx'
app=FastAPI(title='Quote Generator',docs_url=None,redoc_url=None,openapi_url=None)

@app.middleware('http')
async def headers(req,handler):
    response=await handler(req)
    response.headers['X-Content-Type-Options']='nosniff'
    response.headers['Referrer-Policy']='no-referrer'
    response.headers['X-Frame-Options']='DENY'
    response.headers['Cache-Control']='no-store'
    response.headers['Content-Security-Policy']="default-src 'self'; img-src 'self' data:; style-src 'self'; script-src 'self'; connect-src 'self' https:; form-action 'self'; base-uri 'self'; object-src 'none'; frame-src 'self' blob:; frame-ancestors 'none'"
    return response

def auth_conf():
    url=os.environ.get('SUPABASE_URL','').strip().rstrip('/')
    key=os.environ.get('SUPABASE_PUBLISHABLE_KEY','').strip()
    allow={x.strip().lower() for x in os.environ.get('ALLOWED_EMAILS','').split(',') if x.strip()}
    try:
        p=urlparse(url)
        if p.scheme!='https' or not p.netloc or p.path not in ('','/'):return None,None,set()
    except ValueError:return None,None,set()
    if not key.startswith('sb_publishable_') or not allow:return None,None,set()
    return url,key,allow

async def permitted(authorization: str|None=Header(None)):
    url,key,allow=auth_conf()
    if not url:raise HTTPException(503,'Login nyata belum dikonfigurasi.')
    if not authorization or not authorization.startswith('Bearer '):raise HTTPException(401,'Harus login untuk data supplier asli.')
    token=authorization[7:]
    if len(token)<32 or len(token)>4096:raise HTTPException(401,'Sesi tidak valid.')
    try:
        async with httpx.AsyncClient(timeout=8) as client:
            resp=await client.get(url+'/auth/v1/user',headers={'apikey':key,'Authorization':'Bearer '+token})
        if resp.status_code!=200:raise HTTPException(401,'Sesi login kedaluwarsa atau tidak valid.')
        user=resp.json()
    except HTTPException:raise
    except (httpx.HTTPError,ValueError):raise HTTPException(503,'Autentikasi sedang tidak tersedia.')
    email=(user.get('email') or '').strip().lower()
    if email not in allow or not user.get('email_confirmed_at') or not user.get('id'):
        raise HTTPException(403,'Akun tidak memiliki izin memakai quotation supplier asli.')
    return {'id':user['id'],'email':email,'access_token':token}

def safe_error(e):
    if isinstance(e,QuoteError):raise HTTPException(422,str(e))
    raise HTTPException(422,'Validasi file Excel gagal.')

def synthetic(m=15,fx=18300):
    mar,rate=settings(m,fx)
    samples=[('DEMO-SKU-001','Optical transponder sample',2,Decimal('1000')),
             ('DEMO-SKU-002','Pluggable module sample',4,Decimal('250.50')),
             ('DEMO-SKU-003','Chassis fan sample',1,Decimal('85'))]
    items=[]
    for sku,desc,qty,cost in samples:
        unit=amount(cost/(1-mar/100));total=round2(unit*qty)
        items.append({'sku':sku,'description':desc,'qty':qty,'unit':unit,'total':total})
    total=sum((r['total'] for r in items),Decimal(0))
    return mar,rate,items,total

def today():return dt.datetime.now(dt.timezone(dt.timedelta(hours=7))).date().isoformat()

@app.get('/api/status')
async def status():
    url,key,allow=auth_conf()
    return {'version':'2.2.3','state':'PREVIEW','excel_comparison_enabled':True,'visual_document_preview_enabled':True,'demo_ready':True,'real_auth_configured':bool(url),
            'commverge_prepriced_supported':True,
            't01_master_ready':None,'t01_master_location':'private_supabase_storage',
            't02_uses_original_source':True,'real_processing_enabled':bool(url and real_processing_ready()),'today':today()}

@app.get('/api/whoami')
async def whoami(user:dict=Depends(permitted)):
    # Authorization is verified server-side against Supabase and ALLOWED_EMAILS.
    return {'email':user['email'], 'authorized':True}

def real_processing_ready():
    # A configured Supabase project and allowlist are required; both protected
    # endpoints separately validate a current JWT and authorized email.
    url,key,allow=auth_conf()
    return bool(url and key and allow)

@app.get('/api/auth-config')
async def auth_config():
    url,key,_=auth_conf()
    return {'configured':bool(url),'url':url or '', 'publishable_key':key or ''}

async def private_t01_master(access_token:str):
    """Fetch the reusable branded template from private Supabase Storage.
    A user-scoped JWT, not the service-role key, authorizes every download.
    """
    url,key,_=auth_conf()
    if not url:raise HTTPException(503,'Login untuk template belum dikonfigurasi.')
    path=f'{url}/storage/v1/object/authenticated/{T01_PRIVATE_BUCKET}/{T01_PRIVATE_PATH}'
    try:
        async with httpx.AsyncClient(timeout=20) as client:
            async with client.stream('GET',path,headers={
                'apikey':key,'Authorization':'Bearer '+access_token,
                'Accept':'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            }) as resp:
                if resp.status_code==404:return None
                if resp.status_code in (401,403):
                    raise HTTPException(403,'Akses template CommVerge ditolak.')
                if resp.status_code!=200:
                    raise HTTPException(503,'Penyimpanan master CommVerge sedang tidak tersedia.')
                chunks=[]
                size=0
                async for chunk in resp.aiter_bytes():
                    size+=len(chunk)
                    if size>MAX_UPLOAD:raise HTTPException(413,'Master CommVerge melebihi 8 MB.')
                    chunks.append(chunk)
    except HTTPException:raise
    except httpx.HTTPError:
        raise HTTPException(503,'Koneksi penyimpanan master CommVerge gagal.')
    raw=b''.join(chunks)
    if not raw.startswith(b'PK') or len(raw)<200:
        raise HTTPException(422,'Master CommVerge di storage bukan file XLSX valid.')
    return raw

@app.get('/api/master-status')
async def master_status(user:dict=Depends(permitted)):
    master=await private_t01_master(user['access_token'])
    return {'t01_ready':master is not None,'storage':'private','requires_user_upload':False}

@app.get('/api/demo')
async def get_demo(margin:float=15,fx:int=18300):
    try:mar,rate,items,total=synthetic(margin,fx)
    except QuoteError as e:raise HTTPException(422,str(e))
    return {'sample':True,'currency':'USD','margin':float(mar),'fx':int(rate),'total_usd':float(total),
            'total_idr':int((total*rate).quantize(Decimal('1'),rounding=ROUND_HALF_UP)),
            'items':[{'sku':i['sku'],'description':i['description'],'qty':i['qty'],'unit_usd':float(i['unit']),'total_usd':float(i['total'])} for i in items]}

@app.get('/api/demo-export')
async def demo_export(margin:float=15,fx:int=18300,format:str='T01',customer:str='DEMO Customer',project:str='Synthetic Sample'):
    if format not in ('T01','T02','BOTH'):raise HTTPException(422,'Format tidak valid.')
    if len(customer)>90 or len(project)>90 or any(v[:1] in ('=','+','@') for v in [customer,project]):raise HTTPException(422,'Metadata tidak valid.')
    try:mar,rate,items,total=synthetic(margin,fx)
    except QuoteError as e:raise HTTPException(422,str(e))
    def make_sample(code):
        z=workbook((ROOT/'private/DEMO_MASTER.xlsx').read_bytes())
        s,ss=sheet_and_shared(z)
        setval(s,'B4',code+'-DEMO-SAMPLE');setval(s,'C4',today());setval(s,'A7',mar,'number');setval(s,'B7',rate,'number')
        setval(s,'C7',total,'number');setval(s,'D7',int((total*rate).quantize(Decimal('1'),rounding=ROUND_HALF_UP)),'number')
        for i,row in zip(items,range(11,14)):
            setval(s,f'A{row}',i['sku']);setval(s,f'B{row}',i['description']);setval(s,f'C{row}',i['qty'],'number');setval(s,f'D{row}',i['unit'],'number');setval(s,f'E{row}',(f'ROUND(C{row}*D{row},2)',i['total']),'formula')
        setval(s,'E15',('ROUND(SUM(E11:E13),2)',total),'formula')
        setval(s,'A18',customer);setval(s,'C18',project)
        return pack(z,s)
    files=[(f'{code}-DEMO-SAMPLE.xlsx',make_sample(code)) for code in (['T01','T02'] if format=='BOTH' else [format])]
    if len(files)==1:name,data=files[0];mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    else:
        mem=io.BytesIO()
        with zipfile.ZipFile(mem,'w',zipfile.ZIP_DEFLATED) as z:
            for fname,raw in files:z.writestr(fname,raw)
        name,data,mime='DANIEL_DEMO_SAMPLES.zip',mem.getvalue(),'application/zip'
    return Response(data,media_type=mime,headers={'Content-Disposition':f'attachment; filename="{name}"','X-Quotation-Status':'DEMO-SAMPLE-ONLY'})

async def read_quote(file:UploadFile):
    if not file.filename or not file.filename.lower().endswith('.xlsx'):raise HTTPException(422,'Pilih sumber Excel .xlsx')
    buf=await file.read(MAX_UPLOAD+1)
    if len(buf)>MAX_UPLOAD:raise HTTPException(413,'File melebihi 8 MB')
    return buf

@app.post('/api/preview')
async def real_preview(quote:UploadFile=File(...),margin:str=Form('15'),fx:str=Form('18300'),format:str=Form('T01'),details:str=Form('{}'),user:dict=Depends(permitted)):
    if not real_processing_ready():raise HTTPException(503,'Pemrosesan quotation asli belum diaktifkan.')
    raw=await read_quote(quote)
    try:
        result=preview(raw,margin,fx)
        master=None
        if format in ('T01','BOTH'):
            if result['source_type']=='commverge_prepriced':
                master=raw
            else:
                master=await private_t01_master(user['access_token'])
                if master is None:
                    raise QuoteError('Master CommVerge T01 belum tersedia; pilih Smartoptics (T02) atau minta admin memasang master.')
        result['comparison']=compare_workbooks(raw,format,margin,fx,details,master)
        return result
    except Exception as e:safe_error(e)

@app.post('/api/generate')
async def real_generate(quote:UploadFile=File(...),margin:str=Form('15'),fx:str=Form('18300'),format:str=Form('T01'),details:str=Form('{}'),file_type:str=Form('xlsx'),user:dict=Depends(permitted)):
    if not real_processing_ready():raise HTTPException(503,'Pemrosesan quotation asli belum diaktifkan.')
    raw=await read_quote(quote)
    try:
        master=None
        if format in ('T01','BOTH'):
            inspected=source(raw,margin,fx)
            if inspected['source_type']=='commverge_prepriced':
                master=raw  # The supplier's quote has already been priced and branded.
            else:
                master=await private_t01_master(user['access_token'])
                if master is None:
                    raise QuoteError('Master CommVerge T01 belum dipasang di penyimpanan privat.')
        if file_type=='xlsx':
            name,data,mime=generate(raw,format,margin,fx,details,master)
        elif file_type=='pdf':
            name,data,mime=generate_pdf(raw,format,margin,fx,details,master)
        else:
            raise QuoteError('Tipe output harus XLSX atau PDF.')
    except Exception as e:safe_error(e)
    return Response(data,media_type=mime,headers={'Content-Disposition':f'attachment; filename="{name}"','X-Quotation-Status':'DRAFT-REVIEW-REQUIRED'})

@app.post('/api/pdf-preview')
async def pdf_preview(quote:UploadFile=File(...),margin:str=Form('15'),fx:str=Form('18300'),format:str=Form('T01'),details:str=Form('{}'),user:dict=Depends(permitted)):
    if not real_processing_ready():raise HTTPException(503,'Pemrosesan quotation asli belum diaktifkan.')
    raw=await read_quote(quote)
    try:
        if format not in ('T01','T02','BOTH'):raise QuoteError('Format tidak valid.')
        pdf_format='T01' if format=='BOTH' else format
        master=None
        if pdf_format=='T01':
            inspected=source(raw,margin,fx)
            if inspected['source_type']=='commverge_prepriced':
                master=raw
            else:
                master=await private_t01_master(user['access_token'])
                if master is None:
                    raise QuoteError('Master CommVerge T01 belum dipasang di penyimpanan privat.')
        _,data,mime=generate_pdf(raw,pdf_format,margin,fx,details,master)
    except Exception as e:safe_error(e)
    return Response(data,media_type='application/pdf',headers={
        'Content-Disposition':'inline; filename="QUOTATION_PREVIEW_A4.pdf"',
        'X-Quotation-Status':'DRAFT-REVIEW-REQUIRED',
    })

@app.get('/')
async def index():return FileResponse(ROOT/'static/index.html')
app.mount('/static',StaticFiles(directory=ROOT/'static'),name='static')
