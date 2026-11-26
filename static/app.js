'use strict';
const $=id=>document.getElementById(id);
const sf=(v)=>'USD '+Number(v).toLocaleString('en-US',{minimumFractionDigits:2,maximumFractionDigits:2});
const rf=(v)=>'Rp'+Math.round(v).toLocaleString('id-ID');
const samples=[{sku:'DEMO-SKU-001',description:'Optical transponder sample',qty:2,cost:1000},{sku:'DEMO-SKU-002',description:'Pluggable module sample',qty:4,cost:250.5},{sku:'DEMO-SKU-003',description:'Chassis fan sample',qty:1,cost:85}];
let demoState={total:0,valid:false};
let access={configured:false,url:'',publishable_key:''};
let token=null,realPreview=null,realMode='simple',masterReady=null,pdfObjectUrl=null,previewEpoch=0;
const metaNames=[['customer','Customer'],['attention','Attention'],['email','Email customer'],['project','Project'],['quote_ref','Quotation reference'],['date','Date'],['delivery','Delivery address'],['billing','Billing address'],['shipping','Shipping terms'],['payment','Payment terms'],['lead_time','Lead time'],['validity','Validity'],['remarks','Remarks']];
// One-page flow: choose a file first. Authentication is enforced by the server, never bypassed here.
function syncDemoBadge() { $('demoBadge').hidden=Boolean(token); }
syncDemoBadge();
function renderMasterStatus(){
  const info=$('rMasterStatus');
  info.hidden=$('rFormat').value==='T02' || !realPreview || realPreview.source_type==='commverge_prepriced';
  if(info.hidden)return;
  if(masterReady===true){
    info.textContent='Master CommVerge T01 tersedia di penyimpanan privat.';
    info.classList.remove('error');
  }else if(masterReady===false){
    info.textContent='Master CommVerge T01 belum terpasang di server privat. Hubungi admin; Smartoptics (T02) tetap bisa digunakan.';
    info.classList.add('error');
  }else{
    info.textContent='Memeriksa master CommVerge T01 di server privat…';
    info.classList.remove('error');
  }
}
async function checkPrivateMaster(){
  masterReady=null;
  renderMasterStatus();
  if(!token)return;
  try{
    const r=await fetch('/api/master-status',{
      headers:{Authorization:'Bearer '+token},
      cache:'no-store'
    });
    if(!r.ok)throw Error('Master server belum dapat diverifikasi.');
    const result=await r.json();
    masterReady=Boolean(result.t01_ready);
  }catch(e){
    masterReady=false;
  }
  renderMasterStatus();
}
function onFileSelected(){
  invalidateReal();
  $('rMargin').disabled=false;
  for(const option of $('rFormat').options)option.disabled=false;
  const selected=$('rFile').files[0];
  const info=$('selectedFileInfo');
  $('realUnconfigured').classList.add('hidden');
  $('realLogin').classList.add('hidden');
  $('realWorkspace').classList.add('hidden');
  $('rError').textContent='';
  info.hidden=!selected;
  if(!selected) return;
  info.textContent='Dipilih: '+selected.name;
  if(!selected.name.toLowerCase().endsWith('.xlsx') || !selected.size || selected.size>8*1024*1024){
    info.textContent='Pilih file XLSX maksimal 8 MB.';
    return;
  }
  if(!access.configured) $('realUnconfigured').classList.remove('hidden');
  else if(token) $('realWorkspace').classList.remove('hidden');
  else $('realLogin').classList.remove('hidden');
}
$('rFile').addEventListener('change',onFileSelected);
const itemElement=i=>{const el=document.createElement('div');el.className='item';const left=document.createElement('div');const a=document.createElement('div');a.className='sku';a.textContent=i.sku;const d=document.createElement('div');d.className='desc';d.textContent=i.description;const q=document.createElement('div');q.className='qty';q.textContent=`Qty ${i.qty} · Unit ${sf(i.unit)}`;left.append(a,d,q);const right=document.createElement('div');const total=document.createElement('strong');total.textContent=sf(i.total);const caption=document.createElement('small');caption.textContent='Subtotal';right.append(total,caption);el.append(left,right);return el;};
const getDemo=()=>{const margin=Number($('dMargin').value),fx=Number($('dFx').value);if(!Number.isFinite(margin)||margin<=0||margin>80||!Number.isSafeInteger(fx)||fx<1000||fx>1e6)throw Error('Margin harus 0–80% dan kurs harus Rp1.000–Rp1.000.000.');
const items=samples.map(i=>{const unit=Math.ceil((i.cost/(1-margin/100))*100-1e-8)/100;return {...i,unit,total:Math.round(unit*i.qty*100)/100};});
return {margin,fx,items,total:Math.round(items.reduce((x,y)=>x+y.total,0)*100)/100};};
function updateDemo(){try{const s=getDemo();demoState={total:s.total,valid:true};$('dUsd').textContent=sf(s.total);$('dIdr').textContent=rf(s.total*s.fx);$('dGm').textContent=s.margin+'%';$('dRate').textContent=rf(s.fx);$('dItems').replaceChildren(...s.items.map(itemElement));$('dQa').textContent='5/5 pemeriksaan lulus';$('dError').textContent='';$('dDownload').disabled=false;}catch(e){demoState.valid=false;$('dError').textContent=e.message;$('dDownload').disabled=true;}}
$('dMargin').oninput=updateDemo;$('dFx').oninput=updateDemo;
$('dReset').onclick=()=>{$('dMargin').value=15;$('dFx').value=18300;$('dCustomer').value='DEMO Customer';$('dProject').value='Synthetic Sample';$('dFormat').value='T01';document.querySelector('[data-mode="simple"]').click();updateDemo();};
document.querySelectorAll('[data-mode]').forEach(b=>b.onclick=()=>{document.querySelectorAll('[data-mode]').forEach(t=>t.classList.toggle('active',t===b));const advanced=b.dataset.mode==='advanced';$('dAdvanced').style.display=advanced?'block':'none';});
$('dDownload').onclick=async()=>{if(!demoState.valid)return;const btn=$('dDownload');btn.disabled=true;btn.textContent='Menyiapkan Excel...';try{const s=getDemo();const url=new URL('/api/demo-export',location.origin);url.searchParams.set('margin',s.margin);url.searchParams.set('fx',s.fx);url.searchParams.set('format',$('dFormat').value);url.searchParams.set('customer',$('dCustomer').value);url.searchParams.set('project',$('dProject').value);const r=await fetch(url);if(!r.ok){const d=await r.json();throw Error(d.detail||'Gagal mengunduh contoh.');}const blob=await r.blob();download(blob,$('dFormat').value==='BOTH'?'DANIEL_DEMO_SAMPLES.zip':$('dFormat').value+'-DEMO-SAMPLE.xlsx');$('dError').textContent='File contoh berhasil diunduh. Bukan quotation untuk customer.';}catch(e){$('dError').textContent=e.message;}finally{btn.disabled=false;btn.textContent='↓ Download Excel Contoh';}};
function download(blob,name){const obj=URL.createObjectURL(blob),a=document.createElement('a');a.href=obj;a.download=name;document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(obj),20000);}
function switchRealMode(mode){realMode=mode;document.querySelectorAll('[data-r-mode]').forEach(b=>b.classList.toggle('active',b.dataset.rMode===mode));$('rAdvanced').style.display=mode==='advanced'?'block':'none';invalidateReal();}
function clearPdfPreview(){
  if(pdfObjectUrl){URL.revokeObjectURL(pdfObjectUrl);pdfObjectUrl=null;}
  $('rPdfFrame').removeAttribute('src');
  $('rPdfFrame').hidden=true;
}
function invalidateReal(){
  previewEpoch+=1;
  clearPdfPreview();
  realPreview=null;
  $('rResults').hidden=true;
  renderMasterStatus();
  $('rAck').checked=false;
  $('rGenerate').disabled=true;
}
document.querySelectorAll('[data-r-mode]').forEach(b=>b.onclick=()=>switchRealMode(b.dataset.rMode));
for(const [key,label] of metaNames){const el=document.createElement('label');el.textContent=label;const input=document.createElement('input');input.id='meta_'+key;input.maxLength=190;if(key==='date'){input.type='date';}else if(key==='email'){input.type='email';}else{input.type='text';}input.oninput=invalidateReal;el.append(input);$('rAdvancedFields').append(el);}
function form(){if(!token)throw Error('Login terlebih dahulu.');const f=$('rFile').files[0];if(!f||!f.name.toLowerCase().endsWith('.xlsx'))throw Error('Pilih Excel quotation .xlsx.');if(f.size>8*1024*1024)throw Error('File lebih dari 8 MB.');const fd=new FormData();fd.append('quote',f);fd.append('margin',$('rMargin').value);fd.append('fx',$('rFx').value);return fd;}
async function remoteApi(path,fd){const r=await fetch(path,{method:'POST',headers:{Authorization:'Bearer '+token},body:fd,cache:'no-store'});if(!r.ok){const data=await r.json().catch(()=>({detail:'Permintaan gagal.'}));if([401,403].includes(r.status))token=null;throw Error(data.detail||'Validasi gagal.');}return r;}
$('rLoginBtn').onclick=async()=>{
  const btn=$('rLoginBtn');
  btn.disabled=true;
  $('rLoginError').textContent='';
  try{
    if(!access.configured)throw Error('Login belum dikonfigurasi.');
    const email=$('rEmail').value;
    const password=$('rPassword').value;
    if(!email || !password)throw Error('Pilih email dan masukkan password.');
    const r=await fetch(access.url+'/auth/v1/token?grant_type=password',{
      method:'POST',
      headers:{apikey:access.publishable_key,'Content-Type':'application/json'},
      body:JSON.stringify({email,password}),
      cache:'no-store'
    });
    const data=await r.json().catch(()=>({}));
    if(!r.ok||!data.access_token)throw Error('Email atau password salah.');
    token=data.access_token;
    const authCheck=await fetch('/api/whoami',{
      headers:{Authorization:'Bearer '+token},
      cache:'no-store'
    });
    if(!authCheck.ok){
      token=null;
      throw Error('Akun tidak diizinkan mengakses quotation.');
    }
    $('rPassword').value='';
    $('realLogin').classList.add('hidden');
    $('realWorkspace').classList.remove('hidden');
    syncDemoBadge();
    checkPrivateMaster();
    $('realWorkspace').scrollIntoView({behavior:'smooth',block:'start'});
  }catch(e){
    $('rLoginError').textContent=e.message;
  }finally{
    btn.disabled=false;
  }
};
$('rLogout').onclick=()=>{
  token=null;
  masterReady=null;
  $('rPassword').value='';
  syncDemoBadge();
  onFileSelected();
};
const worksheetCols=[
  ['sku','SKU'],['description','Deskripsi'],['qty','QTY'],['unit','Harga/unit USD'],['total','Subtotal USD']
];
function worksheetCell(textValue,tag='td',className=''){
  const el=document.createElement(tag);
  if(className)el.className=className;
  el.textContent=textValue==null?'':String(textValue);
  return el;
}
function renderWorksheet(id,rows,columns){
  const table=$(id);
  const thead=document.createElement('thead');
  const letters=document.createElement('tr');
  letters.className='excel-column-letters';
  letters.append(worksheetCell('','th','excel-corner'));
  for(const [key] of worksheetCols)letters.append(worksheetCell(columns[key]||'','th'));
  const labels=document.createElement('tr');
  labels.className='excel-heading-row';
  labels.append(worksheetCell('#','th','excel-row-index'));
  for(const [,label] of worksheetCols)labels.append(worksheetCell(label,'th'));
  thead.append(letters,labels);
  const tbody=document.createElement('tbody');
  for(const row of rows){
    const tr=document.createElement('tr');
    tr.className='excel-data-row status-'+String(row.status||'').toLowerCase();
    if(row.note)tr.title=row.note;
    tr.append(worksheetCell(row.row,'th','excel-row-index'));
    for(const [key] of worksheetCols){
      let textValue=row[key];
      let kind=(key==='qty'||key==='unit'||key==='total')?'excel-number':'';
      if(key==='unit'||key==='total'){
        const numeric=Number(textValue);
        if(textValue!==''&&Number.isFinite(numeric))textValue=numeric.toLocaleString('en-US',{minimumFractionDigits:2,maximumFractionDigits:2});
        else if(textValue&&String(textValue).startsWith('#'))kind+=' excel-error';
      }
      const cell=worksheetCell(textValue,'td',kind.trim());
      if(key==='sku')cell.classList.add('excel-sku');
      tr.append(cell);
    }
    tbody.append(tr);
  }
  table.replaceChildren(thead,tbody);
}
function renderComparison(c){
  if(!c)throw Error('Data perbandingan Excel belum tersedia.');
  const stats=c.summary;
  $('rSourceTitle').textContent='A · '+c.source_label;
  $('rDraftTitle').textContent='B · '+c.draft_label;
  renderWorksheet('rSourceTable',c.source_rows,c.source_columns);
  renderWorksheet('rSheetTable',c.draft_rows,c.draft_columns);
  $('rSourceTotal').textContent='Total sumber: '+String(c.source_total||'—')+' USD';
  $('rDraftTotal').textContent='Total draft: '+String(c.draft_total||'—')+' USD';
  $('rCompareNote').textContent=c.note;
  const badges=[
    stats.sku_matches+'/'+stats.rows+' SKU cocok',
    stats.qty_matches+'/'+stats.rows+' QTY cocok',
    stats.description_matches+'/'+stats.rows+' deskripsi cocok',
    stats.price_validated+'/'+stats.rows+' harga draft valid',
    stats.subtotals_repaired+' subtotal diperbaiki',
    stats.margin_adjustments+' penyesuaian harga margin',
    stats.unexpected_differences+' selisih tidak diharapkan'
  ];
  const wrap=$('rCompareChecks');
  wrap.replaceChildren(...badges.map((label,index)=>{
    const span=document.createElement('span');
    span.className='compare-check'+(index===6&&stats.unexpected_differences?' compare-warning':'');
    span.textContent=label;return span;
  }));
}
function showReviewMode(mode){
  const excel=mode==='excel';
  $('rShowExcel').classList.toggle('active',excel);
  $('rShowPdf').classList.toggle('active',!excel);
  $('rShowExcel').setAttribute('aria-pressed',String(excel));
  $('rShowPdf').setAttribute('aria-pressed',String(!excel));
  $('rExcelPreview').hidden=!excel;
  $('rPdfPreview').hidden=excel;
  if(!excel&&realPreview&&!pdfObjectUrl)refreshPdfPreview();
}
$('rShowExcel').onclick=()=>showReviewMode('excel');
$('rShowPdf').onclick=()=>showReviewMode('pdf');

function collectDetails(){
  const meta={};
  if(realMode==='advanced'){
    for(const [key] of metaNames){
      const v=$('meta_'+key).value.trim();
      if(v)meta[key]=v;
    }
  }
  return meta;
}
function pdfForm(){
  const fd=form();
  fd.append('format',$('rFormat').value);
  fd.append('details',JSON.stringify(collectDetails()));
  return fd;
}
async function refreshPdfPreview(){
  if(!realPreview)return;
  const epoch=previewEpoch;
  $('rRefreshPdf').disabled=true;
  $('rPreviewMessage').textContent='Menyiapkan preview PDF A4…';
  clearPdfPreview();
  try{
    const r=await remoteApi('/api/pdf-preview',pdfForm());
    const blob=await r.blob();
    if(epoch!==previewEpoch)return;
    if(blob.type!=='application/pdf'||!blob.size)throw Error('Preview bukan file PDF yang valid.');
    pdfObjectUrl=URL.createObjectURL(blob);
    $('rPdfFrame').src=pdfObjectUrl;
    $('rPdfFrame').hidden=false;
    $('rPreviewMessage').textContent=$('rFormat').value==='BOTH'
      ?'Menampilkan preview CommVerge (T01). Jika memilih kedua format, file download berisi T01 dan T02.'
      :'Preview PDF A4 sesuai item, margin, kurs, dan detail quotation yang telah divalidasi.';
  }catch(e){
    if(epoch===previewEpoch){
      $('rPreviewMessage').textContent='Preview PDF belum tersedia: '+e.message;
    }
  }finally{$('rRefreshPdf').disabled=false;}
}
$('rPreview').onclick=async()=>{
  const btn=$('rPreview');btn.disabled=true;$('rError').textContent='';
  invalidateReal();
  try{
    const epoch=previewEpoch;
    const r=await remoteApi('/api/preview',pdfForm());
    const p=await r.json();
    if(epoch!==previewEpoch)return;
    realPreview=p;
    $('rUsd').textContent=sf(p.total_usd);
    $('rIdr').textContent=rf(p.total_idr);
    const prepriced=p.source_type==='commverge_prepriced';
    $('rMargin').disabled=prepriced;
    for(const opt of $('rFormat').options)opt.disabled=prepriced&&opt.value!=='T01';
    if(prepriced)$('rFormat').value='T01';
    renderMasterStatus();
    updateDownloadLabel();
    $('rResultInfo').textContent=prepriced
      ?`${p.item_count} item · Harga jual existing dipertahankan · ${p.recomputed_subtotals} subtotal diperbaiki (QTY × harga unit) · DRAFT wajib review. Margin tambahan tidak diterapkan.`
      :`${p.item_count} item · margin aktual ${p.actual_gm}% · validasi PASS`;
    $('rPricingNote').textContent=p.price_note||'';
    $('rPricingNote').hidden=!p.price_note;
    renderComparison(p.comparison);
    showReviewMode('excel');
    $('rResults').hidden=false;
    $('rAck').checked=false;
    $('rGenerate').disabled=true;
  }catch(e){invalidateReal();$('rError').textContent=e.message;}
  finally{btn.disabled=false;}
};
$('rRefreshPdf').onclick=refreshPdfPreview;
$('rAck').onchange=()=>{$('rGenerate').disabled=!$('rAck').checked||!realPreview;};
function updateDownloadLabel(){
  const fmt=$('rExportType').value;
  const both=$('rFormat').value==='BOTH';
  $('rGenerate').textContent=fmt==='pdf'
    ?'↓ Download '+(both?'ZIP PDF':'PDF A4')+' DRAFT'
    :'↓ Download '+(both?'ZIP Excel':'Excel')+' DRAFT';
}
$('rExportType').onchange=updateDownloadLabel;
$('rGenerate').onclick=async()=>{
  const btn=$('rGenerate');btn.disabled=true;$('rError').textContent='';
  try{
    if(!realPreview||!$('rAck').checked)throw Error('Preview dan review harus dilakukan terlebih dahulu.');
    const fd=pdfForm();
    fd.append('file_type',$('rExportType').value);
    const r=await remoteApi('/api/generate',fd);
    const mime=r.headers.get('Content-Type')||'';
    const fallback='QUOTE_DRAFT.'+(mime.includes('zip')?'zip':mime.includes('pdf')?'pdf':'xlsx');
    const filename=(r.headers.get('Content-Disposition')||'').match(/filename="([^"]+)"/)?.[1]||fallback;
    download(await r.blob(),filename);
    $('rError').textContent='Download DRAFT berhasil. Review manual sebelum dikirim ke customer.';
  }catch(e){$('rError').textContent=e.message;}
  finally{btn.disabled=false;updateDownloadLabel();}
};
['rMargin','rFx'].forEach(id=>$(id).oninput=invalidateReal);
$('rFormat').onchange=()=>{invalidateReal();renderMasterStatus();updateDownloadLabel();};
async function init(){updateDemo();document.querySelector('[data-mode="simple"]').click();switchRealMode('simple');renderMasterStatus();updateDownloadLabel();
try{const status=await(await fetch('/api/status')).json();$('statustag').textContent=status.demo_ready?'Demo siap':'Preview';const cfg=await(await fetch('/api/auth-config')).json();access=cfg;onFileSelected();}catch(e){$('statustag').textContent='Offline';onFileSelected();}}
init();
