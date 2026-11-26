'use strict';
const $=id=>document.getElementById(id);
const sf=(v)=>'USD '+Number(v).toLocaleString('en-US',{minimumFractionDigits:2,maximumFractionDigits:2});
const rf=(v)=>'Rp'+Math.round(v).toLocaleString('id-ID');
let access={configured:false,url:'',publishable_key:''};
let token=null,realPreview=null,realMode='simple',masterReady=null,pdfObjectUrl=null,previewEpoch=0;
const metaNames=[['customer','Customer'],['attention','Attention'],['email','Email customer'],['project','Project'],['quote_ref','Quotation reference'],['date','Date'],['delivery','Delivery address'],['billing','Billing address'],['shipping','Shipping terms'],['payment','Payment terms'],['lead_time','Lead time'],['validity','Validity'],['remarks','Remarks']];
// One-page flow: choose a file first. Authentication is enforced by the server, never bypassed here.
function renderMasterStatus(){
  const info=$('rMasterStatus');
  info.hidden=$('rFormat').value==='T02' || !realPreview || realPreview.source_type==='commverge_prepriced';
  if(info.hidden)return;
  if(masterReady===true){
    info.textContent='Master CommVerge tersedia di penyimpanan privat.';
    info.classList.remove('error');
  }else if(masterReady===false){
    info.textContent='Master CommVerge belum terpasang di server privat. Hubungi admin; format Smartoptics tetap bisa digunakan.';
    info.classList.add('error');
  }else{
    info.textContent='Memeriksa master CommVerge di server privat…';
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
  onFileSelected();
};

function docNode(tag,cls,textValue){
  const n=document.createElement(tag);
  if(cls)n.className=cls;
  if(textValue!==undefined&&textValue!==null)n.textContent=String(textValue);
  return n;
}
function valueText(v){
  if(v===null||v===undefined||String(v).trim()==='')return '—';
  return String(v);
}
function moneyText(v){
  const raw=String(v??'').trim();
  if(!raw)return '—';
  const numeric=Number(raw);
  return Number.isFinite(numeric)?numeric.toLocaleString('en-US',{minimumFractionDigits:2,maximumFractionDigits:2}):raw;
}
function snapshotTable(items,doc){
  const tbl=docNode('table','quote-paper-items');
  const head=docNode('thead');
  const hrow=docNode('tr');
  ['Part #','Description','Qty','Unit USD','Amount USD'].forEach(label=>hrow.append(docNode('th','',label)));
  head.append(hrow);tbl.append(head);
  const body=docNode('tbody');
  items.forEach((item,index)=>{
    if(doc.omitted&&index===doc.first_count){
      const skipped=docNode('tr','preview-skipped-row');
      const cell=docNode('td','',`… ${doc.omitted} baris item di tengah tidak ditampilkan dalam cuplikan. Semuanya ikut divalidasi.`);
      cell.colSpan=5;skipped.append(cell);body.append(skipped);
    }
    const row=docNode('tr','quote-item-line status-'+String(item.status||'').toLowerCase());
    row.append(docNode('td','quote-part',valueText(item.sku)));
    row.append(docNode('td','quote-description',valueText(item.description)));
    row.append(docNode('td','quote-num',valueText(item.qty)));
    row.append(docNode('td','quote-num',moneyText(item.unit)));
    const total=docNode('td','quote-num',moneyText(item.total));
    if(String(item.total||'').startsWith('#'))total.classList.add('quote-cell-error');
    row.append(total);body.append(row);
  });
  tbl.append(body);return tbl;
}
function metaField(label,v){
  const cell=docNode('div','quote-paper-field');
  cell.append(docNode('span','quote-field-label',label));
  cell.append(docNode('span','quote-field-value',valueText(v)));
  return cell;
}
function renderDocumentPreview(nodeId,doc){
  if(!doc)throw Error('Cuplikan visual Excel tidak diterima dari server.');
  const container=$(nodeId);
  const paper=docNode('div','quote-paper-inner');
  const header=docNode('div','quote-paper-letterhead');
  const brand=docNode('div','quote-paper-brand');
  if(doc.logo&&/^data:image\/(png|jpeg);base64,[A-Za-z0-9+/=]+$/.test(doc.logo)&&doc.logo.length<270000){
    const img=document.createElement('img');
    img.alt='Logo dari file Excel';
    img.src=doc.logo;
    img.className='quote-paper-logo';
    brand.append(img);
  }else{
    brand.append(docNode('strong','quote-paper-company',doc.meta.company||doc.template));
  }
  const heading=docNode('div','quote-paper-heading');
  heading.append(docNode('strong','', 'Q U O T A T I O N'));
  heading.append(docNode('span','',doc.original?'FILE DIUPLOAD':'DRAFT · REVIEW'));
  header.append(brand,heading);
  paper.append(header);
  const company=docNode('div','quote-company-caption',valueText(doc.meta.company));
  paper.append(company);
  const metadata=docNode('div','quote-paper-meta');
  [
    ['To / Customer',doc.meta.customer],
    ['Date',doc.meta.date],
    ['Attn',doc.meta.attention],
    ['Our Ref #',doc.meta.reference],
    ['Project',doc.meta.project]
  ].forEach(([label,v])=>metadata.append(metaField(label,v)));
  paper.append(metadata);
  const caption=docNode('div','quote-line-caption',doc.item_count+' item · Cuplikan hasil '+doc.template);
  paper.append(caption,snapshotTable(doc.items||[],doc));
  const totals=docNode('div','quote-paper-totals');
  totals.append(docNode('span','',doc.original?'TOTAL PADA FILE ASLI':'TOTAL HASIL GENERATOR'));
  const amount=docNode('strong','',moneyText(doc.total)+' USD');
  if(String(doc.total||'').startsWith('#')){
    amount.classList.add('quote-cell-error');
    amount.textContent=String(doc.total)+' · subtotal sumber perlu koreksi';
  }
  totals.append(amount);paper.append(totals);
  const terms=docNode('div','quote-paper-terms');
  terms.append(docNode('strong','quote-terms-heading','CONDITIONS / REMARKS'));
  [['Delivery',doc.meta.delivery],['Payment',doc.meta.payment],['Validity',doc.meta.validity],['Remarks',doc.meta.remarks]]
    .forEach(([label,v])=>{if(v&&String(v).trim())terms.append(metaField(label,v));});
  if(terms.children.length===1)terms.append(docNode('span','quote-field-value','Commercial terms: to be confirmed'));
  paper.append(terms);
  paper.append(docNode('div','quote-paper-footnote',
    doc.original?'Cuplikan read-only dari file yang di-upload · tanpa perubahan.'
                :'Cuplikan read-only dari file XLSX hasil generator · DRAFT perlu review.'));
  container.replaceChildren(paper);
}
function renderComparison(c){
  if(!c||!c.draft_document)throw Error('Preview quotation hasil generator belum tersedia.');
  const stats=c.summary;
  $('rDraftTitle').textContent='Quotation hasil generator · '+c.draft_label;
  renderDocumentPreview('rDraftDocument',c.draft_document);
  const checks=[
    ['SKU cocok',stats.sku_matches+'/'+stats.rows],
    ['QTY cocok',stats.qty_matches+'/'+stats.rows],
    ['Deskripsi cocok',stats.description_matches+'/'+stats.rows],
    ['Harga draft valid',stats.price_validated+'/'+stats.rows],
    ['Subtotal diperbaiki',stats.subtotals_repaired],
    ['Penyesuaian harga margin',stats.margin_adjustments],
    ['Selisih tidak diharapkan',stats.unexpected_differences]
  ];
  const wrap=$('rCompareChecks');
  const nodes=checks.map(([label,amount],index)=>{
    const li=docNode('li','qa-number-line');
    li.append(docNode('span','qa-item-index',String(index+1).padStart(2,'0')));
    li.append(docNode('span','qa-item-name',label));
    li.append(docNode('strong','qa-item-amount',amount));
    if(index===6&&stats.unexpected_differences)li.classList.add('qa-item-warning');
    return li;
  });
  wrap.replaceChildren(...nodes);
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
      ?'Preview menampilkan hasil CommVerge. Download Keduanya berisi hasil CommVerge dan Smartoptics.'
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
    // The user chooses the OUTPUT; source detection must never override it.
    for(const opt of $('rFormat').options)opt.disabled=false;
    renderMasterStatus();
    updateDownloadLabel();
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
async function init(){switchRealMode('simple');renderMasterStatus();updateDownloadLabel();
try{const status=await(await fetch('/api/status')).json();$('statustag').textContent=status.real_processing_enabled?'Siap':'Terbatas';const cfg=await(await fetch('/api/auth-config')).json();access=cfg;onFileSelected();}catch(e){$('statustag').textContent='Offline';onFileSelected();}}
init();
