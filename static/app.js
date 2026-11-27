'use strict';
const $=id=>document.getElementById(id);
const sf=(v)=>'USD '+Number(v).toLocaleString('en-US',{minimumFractionDigits:2,maximumFractionDigits:2});
const rf=(v)=>'Rp'+Math.round(v).toLocaleString('id-ID');
let access={configured:false,url:'',publishable_key:''};
let token=null,realPreview=null,realMode='simple',masterReady=null,pdfObjectUrl=null,previewEpoch=0,pdfRequestEpoch=null,pdfSequence=0,reviewMode="excel",mailReady=false,deliveryBusy=false;
const metaNames=[['customer','Customer'],['attention','Attention'],['email','Email customer'],['project','Project'],['quote_ref','Quotation reference'],['date','Date'],['delivery','Delivery address'],['billing','Billing address'],['shipping','Shipping terms'],['payment','Payment terms'],['lead_time','Lead time'],['validity','Validity'],['remarks','Remarks']];
// Login-first flow. Authentication and real quotation access are enforced by the server.
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
  $('rError').textContent='';
  info.hidden=!selected;
  info.textContent='';
  if(!selected)return;
  if(!selected.name.toLowerCase().endsWith('.xlsx') || !selected.size || selected.size>8*1024*1024){
    info.textContent='Pilih file XLSX maksimal 8 MB.';
    $('rFile').value='';
    return;
  }
  info.textContent='Dipilih: '+selected.name;
}
function showLogin(message=''){
  token=null;
  masterReady=null;
  $('real').classList.add('hidden');
  $('rFile').value='';
  $('rPassword').value='';
  $('rEmail').selectedIndex=0;
  for(const [key] of metaNames)$('meta_'+key).value='';
  $('rMargin').value='15';
  $('rFx').value='18300';
  $('rFormat').value='T01';
  $('rExportType').value='xlsx';
  $('rDraftDocument').replaceChildren();
  $('rCompareChecks').replaceChildren();
  $('rUsd').textContent='—';
  $('rIdr').textContent='—';
  switchRealMode('simple');
  onFileSelected();
  $('realLogin').classList.remove('hidden');
  $('rLoginError').textContent=message;
  $('rLoginBtn').disabled=!access.configured;
}
$('rFile').addEventListener('change',onFileSelected);
function download(blob,name){const obj=URL.createObjectURL(blob),a=document.createElement('a');a.href=obj;a.download=name;document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(obj),20000);}
function switchRealMode(mode){realMode=mode;document.querySelectorAll('[data-r-mode]').forEach(b=>b.classList.toggle('active',b.dataset.rMode===mode));$('rAdvanced').style.display=mode==='advanced'?'block':'none';invalidateReal();}
function clearPdfPreview(){
  pdfSequence+=1;pdfRequestEpoch=null;
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
  $('rSendEmail').disabled=true;
  $('rDeliveryStatus').hidden=true;
}
document.querySelectorAll('[data-r-mode]').forEach(b=>b.onclick=()=>switchRealMode(b.dataset.rMode));
for(const [key,label] of metaNames){const el=document.createElement('label');el.textContent=label;const input=document.createElement('input');input.id='meta_'+key;input.maxLength=190;if(key==='date'){input.type='date';}else if(key==='email'){input.type='email';}else{input.type='text';}input.oninput=invalidateReal;el.append(input);$('rAdvancedFields').append(el);}
function form(){if(!token)throw Error('Login terlebih dahulu.');const f=$('rFile').files[0];if(!f||!f.name.toLowerCase().endsWith('.xlsx'))throw Error('Pilih Excel quotation .xlsx.');if(f.size>8*1024*1024)throw Error('File lebih dari 8 MB.');const fd=new FormData();fd.append('quote',f);fd.append('margin',$('rMargin').value);fd.append('fx',$('rFx').value);return fd;}
async function remoteApi(path,fd){const session=token;const r=await fetch(path,{method:'POST',headers:{Authorization:'Bearer '+session},body:fd,cache:'no-store'});if(session!==token)throw Error('Sesi telah berakhir.');if(!r.ok){const data=await r.json().catch(()=>({detail:'Permintaan gagal.'}));if(r.status===401)showLogin('Sesi berakhir. Silakan login kembali.');throw Error(data.detail||'Validasi gagal.');}return r;}
$('realLogin').onsubmit=async(event)=>{
  event.preventDefault();
  const btn=$('rLoginBtn');
  btn.disabled=true;
  $('rLoginError').textContent='';
  try{
    if(!access.configured)throw Error('Login belum dikonfigurasi.');
    const email=$('rEmail').value;
    const password=$('rPassword').value;
    if(!email || !password)throw Error('Pilih user dan masukkan password.');
    const r=await fetch(access.url+'/auth/v1/token?grant_type=password',{
      method:'POST',
      headers:{apikey:access.publishable_key,'Content-Type':'application/json'},
      body:JSON.stringify({email,password}),
      cache:'no-store'
    });
    const data=await r.json().catch(()=>({}));
    if(!r.ok||!data.access_token)throw Error('User atau password salah.');
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
    $('real').classList.remove('hidden');
    checkPrivateMaster();
  }catch(e){
    token=null;
    $('rLoginError').textContent=e.message;
  }finally{
    btn.disabled=false;
  }
};
$('rLogout').onclick=()=>showLogin();

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
      const cell=docNode('td','',`… ${doc.omitted} item tidak ditampilkan.`);
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
  if(!doc)throw Error('Preview Excel tidak tersedia.');
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
  const caption=docNode('div','quote-line-caption',doc.item_count+' item');
  paper.append(caption,snapshotTable(doc.items||[],doc));
  const totals=docNode('div','quote-paper-totals');
  totals.append(docNode('span','','TOTAL'));
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
  container.replaceChildren(paper);
}
function renderComparison(c){
  if(!c||!c.draft_document)throw Error('Preview quotation belum tersedia.');
  const stats=c.summary;
  $('rDraftTitle').textContent='Quotation · '+c.draft_label;
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
    const li=docNode('li','qa-line');
    li.append(docNode('span','qa-item-name',label));
    li.append(docNode('strong','qa-item-amount',amount));
    if(index===6&&stats.unexpected_differences)li.classList.add('qa-item-warning');
    return li;
  });
  wrap.replaceChildren(...nodes);
}
function showReviewMode(mode){
  reviewMode=mode;
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
$('rPreviewFormat').onchange=()=>{if(!realPreview)return;clearPdfPreview();renderComparison(realPreview.comparisons[$('rPreviewFormat').value]);showReviewMode(reviewMode);};

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
function pdfForm(format=$('rFormat').value){
  const fd=form();
  fd.append('format',format);
  fd.append('details',JSON.stringify(collectDetails()));
  return fd;
}
async function refreshPdfPreview(){
  if(!realPreview||pdfRequestEpoch!==null)return;
  const epoch=previewEpoch;
  clearPdfPreview();
  const sequence=pdfSequence;
  pdfRequestEpoch=sequence;
  $('rPreviewMessage').textContent='Menyiapkan preview…';
  $('rPreviewMessage').hidden=false;
  try{
    const selected=$('rFormat').value==='BOTH'?$('rPreviewFormat').value:$('rFormat').value;
    const r=await remoteApi('/api/pdf-preview',pdfForm(selected));
    const blob=await r.blob();
    if(epoch!==previewEpoch||sequence!==pdfSequence)return;
    if(blob.type!=='application/pdf'||!blob.size)throw Error('Preview bukan file PDF yang valid.');
    pdfObjectUrl=URL.createObjectURL(blob);
    $('rPdfFrame').src=pdfObjectUrl+'#page=1&view=FitH';
    $('rPdfFrame').hidden=false;
    $('rPreviewMessage').textContent='';
    $('rPreviewMessage').hidden=true;
    $('rPdfFrame').title='Preview PDF '+(selected==='T01'?'CommVerge':'Smartoptics');
  }catch(e){
    if(epoch===previewEpoch&&sequence===pdfSequence){
      $('rPreviewMessage').textContent='Preview PDF belum tersedia: '+e.message;
      $('rPreviewMessage').hidden=false;
    }
  }finally{if(pdfRequestEpoch===sequence)pdfRequestEpoch=null;}
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
    $('rPreviewFormat').hidden=$('rFormat').value!=='BOTH';
    $('rPreviewFormat').value='T01';
    renderComparison(p.comparison);
    showReviewMode('excel');
    $('rResults').hidden=false;
    $('rAck').checked=false;
    $('rGenerate').disabled=true;
  }catch(e){invalidateReal();$('rError').textContent=e.message;}
  finally{btn.disabled=false;}
};
function syncDelivery(){const disabled=deliveryBusy||!$('rAck').checked||!realPreview;$('rGenerate').disabled=disabled;$('rSendEmail').disabled=disabled||!mailReady;}
$('rAck').onchange=syncDelivery;
function updateDownloadLabel(){
  const fmt=$('rExportType').value;
  const both=$('rFormat').value==='BOTH';
  $('rGenerate').textContent=fmt==='pdf'?'Download PDF':'Download Excel';
  $('rGenerate').title=both?'Download kedua quotation dalam ZIP':'Download quotation';
  $('rSendEmail').title=mailReady?'Kirim file ke user yang login':'Email Collective Revivers belum dikonfigurasi';
  syncDelivery();
}
$('rExportType').onchange=updateDownloadLabel;
$('rGenerate').onclick=async()=>{
  if(deliveryBusy)return;deliveryBusy=true;syncDelivery();$('rError').textContent='';
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
  finally{deliveryBusy=false;updateDownloadLabel();}
};
$('rSendEmail').onclick=async()=>{
  if(deliveryBusy||!realPreview||!$('rAck').checked||!mailReady)return;
  deliveryBusy=true;syncDelivery();
  const status=$('rDeliveryStatus');status.hidden=false;status.textContent='Mengirim email…';
  try{
    const fd=pdfForm();fd.append('file_type',$('rExportType').value);fd.append('reviewed','true');
    const response=await remoteApi('/api/send-email',fd);const result=await response.json();
    status.textContent=result.sent?'Email berhasil dikirim ke user yang login.':'Pengiriman belum terkonfirmasi.';
  }catch(e){status.textContent=e.message;}
  finally{deliveryBusy=false;syncDelivery();}
};
['rMargin','rFx'].forEach(id=>$(id).oninput=invalidateReal);
$('rFormat').onchange=()=>{invalidateReal();renderMasterStatus();updateDownloadLabel();};
async function init(){
  switchRealMode('simple');renderMasterStatus();updateDownloadLabel();
  try{
    const response=await fetch('/api/auth-config',{cache:'no-store'});
    if(!response.ok)throw Error('Login belum tersedia.');
    access=await response.json();
    showLogin(access.configured?'':'Login belum dikonfigurasi.');
  }catch(e){showLogin('Login belum tersedia. Coba muat ulang halaman.');}
  try{
    const status=await(await fetch('/api/status',{cache:'no-store'})).json();
    mailReady=Boolean(status.email_ready);
    updateDownloadLabel();
    $('statustag').textContent=status.real_processing_enabled?'Siap':'Terbatas';
  }catch(e){$('statustag').textContent='Offline';}
}
init();
