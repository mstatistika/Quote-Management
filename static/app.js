'use strict';
const $=id=>document.getElementById(id);
const sf=(v)=>'USD '+Number(v).toLocaleString('en-US',{minimumFractionDigits:2,maximumFractionDigits:2});
const rf=(v)=>'Rp'+Math.round(v).toLocaleString('id-ID');
const samples=[{sku:'DEMO-SKU-001',description:'Optical transponder sample',qty:2,cost:1000},{sku:'DEMO-SKU-002',description:'Pluggable module sample',qty:4,cost:250.5},{sku:'DEMO-SKU-003',description:'Chassis fan sample',qty:1,cost:85}];
let demoState={total:0,valid:false};
let access={configured:false,url:'',publishable_key:''};
let token=null,realPreview=null,realMode='simple',pinRotationRequired=false;
const validPin=value=>/^[0-9]{6}$/.test(value);
const weakPin=value=>validPin(value)&&(new Set(value).size===1||[...value].every((c,i,a)=>i===0||Number(c)===Number(a[i-1])+1)||[...value].every((c,i,a)=>i===0||Number(c)===Number(a[i-1])-1));
const metaNames=[['customer','Customer'],['attention','Attention'],['email','Email customer'],['project','Project'],['quote_ref','Quotation reference'],['date','Date'],['delivery','Delivery address'],['billing','Billing address'],['shipping','Shipping terms'],['payment','Payment terms'],['lead_time','Lead time'],['validity','Validity'],['remarks','Remarks']];
// One-page flow: choose a file first. Authentication is enforced by the server, never bypassed here.
function syncDemoBadge() { $('demoBadge').hidden=Boolean(token); }
syncDemoBadge();
function onFileSelected(){
  invalidateReal();
  const selected=$('rFile').files[0];
  const info=$('selectedFileInfo');
  $('realUnconfigured').classList.add('hidden');
  $('realLogin').classList.add('hidden');
  $('realWorkspace').classList.add('hidden');
  $('rPinSetup').classList.add('hidden');
  $('rError').textContent='';
  info.hidden=!selected;
  if(!selected) return;
  info.textContent='Dipilih: '+selected.name;
  if(!selected.name.toLowerCase().endsWith('.xlsx') || !selected.size || selected.size>8*1024*1024){
    info.textContent='Pilih file XLSX maksimal 8 MB.';
    return;
  }
  if(!access.configured) $('realUnconfigured').classList.remove('hidden');
  else if(token&&pinRotationRequired) $('rPinSetup').classList.remove('hidden');
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
function invalidateReal(){realPreview=null;$('rResults').hidden=true;$('rAck').checked=false;$('rGenerate').disabled=true;}
document.querySelectorAll('[data-r-mode]').forEach(b=>b.onclick=()=>switchRealMode(b.dataset.rMode));
for(const [key,label] of metaNames){const el=document.createElement('label');el.textContent=label;const input=document.createElement('input');input.id='meta_'+key;input.maxLength=190;if(key==='date'){input.type='date';}else if(key==='email'){input.type='email';}else{input.type='text';}input.oninput=invalidateReal;el.append(input);$('rAdvancedFields').append(el);}
function form(){if(!token)throw Error('Login terlebih dahulu.');const f=$('rFile').files[0];if(!f||!f.name.toLowerCase().endsWith('.xlsx'))throw Error('Upload original supplier .xlsx.');if(f.size>8*1024*1024)throw Error('File lebih dari 8 MB.');const fd=new FormData();fd.append('quote',f);fd.append('margin',$('rMargin').value);fd.append('fx',$('rFx').value);return fd;}
async function remoteApi(path,fd){const r=await fetch(path,{method:'POST',headers:{Authorization:'Bearer '+token},body:fd,cache:'no-store'});if(!r.ok){const data=await r.json().catch(()=>({detail:'Permintaan gagal.'}));if([401,403].includes(r.status))token=null;throw Error(data.detail||'Validasi gagal.');}return r;}
$('rLoginBtn').onclick=async()=>{
  const btn=$('rLoginBtn');btn.disabled=true;$('rLoginError').textContent='';
  try{
    if(!access.configured)throw Error('Login belum dikonfigurasi.');
    const email=$('rEmail').value.trim().toLowerCase();
    const pin=$('rPin').value.trim();
    if(!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)||!validPin(pin))throw Error('Isi email dan PIN 6 digit.');
    const r=await fetch(access.url+'/auth/v1/token?grant_type=password',{
      method:'POST',headers:{apikey:access.publishable_key,'Content-Type':'application/json'},
      body:JSON.stringify({email,password:pin}),cache:'no-store'
    });
    const data=await r.json().catch(()=>({}));
    if(!r.ok||!data.access_token)throw Error('Email atau PIN tidak valid, atau akun belum diaktifkan.');
    token=data.access_token;
    const authCheck=await fetch('/api/whoami',{headers:{Authorization:'Bearer '+token},cache:'no-store'});
    if(!authCheck.ok){
      token=null;throw Error('Akun tidak terdaftar atau belum mendapat izin.');
    }
    $('rPin').value='';
    pinRotationRequired=weakPin(pin);
    $('realLogin').classList.add('hidden');
    $('realWorkspace').classList.toggle('hidden',pinRotationRequired);
    $('rPinSetup').classList.toggle('hidden',!pinRotationRequired);
    syncDemoBadge();
  }catch(e){$('rLoginError').textContent=e.message;}
  finally{btn.disabled=false;}
};
$('rChangePinBtn').onclick=async()=>{
  const btn=$('rChangePinBtn');btn.disabled=true;$('rPinChangeError').textContent='';
  let passwordChanged=false;
  try{
    if(!token||!pinRotationRequired)throw Error('Sesi kedaluwarsa. Silakan masuk kembali.');
    const next=$('rNewPin').value.trim();
    const confirmation=$('rConfirmPin').value.trim();
    if(!validPin(next))throw Error('PIN baru wajib berisi tepat 6 angka.');
    if(weakPin(next))throw Error('PIN terlalu mudah ditebak. Hindari angka berurutan atau semuanya sama (misalnya 123456 atau 111111).');
    if(next!==confirmation)throw Error('Kedua PIN berbeda. Periksa kembali kolom konfirmasi.');
    btn.textContent='Menyimpan PIN…';
    const r=await fetch(access.url+'/auth/v1/user',{
      method:'PUT',headers:{apikey:access.publishable_key,Authorization:'Bearer '+token,'Content-Type':'application/json'},
      body:JSON.stringify({password:next}),cache:'no-store'
    });
    if(!r.ok){
      const response=await r.json().catch(()=>({}));
      throw Error(response.msg||response.message||response.error_description||'Gagal menyimpan PIN. Periksa kebijakan password Supabase.');
    }
    passwordChanged=true;
    // Refresh the session using the NEW PIN; never keep relying on a possibly revoked old session.
    const signin=await fetch(access.url+'/auth/v1/token?grant_type=password',{
      method:'POST',headers:{apikey:access.publishable_key,'Content-Type':'application/json'},
      body:JSON.stringify({email:$('rEmail').value.trim().toLowerCase(),password:next}),cache:'no-store'
    });
    const session=await signin.json().catch(()=>({}));
    if(!signin.ok||!session.access_token)throw Error('PIN sudah tersimpan, tetapi sesi baru belum aktif.');
    token=session.access_token;
    const check=await fetch('/api/whoami',{headers:{Authorization:'Bearer '+token},cache:'no-store'});
    if(!check.ok)throw Error('PIN sudah tersimpan, tetapi izin sesi perlu diverifikasi ulang.');
    pinRotationRequired=false;
    $('rNewPin').value='';$('rConfirmPin').value='';
    $('rPinSetup').classList.add('hidden');
    $('realWorkspace').classList.remove('hidden');
    $('rError').textContent='PIN baru tersimpan. Login berhasil. Lanjutkan pengaturan quotation.';
    syncDemoBadge();
    $('realWorkspace').scrollIntoView({behavior:'smooth',block:'start'});
  }catch(e){
    if(passwordChanged){
      token=null;pinRotationRequired=false;
      $('rNewPin').value='';$('rConfirmPin').value='';
      $('rPinSetup').classList.add('hidden');
      $('realWorkspace').classList.add('hidden');
      $('realLogin').classList.remove('hidden');
      $('rPin').value='';
      syncDemoBadge();
      $('rLoginError').textContent='PIN baru berhasil disimpan. Silakan masuk menggunakan PIN baru.';
      $('realLogin').scrollIntoView({behavior:'smooth',block:'start'});
    }else{
      $('rPinChangeError').textContent=e.message;
    }
  }finally{
    btn.disabled=false;btn.textContent='Simpan PIN baru';
  }
};
$('rLogout').onclick=()=>{
  token=null;pinRotationRequired=false;
  $('rNewPin').value='';$('rConfirmPin').value='';
  syncDemoBadge();onFileSelected();
};
$('rPreview').onclick=async()=>{const btn=$('rPreview');btn.disabled=true;$('rError').textContent='';try{const r=await remoteApi('/api/preview',form());const p=await r.json();realPreview=p;$('rUsd').textContent=sf(p.total_usd);$('rIdr').textContent=rf(p.total_idr);$('rResultInfo').textContent=`${p.item_count} item · margin aktual ${p.actual_gm}% · validasi PASS`;$('rItems').replaceChildren(...p.items.map(i=>itemElement({sku:i.sku,description:i.description,qty:i.qty,unit:i.unit_usd,total:i.total_usd})));$('rResults').hidden=false;$('rAck').checked=false;$('rGenerate').disabled=true;}catch(e){invalidateReal();$('rError').textContent=e.message;}finally{btn.disabled=false;}};
$('rAck').onchange=()=>{$('rGenerate').disabled=!$('rAck').checked||!realPreview;};
$('rGenerate').onclick=async()=>{const btn=$('rGenerate');btn.disabled=true;$('rError').textContent='';try{if(!realPreview||!$('rAck').checked)throw Error('Preview dan review harus dilakukan terlebih dahulu.');const fd=form();fd.append('format',$('rFormat').value);if($('rT01').files[0])fd.append('t01_master',$('rT01').files[0]);const meta={};if(realMode==='advanced')for(const [key] of metaNames){const v=$('meta_'+key).value.trim();if(v)meta[key]=v;}fd.append('details',JSON.stringify(meta));const r=await remoteApi('/api/generate',fd);const mime=r.headers.get('Content-Type')||'';const filename=(r.headers.get('Content-Disposition')||'').match(/filename="([^"]+)"/)?.[1]||'QUOTE_DRAFT.'+(mime.includes('zip')?'zip':'xlsx');download(await r.blob(),filename);$('rError').textContent='Download DRAFT berhasil. Review manual sebelum dikirim ke customer.';}catch(e){$('rError').textContent=e.message;}finally{btn.disabled=false;}};
$('rT01').onchange=invalidateReal;['rMargin','rFx','rFormat'].forEach(id=>$(id).oninput=invalidateReal);
async function init(){updateDemo();document.querySelector('[data-mode="simple"]').click();switchRealMode('simple');
try{const status=await(await fetch('/api/status')).json();$('statustag').textContent=status.demo_ready?'Demo siap':'Preview';const cfg=await(await fetch('/api/auth-config')).json();access=cfg;onFileSelected();}catch(e){$('statustag').textContent='Offline';onFileSelected();}}
init();
