/* PIM Maintenance Hub - EIPA Module */
const EIPA_GENERIC_FUNLOC = 'ID-MJK-EIPA-GENERAL';
let eipaRows = [];
let eipaEvidenceZip = null;

function eipaAllowed(){ return ['admin','supervisor'].includes(profile?.role); }
function eipaText(v){ return v === null || v === undefined ? '' : String(v).trim(); }
function eipaDate(v){
  if(!v) return null;
  if(v instanceof Date) return v.toISOString().slice(0,10);
  if(typeof v === 'number' && window.ExcelJS){ const d=new Date(Math.round((v-25569)*86400*1000)); return d.toISOString().slice(0,10); }
  const d=new Date(v); return Number.isNaN(d.getTime()) ? null : d.toISOString().slice(0,10);
}
function setEipaMessage(t, type='info'){
  const box=document.getElementById('eipaMessage'); if(!box)return;
  box.className='eipa-message '+type; box.textContent=t;
}
async function loadEipaDashboard(){
  const {data,error}=await db.from('work_reports').select('id,report_no,status,priority,description,recommendation,eipa_progress,eipa_target_date,eipa_area,eipa_pic_name,created_at').eq('source_type','EIPA').order('report_no');
  if(error){ setEipaMessage('Gagal membaca data EIPA: '+error.message,'error'); return; }
  eipaRows=data||[];
  const closed=eipaRows.filter(x=>['EVALUATED','CLOSED'].includes(eipaText(x.status).toUpperCase())).length;
  const overdue=eipaRows.filter(x=>x.eipa_target_date && new Date(x.eipa_target_date)<new Date() && !['EVALUATED','CLOSED'].includes(eipaText(x.status).toUpperCase())).length;
  document.getElementById('eipaTotal').textContent=eipaRows.length;
  document.getElementById('eipaClosed').textContent=closed;
  document.getElementById('eipaOpen').textContent=eipaRows.length-closed;
  document.getElementById('eipaOverdue').textContent=overdue;
  renderEipaTable();
}
function renderEipaTable(){
  const q=eipaText(document.getElementById('eipaSearch')?.value).toLowerCase();
  const st=eipaText(document.getElementById('eipaStatus')?.value).toUpperCase();
  const rows=eipaRows.filter(x=>(!q||[x.report_no,x.description,x.eipa_area,x.eipa_pic_name].join(' ').toLowerCase().includes(q))&&(!st||eipaText(x.status).toUpperCase()===st));
  const body=document.getElementById('eipaTable');
  if(!rows.length){ body.innerHTML='<tr><td colspan="8" class="empty">Belum ada data EIPA.</td></tr>'; return; }
  body.innerHTML=rows.map(x=>`<tr><td><b>${eipaText(x.report_no)}</b></td><td>${eipaText(x.eipa_area)||'-'}</td><td>${eipaText(x.description)}</td><td>${eipaText(x.priority)}</td><td>${eipaText(x.eipa_pic_name)||'-'}</td><td>${x.eipa_target_date||'-'}</td><td><span class="badge">${eipaText(x.status)}</span></td><td><button class="btn" onclick="openEipa('${x.id}')">Detail</button></td></tr>`).join('');
}
async function openEipa(id){
  const x=eipaRows.find(r=>r.id===id); if(!x)return;
  const {data:photos}=await db.from('report_photos').select('*').eq('report_id',id).order('created_at');
  let photoHtml='';
  for(const p of photos||[]){ const {data}=await db.storage.from('maintenance-photos').createSignedUrl(p.storage_path,3600); if(data?.signedUrl)photoHtml+=`<a href="${data.signedUrl}" target="_blank"><img src="${data.signedUrl}" class="eipa-thumb" alt="Evidence"></a>`; }
  document.getElementById('mt').textContent=x.report_no;
  document.getElementById('mb').innerHTML=`<div class="eipa-detail"><b>Area</b><p>${eipaText(x.eipa_area)||'-'}</p><b>Temuan</b><p>${eipaText(x.description)}</p><b>Rekomendasi</b><p>${eipaText(x.recommendation)||'-'}</p><b>PIC</b><p>${eipaText(x.eipa_pic_name)||'-'}</p><b>Target</b><p>${x.eipa_target_date||'-'}</p><b>Status</b><p>${eipaText(x.status)}</p><b>Evidence</b><div class="eipa-gallery">${photoHtml||'Belum ada foto'}</div></div>`;
  document.getElementById('modal').classList.remove('hide');
}
async function previewEipaExcel(input){
  const file=input.files?.[0]; if(!file)return;
  try{
    const wb=new ExcelJS.Workbook(); await wb.xlsx.load(await file.arrayBuffer());
    const ws=wb.getWorksheet('CM_Import'); if(!ws)throw new Error('Sheet CM_Import tidak ditemukan.');
    const headers={}; ws.getRow(1).eachCell((c,i)=>headers[eipaText(c.value)]=i);
    const required=['Record_ID','Finding_No','Finding_Description','Recommended_Action','Update_Status','Progress'];
    const missing=required.filter(h=>!headers[h]); if(missing.length)throw new Error('Kolom wajib tidak ada: '+missing.join(', '));
    eipaRows=[];
    ws.eachRow((row,n)=>{ if(n===1)return; const get=h=>row.getCell(headers[h]||999).value; const id=eipaText(get('Record_ID')); if(!id)return;
      eipaRows.push({report_no:id,finding_no:Number(get('Finding_No'))||null,area:eipaText(get('Area_Unit')),category:eipaText(get('Sub_Category')),description:eipaText(get('Finding_Description')),recommendation:eipaText(get('Recommended_Action')),status:eipaText(get('Update_Status'))||'NOT STARTED',progress:Number(get('Progress'))||0,target:eipaDate(get('Target_Date')),pic:eipaText(get('PIC')),priority:eipaText(get('Priority'))||'Medium',pdf_page:eipaText(get('Source_PDF_Page')),reference:eipaText(get('Reference_Title'))});
    });
    document.getElementById('eipaPreviewCount').textContent=eipaRows.length;
    document.getElementById('eipaImportButton').disabled=!eipaRows.length;
    document.getElementById('eipaPreview').innerHTML=eipaRows.slice(0,8).map(x=>`<tr><td>${x.report_no}</td><td>${x.area}</td><td>${x.description}</td><td>${x.status}</td></tr>`).join('');
    setEipaMessage(`${eipaRows.length} baris siap divalidasi dan diimpor.`,'success');
  }catch(e){ eipaRows=[]; setEipaMessage(e.message,'error'); }
}
async function getEipaFunloc(){
  const {data,error}=await db.from('functional_locations').select('id').eq('funloc_code',EIPA_GENERIC_FUNLOC).single();
  if(error||!data)throw new Error('Functional Location EIPA belum tersedia. Jalankan eipa-migration.sql di Supabase.'); return data.id;
}
async function importEipa(){
  if(!eipaAllowed()){setEipaMessage('Hanya Admin atau Supervisor yang dapat mengimpor.','error');return;}
  if(!eipaRows.length)return;
  const button=document.getElementById('eipaImportButton'); button.disabled=true;
  try{
    const funloc=await getEipaFunloc(); const {data:{user}}=await db.auth.getUser(); let ok=0,fail=0;
    for(const x of eipaRows){
      const payload={report_no:x.report_no,report_type:'CM',funloc_id:funloc,priority:x.priority,description:x.description,recommendation:x.recommendation,status:x.status,created_by:user.id,source_type:'EIPA',source_reference:x.report_no,eipa_finding_no:x.finding_no,eipa_area:x.area,eipa_category:x.category,eipa_progress:x.progress,eipa_target_date:x.target,eipa_pic_name:x.pic,eipa_pdf_page:x.pdf_page,eipa_reference_title:x.reference};
      const {error}=await db.from('work_reports').upsert(payload,{onConflict:'report_no'}); if(error){console.error(x.report_no,error);fail++;}else ok++;
    }
    setEipaMessage(`Import selesai. Berhasil ${ok}, gagal ${fail}.` ,fail?'error':'success'); await loadEipaDashboard();
  }catch(e){setEipaMessage(e.message,'error');}finally{button.disabled=false;}
}
async function readEipaZip(input){
  const file=input.files?.[0]; if(!file)return;
  if(!window.JSZip){setEipaMessage('Pustaka ZIP belum termuat. Muat ulang halaman.','error');return;}
  try{eipaEvidenceZip=await JSZip.loadAsync(file); const names=Object.keys(eipaEvidenceZip.files).filter(n=>!eipaEvidenceZip.files[n].dir); document.getElementById('eipaZipCount').textContent=names.length; document.getElementById('eipaUploadButton').disabled=!names.length; setEipaMessage(`${names.length} foto ditemukan dalam ZIP.`,'success');}catch(e){setEipaMessage('ZIP tidak dapat dibaca: '+e.message,'error');}
}
async function uploadEipaEvidence(){
  if(!eipaAllowed()||!eipaEvidenceZip)return;
  const button=document.getElementById('eipaUploadButton'); button.disabled=true;
  try{
    const {data:{user}}=await db.auth.getUser(); const {data:reports,error}=await db.from('work_reports').select('id,eipa_finding_no,report_no').eq('source_type','EIPA'); if(error)throw error;
    const byNo=new Map((reports||[]).map(r=>[Number(r.eipa_finding_no),r])); let ok=0,skip=0;
    for(const [name,entry] of Object.entries(eipaEvidenceZip.files)){
      if(entry.dir)continue; const m=name.match(/(?:^|\/)(\d{1,2})(?:\/|$)|EIPA-(\d{3})/i); const no=Number(m?.[1]||m?.[2]); const report=byNo.get(no); if(!report){skip++;continue;}
      const blob=await entry.async('blob'); if(!blob.type.startsWith('image/')&&!/\.(jpg|jpeg|png|webp)$/i.test(name)){skip++;continue;}
      const ext=(name.split('.').pop()||'jpg').toLowerCase(); const path=`${user.id}/eipa/${report.report_no}/${Date.now()}-${ok}.${ext}`;
      const {error:upErr}=await db.storage.from('maintenance-photos').upload(path,blob,{upsert:false,contentType:blob.type||'image/jpeg'}); if(upErr){console.error(upErr);skip++;continue;}
      const {error:metaErr}=await db.from('report_photos').insert({report_id:report.id,storage_path:path,photo_type:'initial',uploaded_by:user.id}); if(metaErr){console.error(metaErr);skip++;continue;} ok++;
    }
    setEipaMessage(`Upload evidence selesai. Berhasil ${ok}, dilewati/gagal ${skip}.`,skip?'info':'success');
  }catch(e){setEipaMessage(e.message,'error');}finally{button.disabled=false;}
}
window.loadEipaDashboard=loadEipaDashboard; window.renderEipaTable=renderEipaTable; window.previewEipaExcel=previewEipaExcel; window.importEipa=importEipa; window.readEipaZip=readEipaZip; window.uploadEipaEvidence=uploadEipaEvidence; window.openEipa=openEipa;
