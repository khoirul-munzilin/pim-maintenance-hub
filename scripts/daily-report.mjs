import { createClient } from '@supabase/supabase-js';
import ExcelJS from 'exceljs';
import sharp from 'sharp';
import { Resend } from 'resend';
const required=['SUPABASE_URL','SUPABASE_SECRET_KEY','RESEND_API_KEY','REPORT_EMAIL_FROM','REPORT_EMAIL_TO'];
for(const k of required) if(!process.env[k]) throw new Error(`Secret ${k} belum diisi`);
const supabase=createClient(process.env.SUPABASE_URL,process.env.SUPABASE_SECRET_KEY,{auth:{persistSession:false}});
const resend=new Resend(process.env.RESEND_API_KEY);
const now=new Date(); const end=new Date(now); end.setUTCHours(0,30,0,0); const start=new Date(end);start.setUTCDate(start.getUTCDate()-1);
const {data:reports,error}=await supabase.from('work_reports').select(`*,functional_locations(funloc_code,description,area),creator:profiles!work_reports_created_by_fkey(full_name),assignee:profiles!work_reports_assigned_to_fkey(full_name),report_photos(*)`).gte('created_at',start.toISOString()).lt('created_at',end.toISOString()).order('created_at');
if(error) throw error;
const wb=new ExcelJS.Workbook();wb.creator='PIM Maintenance Hub';wb.created=now;
const ws=wb.addWorksheet('Maintenance Report',{views:[{state:'frozen',ySplit:1,showGridLines:false}]});
const headers=['No.','Nomor Laporan','Tanggal','Jenis','Functional Location','Area','Prioritas','Deskripsi / Temuan','Dampak Operasional','Status','Pelapor','Teknisi','Penyebab','Tindakan Perbaikan','Spare Part','Catatan Teknisi','Hasil Test','Catatan Verifikasi','Waktu Mulai','Request Verification','Waktu Closed'];
ws.columns=headers.map((h,i)=>({header:h,key:`c${i}`,width:[7,23,22,12,42,18,13,38,30,15,20,22,30,38,25,32,17,32,21,22,21][i]}));
ws.getRow(1).height=32;ws.getRow(1).eachCell(c=>{c.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF163A5F'}};c.font={bold:true,color:{argb:'FFFFFFFF'}};c.alignment={horizontal:'center',vertical:'middle',wrapText:true}});ws.autoFilter={from:'A1',to:'U1'};
const fmt=d=>d?new Date(d).toLocaleString('id-ID',{timeZone:'Asia/Jakarta'}):'';
for(let i=0;i<reports.length;i++){const r=reports[i],row=ws.addRow([i+1,r.report_no,fmt(r.created_at),r.report_type,r.functional_locations?.funloc_code||'',r.functional_locations?.area||'',r.priority,r.description,r.operational_impact||'',r.status,r.creator?.full_name||'',r.assignee?.full_name||'',r.failure_cause||'',r.corrective_action||r.recommendation||'',r.spare_part||'',r.technician_note||'',r.test_result||'',r.verification_note||'',fmt(r.started_at),fmt(r.verification_requested_at),fmt(r.closed_at)]);row.height=45;row.eachCell(c=>c.alignment={vertical:'top',wrapText:true});const st=row.getCell(10);st.font={bold:true,color:{argb:r.status==='CLOSED'?'FF166534':'FF92400E'}};st.fill={type:'pattern',pattern:'solid',fgColor:{argb:r.status==='CLOSED'?'FFDCFCE7':'FFFEF3C7'}}}
const doc=wb.addWorksheet('Dokumentasi',{views:[{showGridLines:false}]});doc.columns=[{width:34},{width:4},{width:34},{width:4},{width:34}];
async function imageBuffer(path){const {data,error}=await supabase.storage.from('maintenance-photos').download(path);if(error)return null;return sharp(Buffer.from(await data.arrayBuffer())).rotate().resize({width:720,height:480,fit:'inside',withoutEnlargement:true}).jpeg({quality:65}).toBuffer()}
for(let i=0;i<reports.length;i++){const r=reports[i],base=i*22+1;doc.mergeCells(base,1,base,5);const title=doc.getCell(base,1);title.value=`${r.report_no} • ${r.functional_locations?.funloc_code||''}`;title.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF163A5F'}};title.font={bold:true,color:{argb:'FFFFFFFF'},size:12};title.alignment={vertical:'middle'};doc.getRow(base).height=25;const groups=[['initial','Kondisi Awal',1],['result','Hasil Pekerjaan',3],['verification','Verifikasi',5]];
 for(const [type,label,col] of groups){const h=doc.getCell(base+1,col);h.value=label;h.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF0F766E'}};h.font={bold:true,color:{argb:'FFFFFFFF'}};h.alignment={horizontal:'center'};for(let rr=base+2;rr<=base+20;rr++)doc.getRow(rr).height=18;const photos=(r.report_photos||[]).filter(p=>p.photo_type===type).slice(0,1);if(!photos.length){doc.getCell(base+10,col).value='Foto belum tersedia';doc.getCell(base+10,col).alignment={horizontal:'center'};continue}const buf=await imageBuffer(photos[0].storage_path);if(buf){const id=wb.addImage({buffer:buf,extension:'jpeg'});doc.addImage(id,{tl:{col:col-1,row:base+1},br:{col:col,row:base+20},editAs:'oneCell'})}}
}
const date=new Date(start).toLocaleDateString('en-CA',{timeZone:'Asia/Jakarta'}),file=`PIM_PM_CM_Report_${date}.xlsx`,buffer=await wb.xlsx.writeBuffer();
const closed=reports.filter(r=>r.status==='CLOSED').length,pm=reports.filter(r=>r.report_type==='PM').length,cm=reports.filter(r=>r.report_type==='CM').length;
await resend.emails.send({from:process.env.REPORT_EMAIL_FROM,to:process.env.REPORT_EMAIL_TO.split(',').map(x=>x.trim()),subject:`PIM PM & CM Report - ${date}`,html:`<p>Berikut laporan PM dan CM periode 07.30 WIB.</p><ul><li>Total: ${reports.length}</li><li>PM: ${pm}</li><li>CM: ${cm}</li><li>Closed: ${closed}</li></ul>`,attachments:[{filename:file,content:Buffer.from(buffer)}]});
console.log(`Terkirim: ${file}, ${reports.length} laporan`);
