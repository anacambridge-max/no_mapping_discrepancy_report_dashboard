"use client";

import React,{useMemo,useState} from "react";
import * as XLSX from "xlsx";
import jsPDF from "jspdf";
import autoTable from "jspdf-autotable";
import {Download,Upload,FileText,CheckCircle2,AlertTriangle,ArrowUpDown} from "lucide-react";
import {RawRow,ParsedReport,OfficerReportRow,buildOfficerRows,formatNumber,formatOfficer,grandOfficer,parseTimestamp,validateColumns,normalizeRows,parseGrandTotal,validateGrandTotal,percentileGroups} from "../lib/report";

const COLS=[
 "S.No.","OFFICER NAME AND DESIGNATION","DESIG.","NO. OF PS","TOTAL NOTICES",
 "TOTAL NO MAPPING NOTICES","NO MAPPING DISPOSED EARLIER","NO MAPPING DISPOSED LATEST","DIFFERENCE",
 "% OF NO MAPPING NOTICES DISPOSED","NO MAPPING DOCUMENTS UPLOADED","PENDING NO MAPPING NOTICES","",
 "TOTAL DISCREPANCY NOTICES","DISCREPANCY DISPOSED EARLIER","DISCREPANCY DISPOSED LATEST","DIFFERENCE",
 "% OF DISCREPANCY NOTICES DISPOSED","BLO LETTER UPLOADED","PENDING DISCREPANCY NOTICES","",
 "TOTAL NOTICES DISPOSED","% OF TOTAL NOTICES DISPOSED","TOTAL NOTICES PENDING"
];

function rowValues(r:OfficerReportRow,rowIndex:number){
 return [
  rowIndex+1,r.officer,r.designation||"Officer",r.psCount,r.totalNotices,r.nmTotal,r.nmEarlier,r.nmLatest,r.nmDifference,
  r.nmPct+"%",r.nmDocs,r.nmPending,"",r.dTotal,r.dEarlier,r.dLatest,r.dDifference,r.dPct+"%",r.dLetters,r.dPending,"",
  r.totalDisposed,r.totalDisposedPct+"%",r.totalPending
 ];
}

const color=(kind:string):[number,number,number]=>kind==="low"?[248,180,180]:kind==="medium"?[255,230,153]:[183,225,161];

export default function Page(){
 const [earlier,setEarlier]=useState<ParsedReport|null>(null);
 const [latest,setLatest]=useState<ParsedReport|null>(null);
 const [manualEarlier,setManualEarlier]=useState("");
 const [manualLatest,setManualLatest]=useState("");
 const [error,setError]=useState("");
 const [sortKey,setSortKey]=useState<keyof OfficerReportRow>("nmDifference");
 const [sortDir,setSortDir]=useState(-1);

 const load=async(file:File,which:"earlier"|"latest",manual:string)=>{
  setError("");
  try{
   const wb=XLSX.read(await file.arrayBuffer(),{type:"array",cellDates:true});
   const sheet=wb.Sheets["FINAL PS WISE REPORT"];
   if(!sheet) throw new Error('Required sheet "FINAL PS WISE REPORT" was not found.');
   const raw=XLSX.utils.sheet_to_json<RawRow>(sheet,{defval:""});
   const missing=validateColumns(raw);
   if(missing.length) throw new Error("Missing required column(s): "+missing.join(", "));
   const grand=parseGrandTotal(raw);
   const rows=normalizeRows(raw);
   const validation=validateGrandTotal(rows,grand);
   const parsed:ParsedReport={
    fileName:file.name,timestamp:parseTimestamp(file.name,file.lastModified,manual),
    rows,grandTotal:grand,validation
   };
   if(which==="earlier")setEarlier(parsed);else setLatest(parsed);
  }catch(e:any){setError(e?.message||"Unable to read Excel file.");}
 };

 const reports=useMemo(()=>buildOfficerRows(latest?.rows||[],earlier?.rows||[]),[latest,earlier]);
 const rows=useMemo(()=>[...reports].sort((a:any,b:any)=>{
  if(sortKey==="officer"||sortKey==="designation") return String(a[sortKey]).localeCompare(String(b[sortKey]))*sortDir;
  return ((Number(a[sortKey])||0)-(Number(b[sortKey])||0))*sortDir;
 }),[reports,sortKey,sortDir]);
 const total=grandOfficer(rows);
 const psCount=new Set((latest?.rows||[]).map(r=>r.psNo)).size;
 const groups={
  nm:useMemo(()=>percentileGroups(rows,r=>r.nmDifference),[rows]),
  docs:useMemo(()=>percentileGroups(rows,r=>r.nmDocsPct),[rows]),
  disc:useMemo(()=>percentileGroups(rows,r=>r.dDifference),[rows]),
  letters:useMemo(()=>percentileGroups(rows,r=>r.dLettersPct),[rows]),
  totalDisposed:useMemo(()=>percentileGroups(rows,r=>r.totalDisposedPct),[rows])
 };

 const sort=(key:keyof OfficerReportRow)=>{
  if(sortKey===key)setSortDir(v=>-v);else{setSortKey(key);setSortDir(-1);}
 };
 const percentage=(a:number,b:number)=>b?Math.round((a/b)*1000)/10:0;

 const buildOfficerPSRows=(officer:string)=>{
  const L=(latest?.rows||[]).filter(r=>r.officer===officer).sort((a,b)=>String(a.psNo).localeCompare(String(b.psNo),undefined,{numeric:true}));
  const E=new Map((earlier?.rows||[]).filter(r=>r.officer===officer).map(r=>[r.psNo,r]));
  return L.map((r,i)=>{
   const e=E.get(r.psNo);
   const totalDisposed=r.nmDisposed+r.dDisposed;
   return [i+1,r.psNo,r.blo,r.supervisor,r.nmTotal+r.dTotal,r.nmTotal,e?.nmDisposed||0,r.nmDisposed,r.nmDisposed-(e?.nmDisposed||0),r.nmDocs,r.dTotal,e?.dDisposed||0,r.dDisposed,r.dDisposed-(e?.dDisposed||0),r.dLetter,totalDisposed,percentage(totalDisposed,r.nmTotal+r.dTotal)+"%"];
  });
 };

 const pdf=()=>{
  if(!latest||!earlier)return;
  const doc=new jsPDF({orientation:"landscape",unit:"mm",format:"a3"});
  const navy:[number,number,number]=[31,56,100],brown:[number,number,number]=[131,60,11],yellow:[number,number,number]=[255,248,225],grid:[number,number,number]=[184,192,204],grand:[number,number,number]=[217,225,242];

  doc.setFillColor(...navy);doc.rect(0,0,420,18,"F");
  doc.setTextColor(255,255,255);doc.setFont("helvetica","bold");
  doc.setFontSize(16);doc.text("OFFICE OF THE ELECTORAL REGISTRATION OFFICER, AC-34, MATIALA",210,7,{align:"center"});
  doc.setFontSize(14);doc.text("SIR-2026: NOTICE DISPOSAL REPORT",210,12,{align:"center"});
  doc.setFontSize(8);doc.text("NO MAPPING & DISCREPANCY NOTICES - OFFICER-WISE (AERO-WISE)",210,16,{align:"center"});
  doc.setTextColor(55,65,80);doc.setFontSize(7.5);
  doc.text("Data as of "+latest.timestamp+" | "+rows.length+" Officers | "+psCount+" Polling Stations",210,23,{align:"center"});

  const cards:Array<[string,string,[number,number,number]]>=[
    [String(psCount),"TOTAL PS",navy],[formatNumber(total.totalNotices),"TOTAL NOTICES",navy],
    [formatNumber(total.nmTotal),"NO MAPPING NOTICES",navy],[formatNumber(total.nmLatest),"NO MAPPING DISPOSED",navy],
    [formatNumber(total.nmDocs),"NO MAPPING DOCS UPLOADED",navy],[formatNumber(total.dTotal),"DISCREPANCY NOTICES",brown],
    [formatNumber(total.dLatest),"DISCREPANCY DISPOSED",brown],[formatNumber(total.dLetters),"BLO LETTERS UPLOADED",brown],
    [total.nmPct+"%","NM % DISPOSED",navy],[total.dPct+"%","DISC. % DISPOSED",navy]
  ];
  const cardW=39.2,gap=2,cardY=27,cardH=18,left=10;
  cards.forEach((c,i)=>{const x=left+i*(cardW+gap);doc.setFillColor(...c[2] as [number,number,number]);doc.roundedRect(x,cardY,cardW,cardH,1,1,"F");doc.setTextColor(255,255,255);doc.setFont("helvetica","bold");doc.setFontSize(11);doc.text(c[0],x+cardW/2,cardY+7,{align:"center"});doc.setFontSize(5.1);doc.text(c[1],x+cardW/2,cardY+13,{align:"center"});});

  const head1=[
    {content:"S.No.",rowSpan:2},{content:"OFFICER NAME AND DESIGNATION",rowSpan:2},{content:"NO. OF PS",rowSpan:2},{content:"TOTAL\nNOTICES",rowSpan:2},
    {content:"NO MAPPING NOTICES",colSpan:7},{content:"",rowSpan:2},
    {content:"DISCREPANCY NOTICES",colSpan:7},{content:"",rowSpan:2},{content:"TOTAL DISPOSAL",colSpan:3}
  ];
  const head2=[
    "TOTAL NO MAPPING","NM DISP. EARLIER","NM DISP. LATEST","NM DIFF.","% NM DISPOSED","NM DOCS UPLOADED","PENDING NM",
    "TOTAL DISCREPANCY","DISC. DISP. EARLIER","DISC. DISP. LATEST","DISC. DIFF.","% DISC. DISPOSED","BLO LETTER UPLOADED","PENDING DISC.",
    "TOTAL NOTICES DISPOSED","% TOTAL DISPOSED","TOTAL NOTICES PENDING"
  ];
  const body=rows.map((r,i)=>[
    i+1,r.officer+"\n("+(r.designation||"Officer")+")",r.psCount,r.totalNotices,
    r.nmTotal,r.nmEarlier,r.nmLatest,r.nmDifference,r.nmPct+"%",r.nmDocs,r.nmPending,
    "",r.dTotal,r.dEarlier,r.dLatest,r.dDifference,r.dPct+"%",r.dLetters,r.dPending,
    "",r.totalDisposed,r.totalDisposedPct+"%",r.totalPending
  ]);
  body.push(["","GRAND TOTAL - AC-34 MATIALA",psCount,total.totalNotices,total.nmTotal,total.nmEarlier,total.nmLatest,total.nmDifference,total.nmPct+"%",total.nmDocs,total.nmPending,"",total.dTotal,total.dEarlier,total.dLatest,total.dDifference,total.dPct+"%",total.dLetters,total.dPending,"",total.totalDisposed,total.totalDisposedPct+"%",total.totalPending]);

  const widths=[7,45,11,18,22,18,18,15,18,21,19,3,22,19,19,15,18,22,19,3,22,20,23];
  const base={theme:"grid" as const,tableWidth:400,margin:{left:10,right:10,top:0,bottom:0},styles:{font:"helvetica",fontSize:5.7,cellPadding:.65,lineColor:grid,lineWidth:.25,textColor:[20,28,38] as [number,number,number],halign:"center" as const,valign:"middle" as const,overflow:"linebreak" as const},headStyles:{font:"helvetica",fontStyle:"bold" as const,fontSize:5.0,cellPadding:.7,minCellHeight:8,halign:"center" as const,valign:"middle" as const,textColor:255}};
  autoTable(doc,{...base,startY:49,head:[head1,head2],body,columnStyles:Object.fromEntries(widths.map((w,i)=>[i,{cellWidth:w,halign:i===1?"left":"center"}])),didParseCell:(data:any)=>{
    if(data.section==="head"){data.cell.styles.fillColor=data.column.index>=12&&data.column.index<=18?brown:navy;if(data.column.index===11||data.column.index===19){data.cell.styles.fillColor=[255,255,255];data.cell.styles.lineWidth=0;}return;}
    if(data.section!=="body")return;
    if(data.row.index===rows.length){data.cell.styles.fillColor=grand;data.cell.styles.fontStyle="bold";return;}
    const r=rows[data.row.index];if(!r)return;
    if(data.column.index===4||data.column.index===12)data.cell.styles.fillColor=yellow;
    if(data.column.index===7||data.column.index===8)data.cell.styles.fillColor=color(groups.nm.get(r.officer)||"medium");
    if(data.column.index===9)data.cell.styles.fillColor=color(groups.docs.get(r.officer)||"medium");
    if(data.column.index===15||data.column.index===16)data.cell.styles.fillColor=color(groups.disc.get(r.officer)||"medium");
    if(data.column.index===17)data.cell.styles.fillColor=color(groups.letters.get(r.officer)||"medium");
    if(data.column.index===21)data.cell.styles.fillColor=color(groups.totalDisposed.get(r.officer)||"medium");
  }});
  const fy=281;doc.setFont("helvetica","bold");doc.setFontSize(7);doc.setTextColor(40,48,58);doc.text("COLOUR CODE:",10,fy);
  let x=35;[["LOW","low"],["MEDIUM","medium"],["HIGH","high"]].forEach(([label,k])=>{const c=color(k);doc.setFillColor(c[0],c[1],c[2]);doc.rect(x,fy-3.5,23,4,"F");doc.setTextColor(40,48,58);doc.text(label,x+25,fy);x+=50;});
  doc.setFont("helvetica","normal");doc.setFontSize(6.5);doc.text("Colour coding applies to Difference and percentage columns. Lowest third = red; highest third = green.",185,fy);
  doc.text("Pending NM = Total No Mapping Notices - Latest NM Disposed | Pending Disc. = Total Discrepancy Notices - Latest Disposed | Total Pending = Pending NM + Pending Disc.",10,fy+7);
  doc.setFontSize(6);doc.text("AC-34 MATIALA | SIR-2026 | OFFICER-WISE NOTICE DISPOSAL REPORT | Data as of "+latest.timestamp,10,292);doc.text("Page 1",410,292,{align:"right"});
  doc.save("Officer_Wise_Report_"+new Date().toISOString().slice(0,10)+".pdf");
 };


 const excel=()=>{
  if(!latest||!earlier)return;
  const data=rows.map((r,i)=>({
   "S.No.":i+1,"OFFICER NAME AND DESIGNATION":formatOfficer(r),"TOTAL NOTICES":r.totalNotices,
   "TOTAL NO MAPPING NOTICES":r.nmTotal,"NO MAPPING DISPOSED EARLIER":r.nmEarlier,
   "NO MAPPING DISPOSED LATEST":r.nmLatest,"DIFFERENCE":r.nmDifference,
   "% OF NO MAPPING NOTICES DISPOSED":r.nmPct+"%","NO MAPPING DOCUMENTS UPLOADED":r.nmDocs,
   "% DOCUMENTS / TOTAL NO MAPPING NOTICES":r.nmDocsPct+"%","PENDING NO MAPPING NOTICES":r.nmPending,"TOTAL DISCREPANCY NOTICES":r.dTotal,
   "DISCREPANCY DISPOSED EARLIER":r.dEarlier,"DISCREPANCY DISPOSED LATEST":r.dLatest,
   "DIFFERENCE (DISCREPANCY)":r.dDifference,"% OF DISCREPANCY NOTICES DISPOSED":r.dPct+"%",
   "BLO LETTER UPLOADED":r.dLetters,"% BLO LETTER / TOTAL DISCREPANCY NOTICES":r.dLettersPct+"%","PENDING DISCREPANCY NOTICES":r.dPending,
   "TOTAL NOTICES DISPOSED":r.totalDisposed,"% OF TOTAL NOTICES DISPOSED":r.totalDisposedPct+"%","TOTAL NOTICES PENDING":r.totalPending
  }));
  const ws=XLSX.utils.json_to_sheet(data);
  ws["!freeze"]={xSplit:0,ySplit:1};
  const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,"Officer Wise");
  XLSX.writeFile(wb,`Officer_Wise_Report_${new Date().toISOString().slice(0,10)}.xlsx`);
 };

 const fileBox=(which:"earlier"|"latest",value:ParsedReport|null,manual:string,setManual:(v:string)=>void)=>
  <div className="upload">
   <div className="uploadtitle"><b>{which==="earlier"?"1. EARLIER REPORT":"2. LATEST REPORT"}</b><span>{value?.fileName||"Upload Excel report"}</span></div>
   <div className="drop"><Upload size={17}/><b>Choose .xlsx / .xls</b><input type="file" accept=".xlsx,.xls" onChange={e=>{const f=e.target.files?.[0];if(f)load(f,which,manual)}}/></div>
   <input className="manual" value={manual} onChange={e=>setManual(e.target.value)} placeholder="Optional date/time"/>
   {value&&<div className="ok"><CheckCircle2 size={14}/>{value.timestamp}{value.validation.matched&&<strong> GRAND TOTAL VALIDATED</strong>}</div>}
  </div>;

 return <main>
  <div className="bar"/>
  <header>
   <h1>OFFICER WISE PROGRESS REPORT</h1>
   <div>No Mapping &amp; Discrepancy Notices - Earlier vs Latest Disposal Status</div>
   <small>Data as of {latest?.timestamp||"—"} | {rows.length||17} Officers | {psCount||430} Polling Stations</small>
  </header>
  <section className="uploads">{fileBox("earlier",earlier,manualEarlier,setManualEarlier)}{fileBox("latest",latest,manualLatest,setManualLatest)}</section>
  {error&&<div className="error"><AlertTriangle size={16}/>{error}</div>}
  {latest&&earlier&&<section className="summaryCards">
   {[
    [psCount||430,"TOTAL PS","navy"],[total.totalNotices,"TOTAL NOTICES","navy"],[total.nmTotal,"NO MAPPING NOTICES","navy"],
    [total.nmLatest,"NO MAPPING DISPOSED","navy"],[total.nmDocs,"NO MAPPING DOCS UPLOADED","navy"],
    [total.dTotal,"DISCREPANCY NOTICES","brown"],[total.dLatest,"DISCREPANCY DISPOSED","brown"],[total.dLetters,"BLO LETTERS UPLOADED","brown"],
    [total.nmPct+"%","NM % DISPOSED","navy"],[total.dPct+"%","DISC. % DISPOSED","navy"]
   ].map(([v,l,c],i)=><div key={i} className={"summaryCard "+c}><b>{typeof v==="string"?v:formatNumber(Number(v))}</b><span>{l}</span></div>)}
  </section>}
  {latest&&earlier?<section className="panel">
   <div className="actions">
    <div><b>OFFICER WISE PROGRESS REPORT</b><span>Officer-wise summary | No Mapping + Discrepancy + Total Disposal</span></div>
    <div className="buttons"><button onClick={excel}><Download size={15}/> Download Excel</button><button className="pdf" onClick={pdf}><FileText size={15}/> Download PDF</button></div>
   </div>
   <div className="tablewrap"><table><thead>
    <tr className="groups"><th colSpan={4}></th><th colSpan={7}>NO MAPPING NOTICES</th><th className="sep"></th><th colSpan={7}>DISCREPANCY NOTICES</th><th className="sep"></th><th colSpan={3}>TOTAL DISPOSAL</th></tr>
    <tr>{[
      "S.No.","OFFICER NAME AND DESIGNATION","NO. OF PS","TOTAL NOTICES","TOTAL NO MAPPING","NM DISP. EARLIER","NM DISP. LATEST","NM DIFF.","% NM DISPOSED","NM DOCS UPLOADED","PENDING NM","",
      "TOTAL DISCREPANCY","DISC. DISP. EARLIER","DISC. DISP. LATEST","DISC. DIFF.","% DISC. DISPOSED","BLO LETTER UPLOADED","PENDING DISC.","","TOTAL NOTICES DISPOSED","% TOTAL DISPOSED","TOTAL NOTICES PENDING"
    ].map((c,i)=>c===""?<th key={i} className="sep"/>:<th key={i} onClick={()=>sort(([
      "x","officer","psCount","totalNotices","nmTotal","nmEarlier","nmLatest","nmDifference","nmPct","nmDocs","nmPending","x",
      "dTotal","dEarlier","dLatest","dDifference","dPct","dLetters","dPending","x","totalDisposed","totalDisposedPct","totalPending"
    ][i]||"nmDifference") as keyof OfficerReportRow)}>{c}<ArrowUpDown size={10}/></th>)}</tr>
   </thead><tbody>
    {rows.map((r,i)=>{const vals=[i+1,formatOfficer(r),r.psCount,r.totalNotices,r.nmTotal,r.nmEarlier,r.nmLatest,r.nmDifference,r.nmPct+"%",r.nmDocs,r.nmPending,"",r.dTotal,r.dEarlier,r.dLatest,r.dDifference,r.dPct+"%",r.dLetters,r.dPending,"",r.totalDisposed,r.totalDisposedPct+"%",r.totalPending];return <tr key={r.officer}>{vals.map((v,j)=>{if(j===11||j===19)return <td key={j} className="sep"/>;const k=j===7||j===8?groups.nm.get(r.officer):j===9?groups.docs.get(r.officer):j===15||j===16?groups.disc.get(r.officer):j===17?groups.letters.get(r.officer):j===21?groups.totalDisposed.get(r.officer):"";const cls=[j===1?"name":"",j===4||j===12?"earlier":"",k?"colour "+k:""].join(" ");return <td key={j} className={cls}>{v}</td>})}</tr>})}
    <tr className="grand"><td></td><td>GRAND TOTAL - AC-34 MATIALA</td><td>{psCount}</td><td>{formatNumber(total.totalNotices)}</td><td>{formatNumber(total.nmTotal)}</td><td>{formatNumber(total.nmEarlier)}</td><td>{formatNumber(total.nmLatest)}</td><td>{formatNumber(total.nmDifference)}</td><td>{total.nmPct}%</td><td>{formatNumber(total.nmDocs)}</td><td>{formatNumber(total.nmPending)}</td><td className="sep"></td><td>{formatNumber(total.dTotal)}</td><td>{formatNumber(total.dEarlier)}</td><td>{formatNumber(total.dLatest)}</td><td>{formatNumber(total.dDifference)}</td><td>{total.dPct}%</td><td>{formatNumber(total.dLetters)}</td><td>{formatNumber(total.dPending)}</td><td className="sep"></td><td>{formatNumber(total.totalDisposed)}</td><td>{total.totalDisposedPct}%</td><td>{formatNumber(total.totalPending)}</td></tr>
   </tbody></table></div>
   <div className="legend"><b>COLOUR CODE:</b><span className="low">LOW (needs attention)</span><span className="medium">MEDIUM</span><span className="high">HIGH (good)</span></div>
   <div className="legendnote">Applied to Difference and percentage columns. Lowest third = red; highest third = green.</div>
   <div className="notes">Notes: Disposed means Approved / Disposed notices. % Disposed = Latest Disposed / Total Notices of that type. Difference = Latest - Earlier. Total Notices Disposed = No Mapping Disposed + Discrepancy Disposed. Total Notices Pending = Pending No Mapping + Pending Discrepancy.</div>
  </section>:<div className="empty"><FileText size={30}/><b>Upload Earlier and Latest Excel reports</b><span>Only the FINAL PS WISE REPORT sheet is used.</span></div>}
  <style jsx global>{`
   *{box-sizing:border-box}body{margin:0;background:#f3f5f8;color:#172033;font-family:Arial,Helvetica,sans-serif}.bar{height:8px;background:#1f3864}header{background:#1f3864;color:#fff;text-align:center;padding:13px 10px 11px}header h1{font-size:25px;margin:0 0 4px;font-weight:800}header div{font-size:12px}header small{display:block;font-size:9px;margin-top:4px}.uploads{display:grid;grid-template-columns:1fr 1fr;gap:12px;padding:14px 22px}.upload{background:#fff;border:1px solid #b8c0cc;padding:11px}.uploadtitle b{display:block;color:#1f3864;font-size:12px}.uploadtitle span{display:block;color:#667085;font-size:10px;margin-top:2px}.drop{height:43px;border:1px dashed #8d9aad;margin-top:8px;display:flex;align-items:center;justify-content:center;gap:6px;position:relative;font-size:11px}.drop input{position:absolute;inset:0;opacity:0;cursor:pointer}.manual{margin-top:7px;padding:6px;border:1px solid #d0d5dd;font-size:10px;width:210px}.ok{margin-top:6px;color:#20733a;font-size:10px;display:flex;align-items:center;gap:5px}.ok strong{background:#d9ead3;padding:2px 5px}.error{margin:0 22px 10px;padding:9px;background:#fde2e2;color:#8f2d2d;border:1px solid #e5a4a4;font-weight:700;display:flex;gap:6px}.panel{margin:0 22px 24px;background:#fff;border:1px solid #b8c0cc}.actions{padding:9px 10px;border-bottom:1px solid #b8c0cc;display:flex;justify-content:space-between;align-items:center}.actions b{display:block;color:#1f3864;font-size:14px}.actions span{display:block;color:#667085;font-size:10px;margin-top:2px}.buttons{display:flex;gap:7px}.buttons button{display:flex;align-items:center;gap:5px;border:1px solid #1f3864;background:#fff;color:#1f3864;padding:7px 10px;font-weight:800;font-size:11px}.buttons .pdf{background:#1f3864;color:#fff}.tablewrap{overflow:auto}table{border-collapse:collapse;width:100%;min-width:2350px}th,td{border:1px solid #b8c0cc;padding:5px 4px;text-align:center;font-size:11px;line-height:1.08}th{background:#1f3864;color:#fff;font-size:8px;font-weight:800;white-space:normal;cursor:pointer}th svg{vertical-align:middle;margin-left:2px}.groups th{font-size:9px;cursor:default;padding:5px}.groups th:nth-child(2){background:#1f3864}.groups th:nth-child(3){background:#833c0b}.groups th:nth-child(4){background:#1f3864}.groups .sep{background:#fff}.name{text-align:left;font-weight:700;white-space:normal;min-width:190px}tbody tr:nth-child(even){background:#f2f5fa}.earlier{background:#fff8e1}.colour{font-weight:800}.low{background:#f8b4b4}.medium{background:#ffe699}.high{background:#b7e1a1}.colour.low{background:#f8b4b4}.colour.medium{background:#ffe699}.colour.high{background:#b7e1a1}.sep{width:7px;min-width:7px;padding:0!important;background:#fff!important;border-left:0!important;border-right:0!important}.grand td{background:#d9e1f2!important;font-weight:800}.legend{padding:7px 10px 2px;display:flex;align-items:center;gap:8px;font-size:10px}.legend span{padding:3px 10px;font-weight:800;border:1px solid #b8c0cc}.legend .low{background:#f8b4b4}.legend .medium{background:#ffe699}.legend .high{background:#b7e1a1}.legendnote,.notes{font-size:9px;color:#475467;padding:3px 10px}.notes{padding-bottom:9px}.empty{margin:20px 22px;padding:40px;background:#fff;border:1px solid #b8c0cc;display:flex;flex-direction:column;align-items:center;gap:8px;color:#667085}.empty b{color:#1f3864}@media(max-width:900px){.summaryCards{grid-template-columns:repeat(2,1fr)}.uploads{grid-template-columns:1fr}.actions{align-items:flex-start;gap:10px;flex-direction:column}}@media print{.uploads,.actions{display:none!important}.panel{margin:0}body{background:#fff}}
  `}</style>
 </main>;
}
