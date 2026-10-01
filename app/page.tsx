"use client";

import React,{useMemo,useState} from "react";
import * as XLSX from "xlsx";
import jsPDF from "jspdf";
import autoTable from "jspdf-autotable";
import {Download,Upload,FileText,CheckCircle2,AlertTriangle,ArrowUpDown} from "lucide-react";
import {RawRow,ParsedReport,OfficerReportRow,buildOfficerRows,formatNumber,formatOfficer,grandOfficer,parseTimestamp,validateColumns,normalizeRows,parseGrandTotal,validateGrandTotal,percentileGroups} from "../lib/report";

const COLS=[
 "S.No.","OFFICER NAME AND DESIGNATION","TOTAL NOTICES",
 "TOTAL NO MAPPING NOTICES","NO MAPPING DISPOSED EARLIER","NO MAPPING DISPOSED LATEST",
 "DIFFERENCE","% OF NO MAPPING NOTICES DISPOSED","NO MAPPING DOCUMENTS UPLOADED",
 "% DOCUMENTS / TOTAL NO MAPPING NOTICES","",
 "TOTAL DISCREPANCY NOTICES","DISCREPANCY DISPOSED EARLIER","DISCREPANCY DISPOSED LATEST",
 "DIFFERENCE","% OF DISCREPANCY NOTICES DISPOSED","BLO LETTER UPLOADED",
 "% BLO LETTER / TOTAL DISCREPANCY NOTICES","TOTAL NOTICES DISPOSED","% OF TOTAL NOTICES DISPOSED"
];

function rowValues(r:OfficerReportRow,rowIndex:number){
 return [
  rowIndex+1,formatOfficer(r),r.totalNotices,r.nmTotal,r.nmEarlier,r.nmLatest,r.nmDifference,
  r.nmPct+"%",r.nmDocs,r.nmDocsPct+"%","",r.dTotal,r.dEarlier,r.dLatest,
  r.dDifference,r.dPct+"%",r.dLetters,r.dLettersPct+"%",r.totalDisposed,r.totalDisposedPct+"%"
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
  const doc=new jsPDF({orientation:"portrait",unit:"mm",format:"a3"});
  const navy:[number,number,number]=[31,56,100],brown:[number,number,number]=[131,60,11],yellow:[number,number,number]=[255,248,225],grid:[number,number,number]=[184,192,204];
  doc.setFillColor(...navy);doc.rect(0,0,297,29,"F");
  doc.setTextColor(255,255,255);doc.setFont("helvetica","bold");doc.setFontSize(17);
  doc.text("OFFICER WISE PROGRESS REPORT",148.5,9,{align:"center"});
  doc.setFont("helvetica","normal");doc.setFontSize(8);
  doc.text("No Mapping & Discrepancy Notices - Earlier vs Latest Disposal Status",148.5,15,{align:"center"});
  doc.text(`Data as of ${latest.timestamp} | ${rows.length} Officers | ${psCount} Polling Stations`,148.5,21,{align:"center"});

  const top:any[]=[
   {content:"S.N\no.",rowSpan:2},
   {content:"OFFICER NAME AND\nDESIGNATION",rowSpan:2},
   {content:"TOTAL\nNOTICES\n(NO\nMAPPING +\nDISCREPANCY)",rowSpan:2},
   {content:"NO MAPPING NOTICES",colSpan:7,styles:{fillColor:navy,textColor:255}},
   {content:"",rowSpan:2,styles:{fillColor:[255,255,255],textColor:[255,255,255],lineWidth:0}},
   {content:"DISCREPANCY NOTICES",colSpan:7,styles:{fillColor:brown,textColor:255}},
   {content:"TOTAL DISPOSAL",colSpan:2,styles:{fillColor:navy,textColor:255}}
  ];
  const sub:any[]=[
   "TOTAL NO\nMAPPING\nNOTICES","NO MAPPING\nDISPOSED\nEARLIER","NO MAPPING\nDISPOSED\nLATEST","DIFFERENCE",
   "% OF NO\nMAPPING\nNOTICES\nDISPOSED","NO MAPPING DOCUMENTS\nUPLOADED","% DOCUMENTS /\nTOTAL NO MAPPING\nNOTICES",
   "TOTAL DISCREPANCY\nNOTICES","DISCREPANCY\nDISPOSED\nEARLIER","DISCREPANCY\nDISPOSED\nLATEST","DIFFERENCE",
   "% OF DISCREPANCY\nNOTICES\nDISPOSED","BLO LETTER\nUPLOADED","% BLO LETTER /\nTOTAL DISCREPANCY\nNOTICES"
  ].map((content,i)=>({content,styles:{fillColor:(i<7?navy:brown) as [number,number,number],textColor:255}}));
  sub.splice(7,0,{content:"",styles:{fillColor:[255,255,255],textColor:[255,255,255],lineWidth:0}});
  sub.push(
   {content:"TOTAL NOTICES\nDISPOSED",styles:{fillColor:navy,textColor:255}},
   {content:"% OF TOTAL\nNOTICES DISPOSED",styles:{fillColor:navy,textColor:255}}
  );

  const body=rows.map((r,rowIndex)=>rowValues(r,rowIndex));
  body.push([
   "","GRAND TOTAL",total.totalNotices,total.nmTotal,total.nmEarlier,total.nmLatest,total.nmDifference,
   total.nmPct+"%",total.nmDocs,total.nmDocsPct+"%","",total.dTotal,total.dEarlier,total.dLatest,
   total.dDifference,total.dPct+"%",total.dLetters,total.dLettersPct+"%",total.totalDisposed,total.totalDisposedPct+"%"
  ]);

  const widths=[7,29,13,14,14,14,12,16,14,16,3,14,14,14,12,16,13,16,14,15];
  autoTable(doc,{
   startY:34,head:[top,sub],body,theme:"grid",tableWidth:277,margin:{left:10,right:10,bottom:18},
   styles:{font:"helvetica",fontSize:7.0,cellPadding:1.05,lineColor:grid,lineWidth:.25,textColor:[20,28,38] as [number,number,number],halign:"center",valign:"middle",overflow:"linebreak"},
   headStyles:{font:"helvetica",fontStyle:"bold",fontSize:6.0,halign:"center",valign:"middle",cellPadding:2.2,minCellHeight:15},
   columnStyles:Object.fromEntries(widths.map((w,i)=>[i,{cellWidth:w,halign:i===1?"left":"center"}])),
   didParseCell:(data:any)=>{
    if(data.section==="head"){
     data.cell.styles.minCellHeight=data.row.index===0?11:16;
    }
    if(data.section==="body"){
     const r=rows[data.row.index];
     if(data.row.index===rows.length){
      data.cell.styles.fillColor=[217,225,242];data.cell.styles.fontStyle="bold";
     }else if(r){
      if(data.column.index===4||data.column.index===12)data.cell.styles.fillColor=yellow;
      if(data.column.index===6)data.cell.styles.fillColor=color(groups.nm.get(r.officer)||"medium");
      if(data.column.index===9)data.cell.styles.fillColor=color(groups.docs.get(r.officer)||"medium");
      if(data.column.index===14)data.cell.styles.fillColor=color(groups.disc.get(r.officer)||"medium");
      if(data.column.index===17)data.cell.styles.fillColor=color(groups.letters.get(r.officer)||"medium");
      if(data.column.index===19)data.cell.styles.fillColor=color(groups.totalDisposed.get(r.officer)||"medium");
      if([5,6,8,9,13,14,16,17,18,19].includes(data.column.index))data.cell.styles.fontStyle="bold";
     }
    }
   },
   didDrawPage:(data:any)=>{
    doc.setFont("helvetica","normal");doc.setFontSize(7);doc.setTextColor(80,88,98);
    doc.text("Officer Wise Progress Report - No Mapping & Discrepancy Notices",10,412);
    doc.text(`Page ${data.pageNumber}`,287,412,{align:"right"});
   }
  });

  let y=((doc as any).lastAutoTable?.finalY||260)+6;
  if(y>395){doc.addPage();y=18;}
  doc.setFont("helvetica","bold");doc.setFontSize(8);doc.setTextColor(40,48,58);doc.text("COLOUR CODE:",10,y);
  let x=34;
  [["LOW (needs attention)","low"],["MEDIUM","medium"],["HIGH (good)","high"]].forEach(([label,k])=>{
   const c=color(k);doc.setFillColor(c[0],c[1],c[2]);doc.rect(x,y-4,27,4,"F");
   doc.setTextColor(40,48,58);doc.text(label,x+29,y);x+=78;
  });
  doc.setFont("helvetica","normal");doc.setFontSize(7);
  doc.text("Applied to Difference and percentage columns. Lowest third = red; highest third = green.",10,y+6);
  doc.text("Notes: Disposed means Approved / Disposed notices. % Disposed = Latest Disposed / Total Notices of that type. Difference = Latest - Earlier.",10,y+11);
  // Detailed PS-wise tables: one section for every officer, using that officer's concerned PS only.
  rows.forEach((officerRow,officerIndex)=>{
   const psBody=buildOfficerPSRows(officerRow.officer);
   if(!psBody.length)return;
   const psTotal=[
    "","OFFICER TOTAL","", "",officerRow.totalNotices,
    officerRow.nmTotal,officerRow.nmEarlier,officerRow.nmLatest,officerRow.nmDifference,officerRow.nmDocs,
    officerRow.dTotal,officerRow.dEarlier,officerRow.dLatest,officerRow.dDifference,officerRow.dLetters,
    officerRow.totalDisposed,officerRow.totalDisposedPct+"%"
   ];
   psBody.push(psTotal);
   doc.addPage();
   doc.setFillColor(...navy);doc.rect(0,0,297,13,"F");
   doc.setTextColor(255,255,255);doc.setFont("helvetica","bold");doc.setFontSize(10);
   doc.text(`${officerIndex+1}. ${officerRow.officer} ${officerRow.designation||"Officer"} - ${psBody.length} PS`,148.5,8,{align:"center"});
   doc.setTextColor(30,40,55);doc.setFont("helvetica","bold");doc.setFontSize(8);
   doc.text(`CONCERNED PS-WISE NO MAPPING & DISCREPANCY DISPOSAL - ${officerRow.officer}`,10,20);
   const pHead=[
    "S.No.","PS No.","BLO Name","BLO Supervisor","TOTAL\nNOTICES",
    "NO MAPPING\nNOTICES","NM DISPOSED\nEARLIER","NM DISPOSED\nLATEST","NM\nDIFF.","NM DOCS",
    "DISCREPANCY\nNOTICES","DISP.\nEARLIER","DISP.\nLATEST","DIFF.","BLO\nLETTER",
    "TOTAL NOTICES\nDISPOSED","% TOTAL\nNOTICES DISPOSED"
   ];
   const pWidths=[6,9,22,20,12,12,12,12,9,11,12,11,11,9,11,14,14];
   const psData=psBody.slice(0,-1);
   const psGroups=(idx:number)=>{
    const vals=psData.map((r:any)=>Number(r[idx])||0).sort((a:number,b:number)=>a-b);
    if(!vals.length)return [] as string[];
    const q=(p:number)=>{const z=(vals.length-1)*p,lo=Math.floor(z),hi=Math.ceil(z);return vals[lo]+(vals[hi]-vals[lo])*(z-lo);};
    const p33=q(1/3),p66=q(2/3);
    return psData.map((r:any)=>{const v=Number(r[idx])||0;return v<=p33?"low":v<=p66?"medium":"high";});
   };
   const psNmDiffGroups=psGroups(8),psDiscDiffGroups=psGroups(13),psTotalPctGroups=psGroups(16);
   autoTable(doc,{
    startY:24,head:[pHead],body:psBody,theme:"grid",tableWidth:277,margin:{left:10,right:10,bottom:15},
    styles:{font:"helvetica",fontSize:7.0,cellPadding:1.05,lineColor:grid,lineWidth:.25,halign:"center",valign:"middle",overflow:"linebreak"},
    headStyles:{font:"helvetica",fontStyle:"bold",fontSize:6.0,fillColor:navy,textColor:255,cellPadding:2.1,minCellHeight:15},
    columnStyles:Object.fromEntries(pWidths.map((w,i)=>[i,{cellWidth:w,halign:[2,3].includes(i)?"left":"center"}])),
    didParseCell:(data:any)=>{
     if(data.section==="body"){
      if(data.row.index===psBody.length-1){
       data.cell.styles.fillColor=[217,225,242];data.cell.styles.fontStyle="bold";
      }else{
       const ri=data.row.index;
       if(data.column.index===8){
        data.cell.styles.fillColor=color(psNmDiffGroups[ri]||"medium");
        data.cell.styles.fontStyle="bold";
       }
       if(data.column.index===13){
        data.cell.styles.fillColor=color(psDiscDiffGroups[ri]||"medium");
        data.cell.styles.fontStyle="bold";
       }
       if(data.column.index===16){
        data.cell.styles.fillColor=color(psTotalPctGroups[ri]||"medium");
        data.cell.styles.fontStyle="bold";
       }
      }
     }
    },
    didDrawPage:(data:any)=>{
     doc.setFont("helvetica","normal");doc.setFontSize(7);doc.setTextColor(80,88,98);
     doc.text("Officer Wise Progress Report - Concerned PS Details",10,412);
     doc.text(`Page ${data.pageNumber}`,287,412,{align:"right"});
    }
   });
  });
  doc.save(`Officer_Wise_Report_${new Date().toISOString().slice(0,10)}.pdf`);
 };

 const excel=()=>{
  if(!latest||!earlier)return;
  const data=rows.map((r,i)=>({
   "S.No.":i+1,"OFFICER NAME AND DESIGNATION":formatOfficer(r),"TOTAL NOTICES":r.totalNotices,
   "TOTAL NO MAPPING NOTICES":r.nmTotal,"NO MAPPING DISPOSED EARLIER":r.nmEarlier,
   "NO MAPPING DISPOSED LATEST":r.nmLatest,"DIFFERENCE":r.nmDifference,
   "% OF NO MAPPING NOTICES DISPOSED":r.nmPct+"%","NO MAPPING DOCUMENTS UPLOADED":r.nmDocs,
   "% DOCUMENTS / TOTAL NO MAPPING NOTICES":r.nmDocsPct+"%","TOTAL DISCREPANCY NOTICES":r.dTotal,
   "DISCREPANCY DISPOSED EARLIER":r.dEarlier,"DISCREPANCY DISPOSED LATEST":r.dLatest,
   "DIFFERENCE (DISCREPANCY)":r.dDifference,"% OF DISCREPANCY NOTICES DISPOSED":r.dPct+"%",
   "BLO LETTER UPLOADED":r.dLetters,"% BLO LETTER / TOTAL DISCREPANCY NOTICES":r.dLettersPct+"%",
   "TOTAL NOTICES DISPOSED":r.totalDisposed,"% OF TOTAL NOTICES DISPOSED":r.totalDisposedPct+"%"
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
  {latest&&earlier?<section className="panel">
   <div className="actions">
    <div><b>OFFICER WISE PROGRESS REPORT</b><span>One report • exact reference layout</span></div>
    <div className="buttons"><button onClick={excel}><Download size={15}/> Download Excel</button><button className="pdf" onClick={pdf}><FileText size={15}/> Download PDF</button></div>
   </div>
   <div className="tablewrap"><table><thead>
    <tr className="groups"><th colSpan={3}></th><th colSpan={7}>NO MAPPING NOTICES</th><th className="sep"></th><th colSpan={7}>DISCREPANCY NOTICES</th><th colSpan={2}>TOTAL DISPOSAL</th></tr>
    <tr>{COLS.map((c,i)=>i===10?<th key={i} className="sep"/>:<th key={i} onClick={()=>sort((["sno","officer","totalNotices","nmTotal","nmEarlier","nmLatest","nmDifference","nmPct","nmDocs","nmDocsPct","x","dTotal","dEarlier","dLatest","dDifference","dPct","dLetters","dLettersPct","totalDisposed","totalDisposedPct"][i]||"nmDifference") as keyof OfficerReportRow)}>{c}<ArrowUpDown size={10}/></th>)}</tr>
   </thead><tbody>
    {rows.map((r,i)=><tr key={r.officer}>
     {COLS.map((_,j)=>{
      if(j===10)return <td key={j} className="sep"/>;
      const k=j===6?groups.nm.get(r.officer):j===9?groups.docs.get(r.officer):j===14?groups.disc.get(r.officer):j===17?groups.letters.get(r.officer):"";
      const cls=[j===4||j===12?"earlier":"",k?"colour "+k:"",j===1?"name":""].join(" ");
      return <td key={j} className={cls}>{rowValues(r,i)[j]}</td>;
     })}
    </tr>)}
    <tr className="grand"><td></td><td>GRAND TOTAL</td><td>{formatNumber(total.totalNotices)}</td><td>{formatNumber(total.nmTotal)}</td><td>{formatNumber(total.nmEarlier)}</td><td>{formatNumber(total.nmLatest)}</td><td>{formatNumber(total.nmDifference)}</td><td>{total.nmPct}%</td><td>{formatNumber(total.nmDocs)}</td><td>{total.nmDocsPct}%</td><td className="sep"></td><td>{formatNumber(total.dTotal)}</td><td>{formatNumber(total.dEarlier)}</td><td>{formatNumber(total.dLatest)}</td><td>{formatNumber(total.dDifference)}</td><td>{total.dPct}%</td><td>{formatNumber(total.dLetters)}</td><td>{total.dLettersPct}%</td><td>{formatNumber(total.totalDisposed)}</td><td>{total.totalDisposedPct}%</td></tr>
   </tbody></table></div>
   <div className="legend"><b>COLOUR CODE:</b><span className="low">LOW (needs attention)</span><span className="medium">MEDIUM</span><span className="high">HIGH (good)</span></div>
   <div className="legendnote">Applied to Difference and upload-percentage columns. Lowest third = red; highest third = green.</div>
   <div className="notes">Notes: Disposed means Approved / Disposed notices. % Disposed = Latest Disposed / Total Notices of that type. Difference = Latest - Earlier. Total Notices Disposed = No Mapping Disposed + Discrepancy Disposed.</div>
  </section>:<div className="empty"><FileText size={30}/><b>Upload Earlier and Latest Excel reports</b><span>Only the FINAL PS WISE REPORT sheet is used.</span></div>}
  <style jsx global>{`
   *{box-sizing:border-box}body{margin:0;background:#f3f5f8;color:#172033;font-family:Arial,Helvetica,sans-serif}.bar{height:8px;background:#1f3864}header{background:#1f3864;color:#fff;text-align:center;padding:13px 10px 11px}header h1{font-size:25px;margin:0 0 4px;font-weight:800}header div{font-size:12px}header small{display:block;font-size:9px;margin-top:4px}.uploads{display:grid;grid-template-columns:1fr 1fr;gap:12px;padding:14px 22px}.upload{background:#fff;border:1px solid #b8c0cc;padding:11px}.uploadtitle b{display:block;color:#1f3864;font-size:12px}.uploadtitle span{display:block;color:#667085;font-size:10px;margin-top:2px}.drop{height:43px;border:1px dashed #8d9aad;margin-top:8px;display:flex;align-items:center;justify-content:center;gap:6px;position:relative;font-size:11px}.drop input{position:absolute;inset:0;opacity:0;cursor:pointer}.manual{margin-top:7px;padding:6px;border:1px solid #d0d5dd;font-size:10px;width:210px}.ok{margin-top:6px;color:#20733a;font-size:10px;display:flex;align-items:center;gap:5px}.ok strong{background:#d9ead3;padding:2px 5px}.error{margin:0 22px 10px;padding:9px;background:#fde2e2;color:#8f2d2d;border:1px solid #e5a4a4;font-weight:700;display:flex;gap:6px}.panel{margin:0 22px 24px;background:#fff;border:1px solid #b8c0cc}.actions{padding:9px 10px;border-bottom:1px solid #b8c0cc;display:flex;justify-content:space-between;align-items:center}.actions b{display:block;color:#1f3864;font-size:14px}.actions span{display:block;color:#667085;font-size:10px;margin-top:2px}.buttons{display:flex;gap:7px}.buttons button{display:flex;align-items:center;gap:5px;border:1px solid #1f3864;background:#fff;color:#1f3864;padding:7px 10px;font-weight:800;font-size:11px}.buttons .pdf{background:#1f3864;color:#fff}.tablewrap{overflow:auto}table{border-collapse:collapse;width:100%;min-width:1700px}th,td{border:1px solid #b8c0cc;padding:6px 5px;text-align:center;font-size:12px;line-height:1.08}th{background:#1f3864;color:#fff;font-size:9px;font-weight:800;white-space:normal;cursor:pointer}th svg{vertical-align:middle;margin-left:2px}.groups th{font-size:10px;cursor:default;padding:5px}.groups th:last-child{background:#833c0b}.groups th:nth-child(2){background:#1f3864}.groups .sep{background:#fff}.name{text-align:left;font-weight:700;white-space:normal;min-width:180px}tbody tr:nth-child(even){background:#f2f5fa}.earlier{background:#fff8e1}.colour{font-weight:800}.low{background:#f8b4b4}.medium{background:#ffe699}.high{background:#b7e1a1}.colour.low{background:#f8b4b4}.colour.medium{background:#ffe699}.colour.high{background:#b7e1a1}.sep{width:7px;min-width:7px;padding:0!important;background:#fff!important;border-left:0!important;border-right:0!important}.grand td{background:#d9e1f2!important;font-weight:800}.legend{padding:7px 10px 2px;display:flex;align-items:center;gap:8px;font-size:10px}.legend span{padding:3px 10px;font-weight:800;border:1px solid #b8c0cc}.legend .low{background:#f8b4b4}.legend .medium{background:#ffe699}.legend .high{background:#b7e1a1}.legendnote,.notes{font-size:9px;color:#475467;padding:3px 10px}.notes{padding-bottom:9px}.empty{margin:20px 22px;padding:40px;background:#fff;border:1px solid #b8c0cc;display:flex;flex-direction:column;align-items:center;gap:8px;color:#667085}.empty b{color:#1f3864}@media(max-width:900px){.uploads{grid-template-columns:1fr}.actions{align-items:flex-start;gap:10px;flex-direction:column}}@media print{.uploads,.actions{display:none!important}.panel{margin:0}body{background:#fff}}
  `}</style>
 </main>;
}
