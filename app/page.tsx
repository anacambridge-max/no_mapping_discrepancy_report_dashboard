[Reading 191 lines from start (total: 191 lines, 0 remaining)]

"use client";
import React,{useMemo,useState} from "react";
import * as XLSX from "xlsx";
import jsPDF from "jspdf";
import autoTable from "jspdf-autotable";
import {Download,Upload,FileText,CheckCircle2,AlertTriangle,ArrowUpDown,BarChart3} from "lucide-react";
import {RawRow,ParsedReport,OfficerReportRow,PSReportRow,buildReports,formatNumber,formatOfficer,grandOfficer,grandPS,n,parseTimestamp,validateColumns,normalizeRows,parseGrandTotal,validateGrandTotal,percentileGroups} from "../lib/report";

const HEADER = ["S.No.","OFFICER NAME AND DESIGNATION","TOTAL NOTICES","TOTAL NO MAPPING NOTICES","NO MAPPING DISPOSED EARLIER","NO MAPPING DISPOSED LATEST","DIFFERENCE","% OF NO MAPPING NOTICES DISPOSED","NO MAPPING DOCUMENTS UPLOADED","% DOCUMENTS / TOTAL NO MAPPING NOTICES","","TOTAL DISCREPANCY NOTICES","DISCREPANCY DISPOSED EARLIER","DISCREPANCY DISPOSED LATEST","DIFFERENCE","% OF DISCREPANCY NOTICES DISPOSED","BLO LETTER UPLOADED","% BLO LETTER / TOTAL DISCREPANCY NOTICES"];
const PS_HEADER = ["S.No.","PS NO.","OFFICER NAME AND DESIGNATION","BLO SUPERVISOR","BLO NAME","TOTAL NOTICES","TOTAL NO MAPPING NOTICES","NO MAPPING DISPOSED EARLIER","NO MAPPING DISPOSED LATEST","DIFFERENCE","% OF NO MAPPING NOTICES DISPOSED","NO MAPPING DOCUMENTS UPLOADED","% DOCUMENTS / TOTAL NO MAPPING NOTICES","","TOTAL DISCREPANCY NOTICES","DISCREPANCY DISPOSED EARLIER","DISCREPANCY DISPOSED LATEST","DIFFERENCE","% OF DISCREPANCY NOTICES DISPOSED","BLO LETTER UPLOADED","% BLO LETTER / TOTAL DISCREPANCY NOTICES"];

function cellValue(r:any,i:number,ps=false){
 if(ps){return [i+1,r.psNo,formatOfficer(r),r.supervisor,r.blo,r.totalNotices,r.nmTotal,r.nmEarlier,r.nmLatest,r.nmDifference,r.nmPct+"%",r.nmDocs,r.nmDocsPct+"%","",r.dTotal,r.dEarlier,r.dLatest,r.dDifference,r.dPct+"%",r.dLetters,r.dLettersPct+"%"][i];}
 return [i+1,formatOfficer(r),r.totalNotices,r.nmTotal,r.nmEarlier,r.nmLatest,r.nmDifference,r.nmPct+"%",r.nmDocs,r.nmDocsPct+"%","",r.dTotal,r.dEarlier,r.dLatest,r.dDifference,r.dPct+"%",r.dLetters,r.dLettersPct+"%"][i];
}

function Header({report,setReport}:{report:"officer"|"ps";setReport:(x:"officer"|"ps")=>void}){
 return <><div className="topbar"/><header className="titleblock">
   <div className="switch"><button className={report==="officer"?"active":""} onClick={()=>setReport("officer")}>OFFICER WISE REPORT</button><button className={report==="ps"?"active":""} onClick={()=>setReport("ps")}>PS WISE REPORT</button></div>
   <h1>{report==="officer"?"OFFICER WISE PROGRESS REPORT":"PS WISE PROGRESS REPORT"}</h1>
   <div className="subtitle">No Mapping &amp; Discrepancy Notices - Earlier vs Latest Disposal Status</div>
 </header></>;
}

export default function Page(){
 const [earlier,setEarlier]=useState<ParsedReport|null>(null),[latest,setLatest]=useState<ParsedReport|null>(null);
 const [manualEarlier,setManualEarlier]=useState(""),[manualLatest,setManualLatest]=useState(""),[report,setReport]=useState<"officer"|"ps">("officer");
 const [error,setError]=useState(""),[sortKey,setSortKey]=useState("nmDifference"),[sortDir,setSortDir]=useState(-1);

 const loadFile=async(file:File,which:"earlier"|"latest",manual:string)=>{
   setError("");
   try{
     const wb=XLSX.read(await file.arrayBuffer(),{type:"array",cellDates:true});
     if(!wb.SheetNames.includes("FINAL PS WISE REPORT")) throw new Error('Required sheet "FINAL PS WISE REPORT" was not found.');
     const raw=XLSX.utils.sheet_to_json<RawRow>(wb.Sheets["FINAL PS WISE REPORT"],{defval:""});
     const missing=validateColumns(raw);
     if(missing.length) throw new Error("Missing required column(s): "+missing.join(", "));
     const grand=parseGrandTotal(raw), rows=normalizeRows(raw), validation=validateGrandTotal(rows,grand);
     const parsed={fileName:file.name,timestamp:parseTimestamp(file.name,file.lastModified,manual),rows,grandTotal:grand,validation};
     which==="earlier"?setEarlier(parsed):setLatest(parsed);
   }catch(e:any){setError(e?.message||"Unable to read Excel file.");}
 };
 const drop=(which:"earlier"|"latest",manual:string)=>(e:React.DragEvent)=>{
   e.preventDefault(); const f=e.dataTransfer.files?.[0]; if(f)loadFile(f,which,manual);
 };

 const reports=useMemo(()=>buildReports(latest?.rows||[],earlier?.rows||[]),[latest,earlier]);
 const officerRows=useMemo(()=>[...reports.officerRows].sort((a:any,b:any)=>{
   if(sortKey==="officer") return a.officer.localeCompare(b.officer)*sortDir;
   if(sortKey==="designation") return a.designation.localeCompare(b.designation)*sortDir;
   return ((a[sortKey]??0)-(b[sortKey]??0))*sortDir;
 }),[reports.officerRows,sortKey,sortDir]);
 const psRows=useMemo(()=>[...reports.psRows].sort((a:any,b:any)=>{
   if(sortKey==="psNo") return Number(a.psNo)-Number(b.psNo);
   if(sortKey==="officer") return a.officer.localeCompare(b.officer)*sortDir;
   return ((a[sortKey]??0)-(b[sortKey]??0))*sortDir;
 }),[reports.psRows,sortKey,sortDir]);
 const active=report==="officer"?officerRows:psRows;
 const total=report==="officer"?grandOfficer(officerRows):grandPS(psRows);
 const psCount=new Set((latest?.rows||[]).map(r=>r.psNo)).size;
 const officerCount=new Set((latest?.rows||[]).map(r=>r.officer)).size;
 const latestTimestamp=latest?.timestamp||"—";

 const nmGroups=useMemo(()=>percentileGroups(active,r=>r.nmDifference),[active]);
 const docsPctGroups=useMemo(()=>percentileGroups(active,r=>r.nmDocsPct),[active]);
 const discGroups=useMemo(()=>percentileGroups(active,r=>r.dDifference),[active]);
 const lettersPctGroups=useMemo(()=>percentileGroups(active,r=>r.dLettersPct),[active]);

 const sort=(k:string)=>{if(sortKey===k)setSortDir(x=>-x);else{setSortKey(k);setSortDir(-1)}};

 const pdf=()=>{
   if(!latest)return;
   const doc=new jsPDF({orientation:"landscape",unit:"mm",format:"a3"});
   const isOfficer=report==="officer";
   const rows:any[]=isOfficer?officerRows:psRows;
   const g:any=total;
   const title=isOfficer?"OFFICER WISE PROGRESS REPORT":"PS WISE PROGRESS REPORT";
   const headers=isOfficer?HEADER:PS_HEADER;
   const keyFor=(r:any,idx:number)=>isOfficer?cellValue(r,idx):cellValue(r,idx,true);
   const groups={nm:nmGroups,docs:docsPctGroups,disc:discGroups,letters:lettersPctGroups};
   const fill=(group:any,r:any)=>group.get(r.officer||r.psNo)==="low"?[248,180,180]:group.get(r.officer||r.psNo)==="medium"?[255,230,153]:[183,225,161];

   doc.setFillColor(31,56,100);doc.rect(0,0,420,29,"F");
   doc.setTextColor(255,255,255);doc.setFont("helvetica","bold");doc.setFontSize(18);doc.text(title,210,10,{align:"center"});
   doc.setFont("helvetica","normal");doc.setFontSize(9);doc.text("No Mapping & Discrepancy Notices - Earlier vs Latest Disposal Status",210,17,{align:"center"});
   doc.text(`Data as of ${latestTimestamp} | ${officerCount} Officers | ${psCount} Polling Stations`,210,23,{align:"center"});

   const baseCount=isOfficer?3:6;
   const head1:any[]=[];
   if(isOfficer){
     head1.push({content:"S.No.",rowSpan:2},{content:"OFFICER NAME AND DESIGNATION",rowSpan:2},{content:"TOTAL NOTICES",rowSpan:2});
   }else{
     head1.push({content:"S.No.",rowSpan:2},{content:"PS NO.",rowSpan:2},{content:"OFFICER NAME AND DESIGNATION",rowSpan:2},{content:"BLO SUPERVISOR",rowSpan:2},{content:"BLO NAME",rowSpan:2},{content:"TOTAL NOTICES",rowSpan:2});
   }
   head1.push({content:"NO MAPPING NOTICES",colSpan:7,styles:{fillColor:[31,56,100],textColor:255,halign:"center",valign:"middle"}});
   head1.push({content:"",rowSpan:2,styles:{fillColor:[255,255,255],textColor:[255,255,255]}});
   head1.push({content:"DISCREPANCY NOTICES",colSpan:7,styles:{fillColor:[131,60,11],textColor:255,halign:"center",valign:"middle"}});

   const sub=[
     "TOTAL NO MAPPING NOTICES","NO MAPPING DISPOSED EARLIER","NO MAPPING DISPOSED LATEST","DIFFERENCE","% OF NO MAPPING NOTICES DISPOSED","NO MAPPING DOCUMENTS UPLOADED","% DOCUMENTS / TOTAL NO MAPPING NOTICES",
     "TOTAL DISCREPANCY NOTICES","DISCREPANCY DISPOSED EARLIER","DISCREPANCY DISPOSED LATEST","DIFFERENCE","% OF DISCREPANCY NOTICES DISPOSED","BLO LETTER UPLOADED","% BLO LETTER / TOTAL DISCREPANCY NOTICES"
   ];
   const head2=sub.map((x,i)=>({content:x,styles:{fillColor:i<7?[31,56,100]:[131,60,11],textColor:255,halign:"center",valign:"middle",fontStyle:"bold"}}));

   const body=rows.map((r:any)=>headers.map((_,i)=>keyFor(r,i)));
   const grand=isOfficer
     ? ["","GRAND TOTAL",g.totalNotices,g.nmTotal,g.nmEarlier,g.nmLatest,g.nmDifference,g.nmPct+"%",g.nmDocs,g.nmDocsPct+"%","",g.dTotal,g.dEarlier,g.dLatest,g.dDifference,g.dPct+"%",g.dLetters,g.dLettersPct+"%"]
     : ["","GRAND TOTAL","","","",g.totalNotices,g.nmTotal,g.nmEarlier,g.nmLatest,g.nmDifference,g.nmPct+"%",g.nmDocs,g.nmDocsPct+"%","",g.dTotal,g.dEarlier,g.dLatest,g.dDifference,g.dPct+"%",g.dLetters,g.dLettersPct+"%"];
   body.push(grand);

   const widths=isOfficer
     ?[9,45,20,20,21,21,18,22,20,24,4,20,21,21,18,22,20,24]
     :[9,13,35,25,23,19,19,21,21,18,21,20,22,4,20,21,21,18,22,20,24];

   autoTable(doc,{
     startY:35,
     head:[head1,head2],
     body,
     theme:"grid",
     tableWidth:400,
     margin:{left:10,right:10,bottom:18},
     styles:{font:"helvetica",fontSize:isOfficer?8.2:7.2,cellPadding:1.5,lineColor:[184,192,204],lineWidth:.25,textColor:[20,28,38],halign:"center",valign:"middle",overflow:"linebreak"},
     columnStyles:Object.fromEntries(widths.map((w,i)=>[i,{cellWidth:w,halign:(i===1||(!isOfficer&&[2,3,4].includes(i)))?"left":"center"}])),
     headStyles:{fontStyle:"bold",fontSize:7.5,textColor:[255,255,255],fillColor:[31,56,100],halign:"center",valign:"middle"},
     alternateRowStyles:{fillColor:[242,245,250]},
     didParseCell:(d:any)=>{
       if(d.section==="body"&&d.row.index===body.length-1){d.cell.styles.fillColor=[217,225,242];d.cell.styles.fontStyle="bold";return;}
       if(d.section!=="body")return;
       const r=rows[d.row.index]; if(!r)return;
       const idx=d.column.index; let c:any=null;
       if(isOfficer){if(idx===6)c=groups.nm;if(idx===9)c=groups.docs;if(idx===14)c=groups.disc;if(idx===17)c=groups.letters;}
       else {if(idx===9)c=groups.nm;if(idx===12)c=groups.docs;if(idx===17)c=groups.disc;if(idx===20)c=groups.letters;}
       if(c)d.cell.styles.fillColor=fill(c,r);
       if((isOfficer&&[4,12].includes(idx))||(!isOfficer&&[7,15].includes(idx)))d.cell.styles.fillColor=[255,248,225];
       if((isOfficer&&[5,6,8,9,13,14,16,17].includes(idx))||(!isOfficer&&[8,9,11,12,16,17,19,20].includes(idx)))d.cell.styles.fontStyle="bold";
     },
     didDrawPage:(d:any)=>{doc.setFont("helvetica","normal");doc.setFontSize(7);doc.setTextColor(85,94,106);doc.text(`${isOfficer?"Officer":"PS"} Wise Progress Report - No Mapping & Discrepancy Notices`,10,290);doc.text(`Page ${d.pageNumber}`,410,290,{align:"right"});}
   });

   let y=(doc as any).lastAutoTable.finalY+7;
   if(y>275){doc.addPage();y=18;}
   doc.setFont("helvetica","bold");doc.setFontSize(8);doc.setTextColor(40,48,58);doc.text("COLOUR CODE:",10,y);
   const legend=[["LOW (needs attention)",[248,180,180]],["MEDIUM",[255,230,153]],["HIGH (good)",[183,225,161]]];
   let x=35;legend.forEach(([label,c]:any)=>{doc.setFillColor(c[0],c[1],c[2]);doc.rect(x,y-4,28,5,"F");doc.setTextColor(40,48,58);doc.text(label,x+31,y);x+=80;});
   doc.setFont("helvetica","normal");doc.setFontSize(7);doc.text("Applied to Difference and upload-percentage columns. Lowest third = red; highest third = green.",10,y+6);
   doc.text("Notes: Disposed means Approved / Disposed notices. % Disposed = Latest Disposed / Total Notices of that type. Difference = Latest - Earlier.",10,y+12);
   doc.save(`${isOfficer?"Officer_Wise_Report":"PS_Wise_Report"}_${new Date().toISOString().slice(0,10)}.pdf`);
 };

 const excel=()=>{
   const data=report==="officer"?officerRows.map((r,i)=>({"S.No.":i+1,"OFFICER NAME AND DESIGNATION":formatOfficer(r),"TOTAL NOTICES":r.totalNotices,"TOTAL NO MAPPING NOTICES":r.nmTotal,"NO MAPPING DISPOSED EARLIER":r.nmEarlier,"NO MAPPING DISPOSED LATEST":r.nmLatest,"DIFFERENCE":r.nmDifference,"% OF NO MAPPING NOTICES DISPOSED":r.nmPct+"%","NO MAPPING DOCUMENTS UPLOADED":r.nmDocs,"% DOCUMENTS / TOTAL NO MAPPING NOTICES":r.nmDocsPct+"%","TOTAL DISCREPANCY NOTICES":r.dTotal,"DISCREPANCY DISPOSED EARLIER":r.dEarlier,"DISCREPANCY DISPOSED LATEST":r.dLatest,"DISCREPANCY DIFFERENCE":r.dDifference,"% OF DISCREPANCY NOTICES DISPOSED":r.dPct+"%","BLO LETTER UPLOADED":r.dLetters,"% BLO LETTER / TOTAL DISCREPANCY NOTICES":r.dLettersPct+"%"})):psRows.map((r,i)=>({"S.No.":i+1,"PS NO.":r.psNo,"OFFICER NAME AND DESIGNATION":formatOfficer(r),"BLO SUPERVISOR":r.supervisor,"BLO NAME":r.blo,"TOTAL NOTICES":r.totalNotices,"TOTAL NO MAPPING NOTICES":r.nmTotal,"NO MAPPING DISPOSED EARLIER":r.nmEarlier,"NO MAPPING DISPOSED LATEST":r.nmLatest,"DIFFERENCE":r.nmDifference,"% OF NO MAPPING NOTICES DISPOSED":r.nmPct+"%","NO MAPPING DOCUMENTS UPLOADED":r.nmDocs,"% DOCUMENTS / TOTAL NO MAPPING NOTICES":r.nmDocsPct+"%","TOTAL DISCREPANCY NOTICES":r.dTotal,"DISCREPANCY DISPOSED EARLIER":r.dEarlier,"DISCREPANCY DISPOSED LATEST":r.dLatest,"DISCREPANCY DIFFERENCE":r.dDifference,"% OF DISCREPANCY NOTICES DISPOSED":r.dPct+"%","BLO LETTER UPLOADED":r.dLetters,"% BLO LETTER / TOTAL DISCREPANCY NOTICES":r.dLettersPct+"%"}));
   const ws=XLSX.utils.json_to_sheet(data);ws["!freeze"]={xSplit:0,ySplit:1};const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,report==="officer"?"Officer Wise":"PS Wise");XLSX.writeFile(wb,`${report==="officer"?"Officer_Wise_Report":"PS_Wise_Report"}_${new Date().toISOString().slice(0,10)}.xlsx`);
 };

 const fileBox=(which:"earlier"|"latest",value:ParsedReport|null,setManual:(s:string)=>void,manual:string)=><div className="upload"><div className="upload-title"><span>{which==="earlier"?"1":"2"}</span><div><b>{which==="earlier"?"EARLIER REPORT":"LATEST REPORT"}</b><small>{value?.fileName||"Drop Excel file here or click to browse"}</small></div></div><div className="drop" onDragOver={e=>e.preventDefault()} onDrop={drop(which,manual)}><Upload size={18}/><label>Choose .xlsx / .xls<input type="file" accept=".xlsx,.xls" onChange={e=>{const f=e.target.files?.[0];if(f)loadFile(f,which,manual)}}/></label></div><input className="manual" placeholder="Optional date/time" value={manual} onChange={e=>setManual(e.target.value)}/>{value&&<div className="fileok"><CheckCircle2 size={14}/> {value.timestamp} {value.validation.matched&&<span>Validated</span>}</div>}</div>;

 const sortHeader=(label:string,key:string,earlier=false)=><th className={earlier?"earlier":""} onClick={()=>sort(key)}>{label}<ArrowUpDown size={11}/></th>;
 return <main>
  <Header report={report} setReport={setReport}/>
  <div className="meta">Data as of {latestTimestamp} | {officerCount} Officers | {psCount} Polling Stations</div>
  <section className="uploads">{fileBox("earlier",earlier,setManualEarlier,manualEarlier)}{fileBox("latest",latest,setManualLatest,manualLatest)}</section>
  {error&&<div className="error"><AlertTriangle size={17}/>{error}</div>}
  {latest&&earlier&&<><section className="kpis">
    <K label="TOTAL NOTICES" value={formatNumber(total.totalNotices)} tone="navy"/>
    <K label="NO MAPPING DISPOSED (LATEST)" value={formatNumber(total.nmLatest)} tone="navy"/>
    <K label="DISCREPANCY DISPOSED (LATEST)" value={formatNumber(total.dLatest)} tone="brown"/>
    <K label="IMPROVEMENT SINCE EARLIER" value={`${total.nmDifference+total.dDifference>=0?"+":""}${formatNumber(total.nmDifference+total.dDifference)}`} tone="green"/>
    <K label="NO MAPPING DOCUMENTS UPLOADED" value={formatNumber(total.nmDocs)} tone="navy"/>
    <K label="BLO LETTERS UPLOADED" value={formatNumber(total.dLetters)} tone="brown"/>
  </section>
  <div className="controls"><div><label><input type="checkbox" checked={false} disabled/> Upload percentage colour coding</label><small>Documents ÷ Total No Mapping Notices | BLO Letters ÷ Total Discrepancy Notices</small></div><div className="actions"><button onClick={excel}><Download size={15}/> Download Excel</button><button className="pdf" onClick={pdf}><FileText size={15}/> Download PDF</button></div></div>
  <section className="report"><div className="reporthead"><div><b>{report==="officer"?"OFFICER WISE PROGRESS REPORT":"PS WISE PROGRESS REPORT"}</b><span>Earlier vs Latest Disposal Status</span></div><div className={latest.validation.matched?"validated":"notvalidated"}>{latest.validation.matched?<><CheckCircle2 size={14}/> GRAND TOTAL VALIDATED</>:<><AlertTriangle size={14}/> GRAND TOTAL NOT MATCHED</>}</div></div>
   <div className="table-scroll"><table className="reporttable"><thead><tr className="sectionrow"><th colSpan={report==="officer"?3:6}></th><th colSpan={6}>NO MAPPING NOTICES</th><th className="separator"></th><th colSpan={7}>DISCREPANCY NOTICES</th></tr><tr>
     {report==="officer"?<>{sortHeader("S.No.","sno")}{sortHeader("OFFICER NAME AND DESIGNATION","officer")}{sortHeader("TOTAL NOTICES","totalNotices")}{sortHeader("TOTAL NO MAPPING NOTICES","nmTotal")}{sortHeader("NO MAPPING DISPOSED EARLIER","nmEarlier",true)}{sortHeader("NO MAPPING DISPOSED LATEST","nmLatest")}{sortHeader("DIFFERENCE","nmDifference")}{sortHeader("% OF NO MAPPING NOTICES DISPOSED","nmPct")}{sortHeader("NO MAPPING DOCUMENTS UPLOADED","nmDocs")}{<th className="separator"></th>}{sortHeader("TOTAL DISCREPANCY NOTICES","dTotal")}{sortHeader("DISCREPANCY DISPOSED EARLIER","dEarlier",true)}{sortHeader("DISCREPANCY DISPOSED LATEST","dLatest")}{sortHeader("DIFFERENCE","dDifference")}{sortHeader("% OF DISCREPANCY NOTICES DISPOSED","dPct")}{sortHeader("BLO LETTER UPLOADED","dLetters")}</>:<>{sortHeader("S.No.","sno")}{sortHeader("PS NO.","psNo")}{sortHeader("OFFICER NAME AND DESIGNATION","officer")}{sortHeader("BLO SUPERVISOR","supervisor")}{sortHeader("BLO NAME","blo")}{sortHeader("TOTAL NOTICES","totalNotices")}{sortHeader("TOTAL NO MAPPING NOTICES","nmTotal")}{sortHeader("NO MAPPING DISPOSED EARLIER","nmEarlier",true)}{sortHeader("NO MAPPING DISPOSED LATEST","nmLatest")}{sortHeader("DIFFERENCE","nmDifference")}{sortHeader("% OF NO MAPPING NOTICES DISPOSED","nmPct")}{sortHeader("NO MAPPING DOCUMENTS UPLOADED","nmDocs")}{<th className="separator"></th>}{sortHeader("TOTAL DISCREPANCY NOTICES","dTotal")}{sortHeader("DISCREPANCY DISPOSED EARLIER","dEarlier",true)}{sortHeader("DISCREPANCY DISPOSED LATEST","dLatest")}{sortHeader("DIFFERENCE","dDifference")}{sortHeader("% OF DISCREPANCY NOTICES DISPOSED","dPct")}{sortHeader("BLO LETTER UPLOADED","dLetters")}</>}
   </tr></thead>
   <tbody>{active.map((r:any,i:number)=><tr key={r.officer||r.psNo}><td>{i+1}</td>{report==="officer"?<><td className="officer">{formatOfficer(r)}</td><td>{formatNumber(r.totalNotices)}</td><td>{formatNumber(r.nmTotal)}</td><td className="earlier">{formatNumber(r.nmEarlier)}</td><td className="latest">{formatNumber(r.nmLatest)}</td><td className={"colour "+nmGroups.get(r.officer)}>{formatNumber(r.nmDifference)}</td><td>{r.nmPct}%</td><td >{formatNumber(r.nmDocs)}</td><td className={"colour "+docsPctGroups.get(r.officer)}>{r.nmDocsPct}%</td><td className="separator"></td><td>{formatNumber(r.dTotal)}</td><td className="earlier">{formatNumber(r.dEarlier)}</td><td className="latest">{formatNumber(r.dLatest)}</td><td className={"colour "+discGroups.get(r.officer)}>{formatNumber(r.dDifference)}</td><td>{r.dPct}%</td><td >{formatNumber(r.dLetters)}</td><td className={"colour "+lettersPctGroups.get(r.officer)}>{r.dLettersPct}%</td></>:<><td>{r.psNo}</td><td className="officer">{formatOfficer(r)}</td><td>{r.supervisor}</td><td>{r.blo}</td><td>{formatNumber(r.totalNotices)}</td><td>{formatNumber(r.nmTotal)}</td><td className="earlier">{formatNumber(r.nmEarlier)}</td><td className="latest">{formatNumber(r.nmLatest)}</td><td className={"colour "+nmGroups.get(r.psNo)}>{formatNumber(r.nmDifference)}</td><td>{r.nmPct}%</td><td >{formatNumber(r.nmDocs)}</td><td className={"colour "+docsPctGroups.get(r.psNo)}>{r.nmDocsPct}%</td><td className="separator"></td><td>{formatNumber(r.dTotal)}</td><td className="earlier">{formatNumber(r.dEarlier)}</td><td className="latest">{formatNumber(r.dLatest)}</td><td className={"colour "+discGroups.get(r.psNo)}>{formatNumber(r.dDifference)}</td><td>{r.dPct}%</td><td >{formatNumber(r.dLetters)}</td><td className={"colour "+lettersPctGroups.get(r.psNo)}>{r.dLettersPct}%</td></>}</tr>)}
   <tr className="grand">{report==="officer"?<><td></td><td>GRAND TOTAL</td><td>{formatNumber(total.totalNotices)}</td><td>{formatNumber(total.nmTotal)}</td><td>{formatNumber(total.nmEarlier)}</td><td>{formatNumber(total.nmLatest)}</td><td>{formatNumber(total.nmDifference)}</td><td>{total.nmPct}%</td><td>{formatNumber(total.nmDocs)}</td><td className="colour">{total.nmDocsPct}%</td><td className="separator"></td><td>{formatNumber(total.dTotal)}</td><td>{formatNumber(total.dEarlier)}</td><td>{formatNumber(total.dLatest)}</td><td>{formatNumber(total.dDifference)}</td><td>{total.dPct}%</td><td>{formatNumber(total.dLetters)}</td><td className="colour">{total.dLettersPct}%</td></>:<><td></td><td>GRAND TOTAL</td><td></td><td></td><td></td><td>{formatNumber(total.totalNotices)}</td><td>{formatNumber(total.nmTotal)}</td><td>{formatNumber(total.nmEarlier)}</td><td>{formatNumber(total.nmLatest)}</td><td>{formatNumber(total.nmDifference)}</td><td>{total.nmPct}%</td><td>{formatNumber(total.nmDocs)}</td><td className="separator"></td><td>{formatNumber(total.dTotal)}</td><td>{formatNumber(total.dEarlier)}</td><td>{formatNumber(total.dLatest)}</td><td>{formatNumber(total.dDifference)}</td><td>{total.dPct}%</td><td>{formatNumber(total.dLetters)}</td><td className="colour">{total.dLettersPct}%</td></>}</tr>
   </tbody></table></div>
   <div className="legend"><b>COLOUR CODE:</b><span className="legendlow">LOW (needs attention)</span><span className="legendmid">MEDIUM</span><span className="legendhigh">HIGH (good)</span><span>Applied on Difference and upload-percentage columns. Lowest third = red.</span></div>
   <div className="notes">Notes: "Disposed" means Approved / Disposed notices. % Disposed = Latest Disposed / Total Notices of that type. Difference = Latest - Earlier.</div>
  </section>
  {report==="officer"&&<section className="charts"><Chart title="DIFFERENCE BY OFFICER" rows={officerRows} field="nmDifference"/><Chart title="% DISPOSED BY OFFICER" rows={officerRows} field="nmPct" percent/></section>}
  </>}
  {!latest||!earlier?<div className="empty"><FileText size={30}/><b>Upload Earlier and Latest Excel reports</b><span>Only the FINAL PS WISE REPORT sheet is used.</span></div>:null}
 <style jsx global>{`
@page{size:A3 landscape;margin:8mm}*{box-sizing:border-box}body{margin:0;background:#f3f5f8;color:#172033;font-family:Arial,Helvetica,sans-serif}.topbar{height:9px;background:#1f3864}.titleblock{background:#fff;text-align:center;padding:14px 20px 9px;border-bottom:1px solid #b8c0cc}.titleblock h1{font-size:30px;margin:0;color:#1f3864;font-weight:800}.subtitle{font-size:15px;margin-top:4px;color:#344054}.switch{position:absolute;right:18px;top:25px;display:flex;gap:4px}.switch button{border:1px solid #b8c0cc;background:#fff;color:#1f3864;padding:7px 10px;font-size:11px;font-weight:700}.switch button.active{background:#1f3864;color:#fff}.meta{text-align:center;background:#fff;color:#475467;font-size:13px;padding:3px 0 10px}.uploads{display:grid;grid-template-columns:1fr 1fr;gap:12px;padding:12px 22px}.upload{background:#fff;border:1px solid #b8c0cc;padding:12px;border-radius:4px}.upload-title{display:flex;gap:9px;align-items:center}.upload-title>span{width:28px;height:28px;background:#1f3864;color:#fff;display:grid;place-items:center;font-weight:800}.upload-title b{display:block;color:#1f3864;font-size:12px}.upload-title small{display:block;color:#667085;margin-top:2px}.drop{height:44px;border:1px dashed #8d9aad;margin-top:8px;display:flex;align-items:center;justify-content:center;gap:7px;color:#344054;font-size:12px;position:relative}.drop input{position:absolute;inset:0;opacity:0;cursor:pointer}.manual{margin-top:7px;border:1px solid #d0d5dd;padding:6px;width:210px;font-size:11px}.fileok{margin-top:6px;color:#20733a;font-size:11px;display:flex;gap:5px;align-items:center}.fileok span{background:#d9ead3;padding:2px 6px;font-weight:700}.error{margin:0 22px 10px;padding:10px;background:#fde2e2;color:#8f2d2d;border:1px solid #e5a4a4;display:flex;gap:7px;align-items:center;font-weight:700}.kpis{display:grid;grid-template-columns:repeat(6,1fr);gap:9px;padding:0 22px 10px}.kpi{background:#fff;border:1px solid #b8c0cc;border-top:4px solid #1f3864;padding:10px 12px}.kpi.brown{border-top-color:#833c0b}.kpi.green{border-top-color:#4f7f3f}.kpi span{display:block;font-size:10px;color:#475467;font-weight:800}.kpi b{font-size:22px;display:block;margin-top:3px}.controls{margin:0 22px 10px;background:#fff;border:1px solid #b8c0cc;padding:8px 11px;display:flex;justify-content:space-between;align-items:center}.controls label{font-size:12px;font-weight:700}.controls small{display:block;color:#667085;font-size:10px;margin-left:20px}.actions{display:flex;gap:7px}.actions button{border:1px solid #1f3864;background:#fff;color:#1f3864;padding:8px 12px;font-weight:800}.actions .pdf{background:#1f3864;color:#fff}.report{margin:0 22px;background:#fff;border:1px solid #b8c0cc}.reporthead{display:flex;justify-content:space-between;padding:9px 11px;border-bottom:1px solid #b8c0cc}.reporthead b{color:#1f3864;font-size:14px}.reporthead span{display:block;color:#667085;font-size:10px;margin-top:2px}.validated,.notvalidated{font-size:10px;font-weight:800;display:flex;gap:5px;align-items:center;padding:4px 7px}.validated{background:#d9ead3;color:#236b36}.notvalidated{background:#fde2e2;color:#8f2d2d}.table-scroll{overflow:auto}.reporttable{border-collapse:collapse;width:100%;min-width:1550px;table-layout:auto}.reporttable th,.reporttable td{border:1px solid #b8c0cc;padding:7px 6px;text-align:center;font-size:13px;line-height:1.1;white-space:nowrap}.reporttable th{background:#1f3864;color:#fff;font-size:11px;font-weight:800;position:sticky;top:0;z-index:2}.reporttable th svg{vertical-align:middle;margin-left:3px}.reporttable th:nth-last-child(-n+6){background:#833c0b}.reporttable .sectionrow th{position:static;font-size:11px;padding:5px}.reporttable .sectionrow th:nth-child(2){background:#1f3864}.reporttable .sectionrow th:last-child{background:#833c0b}.reporttable td.officer{text-align:left;font-weight:700}.reporttable tbody tr:nth-child(even){background:#f2f5fa}.reporttable td.earlier{background:#fff8e1}.reporttable td.latest{font-weight:800}.reporttable td.colour{font-weight:800}.reporttable td.low{background:#f8b4b4;color:#20242b}.reporttable td.medium{background:#ffe699;color:#20242b}.reporttable td.high{background:#b7e1a1;color:#20242b}.separator{width:7px;min-width:7px;padding:0!important;background:#fff!important;border-left:0!important;border-right:0!important}.grand td{background:#d9e1f2!important;font-weight:800}.legend{padding:8px 10px;display:flex;align-items:center;gap:8px;border-top:1px solid #b8c0cc;font-size:10px}.legendlow,.legendmid,.legendhigh{padding:4px 12px;font-weight:800;border:1px solid #b8c0cc}.legendlow{background:#f8b4b4}.legendmid{background:#ffe699}.legendhigh{background:#b7e1a1}.notes{font-size:10px;color:#475467;padding:2px 10px 9px}.charts{display:grid;grid-template-columns:1fr 1fr;gap:12px;margin:12px 22px 30px}.chart{background:#fff;border:1px solid #b8c0cc;padding:12px}.chart h3{font-size:12px;color:#1f3864;margin:0 0 8px}.barrow{display:grid;grid-template-columns:145px 1fr 55px;gap:7px;align-items:center;margin:4px 0;font-size:10px}.barrow label{overflow:hidden;text-overflow:ellipsis;white-space:nowrap}.barrow div{height:10px;background:#edf0f4}.barrow i{display:block;height:100%;background:#1f3864}.barrow b{text-align:right}.empty{margin:20px 22px;padding:40px;background:#fff;border:1px solid #b8c0cc;display:flex;flex-direction:column;align-items:center;gap:8px;color:#667085}.empty b{color:#1f3864}@media(max-width:1000px){.switch{position:static;justify-content:center;margin-bottom:8px}.titleblock h1{font-size:23px}.uploads,.kpis{grid-template-columns:1fr}.controls{align-items:flex-start;gap:10px;flex-direction:column}.charts{grid-template-columns:1fr}.reporttable th,.reporttable td{font-size:12px}}@media print{body{background:#fff}.uploads,.controls,.switch,.charts{display:none!important}.report{margin:0}.titleblock{padding-top:8px}}`}</style></main>;
}

function K({label,value,tone}:{label:string;value:string;tone:string}){return <div className={"kpi "+tone}><span>{label}</span><b>{value}</b></div>}
function Chart({title,rows,field,percent=false}:{title:string;rows:any[];field:string;percent?:boolean}){const max=Math.max(...rows.map(r=>Number(r[field])||0),1);return <div className="chart"><h3>{title}</h3><div>{rows.map(r=><div className="barrow" key={r.officer}><label>{r.officer}</label><div><i style={{width:`${Math.max(2,(Number(r[field])||0)/max*100)}%`}}></i></div><b>{percent?r[field]+"%":formatNumber(r[field])}</b></div>)}</div></div>}


[executed on device: Akashs-MacBook-Air.local (4036675e-2011-40ea-be95-0a3a36abe389)]