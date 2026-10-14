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

const color=(v:any):[number,number,number]=>{
  if(Array.isArray(v))return v as [number,number,number];
  return [255,235,132];
};
const rgbCss=(rgb:[number,number,number])=>"rgb("+rgb.join(",")+")";
// Fixed percentage bands requested for all percentage cells (grand-total row stays unchanged).
// Shared fixed-range palette for both the web preview and exported PDF.
const percentageColor=(value:number):[number,number,number]=>{
  if(value<20)return [201,76,76];       // #C94C4C deep red
  if(value<30)return [229,115,115];     // #E57373 soft red
  if(value<40)return [237,138,90];      // #ED8A5A muted coral
  if(value<60)return [242,176,94];      // #F2B05E soft orange
  if(value<70)return [243,214,117];     // #F3D675 warm yellow
  if(value<80)return [209,220,120];     // #D1DC78 yellow green
  if(value<90)return [145,201,139];     // #91C98B soft light green
  return [63,166,107];                  // #3FA66B professional green
};
const percentageTextColor=(value:number)=>value<20||value>=90?"#FFFFFF":"#111827";
// Parse numeric percentages consistently, including strings such as "80.5%".
const getPercentageStyle=(input:any)=>{
  const value=Number.parseFloat(String(input??"").replace("%","").trim());
  if(!Number.isFinite(value))return {backgroundColor:"#FFFFFF",color:"#1F2937"};
  const rgb=percentageColor(value);
  const backgroundColor="#"+rgb.map(v=>v.toString(16).padStart(2,"0")).join("").toUpperCase();
  return {backgroundColor,color:percentageTextColor(value)};
};
const hexRgb=(hex:string):[number,number,number]=>{
  const h=hex.replace("#","");
  return [parseInt(h.slice(0,2),16),parseInt(h.slice(2,4),16),parseInt(h.slice(4,6),16)];
};

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
  const isAnuja=(r:any)=>/anuja\s+trivedi/i.test(String(r.officer||""));
  const ap=isAnuja(a),bp=isAnuja(b);
  if(ap&&!bp)return -1;
  if(!ap&&bp)return 1;
  if(sortKey==="officer"||sortKey==="designation") return String(a[sortKey]).localeCompare(String(b[sortKey]))*sortDir;
  return ((Number(a[sortKey])||0)-(Number(b[sortKey])||0))*sortDir;
 }),[reports,sortKey,sortDir]);
 const pdfRows=useMemo(()=>[...reports].sort((a:any,b:any)=>{
  const isAnuja=(r:any)=>/anuja\s+trivedi/i.test(String(r.officer||""));
  const ap=isAnuja(a),bp=isAnuja(b);
  if(ap&&!bp)return -1;
  if(!ap&&bp)return 1;
  return String(a.officer||"").localeCompare(String(b.officer||""),undefined,{sensitivity:"base",numeric:false});
 }),[reports]);
 const total=grandOfficer(rows);
 const psCount=new Set((latest?.rows||[]).map(r=>r.psNo)).size;
 const groups={
  nm:useMemo(()=>percentileGroups(rows,r=>r.nmDifference),[rows]),
  disc:useMemo(()=>percentileGroups(rows,r=>r.dDifference),[rows]),
  totalDisposed:useMemo(()=>percentileGroups(rows,r=>r.totalDisposedPct),[rows])
 };

 const sort=(key:keyof OfficerReportRow)=>{
  if(sortKey===key)setSortDir(v=>-v);else{setSortKey(key);setSortDir(-1);}
 };
 const percentage=(a:number,b:number)=>b?Math.round((a/b)*1000)/10:0;

 const buildOfficerPSRows=(officer:string)=>{
  const L=(latest?.rows||[]).filter(r=>r.officer===officer && String(r.psNo??"").trim()!=="").sort((a,b)=>String(a.psNo).localeCompare(String(b.psNo),undefined,{numeric:true}));
  const E=new Map((earlier?.rows||[]).filter(r=>r.officer===officer).map(r=>[r.psNo,r]));
  return L.map((r,i)=>{
   const e=E.get(r.psNo);
   const totalDisposed=r.nmDisposed+r.dDisposed;
   const nmPending=r.nmTotal-r.nmDisposed;
   const dPending=r.dTotal-r.dDisposed;
   return [i+1,r.psNo,r.blo,r.supervisor,r.nmTotal+r.dTotal,r.nmTotal,r.nmDisposed,r.nmDisposed-(e?.nmDisposed||0),percentage(r.nmDisposed,r.nmTotal).toFixed(1)+"%",nmPending,r.nmDocs,r.dTotal,r.dDisposed,r.dDisposed-(e?.dDisposed||0),percentage(r.dDisposed,r.dTotal).toFixed(1)+"%",dPending,r.dLetter,totalDisposed,percentage(totalDisposed,r.nmTotal+r.dTotal).toFixed(1)+"%"];
  });
 };
 const buildOfficerPSTotal=(officer:string)=>{
  const L=(latest?.rows||[]).filter(r=>r.officer===officer && String(r.psNo??"").trim()!=="");
  const E=new Map((earlier?.rows||[]).filter(r=>r.officer===officer).map(r=>[r.psNo,r]));
  const psCount=new Set(L.map(r=>String(r.psNo).trim()).filter(Boolean)).size;
  const sum=(fn:(r:(typeof L)[number])=>number)=>L.reduce((a,r)=>a+(Number(fn(r))||0),0);
  const totalNotices=sum(r=>r.nmTotal+r.dTotal);
  const nmTotal=sum(r=>r.nmTotal),nmLatest=sum(r=>r.nmDisposed),nmDocs=sum(r=>r.nmDocs);
  const nmDifference=L.reduce((a,r)=>a+r.nmDisposed-(E.get(r.psNo)?.nmDisposed||0),0);
  const nmPending=sum(r=>r.nmTotal-r.nmDisposed);
  const nmEarlier=L.reduce((a,r)=>a+(E.get(r.psNo)?.nmDisposed||0),0);
  const dTotal=sum(r=>r.dTotal),dLatest=sum(r=>r.dDisposed),dLetters=sum(r=>r.dLetter);
  const dDifference=L.reduce((a,r)=>a+r.dDisposed-(E.get(r.psNo)?.dDisposed||0),0);
  const dEarlier=L.reduce((a,r)=>a+(E.get(r.psNo)?.dDisposed||0),0);
  const dPending=sum(r=>r.dTotal-r.dDisposed);
  const totalDisposed=nmLatest+dLatest,totalPending=totalNotices-totalDisposed;
  const nmPct=nmTotal>0?nmLatest/nmTotal*100:0;
  const dPct=dTotal>0?dLatest/dTotal*100:0;
  const totalPct=totalNotices>0?totalDisposed/totalNotices*100:0;
  const summary=reports.find(r=>r.officer===officer);
  const checks:Record<string,[number,number]> = {
   "NO. OF PS":[psCount,summary?.psCount||0],"TOTAL NOTICES":[totalNotices,summary?.totalNotices||0],
   "TOTAL NO MAPPING":[nmTotal,summary?.nmTotal||0],"NM DISP. LATEST":[nmLatest,summary?.nmLatest||0],
   "NM DIFF.":[nmDifference,summary?.nmDifference||0],"PENDING NO MAPPING CASES":[nmPending,summary?.nmPending||0],
   "NM DOCS UPLOADED":[nmDocs,summary?.nmDocs||0],"TOTAL DISCREPANCY":[dTotal,summary?.dTotal||0],
   "DISC. DISP. LATEST":[dLatest,summary?.dLatest||0],"DISC. DIFF.":[dDifference,summary?.dDifference||0],
   "PENDING DISCREPANCY CASES":[dPending,summary?.dPending||0],"BLO LETTER UPLOADED":[dLetters,summary?.dLetters||0],
   "TOTAL NOTICES DISPOSED":[totalDisposed,summary?.totalDisposed||0],"TOTAL NOTICES PENDING":[totalPending,summary?.totalPending||0],
   "NM DISPOSED PERCENT":[percentage(nmLatest,nmTotal),summary?.nmPct||0],
   "DISC. DISPOSED PERCENT":[percentage(dLatest,dTotal),summary?.dPct||0],
   "TOTAL DISPOSED PERCENT":[percentage(totalDisposed,totalNotices),summary?.totalDisposedPct||0]
  };
  const mismatches=Object.entries(checks).filter(([,v])=>v[0]!==v[1]);
  return {psCount,totalNotices,nmTotal,nmLatest,nmDifference,nmPending,nmDocs,nmEarlier,dTotal,dLatest,dDifference,dPending,dLetters,dEarlier,totalDisposed,totalPending,nmPct,dPct,totalPct,mismatches,
   pdfRow:[psCount,{content:"TOTAL - "+officer,colSpan:3},totalNotices,nmTotal,nmLatest,nmDifference,nmPct.toFixed(1)+"%",nmPending,nmDocs,dTotal,dLatest,dDifference,dPct.toFixed(1)+"%",dPending,dLetters,totalDisposed,totalPct.toFixed(1)+"%"]};
 };

 const pdf=()=>{
  if(!latest||!earlier)return;
  const integrity=pdfRows.map(r=>({officer:r.officer,total:buildOfficerPSTotal(r.officer)})).filter(x=>x.total.mismatches.length);
  if(integrity.length){
   const details=integrity.map(x=>x.officer+": "+x.total.mismatches.map(([k,v])=>k+" (PS total "+v[0]+", summary "+v[1]+")").join(", ")).join(" | ");
   setError("PDF not generated: officer PS-wise totals do not match the Officer-Wise Summary. "+details);
   return;
  }
  setError("");
  const doc=new jsPDF({orientation:"landscape",unit:"mm",format:"a3"});
  const navy:[number,number,number]=[31,56,100],brown:[number,number,number]=[131,60,11],yellow:[number,number,number]=[255,248,225],grid:[number,number,number]=[184,192,204],grand:[number,number,number]=[217,225,242];

  doc.setFillColor(...navy);doc.rect(0,0,420,18,"F");
  doc.setTextColor(255,255,255);doc.setFont("helvetica","bold");
  doc.setFontSize(16);doc.text("OFFICE OF THE ELECTORAL REGISTRATION OFFICER, AC-34, MATIALA",210,7,{align:"center"});
  doc.setFontSize(14);doc.text("SIR-2026: NOTICE DISPOSAL REPORT",210,12,{align:"center"});
  doc.setFontSize(8);doc.text("NO MAPPING & DISCREPANCY NOTICES - OFFICER-WISE (AERO-WISE)",210,16,{align:"center"});
  doc.setTextColor(55,65,80);doc.setFontSize(7.5);
  doc.text("Data as of "+latest.timestamp+" | "+pdfRows.length+" Officers | "+psCount+" Polling Stations",210,23,{align:"center"});

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
    {content:"NO MAPPING NOTICES",colSpan:6},{content:"",rowSpan:2},
    {content:"DISCREPANCY NOTICES",colSpan:6},{content:"",rowSpan:2},{content:"TOTAL DISPOSAL",colSpan:3}
  ];
  const head2=[
    "TOTAL NO MAPPING","NM DISP. LATEST","NM DIFF.","% NM DISPOSED","PENDING NO MAPPING CASES","NM DOCS UPLOADED",
    "TOTAL DISCREPANCY","DISC. DISP. LATEST","DISC. DIFF.","% DISC. DISPOSED","PENDING DISCREPANCY CASES","BLO LETTER UPLOADED",
    "TOTAL NOTICES DISPOSED","% TOTAL DISPOSED","TOTAL NOTICES PENDING"
  ];
  const body=pdfRows.map((r,i)=>[
    i+1,r.officer+"\n("+(r.designation||"Officer")+")",r.psCount,r.totalNotices,
    r.nmTotal,r.nmLatest,r.nmDifference,r.nmPct+"%",r.nmPending,r.nmDocs,
    "",r.dTotal,r.dLatest,r.dDifference,r.dPct+"%",r.dPending,r.dLetters,
    "",r.totalDisposed,r.totalDisposedPct+"%",r.totalPending
  ]);
  body.push(["","GRAND TOTAL - AC-34 MATIALA",psCount,total.totalNotices,total.nmTotal,total.nmLatest,total.nmDifference,total.nmPct+"%",total.nmPending,total.nmDocs,"",total.dTotal,total.dLatest,total.dDifference,total.dPct+"%",total.dPending,total.dLetters,"",total.totalDisposed,total.totalDisposedPct+"%",total.totalPending]);

  const widths=[6,43,9.5,16,20,16,14,16,36.5,19,3,20,17,14,16,36.5,20,3,20,18,21];
  const base={theme:"grid" as const,tableWidth:385,margin:{left:17.5,right:17.5,top:0,bottom:0},styles:{font:"helvetica",fontStyle:"bold" as const,fontSize:9.0,cellPadding:1.0,lineColor:grid,lineWidth:.25,textColor:[20,28,38] as [number,number,number],halign:"center" as const,valign:"middle" as const,overflow:"linebreak" as const},headStyles:{font:"helvetica",fontStyle:"bold" as const,fontSize:7.8,cellPadding:1.0,minCellHeight:11.5,halign:"center" as const,valign:"middle" as const,textColor:255}};
  autoTable(doc,{...base,startY:49,head:[head1,head2],body,columnStyles:Object.fromEntries(widths.map((w,i)=>[i,{cellWidth:w,halign:i===1?"left":"center"}])),didParseCell:(data:any)=>{
    if(data.section==="head"){data.cell.styles.fillColor=data.column.index>=11&&data.column.index<=16?brown:navy;if(data.column.index===10||data.column.index===17){data.cell.styles.fillColor=[255,255,255];data.cell.styles.lineWidth=0;}return;}
    if(data.section!=="body")return;
    if(data.row.index===pdfRows.length){data.cell.styles.fillColor=grand;data.cell.styles.fontStyle="bold";return;}
    const r=pdfRows[data.row.index];if(!r)return;
    if(data.column.index===7||data.column.index===14||data.column.index===19){
      const style=getPercentageStyle(data.cell.raw);
      data.cell.styles.fillColor=hexRgb(style.backgroundColor);
      data.cell.styles.textColor=hexRgb(style.color);
      data.cell.styles.fontStyle="bold";
      data.cell.styles.halign="center";
    }
  }});
  // PS-wise detail section: all concerned PS are listed officer-by-officer in the same report.
  // IMPORTANT: Officer-wise summary table above is intentionally untouched.
  doc.addPage();
  const detailPageHeader=(title:string)=>{
    doc.setFillColor(...navy);doc.rect(0,0,420,13,"F");
    doc.setTextColor(255,255,255);doc.setFont("helvetica","bold");doc.setFontSize(9.5);
    doc.text(title,210,8.2,{align:"center"});
  };
  let detailY=18;
  const psHead=["S.No.","PS No.","BLO NAME","BLO SUPERVISOR","TOTAL NOTICES","TOTAL NO MAPPING","NM DISP. LATEST","NM DIFF.","% NM DISPOSED","PENDING NO MAPPING CASES","NM DOCS UPLOADED","TOTAL DISCREPANCY","DISC. DISP. LATEST","DISC. DIFF.","% DISC. DISPOSED","PENDING DISCREPANCY CASES","BLO LETTER UPLOADED","TOTAL NOTICES DISPOSED","% TOTAL DISPOSED"];
  // Full A3 width: the PS-wise table uses the same visual scale and colour language as the Officer-wise table.
  const psWidths=[7,12,34,34,20,20,20,16,15,24,20,20,20,16,15,24,20,22,16];
  const psGroup=[
    {content:"",colSpan:5,rowSpan:1},
    {content:"NO MAPPING NOTICES",colSpan:6},
    {content:"DISCREPANCY NOTICES",colSpan:6},
    {content:"TOTAL DISPOSAL",colSpan:2}
  ];
  const allPS=pdfRows.flatMap(r=>buildOfficerPSRows(r.officer));
  pdfRows.forEach((officerRow,oi)=>{
    const psRows=buildOfficerPSRows(officerRow.officer);
    // Each officer starts a distinct section. If a long table continues, the officer heading repeats.
    if(oi>0) doc.addPage();
    const officerTitle=(oi+1)+". "+officerRow.officer+" ("+(officerRow.designation||"Officer")+") — "+officerRow.psCount+" PS";
    detailPageHeader(officerTitle);
    detailY=18;
    autoTable(doc,{
      theme:"grid",startY:detailY,head:[psGroup,psHead],body:[...psRows,buildOfficerPSTotal(officerRow.officer).pdfRow],
      tableWidth:385,margin:{left:17.5,right:17.5,top:18,bottom:12},
      showHead:"everyPage",pageBreak:"auto",rowPageBreak:"avoid",
      styles:{font:"helvetica",fontStyle:"bold" as const,fontSize:8.5,cellPadding:1.0,lineColor:grid,lineWidth:.25,textColor:[20,28,38] as [number,number,number],halign:"center" as const,valign:"middle" as const,overflow:"linebreak" as const},
      headStyles:{font:"helvetica",fontStyle:"bold" as const,fontSize:7.4,cellPadding:1.0,minCellHeight:10.5,fillColor:navy,textColor:255,halign:"center" as const,valign:"middle" as const},
      columnStyles:Object.fromEntries(psWidths.map((w,i)=>[i,{cellWidth:w,halign:i===2||i===3?"left":"center"}])),
      willDrawPage:()=>{
        detailPageHeader(officerTitle);
      },
      didParseCell:(data:any)=>{
        if(data.section==="head"){
          data.cell.styles.fillColor=(data.row.index===0&&data.column.index>=5&&data.column.index<=10)||(data.row.index===1&&data.column.index>=5&&data.column.index<=10)?navy:
            (data.row.index===0&&data.column.index>=11&&data.column.index<=16)||(data.row.index===1&&data.column.index>=11&&data.column.index<=16)?brown:navy;
          return;
        }
        if(data.section!=="body")return;
        if(data.row.index===psRows.length){
          data.cell.styles.fillColor=grand;
          data.cell.styles.textColor=navy;
          data.cell.styles.fontStyle="bold";
          data.cell.styles.lineColor=navy;
          data.cell.styles.lineWidth={top:0.6,right:0.25,bottom:0.25,left:0.25};
          data.cell.styles.minCellHeight=9;
          if(data.column.index===8||data.column.index===14||data.column.index===18){
            const style=getPercentageStyle(data.cell.raw);
            data.cell.styles.fillColor=hexRgb(style.backgroundColor);
            data.cell.styles.textColor=hexRgb(style.color);
            data.cell.styles.fontStyle="bold";
          }
          return;
        }
        const p=psRows[data.row.index] as any[]|undefined;
        if(!p)return;
        if(data.column.index===8||data.column.index===14||data.column.index===18){
          const style=getPercentageStyle(data.cell.raw);
          data.cell.styles.fillColor=hexRgb(style.backgroundColor);
          data.cell.styles.textColor=hexRgb(style.color);
          data.cell.styles.fontStyle="bold";
        }
      }
    });
    detailY=((doc as any).lastAutoTable?.finalY||detailY+10)+7;
  });
  const fy=((doc as any).lastAutoTable?.finalY||120)+10;
  const pageHeight=doc.internal.pageSize.getHeight();
  if(fy>pageHeight-24) { doc.addPage(); }
  const legendY=(doc.internal.pages.length-1)>1 && fy>pageHeight-24?18:fy;
  doc.setFont("helvetica","bold");doc.setFontSize(7);doc.setTextColor(40,48,58);doc.text("COLOUR CODE:",10,legendY);
  let x=35;[["LOW","low"],["MEDIUM","medium"],["HIGH","high"]].forEach(([label,k])=>{const c=color(k);doc.setFillColor(c[0],c[1],c[2]);doc.rect(x,legendY-3.5,23,4,"F");doc.setTextColor(40,48,58);doc.text(label,x+25,legendY);x+=50;});
  doc.setFont("helvetica","normal");doc.setFontSize(6.5);doc.text("Colour coding applies only to the three disposal percentage columns. Lowest third = red; highest third = green.",185,legendY);
  doc.text("Pending NM = Total No Mapping Notices - Latest NM Disposed | Pending Disc. = Total Discrepancy Notices - Latest Disposed | Total Pending = Pending NM + Pending Disc.",10,legendY+7);
  const pageCount=doc.internal.pages.length-1;
  for(let pno=2;pno<=pageCount;pno++){
    doc.setPage(pno);doc.setFont("helvetica","normal");doc.setFontSize(6);doc.setTextColor(40,48,58);
    doc.text("AC-34 MATIALA | SIR-2026 | OFFICER-WISE NOTICE DISPOSAL REPORT | Data as of "+latest.timestamp,10,pageHeight-5);
    doc.text("Page "+pno+" of "+pageCount,410,pageHeight-5,{align:"right"});
  }
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
    {rows.map((r,i)=>{const vals=[i+1,formatOfficer(r),r.psCount,r.totalNotices,r.nmTotal,r.nmEarlier,r.nmLatest,r.nmDifference,r.nmPct+"%",r.nmDocs,r.nmPending,"",r.dTotal,r.dEarlier,r.dLatest,r.dDifference,r.dPct+"%",r.dLetters,r.dPending,"",r.totalDisposed,r.totalDisposedPct+"%",r.totalPending];return <React.Fragment key={r.officer}><tr>{vals.map((v,j)=>{if(j===11||j===19)return <td key={j} className="sep"/>;const pctValue=j===8?r.nmPct:j===16?r.dPct:j===21?r.totalDisposedPct:null;const k=j===7?groups.nm.get(r.officer):j===15?groups.disc.get(r.officer):"";const cls=[j===1?"name":"",j===4||j===12?"earlier":"",k?"colour "+k:""].join(" ");const pctStyle=pctValue===null?{}:{background:rgbCss(percentageColor(pctValue)),color:percentageTextColor(pctValue),fontWeight:850,textAlign:"center" as const};return <td key={j} className={cls} style={pctStyle}>{v}</td>})}</tr><tr className="psDetailRow"><td colSpan={23}><div className="psTitle">PS-WISE DETAILS — {formatOfficer(r)}</div><div className="psTableWrap"><table className="psTable"><thead><tr>{["S.No.","PS No.","BLO NAME","BLO SUPERVISOR","TOTAL NOTICES","TOTAL NO MAPPING","NM DISP. LATEST","NM DIFF.","% NM DISPOSED","PENDING NO MAPPING CASES","NM DOCS UPLOADED","TOTAL DISCREPANCY","DISC. DISP. LATEST","DISC. DIFF.","% DISC. DISPOSED","PENDING DISCREPANCY CASES","BLO LETTER UPLOADED","TOTAL NOTICES DISPOSED","% TOTAL DISPOSED"].map(h=><th key={h}>{h}</th>)}</tr></thead><tbody>{buildOfficerPSRows(r.officer).map((p:any)=><tr key={String(p[1])}>{p.map((v:any,j:number)=>{const pct=j===8||j===14||j===18?Number(String(v).replace("%","")):null;return <td key={j} style={pct===null?{}:{background:rgbCss(percentageColor(pct)),color:percentageTextColor(pct),fontWeight:850,textAlign:"center" as const}}>{v}</td>})}</tr>)}{(()=>{const t=buildOfficerPSTotal(r.officer);return <tr className="officerTotalRow"><td>{t.psCount}</td><td colSpan={3}>TOTAL - {r.officer}</td><td>{formatNumber(t.totalNotices)}</td><td>{formatNumber(t.nmTotal)}</td><td>{formatNumber(t.nmLatest)}</td><td>{formatNumber(t.nmDifference)}</td><td className="percentageCell" style={{background:rgbCss(percentageColor(t.nmPct)),color:percentageTextColor(t.nmPct)}}>{t.nmPct.toFixed(1)}%</td><td>{formatNumber(t.nmPending)}</td><td>{formatNumber(t.nmDocs)}</td><td>{formatNumber(t.dTotal)}</td><td>{formatNumber(t.dLatest)}</td><td>{formatNumber(t.dDifference)}</td><td className="percentageCell" style={{background:rgbCss(percentageColor(t.dPct)),color:percentageTextColor(t.dPct)}}>{t.dPct.toFixed(1)}%</td><td>{formatNumber(t.dPending)}</td><td>{formatNumber(t.dLetters)}</td><td>{formatNumber(t.totalDisposed)}</td><td className="percentageCell" style={{background:rgbCss(percentageColor(t.totalPct)),color:percentageTextColor(t.totalPct)}}>{t.totalPct.toFixed(1)}%</td></tr>})()}</tbody></table></div></td></tr></React.Fragment>})}
    <tr className="grand"><td></td><td>GRAND TOTAL - AC-34 MATIALA</td><td>{psCount}</td><td>{formatNumber(total.totalNotices)}</td><td>{formatNumber(total.nmTotal)}</td><td>{formatNumber(total.nmEarlier)}</td><td>{formatNumber(total.nmLatest)}</td><td>{formatNumber(total.nmDifference)}</td><td>{total.nmPct}%</td><td>{formatNumber(total.nmDocs)}</td><td>{formatNumber(total.nmPending)}</td><td className="sep"></td><td>{formatNumber(total.dTotal)}</td><td>{formatNumber(total.dEarlier)}</td><td>{formatNumber(total.dLatest)}</td><td>{formatNumber(total.dDifference)}</td><td>{total.dPct}%</td><td>{formatNumber(total.dLetters)}</td><td>{formatNumber(total.dPending)}</td><td className="sep"></td><td>{formatNumber(total.totalDisposed)}</td><td>{total.totalDisposedPct}%</td><td>{formatNumber(total.totalPending)}</td></tr>
   </tbody></table></div>
   <div className="legend"><b>COLOUR CODE:</b><span className="low">LOW (needs attention)</span><span className="medium">MEDIUM</span><span className="high">HIGH (good)</span></div>
   <div className="legendnote">Applied to Difference and percentage columns. Lowest third = red; highest third = green.</div>
   <div className="notes">Notes: Disposed means Approved / Disposed notices. % Disposed = Latest Disposed / Total Notices of that type. Difference = Latest - Earlier. Total Notices Disposed = No Mapping Disposed + Discrepancy Disposed. Total Notices Pending = Pending No Mapping + Pending Discrepancy.</div>
  </section>:<div className="empty"><FileText size={30}/><b>Upload Earlier and Latest Excel reports</b><span>Only the FINAL PS WISE REPORT sheet is used.</span></div>}
  <style jsx global>{`
   *{box-sizing:border-box}
   :root{color-scheme:dark}
   html{background:#070b13}
   body{margin:0;background:radial-gradient(circle at 50% -10%,#182846 0,#0b1220 34%,#070b13 72%);color:#e7edf7;font-family:Inter,ui-sans-serif,system-ui,-apple-system,BlinkMacSystemFont,"Segoe UI",Arial,sans-serif;min-height:100vh}
   .bar{height:5px;background:linear-gradient(90deg,#4d8dff,#1f3864 45%,#8b5cf6)}
   header{position:relative;background:linear-gradient(135deg,#101b30 0%,#162b50 52%,#0e1728 100%);color:#fff;text-align:center;padding:24px 20px 20px;border-bottom:1px solid #263a5b;box-shadow:0 12px 40px rgba(0,0,0,.28)}
   header:after{content:"";position:absolute;left:12%;right:12%;bottom:0;height:1px;background:linear-gradient(90deg,transparent,#4d8dff,transparent);opacity:.7}
   header h1{font-size:29px;letter-spacing:.4px;margin:0 0 7px;font-weight:850;text-shadow:0 2px 18px rgba(77,141,255,.2)}
   header div{font-size:12px;color:#b9c7dc;letter-spacing:.2px}
   header small{display:block;font-size:9px;margin-top:7px;color:#8fa2bf}
   .uploads{display:grid;grid-template-columns:1fr 1fr;gap:16px;padding:18px 24px 14px;max-width:1800px;margin:auto}
   .upload{position:relative;background:linear-gradient(145deg,#111a2b,#0d1523);border:1px solid #263750;border-radius:12px;padding:15px;box-shadow:0 10px 28px rgba(0,0,0,.2);overflow:hidden}
   .upload:before{content:"";position:absolute;left:0;top:0;bottom:0;width:3px;background:#4d8dff}
   .uploadtitle b{display:block;color:#f3f7ff;font-size:12px;letter-spacing:.4px}
   .uploadtitle span{display:block;color:#8191aa;font-size:10px;margin-top:3px}
   .drop{height:52px;border:1px dashed #405473;border-radius:8px;margin-top:10px;background:#0a1220;display:flex;align-items:center;justify-content:center;gap:7px;position:relative;font-size:11px;color:#c8d4e7;transition:.18s}
   .drop:hover{border-color:#5d8fe8;background:#0d1829;box-shadow:0 0 0 3px rgba(77,141,255,.08)}
   .drop input{position:absolute;inset:0;opacity:0;cursor:pointer}
   .manual{margin-top:8px;padding:7px 9px;border:1px solid #2d405d;border-radius:6px;background:#09111e;color:#dce6f4;font-size:10px;width:220px;outline:none}
   .manual::placeholder{color:#64748b}.manual:focus{border-color:#4d8dff;box-shadow:0 0 0 3px rgba(77,141,255,.1)}
   .ok{margin-top:8px;color:#7dd3a0;font-size:10px;display:flex;align-items:center;gap:5px}.ok strong{background:#173d2b;color:#9ee6b7;border:1px solid #286445;padding:3px 6px;border-radius:4px}
   .error{margin:0 24px 12px;padding:10px 12px;background:#3a171b;color:#ffb4b4;border:1px solid #74343a;border-radius:8px;font-weight:700;display:flex;gap:7px}
   .summaryCards{display:grid;grid-template-columns:repeat(5,minmax(150px,1fr));gap:10px;padding:8px 24px 18px;max-width:1800px;margin:auto}
   .summaryCard{position:relative;min-height:76px;padding:13px 14px;border:1px solid #293b57;border-radius:10px;background:linear-gradient(145deg,#111b2c,#0d1523);box-shadow:0 8px 24px rgba(0,0,0,.18);overflow:hidden}
   .summaryCard:after{content:"";position:absolute;right:-18px;top:-18px;width:58px;height:58px;border-radius:50%;background:rgba(77,141,255,.07)}
   .summaryCard.navy{border-top:2px solid #4d8dff}.summaryCard.brown{border-top:2px solid #c27a43}
   .summaryCard b{display:block;font-size:20px;line-height:1.1;color:#f4f7fb;letter-spacing:.2px}.summaryCard span{display:block;margin-top:8px;color:#8394ad;font-size:9px;font-weight:750;letter-spacing:.55px;text-transform:uppercase}
   .panel{margin:0 24px 28px;background:#0c1422;border:1px solid #263750;border-radius:12px;box-shadow:0 18px 50px rgba(0,0,0,.25);overflow:hidden}
   .actions{padding:13px 15px;border-bottom:1px solid #263750;display:flex;justify-content:space-between;align-items:center;background:linear-gradient(180deg,#111c2f,#0e1727)}
   .actions b{display:block;color:#eef4ff;font-size:14px;letter-spacing:.3px}.actions span{display:block;color:#7f90aa;font-size:10px;margin-top:3px}
   .buttons{display:flex;gap:8px}.buttons button{display:flex;align-items:center;gap:6px;border:1px solid #365174;border-radius:7px;background:#101c2d;color:#bcd0ea;padding:8px 11px;font-weight:800;font-size:11px;cursor:pointer;transition:.18s}.buttons button:hover{border-color:#5d8fe8;background:#14243a;transform:translateY(-1px)}.buttons .pdf{background:#1f3864;border-color:#4d73ac;color:#fff}.buttons .pdf:hover{background:#28497d}
   .tablewrap{overflow:auto;scrollbar-color:#344a69 #0b1220;scrollbar-width:thin}
   table{border-collapse:collapse;width:100%;min-width:2350px}
   th,td{border:1px solid #293b55;padding:6px 5px;text-align:center;font-size:11px;line-height:1.08}
   th{background:#142541;color:#dce8fa;font-size:8px;font-weight:800;white-space:normal;cursor:pointer;position:relative}th:hover{background:#1a3153}th svg{vertical-align:middle;margin-left:2px;color:#7fa8e8}.groups th{font-size:9px;cursor:default;padding:6px}.groups th:nth-child(2){background:#193766}.groups th:nth-child(3){background:#63320f}.groups th:nth-child(4){background:#193766}.groups .sep{background:#0c1422}
   tbody tr:not(.psDetailRow):nth-child(even){background:#0f1928}tbody tr:not(.psDetailRow):hover{background:#14243a}
   td{color:#cbd6e5;background:#0d1625}.name{text-align:left;font-weight:750;color:#eef4ff;white-space:normal;min-width:190px}
   .earlier{background:#2c291d!important;color:#f5e7b1}.colour{font-weight:850}.low{background:#5a252c!important;color:#ffd4d7}.medium{background:#5a481d!important;color:#ffe9a7}.high{background:#214d35!important;color:#c5f3d5}
   .colour.low{background:#5a252c!important}.colour.medium{background:#5a481d!important}.colour.high{background:#214d35!important}
   .sep{width:7px;min-width:7px;padding:0!important;background:#0c1422!important;border-left:0!important;border-right:0!important}
   .grand td{background:#172a47!important;color:#f1f6ff!important;font-weight:850;border-color:#3b5577}
   .psDetailRow>td{padding:0!important;background:#0a111d!important}.psTitle{background:linear-gradient(90deg,#111f35,#0e1827);color:#8fb6f4;font-weight:800;text-align:left;padding:7px 10px;font-size:10px;border-bottom:1px solid #2b405d;letter-spacing:.25px}.psTableWrap{overflow:auto}.psTable{min-width:1900px;width:100%;border-collapse:collapse}.psTable th,.psTable td{font-size:9px;padding:5px 4px;border-color:#293b55}.psTable th{background:#67350f;color:#fff;font-size:7.5px}.psTable td{font-weight:700;color:#cbd6e5;background:#0d1725}.psTable tr:nth-child(even) td{background:#101c2d}.psTable tr:hover td{background:#15263d}.psTable td:nth-child(9),.psTable td:nth-child(14){background:#2c291d;color:#f5e7b1}.psTable .officerTotalRow td{background:#D9E1F2;color:#1F3864;font-weight:900;border-top:2px solid #1F3864;padding-top:7px;padding-bottom:7px}.psTable .officerTotalRow td.percentageCell{font-weight:900}
   .legend{padding:9px 12px 3px;display:flex;align-items:center;gap:8px;font-size:10px;background:#0d1725}.legend span{padding:4px 10px;font-weight:800;border:1px solid #3b4d66;border-radius:4px}.legend .low{background:#5a252c}.legend .medium{background:#5a481d}.legend .high{background:#214d35}.legendnote,.notes{font-size:9px;color:#71829b;padding:4px 12px;background:#0d1725}.notes{padding-bottom:12px}
   .empty{margin:24px;padding:52px;background:linear-gradient(145deg,#111a2b,#0d1523);border:1px solid #293c59;border-radius:12px;display:flex;flex-direction:column;align-items:center;gap:9px;color:#7f90aa;box-shadow:0 15px 40px rgba(0,0,0,.2)}.empty b{color:#dce8fa}
   @media(max-width:1100px){.summaryCards{grid-template-columns:repeat(3,1fr)}}@media(max-width:900px){.summaryCards{grid-template-columns:repeat(2,1fr)}.uploads{grid-template-columns:1fr}.actions{align-items:flex-start;gap:10px;flex-direction:column}.panel{margin-left:12px;margin-right:12px}.uploads,.summaryCards{padding-left:12px;padding-right:12px}}
   @media print{.uploads,.actions{display:none!important}.panel{margin:0}body{background:#fff;color:#111}
  `}</style> </main>;
}
