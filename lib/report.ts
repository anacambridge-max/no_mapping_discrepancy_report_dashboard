export type RawRow = Record<string, any>;

export type NormalizedRow = {
  psNo:string; officer:string; designation:string; supervisor:string; blo:string;
  nmTotal:number; nmPending:number; nmDisposed:number; nmDocs:number;
  dTotal:number; dPending:number; dDisposed:number; dLetter:number;
};

export type OfficerReportRow = {
  officer:string; designation:string; totalNotices:number;
  nmTotal:number; nmEarlier:number; nmLatest:number; nmDifference:number; nmPct:number;
  nmDocs:number; nmDocsPct:number;
  dTotal:number; dEarlier:number; dLatest:number; dDifference:number; dPct:number;
  dLetters:number; dLettersPct:number; psCount:number;
  totalDisposed:number; totalDisposedPct:number;
};

export type ParsedReport = {
  fileName:string; timestamp:string; rows:NormalizedRow[]; grandTotal:RawRow|null;
  validation:{matched:boolean; calculated:Record<string,number>; source:Record<string,number>};
};

const aliases:Record<string,string[]> = {
  ps:["PS NO.","PS No","PS","PART NO.","Part No"],
  nmTotal:["No Mapping Notices"],
  nmPending:["Pending EPIC - No Mapping"],
  nmDisposed:["No Mapping Approved / Disposed"],
  nmDocs:["Documents Uploaded by BLO - No Mapping"],
  dTotal:["Anomaly Notices","Discrepancy Notices"],
  dPending:["Pending EPIC - Discrepancy"],
  dDisposed:["Discrepancy Approved / Disposed"],
  dLetter:["BLO Letter Uploaded - Discrepancy"],
  officer:["AERO / Ad.AERO"],
  designation:["Designation"],
  supervisor:["BLO Supervisor"],
  blo:["BLO Name"]
};

const norm=(v:any)=>String(v??"").toLowerCase().replace(/[\s_.\-/()]+/g,"").trim();
const key=(row:RawRow,names:string[])=>{
  for(const name of names){
    const found=Object.keys(row).find(k=>norm(k)===norm(name));
    if(found!==undefined) return row[found];
  }
  return "";
};
export const n=(v:any)=>{
  const x=Number(String(v??"").replace(/,/g,""));
  return Number.isFinite(x)?x:0;
};
export const percentage=(a:number,b:number)=>b?Math.round((a/b)*1000)/10:0;

export function parseTimestamp(name:string,lastModified:number,manual?:string){
  if(manual?.trim()) return manual.trim();
  const iso=name.match(/(20\d{2})[.-](\d{1,2})[.-](\d{1,2})/);
  const dmatch=name.replace(/[_-]/g," ").match(/(\d{1,2})[\s./](\d{1,2})[\s./](20\d{2})/);
  let d:Date;
  if(iso) d=new Date(Number(iso[1]),Number(iso[2])-1,Number(iso[3]));
  else if(dmatch) d=new Date(Number(dmatch[3]),Number(dmatch[2])-1,Number(dmatch[1]));
  else d=new Date(lastModified||Date.now());
  const tm=name.match(/(\d{1,2})[.:_-](\d{2})\s*(AM|PM)?/i);
  if(tm){
    let h=Number(tm[1]); const m=Number(tm[2]); const ap=tm[3]?.toUpperCase();
    if(ap==="PM"&&h<12) h+=12; if(ap==="AM"&&h===12) h=0; d.setHours(h,m,0,0);
  }
  return d.toLocaleString("en-IN",{day:"2-digit",month:"short",year:"numeric",hour:"2-digit",minute:"2-digit"});
}

export const REQUIRED_COLUMNS=[
  "PS NO.","No Mapping Notices","Pending EPIC - No Mapping","No Mapping Approved / Disposed",
  "Documents Uploaded by BLO - No Mapping","Anomaly Notices","Pending EPIC - Discrepancy",
  "Discrepancy Approved / Disposed","BLO Letter Uploaded - Discrepancy","AERO / Ad.AERO",
  "Designation","BLO Supervisor","BLO Name"
];

export function validateColumns(rows:RawRow[]){
  const sample=rows.find(r=>Object.keys(r).length>0)||{};
  const keys=Object.keys(sample);
  return REQUIRED_COLUMNS.filter(required=>!keys.some(k=>norm(k)===norm(required)));
}

export function normalizeRows(rows:RawRow[]):NormalizedRow[]{
  return rows
    .filter(r=>{
      const ps=String(key(r,aliases.ps)).trim();
      return ps!=="" && ps.toUpperCase()!=="GRAND TOTAL";
    })
    .map(r=>({
      psNo:String(key(r,aliases.ps)).trim(),
      officer:String(key(r,aliases.officer)).trim()||"UNMAPPED",
      designation:String(key(r,aliases.designation)).trim(),
      supervisor:String(key(r,aliases.supervisor)).trim(),
      blo:String(key(r,aliases.blo)).trim(),
      nmTotal:n(key(r,aliases.nmTotal)),
      nmPending:n(key(r,aliases.nmPending)),
      nmDisposed:n(key(r,aliases.nmDisposed)),
      nmDocs:n(key(r,aliases.nmDocs)),
      dTotal:n(key(r,aliases.dTotal)),
      dPending:n(key(r,aliases.dPending)),
      dDisposed:n(key(r,aliases.dDisposed)),
      dLetter:n(key(r,aliases.dLetter))
    }));
}

export function parseGrandTotal(rows:RawRow[]){
  return rows.find(r=>String(key(r,aliases.ps)).trim().toUpperCase()==="GRAND TOTAL")||null;
}

type MetricTotals={
  nmTotal:number; nmPending:number; nmDisposed:number; nmDocs:number;
  dTotal:number; dPending:number; dDisposed:number; dLetter:number;
};

function totals(rows:NormalizedRow[]):MetricTotals{
  const out:MetricTotals={
    nmTotal:0,nmPending:0,nmDisposed:0,nmDocs:0,
    dTotal:0,dPending:0,dDisposed:0,dLetter:0
  };
  for(const r of rows){
    out.nmTotal+=r.nmTotal; out.nmPending+=r.nmPending; out.nmDisposed+=r.nmDisposed; out.nmDocs+=r.nmDocs;
    out.dTotal+=r.dTotal; out.dPending+=r.dPending; out.dDisposed+=r.dDisposed; out.dLetter+=r.dLetter;
  }
  return out;
}

export function validateGrandTotal(rows:NormalizedRow[],source:RawRow|null){
  const calculated=totals(rows);
  const sourceTotals:Record<string,number>={
    nmTotal:n(key(source||{},aliases.nmTotal)),
    nmPending:n(key(source||{},aliases.nmPending)),
    nmDisposed:n(key(source||{},aliases.nmDisposed)),
    nmDocs:n(key(source||{},aliases.nmDocs)),
    dTotal:n(key(source||{},aliases.dTotal)),
    dPending:n(key(source||{},aliases.dPending)),
    dDisposed:n(key(source||{},aliases.dDisposed)),
    dLetter:n(key(source||{},aliases.dLetter))
  };
  const matched=!!source && Object.keys(sourceTotals).every(k=>sourceTotals[k]===calculated[k as keyof MetricTotals]);
  return {matched,calculated,source:sourceTotals};
}

export function buildOfficerRows(latest:NormalizedRow[],earlier:NormalizedRow[]):OfficerReportRow[]{
  const names=[...new Set([...latest.map(r=>r.officer),...earlier.map(r=>r.officer)])].filter(Boolean);
  return names.map(officer=>{
    const L=totals(latest.filter(r=>r.officer===officer));
    const E=totals(earlier.filter(r=>r.officer===officer));
    return {
      officer,
      designation:latest.find(r=>r.officer===officer)?.designation||earlier.find(r=>r.officer===officer)?.designation||"",
      totalNotices:L.nmTotal+L.dTotal,
      nmTotal:L.nmTotal,nmEarlier:E.nmDisposed,nmLatest:L.nmDisposed,nmDifference:L.nmDisposed-E.nmDisposed,
      nmPct:percentage(L.nmDisposed,L.nmTotal),nmDocs:L.nmDocs,nmDocsPct:percentage(L.nmDocs,L.nmTotal),
      dTotal:L.dTotal,dEarlier:E.dDisposed,dLatest:L.dDisposed,dDifference:L.dDisposed-E.dDisposed,
      dPct:percentage(L.dDisposed,L.dTotal),dLetters:L.dLetter,dLettersPct:percentage(L.dLetter,L.dTotal),
      psCount:latest.filter(r=>r.officer===officer).length,
      totalDisposed:L.nmDisposed+L.dDisposed,
      totalDisposedPct:percentage(L.nmDisposed+L.dDisposed,L.nmTotal+L.dTotal)
    };
  }).sort((a,b)=>b.nmDifference-a.nmDifference||a.officer.localeCompare(b.officer));
}

export function grandOfficer(rows:OfficerReportRow[]){
  const out={totalNotices:0,nmTotal:0,nmEarlier:0,nmLatest:0,nmDifference:0,nmDocs:0,dTotal:0,dEarlier:0,dLatest:0,dDifference:0,dLetters:0,psCount:0,totalDisposed:0};
  for(const r of rows){
    out.totalNotices+=r.totalNotices; out.nmTotal+=r.nmTotal; out.nmEarlier+=r.nmEarlier; out.nmLatest+=r.nmLatest;
    out.nmDifference+=r.nmDifference; out.nmDocs+=r.nmDocs; out.dTotal+=r.dTotal; out.dEarlier+=r.dEarlier;
    out.dLatest+=r.dLatest; out.dDifference+=r.dDifference; out.dLetters+=r.dLetters; out.psCount+=r.psCount; out.totalDisposed+=r.totalDisposed;
  }
  return {...out,nmPct:percentage(out.nmLatest,out.nmTotal),nmDocsPct:percentage(out.nmDocs,out.nmTotal),dPct:percentage(out.dLatest,out.dTotal),dLettersPct:percentage(out.dLetters,out.dTotal),totalDisposedPct:percentage(out.totalDisposed,out.totalNotices)};
}

export function percentileGroups(rows:OfficerReportRow[],value:(r:OfficerReportRow)=>number){
  const vals=rows.map(value).sort((a,b)=>a-b);
  if(!vals.length)return new Map<string,"low"|"medium"|"high">();
  const q=(p:number)=>{
    const i=(vals.length-1)*p; const lo=Math.floor(i),hi=Math.ceil(i);
    return vals[lo]+(vals[hi]-vals[lo])*(i-lo);
  };
  const p33=q(1/3),p66=q(2/3);
  const out=new Map<string,"low"|"medium"|"high">();
  rows.forEach(r=>out.set(r.officer,value(r)<=p33?"low":value(r)<=p66?"medium":"high"));
  return out;
}

export const formatNumber=(v:number)=>v.toLocaleString("en-IN");
export const formatOfficer=(r:OfficerReportRow)=>`${r.officer} (${r.designation||"Officer"})`;
