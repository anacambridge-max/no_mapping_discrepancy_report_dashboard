export type RawRow = Record<string, any>;

export type OfficerReportRow = {
  officer: string;
  designation: string;
  totalNotices: number;
  nmTotal: number;
  nmEarlier: number;
  nmLatest: number;
  nmDifference: number;
  nmPct: number;
  nmDocs: number;
  nmDocsPct: number;
  dTotal: number;
  dEarlier: number;
  dLatest: number;
  dDifference: number;
  dPct: number;
  dLetters: number;
  dLettersPct: number;
  psCount: number;
};

export type PSReportRow = OfficerReportRow & {
  psNo: string;
  supervisor: string;
  blo: string;
};

export type ParsedReport = {
  fileName: string;
  timestamp: string;
  rows: RawRow[];
  grandTotal: RawRow | null;
  validation: { matched: boolean; calculated: Record<string, number>; source: Record<string, number> };
};

export const REQUIRED_COLUMNS = [
  "PS NO.","No Mapping Notices","Pending EPIC - No Mapping","No Mapping Approved / Disposed",
  "Documents Uploaded by BLO - No Mapping","Anomaly Notices","Pending EPIC - Discrepancy",
  "Discrepancy Approved / Disposed","BLO Letter Uploaded - Discrepancy","AERO / Ad.AERO",
  "Designation","BLO Supervisor","BLO Name"
];

const aliases: Record<string,string[]> = {
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
  const ks=Object.keys(row);
  for(const n of names){const found=ks.find(k=>norm(k)===norm(n)); if(found!==undefined)return row[found];}
  return "";
};
export const n=(v:any)=>{const x=Number(String(v??"").replace(/,/g,"")); return Number.isFinite(x)?x:0;};
export const percentage=(a:number,b:number)=>b?Math.round((a/b)*1000)/10:0;

export function parseTimestamp(name:string,lastModified:number,manual?:string){
  if(manual?.trim()) return manual.trim();
  const clean=name.replace(/[_-]/g," ");
  const dateMatch=clean.match(/(\d{1,2})[\s./](\d{1,2})[\s./](20\d{2})/);
  const isoMatch=name.match(/(20\d{2})[.-](\d{1,2})[.-](\d{1,2})/);
  let d:Date;
  if(isoMatch) d=new Date(Number(isoMatch[1]),Number(isoMatch[2])-1,Number(isoMatch[3]));
  else if(dateMatch) d=new Date(Number(dateMatch[3]),Number(dateMatch[2])-1,Number(dateMatch[1]));
  else d=new Date(lastModified||Date.now());
  const time=name.match(/(\d{1,2})[.:_-](\d{2})\s*(AM|PM)?/i);
  if(time){let h=Number(time[1]); const m=Number(time[2]); const ap=time[3]?.toUpperCase(); if(ap==="PM"&&h<12)h+=12;if(ap==="AM"&&h===12)h=0;d.setHours(h,m,0,0);}
  return d.toLocaleString("en-IN",{day:"2-digit",month:"short",year:"numeric",hour:"2-digit",minute:"2-digit"});
}

export function normalizeRows(rows:RawRow[]){
  return rows.filter(r=>String(key(r,aliases.ps)).trim() && String(key(r,aliases.ps)).trim().toUpperCase()!=="GRAND TOTAL")
    .map(r=>({
      psNo:String(key(r,aliases.ps)).trim(),
      officer:String(key(r,aliases.officer)).trim()||"UNMAPPED",
      designation:String(key(r,aliases.designation)).trim(),
      supervisor:String(key(r,aliases.supervisor)).trim(),
      blo:String(key(r,aliases.blo)).trim(),
      nmTotal:n(key(r,aliases.nmTotal)), nmPending:n(key(r,aliases.nmPending)), nmDisposed:n(key(r,aliases.nmDisposed)), nmDocs:n(key(r,aliases.nmDocs)),
      dTotal:n(key(r,aliases.dTotal)), dPending:n(key(r,aliases.dPending)), dDisposed:n(key(r,aliases.dDisposed)), dLetter:n(key(r,aliases.dLetter))
    }));
}

export function validateColumns(rows:RawRow[]){
  const sample=rows.find(r=>Object.keys(r).length>0)||{};
  const keys=Object.keys(sample);
  return REQUIRED_COLUMNS.filter(required=>!keys.some(k=>norm(k)===norm(required)));
}

export function parseGrandTotal(rows:RawRow[]){
  return rows.find(r=>String(key(r,aliases.ps)).trim().toUpperCase()==="GRAND TOTAL")||null;
}

type MetricTotals = { nmTotal:number; nmPending:number; nmDisposed:number; nmDocs:number; dTotal:number; dPending:number; dDisposed:number; dLetter:number; };\n\nfunction totals(rows:any[]):MetricTotals{
  return rows.reduce((a,r)=>{
    a.nmTotal+=r.nmTotal;a.nmPending+=r.nmPending;a.nmDisposed+=r.nmDisposed;a.nmDocs+=r.nmDocs;
    a.dTotal+=r.dTotal;a.dPending+=r.dPending;a.dDisposed+=r.dDisposed;a.dLetter+=r.dLetter;return a;
  },{nmTotal:0,nmPending:0,nmDisposed:0,nmDocs:0,nmDocsPct:0,dTotal:0,dPending:0,dDisposed:0,dLetter:0});
}

export function validateGrandTotal(rows:any[],source:RawRow|null){
  const calculated=totals(rows);
  const sourceTotals={
    nmTotal:n(key(source||{},aliases.nmTotal)),nmPending:n(key(source||{},aliases.nmPending)),nmDisposed:n(key(source||{},aliases.nmDisposed)),nmDocs:n(key(source||{},aliases.nmDocs)),
    dTotal:n(key(source||{},aliases.dTotal)),dPending:n(key(source||{},aliases.dPending)),dDisposed:n(key(source||{},aliases.dDisposed)),dLetter:n(key(source||{},aliases.dLetter))
  };
  const matched=!!source && Object.keys(sourceTotals).every(k=>sourceTotals[k as keyof typeof sourceTotals]===calculated[k as keyof typeof calculated]);
  return {matched,calculated,source:sourceTotals};
}

export function buildReports(latest:any[],earlier:any[]){
  const earlierByPS=new Map(earlier.map(r=>[r.psNo,r]));
  const latestByPS=new Map(latest.map(r=>[r.psNo,r]));
  const officerNames=[...new Set([...latest.map(r=>r.officer),...earlier.map(r=>r.officer)])].filter(Boolean);
  const officerRows:OfficerReportRow[]=officerNames.map(officer=>{
    const l=latest.filter(r=>r.officer===officer); const e=earlier.filter(r=>r.officer===officer);
    const L=totals(l),E=totals(e);
    return {officer,designation:l[0]?.designation||e[0]?.designation||"",totalNotices:L.nmTotal+L.dTotal,nmTotal:L.nmTotal,nmEarlier:E.nmDisposed,nmLatest:L.nmDisposed,nmDifference:L.nmDisposed-E.nmDisposed,nmPct:percentage(L.nmDisposed,L.nmTotal),nmDocs:L.nmDocs,nmDocsPct:percentage(L.nmDocs,L.nmTotal),dTotal:L.dTotal,dEarlier:E.dDisposed,dLatest:L.dDisposed,dDifference:L.dDisposed-E.dDisposed,dPct:percentage(L.dDisposed,L.dTotal),dLetters:L.dLetter,dLettersPct:percentage(L.dLetter,L.dTotal),psCount:l.length};
  }).sort((a,b)=>b.nmDifference-a.nmDifference||a.officer.localeCompare(b.officer));

  const psNos=[...new Set([...latest.map(r=>r.psNo),...earlier.map(r=>r.psNo)])];
  const psRows:PSReportRow[]=psNos.map(psNo=>{
    const l=latestByPS.get(psNo); const e=earlierByPS.get(psNo);
    const L:any=l||{nmTotal:0,nmDisposed:0,nmDocs:0,dTotal:0,dDisposed:0,dLetter:0,nmPending:0,dPending:0};
    const E:any=e||{nmDisposed:0,dDisposed:0};
    return {psNo,officer:l?.officer||e?.officer||"UNMAPPED",designation:l?.designation||e?.designation||"",supervisor:l?.supervisor||e?.supervisor||"",blo:l?.blo||e?.blo||"",totalNotices:L.nmTotal+L.dTotal,nmTotal:L.nmTotal,nmEarlier:E.nmDisposed,nmLatest:L.nmDisposed,nmDifference:L.nmDisposed-E.nmDisposed,nmPct:percentage(L.nmDisposed,L.nmTotal),nmDocs:L.nmDocs,nmDocsPct:percentage(L.nmDocs,L.nmTotal),dTotal:L.dTotal,dEarlier:E.dDisposed,dLatest:L.dDisposed,dDifference:L.dDisposed-E.dDisposed,dPct:percentage(L.dDisposed,L.dTotal),dLetters:L.dLetter,dLettersPct:percentage(L.dLetter,L.dTotal),psCount:1};
  }).sort((a,b)=>Number(a.psNo)-Number(b.psNo));
  return {officerRows,psRows};
}

export function grandOfficer(rows:OfficerReportRow[]){
  const x=rows.reduce((a,r)=>{a.totalNotices+=r.totalNotices;a.nmTotal+=r.nmTotal;a.nmEarlier+=r.nmEarlier;a.nmLatest+=r.nmLatest;a.nmDifference+=r.nmDifference;a.nmDocs+=r.nmDocs;a.nmDocsPct+=0;a.dTotal+=r.dTotal;a.dEarlier+=r.dEarlier;a.dLatest+=r.dLatest;a.dDifference+=r.dDifference;a.dLetters+=r.dLetters;a.dLettersPct+=0;a.psCount+=r.psCount;return a;},{totalNotices:0,nmTotal:0,nmEarlier:0,nmLatest:0,nmDifference:0,nmDocs:0,nmDocsPct:0,dTotal:0,dEarlier:0,dLatest:0,dDifference:0,dLetters:0,dLettersPct:0,psCount:0} as {totalNotices:number;nmTotal:number;nmEarlier:number;nmLatest:number;nmDifference:number;nmDocs:number;nmDocsPct:number;dTotal:number;dEarlier:number;dLatest:number;dDifference:number;dLetters:number;dLettersPct:number;psCount:number});
  return {...x,nmPct:percentage(x.nmLatest,x.nmTotal),nmDocsPct:percentage(x.nmDocs,x.nmTotal),dPct:percentage(x.dLatest,x.dTotal),dLettersPct:percentage(x.dLetters,x.dTotal)};
}

export function grandPS(rows:PSReportRow[]){ return grandOfficer(rows); }

export function percentileGroups(rows:any[],value:(r:any)=>number){
  const vals=rows.map(value).sort((a,b)=>a-b); if(!vals.length)return new Map();
  const q=(p:number)=>{const i=(vals.length-1)*p;const lo=Math.floor(i),hi=Math.ceil(i);return vals[lo]+(vals[hi]-vals[lo])*(i-lo);};
  const p33=q(.3333333333),p66=q(.6666666667); const map=new Map<string,"low"|"medium"|"high">();
  rows.forEach(r=>{const v=value(r);map.set(r.officer||r.psNo,v<=p33?"low":v<=p66?"medium":"high");}); return map;
}

export function formatOfficer(r:OfficerReportRow){return `${r.officer} (${r.designation||"Officer"})`;}
export function formatNumber(v:number){return v.toLocaleString("en-IN");}
