// Diamond Jarvis -> KV HISAB auto-sync v3 (Nov 8, 2026): v2 + DELETE (app entry id marker in col P)
// Entries write DIRECTLY into the month tab's real sections:
//   SALE -> SELL (Vechan) | LOT_ENTRY / OUTSIDE_PURCHASE -> PURCHASE (Kharidi)
//   UPAD -> JAVAK UPAD col (partner KV/VF/SHAILESH) or AAVAK row (jama)
//   EXPENSE -> JAVAK row | AAVAK / JAVAK -> top section (INVESTMENT/UCHHINA subtype)
// Quality goes in a QUALITY column of SELL/PURCHASE.
// A new month tab is created automatically from the previous month, with
// OPENING BALANCE (rupees, D3) and OPENING STOCK (carat) carried by formula.
// One-time setup: paste over Code.gs, save, Deploy > Manage deployments >
// pencil > New version > Deploy. Execute as Me, access Anyone.

const SHEET_ID="1RWcEmDj4qt2QGmGWFL_b0AQh3YcG5DH83AwIbJ50DII";
const MONTHS=["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];

function doGet(){return json_({ok:true,msg:"Diamond Jarvis sync v3 running"});}

function doPost(e){
  try{
    const p=JSON.parse((e&&e.postData&&e.postData.contents)||"{}");
    if(!p||!p.type||p.type==="QUERY")return json_({ok:true,skipped:true});
    const ss=SpreadsheetApp.openById(SHEET_ID);
    if(p.type==="DELETE"){
      const sh=monthSheet_(ss,p.date);
      return json_({ok:true,deleted:deleteById_(sh,p.id)});
    }
    const sh=monthSheet_(ss,p.date);
    const where=writeEntry_(sh,p);
    const m=/row (\d+)/.exec(where);
    ensureColP_(sh);
    if(m&&p.id)sh.getRange(Number(m[1]),16).setValue(String(p.id));
    return json_({ok:true,wrote:where});
  }catch(err){
    return json_({ok:false,error:String(err&&err.message||err)});
  }
}
function json_(o){return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);}

/* ---------- month tab ---------- */
function monthSheet_(ss,dateStr){
  const parts=String(dateStr||"").split("-");
  const now=new Date();
  const y=parseInt(parts[0],10)||now.getFullYear();
  let m=parseInt(parts[1],10);if(!(m>=1&&m<=12))m=now.getMonth()+1;
  const title=MONTHS[m-1]+" "+y;
  let sh=ss.getSheetByName(title);
  if(sh)return sh;
  let prev=null,prevTitle="";
  const d=new Date(Date.UTC(y,m-1,1));
  for(let i=0;i<24&&!prev;i++){
    d.setUTCMonth(d.getUTCMonth()-1);
    const t=MONTHS[d.getUTCMonth()]+" "+d.getUTCFullYear();
    const s=ss.getSheetByName(t);
    if(s){prev=s;prevTitle=t;}
  }
  if(!prev){sh=ss.insertSheet(title);buildBlank_(sh);return sh;}
  sh=prev.copyTo(ss);
  sh.setName(title);
  const a=anchors_(sh),pa=anchors_(prev);
  if(a.aavakTotal>4)sh.getRange(4,1,a.aavakTotal-4,14).clearContent();
  if(a.purTotal>a.purDataStart)sh.getRange(a.purDataStart,1,a.purTotal-a.purDataStart,10).clearContent();
  if(a.sellTotal>a.sellDataStart)sh.getRange(a.sellDataStart,1,a.sellTotal-a.sellDataStart,10).clearContent();
  sh.getRange("D3").setValue("='"+prevTitle+"'!D"+pa.tbRow);
  sh.getRange(a.openingRow,4).setValue("='"+prevTitle+"'!D"+pa.closingRow);
  return sh;
}

/* ---------- anchors: find real rows of each section ---------- */
function anchors_(sh){
  const v=sh.getRange(1,1,80,14).getDisplayValues();
  const g=i=>v[i-1]||[];
  const a={aavakTotal:0,tbRow:0,purTitle:0,purHeader:0,purTotal:0,purDataStart:0,
           sellTitle:0,sellHeader:0,sellTotal:0,sellDataStart:0,openingRow:0,closingRow:0};
  let t=0;
  for(let i=1;i<=80;i++){
    const c=String(g(i)[0]||"").trim();
    if(!a.tbRow&&/^TOTAL BALANCE/i.test(c))a.tbRow=i;
    if(!a.purTitle&&/^PURCHASE/i.test(c))a.purTitle=i;
    if(!a.sellTitle&&/^SELL/i.test(c))a.sellTitle=i;
    if(!a.openingRow&&/^OPENING STOCK/i.test(c))a.openingRow=i;
    if(!a.closingRow&&/^CLOSING STOCK/i.test(c))a.closingRow=i;
    if(c==="TOTAL"){t++;if(t===1)a.aavakTotal=i;else if(t===2&&!a.purTotal)a.purTotal=i;else if(t===3&&!a.sellTotal)a.sellTotal=i;}
  }
  if(!a.aavakTotal)a.aavakTotal=25;
  if(!a.tbRow)a.tbRow=a.aavakTotal+1;
  if(a.purTitle){a.purHeader=a.purTitle+1;}
  if(a.sellTitle){a.sellHeader=a.sellTitle+1;}
  a.aavakDataStart=rangeStart_(sh,a.aavakTotal,"B")||4;
  if(a.purTotal)a.purDataStart=rangeStart_(sh,a.purTotal,"B")||a.purHeader+1;
  if(a.sellTotal)a.sellDataStart=rangeStart_(sh,a.sellTotal,"B")||a.sellHeader+1;
  if(!a.purDataStart&&a.purHeader)a.purDataStart=a.purHeader+1;
  if(!a.sellDataStart&&a.sellHeader)a.sellDataStart=a.sellHeader+1;
  if(!a.openingRow)a.openingRow=58;
  if(!a.closingRow)a.closingRow=59;
  return a;
}
function rangeStart_(sh,totalRow,col){
  if(!totalRow)return null;
  const f=String(sh.getRange(totalRow,2).getFormula()||"").toUpperCase();
  const at=f.indexOf("SUM");
  if(at<0)return null;
  const s=f.indexOf(col,at),e=f.indexOf(":",s);
  if(s<0||e<0)return null;
  let d="";
  for(const ch of f.substring(s+1,e)){if(ch>="0"&&ch<="9")d+=ch;}
  return d?parseInt(d,10):null;
}

/* ---------- first empty row, insert if full ---------- */
function nextRow_(sh,col,start,end){
  const idx=col.charCodeAt(0)-64;
  const v=sh.getRange(start,idx,Math.max(1,end-start),1).getDisplayValues();
  for(let i=0;i<v.length;i++){if(!String(v[i][0]).trim())return start+i;}
  sh.insertRowsBefore(end,1);
  return end;
}

function quality_(sh,headerRow,col,q){
  if(!q||!headerRow)return;
  if(!String(sh.getRange(headerRow,col).getDisplayValue()||"").trim())sh.getRange(headerRow,col).setValue("QUALITY");
}

/* ---------- ensure marker column P exists ---------- */
function ensureColP_(sh){
  if(sh.getMaxColumns()<16)sh.insertColumnsAfter(sh.getMaxColumns(),16-sh.getMaxColumns());
}

/* ---------- delete one entry by app id (marker col P) ---------- */
function deleteById_(sh,id){
  if(id==null)return "no id";
  ensureColP_(sh);
  const want=String(id);
  const rows=Math.min(sh.getMaxRows(),400);
  const v=sh.getRange(1,16,rows,1).getDisplayValues();
  for(let i=0;i<v.length;i++){
    if(String(v[i][0]).trim()===want){
      const r=i+1;
      const left=String(sh.getRange(r,1).getDisplayValue()||"").trim();
      const right=String(sh.getRange(r,8).getDisplayValue()||"").trim();
      if(left)sh.getRange(r,1,1,10).clearContent();
      if(right)sh.getRange(r,8,1,7).clearContent();
      sh.getRange(r,16).clearContent();
      return "row "+r;
    }
  }
  return "not found";
}

/* ---------- write one entry ---------- */
function writeEntry_(sh,p){
  const ss=sh.getParent();
  const a=anchors_(sh);
  const n=x=>{const v=parseFloat(x);return isNaN(v)?null:v;};
  const t=p.type;
  const q=String(p.quality||"").trim();
  const d=toDate_(ss,p.date);
  const putD=col=>{sh.getRange(r_,col).setValue(d);sh.getRange(r_,col).setNumberFormat("dd/mm/yyyy");};
  const txt=(col,v)=>{const s=String(v==null?"":v).trim();if(s)sh.getRange(r_,col).setValue(s);};
  let r_;
  const put=(col,v)=>{sh.getRange(r_,col).setValue(v);};
  if(t==="SALE"){
    r_=nextRow_(sh,"A",a.sellDataStart,a.sellTotal);
    putD(1);put(2,n(p.carats));put(3,n(p.rate));put(4,n(p.terms));
    txt(5,p.dalali);txt(8,p.dalal);txt(9,p.party);
    quality_(sh,a.sellHeader,10,q);if(q)put(10,q);
    return "SELL row "+r_;
  }
  if(t==="LOT_ENTRY"||t==="OUTSIDE_PURCHASE"){
    r_=nextRow_(sh,"A",a.purDataStart,a.purTotal);
    putD(1);put(2,n(p.carats));put(3,n(p.rate));put(4,n(p.terms));
    txt(7,p.dalal);txt(8,p.party);
    quality_(sh,a.purHeader,9,q);if(q)put(9,q);
    return "PURCHASE row "+r_;
  }
  if(t==="UPAD"){
    if(p.direction==="jama"){
      r_=nextRow_(sh,"A",a.aavakDataStart,a.aavakTotal);
      putD(1);put(2,n(p.amount));txt(4,p.partner||p.party);
      return "AAVAK (jama) row "+r_;
    }
    r_=nextRow_(sh,"H",a.aavakDataStart,a.aavakTotal);
    putD(8);put(9,n(p.amount));txt(11,p.partner||p.party);txt(14,p.partner);
    return "JAVAK UPAD row "+r_;
  }
  if(t==="EXPENSE"){
    r_=nextRow_(sh,"H",a.aavakDataStart,a.aavakTotal);
    putD(8);put(9,n(p.amount));txt(11,p.note||p.party);
    return "JAVAK (kharcha) row "+r_;
  }
  if(t==="AAVAK"){
    r_=nextRow_(sh,"A",a.aavakDataStart,a.aavakTotal);
    putD(1);
    put(p.subtype==="INVESTMENT"?7:(p.subtype==="UCHHINA"?6:2),n(p.amount));
    txt(4,p.party);txt(5,p.note);
    return "AAVAK row "+r_;
  }
  if(t==="JAVAK"){
    r_=nextRow_(sh,"H",a.aavakDataStart,a.aavakTotal);
    putD(8);
    put(p.subtype==="UCHHINA"?13:9,n(p.amount));
    txt(11,p.party);txt(12,p.note);
    return "JAVAK row "+r_;
  }
  return "unknown type "+t;
}

function toDate_(ss,s){
  const m=String(s||"").match(/^(\d{4})-(\d{2})-(\d{2})/);
  if(!m)return new Date();
  const y=+m[1],mo=+m[2],dy=+m[3];
  const tz=ss.getSpreadsheetTimeZone();
  let d=new Date(Date.UTC(y,mo-1,dy,12,0,0));
  const want=y+"-"+("0"+mo).slice(-2)+"-"+("0"+dy).slice(-2);
  for(let k=0;k<3;k++){
    const wall=Utilities.formatDate(d,tz,"yyyy-MM-dd");
    if(wall===want)break;
    const wt=Date.UTC(+wall.slice(0,4),+wall.slice(5,7)-1,+wall.slice(8,10),12,0,0);
    d=new Date(d.getTime()+(Date.UTC(y,mo-1,dy,12,0,0)-wt));
  }
  return d;
}

/* ---------- blank structure (rare fallback) ---------- */
function buildBlank_(sh){
  sh.getRange("A1").setValue("AAVAK (Jama Thayelu)");
  sh.getRange("H1").setValue("JAVAK (Aapyu Rakam)");
  sh.getRange("A2:G2").setValues([["DATE","PAYMENT","CREDIT","PARTY","NAME","UCHHINA","INVESTMENT"]]);
  sh.getRange("H2:N2").setValues([["DATE","PAYMENT","DEBIT","PARTY","NAME","UCHHINA","UPAD (KV/VF/SHAILESH)"]]);
  sh.getRange("A3").setValue("OPENING BALANCE");sh.getRange("D3").setValue(0);
  sh.getRange("A25").setValue("TOTAL");sh.getRange("H25").setValue("TOTAL");
  sh.getRange("B25").setFormula("=SUM(B4:B24)");sh.getRange("C25").setFormula("=SUM(C4:C24)");
  sh.getRange("F25").setFormula("=SUM(F4:F24)");sh.getRange("G25").setFormula("=SUM(G4:G24)");
  sh.getRange("I25").setFormula("=SUM(I4:I24)");sh.getRange("J25").setFormula("=SUM(J4:J24)");
  sh.getRange("M25").setFormula("=SUM(M4:M24)");sh.getRange("N25").setFormula("=SUM(N4:N24)");
  sh.getRange("A26").setValue("TOTAL BALANCE (AAVAK - JAVAK)");sh.getRange("D26").setFormula("=D3+B25-I25");
  sh.getRange("A27").setValue("PURCHASE (Kharidi)");
  sh.getRange("A28:H28").setValues([["DATE","CARAT","PRICE","TERMS","PAYMENT","DEBIT","BROKER","PARTY"]]);
  sh.getRange("A41").setValue("TOTAL");
  sh.getRange("B41").setFormula("=SUM(B30:B40)");sh.getRange("C41").setFormula("=SUM(C30:C40)");
  sh.getRange("E41").setFormula("=SUM(E30:E40)");sh.getRange("F41").setFormula("=SUM(F30:F40)");
  sh.getRange("A42").setValue("SELL (Vechan)");
  sh.getRange("A43:I43").setValues([["DATE","CARAT","PRICE","TERMS","DALALI","PAYMENT","CREDIT","BROKER","PARTY"]]);
  sh.getRange("A56").setValue("TOTAL");
  sh.getRange("B56").setFormula("=SUM(B45:B55)");sh.getRange("C56").setFormula("=SUM(C45:C55)");
  sh.getRange("E56").setFormula("=SUM(E45:E55)");sh.getRange("F56").setFormula("=SUM(F45:F55)");
  sh.getRange("G56").setFormula("=SUM(G45:G55)");
  sh.getRange("A58").setValue("OPENING STOCK (CARAT)");sh.getRange("D58").setValue(0);
  sh.getRange("A59").setValue("CLOSING STOCK (OPENING + PURCHASE - SELL)");
  sh.getRange("D59").setFormula("=D58+B41-B56");
}
