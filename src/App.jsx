import { useState, useEffect, useRef, useCallback, Fragment } from "react";
import { auth, db } from "./firebase";
import { onAuthStateChanged, signInWithEmailAndPassword, signOut } from "firebase/auth";
import { collection, doc, onSnapshot, setDoc, updateDoc, deleteDoc, writeBatch, getDoc, getDocs, query, where } from "firebase/firestore";
// NOTE: BUR items are scoped per project via a `pid` field.
import burSeed from "./burSeed.json"; // master BUR list imported from BUR.xlsx
import SheetGrid from "./SheetGrid.jsx";

const SECTIONS=[{id:"prelim",name:"Preliminaries"},{id:"building",name:"Building Works"},{id:"external",name:"External Works"},{id:"mande",name:"M&E Works"},{id:"fees",name:"Professional Fees"}];
const UNITS=["m²","m³","m","no.","sum","lot","kg","t","m run","%","item"]
// Dropdown options for a unit box: always include the item's own unit (e.g. "m2" or "nr" from older data) so the box shows what is really stored.
const unitOpts=cur=>!cur||UNITS.includes(cur)?UNITS:[cur,...UNITS];
const STATUSES=["Draft","Under Review","Confirmed"];
const SS={Draft:{bg:"#f1f5f9",c:"#475569"},"Under Review":{bg:"#fef9c3",c:"#b45309"},Confirmed:{bg:"#dcfce7",c:"#15803d"}};
const ROLES=["Lead QS","Estimator","PM","Client"];
const TABS=[{id:"boq",label:"📋 BOQ"},{id:"bur",label:"📚 BUR"},{id:"compare",label:"🆚 Compare Quotes"},{id:"summary",label:"📊 Summary"},{id:"log",label:"📝 Activity"}];
const ALL_COLS=[{id:"ref",label:"Ref",w:48},{id:"desc",label:"Description",w:190},{id:"unit",label:"Unit",w:64},{id:"qty",label:"Qty",w:56,num:1},{id:"rA",label:"Rate A",w:74,num:1},{id:"amtA",label:"Amt A (S$)",w:100,num:1},{id:"rB",label:"Rate B",w:74,num:1},{id:"amtB",label:"Amt B (S$)",w:100,num:1},{id:"code",label:"BUR Code",w:90},{id:"remarks",label:"Remarks",w:130},{id:"by",label:"By",w:56}];
const DEF_COLS=new Set(["ref","desc","unit","qty","rA","amtA","rB","amtB","code","remarks","by"]);
const DEF_CODES=[{id:"dc1",code:"PRELIM",desc:"Preliminaries",cat:"Prelim"},{id:"dc2",code:"BLDG-A",desc:"Building Works Phase A",cat:"Building"},{id:"dc3",code:"BLDG-B",desc:"Building Works Phase B",cat:"Building"},{id:"dc4",code:"EXT-A",desc:"External Works Phase A",cat:"External"},{id:"dc5",code:"EXT-B",desc:"External Works Phase B",cat:"External"},{id:"dc6",code:"ME-ACMV",desc:"Air Conditioning & Mech Ventilation",cat:"M&E"},{id:"dc7",code:"ME-ELV",desc:"Electrical & Low Voltage",cat:"M&E"},{id:"dc8",code:"ME-FP",desc:"Fire Protection",cat:"M&E"},{id:"dc9",code:"ME-STP",desc:"Sewage Treatment Plant",cat:"M&E"},{id:"dc10",code:"FEES",desc:"Professional Fees",cat:"Fees"}];
const DEF_CATS=["Concrete Works","Formwork","Reinforcement","Brickwork & Blockwork","Waterproofing","Structural Steelwork","Floor & Wall Finishes","Roofing","Doors, Windows & Glazing","External Works","M&E Works","Preliminaries","Others"].map((name,i)=>({id:`cat${i+1}`,name}));
const MASTER_CAT={id:"cat_master",name:"Master BUR"};
// Keyword → category id rules for auto-sorting the master list. First match wins; no match → Others (cat13).
const CATEGORY_RULES=[
  {cat:"cat5",kw:["waterproof","membrane","tanking","drainage cell","drainage board","bituminous","damp proof"]},
  {cat:"cat8",kw:["roofing","roof ","metal roof","gutter","downpipe","skylight","rainwater"]},
  {cat:"cat9",kw:["door","window","glazing","glass","louvre","louver","shutter","ironmonger","mirror","balustrade","handrail","curtain wall","skylight"]},
  {cat:"cat3",kw:["reinforc","rebar","mesh","brc","column cage","cage","ductility","steel bar","high tensile bar","y bar","r bar"]},
  {cat:"cat6",kw:["structural steel","steelwork","hollow section","ms plate","steel beam","steel column","metalwork","metal work","steel truss"]},
  {cat:"cat1",kw:["concrete","screed","blinding","grade c","grouting","grout","cement sand","precast","r.c.","rc slab"]},
  {cat:"cat2",kw:["formwork","plywood form","form work"]},
  {cat:"cat4",kw:["brick","block","alc","aac","partition","drywall","stud wall","hollow core panel","masonry"]},
  {cat:"cat7",kw:["tile","tiling","floor","ceiling","skim","plaster","paint","homog","granite","marble","vinyl","carpet","acoustic","pelmet","cornice","laminate","wallpaper","cladding","wall finish","screed finish","terrazzo"]},
  {cat:"cat11",kw:["electric","acmv","air-cond","aircon","ventilation","fire ","sprinkler","plumb","sanitary"," pipe","piping","cable","light fitting","lighting","pump","m&e","mechanical","ductwork","duct ","switch","socket","conduit","fcu","ahu","exhaust"]},
  {cat:"cat10",kw:["kerb","road","turf","planter","landscape","fence","gate","carpark","car park","paver","apron","drain","external","linkway","pavement","sump","precast drain","bollard"]},
  {cat:"cat12",kw:["prelim","hoarding","site office","scaffold","insurance","temporary","mobilis","pile integrity","load test","survey","testing"]},
];
const COMPS=["labour","material","plant","subcon"];
const CLABEL={labour:"Labour",material:"Material",plant:"Plant",subcon:"Subcon / Supplier"};

// ── Quote comparison (QS-style side-by-side subcontractor bid comparison) ──
const CMP_STATUSES=["normal","incl","norate","excluded"];
const CMP_STATUS_LABEL={normal:"Priced",incl:"Incl",norate:"No Rate",excluded:"Excl"};
const QTY_MISMATCH_PCT=0.05;   // qty differs from WH qty by more than this fraction -> flag
const OUTLIER_PCT=0.25;        // rate differs from the row's median priced rate by more than this fraction -> flag
const SEV_COLOR={High:{bg:"#fecaca",c:"#991b1b"},Medium:{bg:"#fed7aa",c:"#9a3412"},Low:{bg:"#fef08a",c:"#854d0e"}};
const newComparison=()=>({project:"",trade:"",reference:"",preparedDate:"",bidders:[],sections:[],commercialTerms:{},riskFlags:[],ts:Date.now()});
const newCmpItem=()=>({id:uid(),no:"",desc:"New item",unit:"no",whQty:"",bidders:{}});
const cmpBidderCell=(item,name)=>item.bidders[name]||{qty:"",rate:"",status:"normal",remarks:""};
const median=arr=>{if(!arr.length)return null;const s=[...arr].sort((a,b)=>a-b),m=Math.floor(s.length/2);return s.length%2?s[m]:(s[m-1]+s[m])/2;};
// Returns the numeric amount for a bidder on an item (whQty x rate), or null if not computable (incl/norate/excluded/no rate/no whQty).
const cmpAmount=(item,name)=>{const b=cmpBidderCell(item,name);if(b.status!=="normal")return null;const wq=+item.whQty,r=+b.rate;if(!wq||!b.rate)return null;return wq*r;};
// Computes per-item-per-bidder auto flags (qty mismatch / outlier rate) without mutating anything.
const cmpAutoFlags=(section,bidders)=>{
  const flags={}; // flags[itemId] = { [bidderName]: {qty:bool, rate:bool} }
  section.items.forEach(item=>{
    const wq=+item.whQty||0;
    const priced=bidders.map(name=>({name,rate:+cmpBidderCell(item,name).rate})).filter(x=>cmpBidderCell(item,x.name).status==="normal"&&x.rate>0);
    const med=median(priced.map(x=>x.rate));
    const f={};
    bidders.forEach(name=>{
      const b=cmpBidderCell(item,name);
      const qtyFlag=wq>0&&b.qty!==""&&b.qty!=null&&Math.abs((+b.qty-wq)/wq)>QTY_MISMATCH_PCT;
      const rateFlag=med&&b.status==="normal"&&+b.rate>0&&Math.abs((+b.rate-med)/med)>OUTLIER_PCT;
      f[name]={qty:!!qtyFlag,rate:!!rateFlag};
    });
    flags[item.id]=f;
  });
  return flags;
};

const newSections=()=>SECTIONS.map(s=>({...s,items:[]}));
const fmt=n=>(n||0).toLocaleString("en-SG",{minimumFractionDigits:2,maximumFractionDigits:2});
// Resize/compress an image File or Blob to a small JPEG data URL (for storing quote photos).
// Aggressively shrink a data-URL image for embedding into Excel (small dimensions + low quality).
// Work out a typed calculation like "((1+1)*2)+((3*1)+2)" or "=12.5*1.1". Only numbers, + - * / and brackets; returns NaN if invalid.
// A BUR code inside a formula stands for that item's current rate, e.g. "=rebar+0.13". The app sets exprLookup (code -> rate).
let exprLookup=null;
const exprRefs=str=>(String(str??"").replace(/^\s*=/,"").match(/[A-Za-z_][A-Za-z0-9_.&]*/g)||[]); // codes may contain "&" (e.g. Door_S&I_no)
const calcExpr=str=>{ const s=String(str??"").replace(/^\s*=/,"").replace(/\s+/g,""); if(!s||/[^0-9A-Za-z_.&+\-*/()]/.test(s))return NaN; let i=0;
  const num=()=>{ const m=/^\d+\.?\d*|^\.\d+/.exec(s.slice(i)); if(!m)return NaN; i+=m[0].length; return parseFloat(m[0]); };
  const fac=()=>{ if(/[A-Za-z_]/.test(s[i]||"")){ const m=/^[A-Za-z_][A-Za-z0-9_.&]*/.exec(s.slice(i)); i+=m[0].length; return exprLookup?+exprLookup(m[0]):NaN; } if(s[i]==="-"){i++;return -fac();} if(s[i]==="+"){i++;return fac();} if(s[i]==="("){i++; const v=sum(); if(s[i]!==")")return NaN; i++; return v;} return num(); };
  const prod=()=>{ let v=fac(); while(s[i]==="*"||s[i]==="/"){ const op=s[i++], r=fac(); v=op==="*"?v*r:v/r; } return v; };
  const sum=()=>{ let v=prod(); while(s[i]==="+"||s[i]==="-"){ const op=s[i++], r=prod(); v=op==="+"?v+r:v-r; } return v; };
  const v=sum(); return i===s.length&&isFinite(v)?v:NaN; };
// A cost-data entry priced in a different unit is converted to the item's unit: rate ÷ cover, where cover = how many item-units
// (e.g. m²) one entry-unit covers. Example: plant S$ 3,500 per no., 1 no. covers 50 m² → S$ 70 / m². cover may be a formula (e.g. 1.2*2.4*50).
const cdCover=e=>{ const c=calcExpr(e&&e.cover); return isFinite(c)&&c>0?c:1; };
const cdRate=e=>(+e.rate||0)/cdCover(e);
const imgSize=dataUrl=>new Promise(res=>{ const im=new Image(); im.onload=()=>res({w:im.naturalWidth||im.width||210,h:im.naturalHeight||im.height||160}); im.onerror=()=>res({w:210,h:160}); im.src=dataUrl; });
const shrinkDataURL=(dataUrl,max=460,q=0.45)=>new Promise(res=>{ if(!dataUrl){res("");return;} const im=new Image(); im.onload=()=>{ let w=im.width,h=im.height; if(w>max){h=h*max/w;w=max;} if(h>max){w=w*max/h;h=max;} const cv=document.createElement("canvas"); cv.width=w; cv.height=h; cv.getContext("2d").drawImage(im,0,0,w,h); try{res(cv.toDataURL("image/jpeg",q));}catch{res(dataUrl);} }; im.onerror=()=>res(dataUrl); im.src=dataUrl; });
const imgToDataURL=blob=>new Promise(res=>{ if(!blob){res("");return;} const fr=new FileReader(); fr.onload=()=>{ const im=new Image(); im.onload=()=>{ const max=1100; let w=im.width,h=im.height; if(w>max){h=h*max/w;w=max;} if(h>max){w=w*max/h;h=max;} const cv=document.createElement("canvas"); cv.width=w; cv.height=h; cv.getContext("2d").drawImage(im,0,0,w,h); try{res(cv.toDataURL("image/jpeg",0.7));}catch{res("");} }; im.onerror=()=>res(""); im.src=fr.result; }; fr.onerror=()=>res(""); fr.readAsDataURL(blob); });
const uid=()=>"_"+Math.random().toString(36).slice(2,10);
// BUR codes carry their unit as a suffix, e.g. Acoustic_feltCeiling_m2. Lookups accept the code with or without the suffix,
// so older BOQ rows / Master BQ templates that still use the bare code keep matching.
const unitTag=u=>String(u||"").trim().replace(/²/g,"2").replace(/³/g,"3").replace(/%/g,"pct").replace(/[\s.]+/g,""); // "no." -> "no", "m run" -> "mrun"
const codeBase=(code,unit)=>{ const c=String(code||"").trim(), t=unitTag(unit); return t&&c.toLowerCase().endsWith("_"+t.toLowerCase())?c.slice(0,-(t.length+1)):c; };
const codeWithUnit=(code,unit)=>{ const c=String(code||"").trim(), t=unitTag(unit); if(!c||!t)return c; return c.toLowerCase().endsWith("_"+t.toLowerCase())?c:`${c}_${t}`; };
const swapUnitTag=(code,oldU,newU)=>String(code||"").trim()?codeWithUnit(codeBase(code,oldU),newU):code;
// Read a master BUR workbook (layout of BUR.xlsx): orange cells in column A start a category (a plain column-A heading also
// starts one when it matches an existing category name); every row with a Code in D is an item: B desc, C unit, D code, E rate.
// Supplier names sit in R onwards on heading rows (e.g. "AID\n(SLA Sep 26)"); numbers under them on item rows become quotes.
const parseBurWorkbook=(wb,catNames=[])=>{
  const ws=wb.worksheets[0]; const known=new Set(catNames.map(n=>String(n).trim().toLowerCase()));
  const val=c=>{ const v=c.value; if(v&&typeof v==="object"){ if("result" in v)return v.result; if(v.richText)return v.richText.map(t=>t.text).join(""); if(v.text!=null)return v.text; return undefined; } return v; };
  const str=c=>{ const v=val(c); return v==null?"":String(v).trim(); };
  const SUP0=18, SUPN=28; let cat="(Uncategorised)", hdr={}, sub=""; const items=[];
  const DATE_TAG=/\((SLA|PS|CAG|WAS|Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)/i;
  for(let r=1;r<=ws.rowCount;r++){
    const A=str(ws.getCell(r,1)),B=str(ws.getCell(r,2)); let D=str(ws.getCell(r,4));
    if(r<6&&/^description$/i.test(A))continue;
    const orange=(ws.getCell(r,1).fill?.fgColor?.argb||"").toUpperCase()==="FFFFC000";
    if(orange&&A)D=""; // an orange category heading is never an item, even with text (e.g. "Mortar") in the Code column
    if(!D&&A&&(orange||(!B&&known.has(A.toLowerCase())))){ cat=A; hdr={}; sub=""; }
    if(!D&&B){ const bc=ws.getCell(r,2), rt=bc.value&&bc.value.richText; if(bc.font?.underline||(rt&&rt.some(t=>t.font&&t.font.underline))) sub=B.replace(/\s+/g," ").replace(/[-:]+$/,"").trim(); }
    if(!D){
      // supplier-name row: a category row, 2+ names, or a single pale-yellow / date-tagged name. Lone notes (sizes, remarks) are ignored.
      const texts=[]; for(let c=SUP0;c<=SUPN;c++){ const cell=ws.getCell(r,c), v=val(cell); if(typeof v==="string"&&v.trim())texts.push({c,v:v.trim(),yellow:(cell.fill?.fgColor?.argb||"").toUpperCase()==="FFFFFFCC"}); }
      const isHdr=cat===A||texts.length>=2||texts.some(t=>t.yellow||DATE_TAG.test(t.v));
      if(isHdr) texts.forEach(t=>{ const one=t.v.replace(/\s*\r?\n\s*/g," "); const m=/^(.*?)\s*\(([^)]*)\)?\s*$/.exec(one); hdr[t.c]=m&&m[1]?{supplier:m[1].trim(),date:m[2].trim()}:{supplier:one,date:""}; });
      continue;
    }
    const num=c=>{ const v=val(ws.getCell(r,c)); return typeof v==="number"&&isFinite(v)?v:null; };
    const rate=num(5)??num(6)??0;
    const quotes=[]; for(let c=SUP0;c<=SUPN;c++){ const q=num(c); if(q!=null&&q!==0&&hdr[c])quotes.push({supplier:hdr[c].supplier,date:hdr[c].date,rate:Math.round(q*100)/100}); }
    items.push({cat,group:sub,ord:r,code:D,desc:B||D,unit:str(ws.getCell(r,3))||"sum",rate:Math.round(rate*100)/100,quotes});
  }
  return items;
};
// Group a category's items under their sub-category (item.group). Items with no sub-category come first; sub-categories follow
// in the order they appear in the master Excel (item.ord), then A–Z. Order inside each sub-category is kept as given.
const orderBySubcat=list=>{ const first={}, minOrd={}; list.forEach((b,i)=>{ const g=String(b.group||"").trim(); if(!(g in first))first[g]=i; if(b.ord!=null)minOrd[g]=Math.min(minOrd[g]??Infinity,+b.ord); });
  const keys=Object.keys(first).sort((a,b)=>a===""?-1:b===""?1:((minOrd[a]??Infinity)-(minOrd[b]??Infinity))||a.localeCompare(b,undefined,{numeric:true}));
  return keys.flatMap(k=>list.filter(b=>String(b.group||"").trim()===k)); };
// Codes with spaces or "/" (e.g. "cage_2SH/LH 10100") can't be used in formulas, so they get "_" instead; matching treats them as the same.
const cleanCode=c=>String(c||"").trim().replace(/[\s/]+/g,"_");
const findBurByCode=(items,code)=>{ const k=cleanCode(code).toLowerCase(); if(!k)return undefined; const lc=s=>cleanCode(s).toLowerCase();
  return items.find(b=>lc(b.code)===k)||items.find(b=>b.code&&(lc(codeBase(b.code,b.unit))===k||lc(codeWithUnit(b.code,b.unit))===k)); };
const ROUND_OPTS=[["","Exact"],["0","Whole $"],["1","1 decimal"],["2","2 decimals"]];
const roundLabel=v=>({0:"whole $",1:"1 decimal",2:"2 decimals"})[+v]||"";
const hasRoundUp=b=>b&&b.roundUp!==undefined&&b.roundUp!==null&&b.roundUp!=="";
// Rate = (labour+material+plant+subcon) + OH% + profit%, then rounded UP if the item asks for it (like Excel ROUNDUP(x,0) / ROUNDUP(x,1)).
// Rate build-up (same idea as BUR.xlsx columns H–Q): direct (labour+material+plant+subcon) + waste % of direct + extra lump-sum items.
// item.waste = waste %, item.extras = [{id,name,value}] each a lump-sum S$ per unit (older "%"-type lines are still honoured). No OH / profit.
const EXTRA_PRESETS=[["Mortar","$"],["Mesh","$"],["Accessories","$"]];
const baseDirect=b=>(+b.labour||0)+(+b.material||0)+(+b.plant||0)+(+b.subcon||0);
const extraAmt=(b,x)=>x.type==="%"?baseDirect(b)*(+x.value||0)/100:(+x.value||0);
const extrasTot=b=>(b.extras||[]).reduce((t,x)=>t+extraAmt(b,x),0);
// item.discount = % taken off the direct cost first; Waste/ OH % is then worked out on the discounted amount.
const discAmt=b=>baseDirect(b)*(+b.discount||0)/100;
const netDirect=b=>baseDirect(b)-discAmt(b);
const wasteAmt=b=>netDirect(b)*(+b.waste||0)/100;
const bTot=b=>{const t=netDirect(b)+wasteAmt(b)+extrasTot(b);if(!hasRoundUp(b))return t;const f=10**(+b.roundUp);return Math.ceil(t*f-1e-9)/f;};
// Custom BOQ columns: compute each column's value per row. Formula columns can reference
// qty, rateA, rateB, amtA, amtB, and earlier custom columns (by sanitised name).
// Compute custom BOQ columns. Formulas may use natural names: Qty, Rate A, Rate B,
// Amt A, Amt B, and any other column's name (spaces ok, case-insensitive).
function computeCols(item,cols){
  const qty=+item.qty||0, rA=+item.rA||0, rB=+item.rB||0, amtA=qty*rA, amtB=qty*rB;
  const map={}; const put=(n,v)=>{ if(n)map[String(n).trim().toLowerCase()]=v; };
  put("qty",qty); put("quantity",qty);
  put("rate a",rA); put("ratea",rA); put("rate",rA);
  put("rate b",rB); put("rateb",rB);
  put("amt a",amtA); put("amount a",amtA); put("amta",amtA); put("amt",amtA);
  put("amt b",amtB); put("amount b",amtB); put("amtb",amtB);
  const evalNamed=f=>{
    let e=String(f).replace(/^\s*=/,"");
    const names=Object.keys(map).sort((a,b)=>b.length-a.length);
    for(const nm of names){ const pat=nm.replace(/[.*+?^${}()|[\]\\]/g,"\\$&").replace(/\s+/g,"\\s*"); e=e.replace(new RegExp(pat,"gi"),"("+(map[nm]||0)+")"); }
    try{ const r=Function('"use strict";return ('+e+');')(); return (typeof r==="number"&&isFinite(r))?r:"ERR"; }catch{ return "ERR"; }
  };
  // Pre-register every column name (init 0 / manual value) so any reference resolves — no ERR from names.
  for(const c of (cols||[])){ if(c.formula)put(c.label,0); else put(c.label,parseFloat(item.cx?.[c.id]??"")||0); }
  const out={};
  // A few passes so columns can reference each other in any order.
  for(let pass=0;pass<4;pass++){
    for(const c of (cols||[])){
      if(c.formula){ const v=evalNamed(c.formula); out[c.id]=v; put(c.label,typeof v==="number"?v:0); }
      else { const val=item.cx?.[c.id]??""; out[c.id]=val; }
    }
  }
  return out;
}
const authErrMsg=e=>{const c=(e&&e.code)||"";if(c.includes("invalid-cred")||c.includes("wrong-password")||c.includes("user-not-found"))return"Incorrect email or password.";if(c.includes("invalid-email"))return"That doesn't look like a valid email.";if(c.includes("too-many"))return"Too many attempts — try again later.";if(c.includes("network"))return"Network error — check your connection.";return"Sign-in failed. "+(e&&e.message||"");};

export default function App(){
  const [authUser,setAuthUser]=useState(undefined); // undefined=loading, null=signed out, object=signed in
  const [profile,setProfile]=useState(null);         // {name,role,email}
  const [email,setEmail]=useState(""); const [pw,setPw]=useState(""); const [authErr,setAuthErr]=useState(""); const [authBusy,setAuthBusy]=useState(false);
  const [nameIn,setNameIn]=useState(""); const [roleIn,setRoleIn]=useState("Estimator");

  const [projects,setProjects]=useState(null);        // null=loading, []=none
  const [pid,setPid]=useState(null);

  const [tab,setTab]=useState("boq");
  const [data,setData]=useState(null);                // current project's BOQ: {sections, ts}
  const [selSec,setSelSec]=useState("prelim");
  const [visCols,setVisCols]=useState(DEF_COLS);
  const [showColPick,setShowColPick]=useState(false);
  const [rateSugg,setRateSugg]=useState(null);
  const [cats,setCats]=useState(DEF_CATS);
  const [burItems,setBurItems]=useState([]);
  const [selCat,setSelCat]=useState(DEF_CATS[0].id);
  const [expBur,setExpBur]=useState(null);
  const [editExpr,setEditExpr]=useState(null); // id of the rate / amount box being typed in
  const [syncState,setSyncState]=useState(null); // "Sync from project" window: {src, loading, rows:[{key,kind,cur,src,diffs}], sel:Set, opts}
  const [collapsedSubs,setCollapsedSubs]=useState(()=>{ try{ return new Set(JSON.parse(localStorage.getItem("burCollapsedSubs")||"[]")); }catch{ return new Set(); } }); // "catId|sub-category" keys folded in the BUR table
  const [burSearch,setBurSearch]=useState(""); const [burSearchCat,setBurSearchCat]=useState(false);
  const [sortBy,setSortBy]=useState("code"); const [sortDir,setSortDir]=useState(1);
  const [burView,setBurView]=useState("tabs"); const [catSort,setCatSort]=useState("none");
  const [selBur,setSelBur]=useState(()=>new Set());
  const [boqMode,setBoqMode]=useState("structured");
  const [burColW,setBurColW]=useState(()=>{try{return JSON.parse(localStorage.getItem("burColW")||"{}")}catch{return{}}});
  const [pasteOpen,setPasteOpen]=useState(false); const [pasteText,setPasteText]=useState(""); const [pasteCat,setPasteCat]=useState("");
  const [costModal,setCostModal]=useState(null);
  const [cType,setCType]=useState("subcon");
  const [cForm,setCForm]=useState({supplier:"",rate:"",date:"",location:"",note:"",imgs:[]});
  const [imgView,setImgView]=useState(null);
  const [ceSort,setCeSort]=useState({by:"none",dir:1}); const [showGraph,setShowGraph]=useState(false);
  const [ccomps,setCcomps]=useState([]);
  const [newCat,setNewCat]=useState(""); const [showNewCat,setShowNewCat]=useState(false);
  const [codes,setCodes]=useState(DEF_CODES);
  const [editCId,setEditCId]=useState(null); const [codeForm,setCodeForm]=useState({code:"",desc:"",cat:""}); const [showAddC,setShowAddC]=useState(false);
  const [rbu,setRbu]=useState({labour:"",material:"",plant:"",subcon:"",oh:15,profit:10});
  const [log,setLog]=useState([]);
  const [toast,setToast]=useState(null);
  const [bqTplB64,setBqTplB64]=useState(null); const [bqTplName,setBqTplName]=useState(""); const [bqResult,setBqResult]=useState(null);

  const [comparisons,setComparisons]=useState([]);
  const [selCompId,setSelCompId]=useState(null);
  const [newBidderName,setNewBidderName]=useState("");
  const [newFlag,setNewFlag]=useState({severity:"Medium",type:"",item:"",bidder:"",detail:""});

  const sTimer=useRef(null); const uRef=useRef(null); const pidRef=useRef(null); const dirtyRef=useRef(false); const pendingRef=useRef(null); const logRef=useRef([]); const burTimers=useRef({});
  // Compare Quotes: same optimistic-edit + debounce idea as the BOQ project doc (writeProject/updItem above),
  // but scoped to whichever single comparison is currently open — cmpDirtyRef guards only that one id against
  // being clobbered by our own collection-wide onSnapshot while a debounced write is in flight.
  const cmpTimer=useRef(null); const cmpDirtyRef=useRef(null); const cmpPendingRef=useRef(null);
  const burDragRef=useRef(null); const burColWRef=useRef(burColW); useEffect(()=>{burColWRef.current=burColW;},[burColW]);
  const burItemsRef=useRef([]); useEffect(()=>{burItemsRef.current=burItems;});
  useEffect(()=>{ const mm=e=>{ const d=burDragRef.current; if(!d)return; const w=Math.max(40,d.startW+(e.clientX-d.start)); setBurColW(p=>({...p,[d.k]:w})); }; const mu=()=>{ if(burDragRef.current){ try{localStorage.setItem("burColW",JSON.stringify(burColWRef.current));}catch{} burDragRef.current=null; } }; window.addEventListener("mousemove",mm); window.addEventListener("mouseup",mu); return ()=>{window.removeEventListener("mousemove",mm);window.removeEventListener("mouseup",mu);}; },[]);
  const user=authUser&&profile?{uid:authUser.uid,email:authUser.email,name:profile.name,role:profile.role}:null;
  useEffect(()=>{uRef.current=user;});
  useEffect(()=>{pidRef.current=pid;},[pid]);
  const toast_=useCallback(msg=>{setToast(msg);setTimeout(()=>setToast(null),3500);},[]);
  const ready=!!authUser&&!!profile;

  // ── Auth ──────────────────────────────────────────────────────────────────
  useEffect(()=>onAuthStateChanged(auth,u=>setAuthUser(u||null)),[]);
  useEffect(()=>{
    if(!authUser){setProfile(null);return;}
    return onSnapshot(doc(db,"users",authUser.uid),s=>setProfile(s.exists()?s.data():null),()=>setProfile(null));
  },[authUser]);

  const doLogin=async()=>{
    if(!email.trim()||!pw)return; setAuthBusy(true);setAuthErr("");
    try{await signInWithEmailAndPassword(auth,email.trim(),pw);}catch(e){setAuthErr(authErrMsg(e));}
    setAuthBusy(false);
  };
  const saveProfile=async()=>{ if(!nameIn.trim()||!authUser)return; try{await setDoc(doc(db,"users",authUser.uid),{name:nameIn.trim(),role:roleIn,email:authUser.email});}catch(e){toast_("⚠️ "+e.message);} };

  // ── Activity log ────────────────────────────────────────────────────────────
  const addLogEntry=useCallback(async action=>{
    if(!action||!uRef.current)return;
    const e={id:uid(),user:uRef.current.name,role:uRef.current.role,action,time:new Date().toLocaleString("en-SG")};
    const nl=[e,...logRef.current].slice(0,80); logRef.current=nl; setLog(nl);
    try{await setDoc(doc(db,"meta","log"),{entries:nl});}catch{}
  },[]);

  // ── Live subscriptions (projects, BUR library, codes, cats, log) ────────────
  useEffect(()=>{ if(!ready)return; return onSnapshot(collection(db,"projects"),snap=>{
    const list=snap.docs.map(d=>({id:d.id,name:d.data().name||"Untitled",createdAt:d.data().createdAt||0})).sort((a,b)=>a.createdAt-b.createdAt);
    setProjects(list); setPid(p=>p||(list[0]?list[0].id:null));
  }); },[ready]);

  const migratedRef=useRef(false);
  useEffect(()=>{ if(!ready||!pid)return; return onSnapshot(collection(db,"bur"),async snap=>{
    const all=snap.docs.map(d=>({id:d.id,...d.data()}));
    // Show this project's items (and any legacy items without a project) ...
    setBurItems(all.filter(b=>b.pid===pid||!b.pid));
    // ... then one-time adopt legacy (pid-less) items into the current project so they become project-scoped.
    const legacy=all.filter(b=>!b.pid);
    if(legacy.length&&!migratedRef.current){ migratedRef.current=true; try{ for(let i=0;i<legacy.length;i+=400){ const batch=writeBatch(db); legacy.slice(i,i+400).forEach(b=>batch.update(doc(db,"bur",b.id),{pid})); await batch.commit(); } }catch{} }
  }); },[ready,pid]);

  useEffect(()=>{ if(!ready||!pid)return; return onSnapshot(collection(db,"comparisons"),snap=>{
    const all=snap.docs.map(d=>({id:d.id,...d.data()})).filter(c=>c.pid===pid);
    // Keep our own in-flight edit on screen instead of the (stale, pre-write) server copy for that one doc.
    if(cmpDirtyRef.current&&cmpPendingRef.current){ const i=all.findIndex(c=>c.id===cmpDirtyRef.current); if(i>=0)all[i]=cmpPendingRef.current; }
    setComparisons(all); setSelCompId(prev=>prev&&all.some(c=>c.id===prev)?prev:(all[0]?.id||null));
  }); },[ready,pid]);
  useEffect(()=>{ if(!ready)return; const ref=doc(db,"meta","codes"); return onSnapshot(ref,s=>{ if(s.exists())setCodes(s.data().list||[]); else{setCodes(DEF_CODES);setDoc(ref,{list:DEF_CODES}).catch(()=>{});} }); },[ready]);
  useEffect(()=>{ if(!ready)return; const ref=doc(db,"meta","cats"); return onSnapshot(ref,s=>{ if(s.exists())setCats(s.data().list||[]); else{setCats(DEF_CATS);setDoc(ref,{list:DEF_CATS}).catch(()=>{});} }); },[ready]);
  // Keep the selected category valid when the category list changes (e.g. after loading the master list).
  useEffect(()=>{ if(cats.length&&!cats.some(c=>c.id===selCat))setSelCat(cats[0].id); },[cats]); // eslint-disable-line
  // Per-project master BQ template override (uploaded). Falls back to the bundled CAG template.
  useEffect(()=>{ if(!ready||!pid){setBqTplB64(null);setBqTplName("");return;} return onSnapshot(doc(db,"boqtemplate",pid),s=>{ if(s.exists()){setBqTplB64(s.data().b64||null);setBqTplName(s.data().name||"");}else{setBqTplB64(null);setBqTplName("");} }); },[ready,pid]);
  useEffect(()=>{ if(!ready)return; return onSnapshot(doc(db,"meta","log"),s=>{ const e=s.exists()?(s.data().entries||[]):[]; logRef.current=e; setLog(e); }); },[ready]);
  useEffect(()=>{ if(!ready)return; return onSnapshot(doc(db,"meta","ccomps"),s=>setCcomps(s.exists()?(s.data().list||[]):[])); },[ready]);
  const addComponentTab=async()=>{ const name=prompt("New cost-data tab name (e.g. Equipment, Transport, Specialist):"); if(!name||!name.trim())return; const key="cc_"+Math.random().toString(36).slice(2,8); try{await setDoc(doc(db,"meta","ccomps"),{list:[...ccomps,{key,label:name.trim()}]});setCType(key);}catch(e){toast_("⚠️ "+e.message);} };
  const delComponentTab=async key=>{ if(!confirm("Remove this tab? Its entries are hidden and no longer count in any rate."))return; try{await setDoc(doc(db,"meta","ccomps"),{list:ccomps.filter(c=>c.key!==key)});}catch{} if(cType===key)setCType("subcon");
    const hit=burItemsRef.current.filter(b=>(b.costData||[]).some(e=>e.inRate&&e.component===key)); if(!hit.length)return;
    try{ recordFields(hit.map(b=>b.id),["costData","labour","material","plant","subcon","extras","preTick"],`Remove tab from ${hit.length} rate(s)`);
      for(let i=0;i<hit.length;i+=400){ const batch=writeBatch(db); hit.slice(i,i+400).forEach(b=>{ const cd=b.costData.map(e=>e.component===key&&e.inRate?{...e,inRate:false}:e); batch.update(doc(db,"bur",b.id),{costData:cd,...tickPatch(b,cd)}); }); await batch.commit(); }
      toast_(`✅ Tab removed — ${hit.length} item rate(s) recalculated`); }catch(e){ toast_("⚠️ "+e.message); } };

  // ── Quote comparison ops ────────────────────────────────────────────────────
  const selComp=comparisons.find(c=>c.id===selCompId)||null;
  // Structural changes (add/remove section, bidder, item) write immediately; per-keystroke edits
  // (rate/qty/remarks) go through this debounced path, same shape as writeProject/updItem for BOQ.
  const writeComp=nd=>{ if(!nd?.id)return; const {id,...rest}=nd; setDoc(doc(db,"comparisons",id),rest,{merge:true}).catch(()=>{}); };
  const pushComp=useCallback(nd=>{ cmpDirtyRef.current=null; cmpPendingRef.current=null; setComparisons(prev=>prev.map(c=>c.id===nd.id?nd:c)); writeComp(nd); },[]);
  const updComp=useCallback(patch=>{
    if(!selComp)return;
    cmpDirtyRef.current=selComp.id;
    setComparisons(prev=>prev.map(c=>{ if(c.id!==selComp.id)return c; const nd={...c,...patch,ts:Date.now()}; cmpPendingRef.current=nd; return nd; }));
    if(cmpTimer.current)clearTimeout(cmpTimer.current);
    cmpTimer.current=setTimeout(()=>{ const nd=cmpPendingRef.current; if(nd){ writeComp(nd); cmpDirtyRef.current=null; cmpPendingRef.current=null; } },700);
  },[selComp]);
  const addComparison=useCallback(async()=>{
    if(!pid)return;
    const ref=doc(collection(db,"comparisons"));
    try{ await setDoc(ref,{...newComparison(),pid}); setSelCompId(ref.id); addLogEntry("Created a new quote comparison"); }catch(e){toast_("⚠️ "+e.message);}
  },[pid,addLogEntry,toast_]);
  const delComparison=useCallback(id=>{ if(!confirm("Delete this comparison? This cannot be undone."))return; deleteDoc(doc(db,"comparisons",id)).catch(e=>toast_("⚠️ "+e.message)); if(selCompId===id)setSelCompId(null); },[selCompId,toast_]);
  const addBidder=useCallback(()=>{ const name=newBidderName.trim(); if(!name||!selComp||selComp.bidders.includes(name))return; updComp({bidders:[...selComp.bidders,name]}); setNewBidderName(""); },[newBidderName,selComp,updComp]);
  const removeBidder=useCallback(name=>{
    if(!selComp)return; if(!confirm(`Remove ${name} from this comparison? Their quoted rates/qty on every item will be lost.`))return;
    const sections=selComp.sections.map(s=>({...s,items:s.items.map(it=>{const b={...it.bidders};delete b[name];return{...it,bidders:b};})}));
    const ct={...selComp.commercialTerms};delete ct[name];
    updComp({bidders:selComp.bidders.filter(b=>b!==name),sections,commercialTerms:ct,riskFlags:selComp.riskFlags.filter(f=>f.bidder!==name)});
  },[selComp,updComp]);
  const addSection=useCallback(()=>{ if(!selComp)return; updComp({sections:[...selComp.sections,{id:uid(),name:"New Section",items:[]}]}); },[selComp,updComp]);
  const delSection=useCallback(sid=>{ if(!selComp||!confirm("Delete this section and all its items?"))return; updComp({sections:selComp.sections.filter(s=>s.id!==sid)}); },[selComp,updComp]);
  const renameSection=useCallback((sid,name)=>{ if(!selComp)return; updComp({sections:selComp.sections.map(s=>s.id===sid?{...s,name}:s)}); },[selComp,updComp]);
  const addCmpItem=useCallback(sid=>{ if(!selComp)return; updComp({sections:selComp.sections.map(s=>s.id===sid?{...s,items:[...s.items,newCmpItem()]}:s)}); },[selComp,updComp]);
  const delCmpItem=useCallback((sid,iid)=>{ if(!selComp)return; updComp({sections:selComp.sections.map(s=>s.id===sid?{...s,items:s.items.filter(i=>i.id!==iid)}:s)}); },[selComp,updComp]);
  const updCmpItem=useCallback((sid,iid,patch)=>{ if(!selComp)return; updComp({sections:selComp.sections.map(s=>s.id===sid?{...s,items:s.items.map(i=>i.id===iid?{...i,...patch}:i)}:s)}); },[selComp,updComp]);
  const updCmpBidderCell=useCallback((sid,iid,name,patch)=>{
    if(!selComp)return;
    updComp({sections:selComp.sections.map(s=>s.id===sid?{...s,items:s.items.map(i=>i.id===iid?{...i,bidders:{...i.bidders,[name]:{...cmpBidderCell(i,name),...patch}}}:i)}:s)});
  },[selComp,updComp]);
  const updCommercialTerm=useCallback((name,patch)=>{ if(!selComp)return; updComp({commercialTerms:{...selComp.commercialTerms,[name]:{...(selComp.commercialTerms[name]||{}),...patch}}}); },[selComp,updComp]);
  const addRiskFlag=useCallback(()=>{ if(!selComp||!newFlag.type.trim()||!newFlag.detail.trim())return; updComp({riskFlags:[...selComp.riskFlags,{id:uid(),...newFlag}]}); setNewFlag({severity:"Medium",type:"",item:"",bidder:"",detail:""}); },[selComp,newFlag,updComp]);
  const delRiskFlag=useCallback(fid=>{ if(!selComp)return; updComp({riskFlags:selComp.riskFlags.filter(f=>f.id!==fid)}); },[selComp,updComp]);

  // Per-bidder grand total across every section (null-amount items excluded, so an all-unpriced bidder totals 0 rather than looking cheapest).
  const cmpBidderTotal=useCallback((comp,name)=>comp.sections.reduce((sum,s)=>sum+s.items.reduce((t,i)=>t+(cmpAmount(i,name)||0),0),0),[]);
  const cmpUnpricedCount=useCallback((comp,name)=>comp.sections.reduce((n,s)=>n+s.items.filter(i=>{const b=cmpBidderCell(i,name);return b.status==="norate"||(b.status==="normal"&&!(+i.whQty&&+b.rate));}).length,0),[]);
  const exportComparisonCSV=useCallback(comp=>{
    const rows=[["Project",comp.project],["Trade",comp.trade],[]];
    const header=["No","Description","Unit","WH Qty",...comp.bidders.flatMap(b=>[`${b} Qty`,`${b} Rate`,`${b} Amount`,`${b} Remarks`])];
    rows.push(header);
    comp.sections.forEach(s=>{
      rows.push([s.name]);
      s.items.forEach(i=>{
        const row=[i.no,i.desc,i.unit,i.whQty];
        comp.bidders.forEach(name=>{const b=cmpBidderCell(i,name);const amt=cmpAmount(i,name);row.push(b.qty,b.rate,amt!=null?amt.toFixed(2):CMP_STATUS_LABEL[b.status],b.remarks);});
        rows.push(row);
      });
    });
    rows.push([]); rows.push(["TOTAL",...comp.bidders.flatMap(b=>["","",cmpBidderTotal(comp,b).toFixed(2),""])]);
    const csv=rows.map(r=>r.map(v=>{const s=String(v??"");return /[",\n]/.test(s)?`"${s.replace(/"/g,'""')}"`:s;}).join(",")).join("\n");
    const blob=new Blob([csv],{type:"text/csv;charset=utf-8;"});
    const a=document.createElement("a"); a.href=URL.createObjectURL(blob); a.download=`Comparison_${(comp.trade||comp.project||"quotes").replace(/[^a-z0-9]+/gi,"_")}.csv`; a.click(); URL.revokeObjectURL(a.href);
  },[cmpBidderTotal]);

  // Current project's BOQ
  useEffect(()=>{ if(!ready||!pid){setData(null);return;} return onSnapshot(doc(db,"projects",pid),s=>{ if(!s.exists())return; if(dirtyRef.current)return; const d=s.data(); setData({sections:d.sections||newSections(),cols:d.cols||[],ts:d.ts||0}); }); },[ready,pid]);

  // ── Project ops ─────────────────────────────────────────────────────────────
  const createProject=async()=>{
    const name=prompt("New project / tender name:"); if(!name||!name.trim())return;
    let sections=newSections(), cols=[];
    const cur=projects&&projects.find(p=>p.id===pid);
    if(data&&cur&&confirm(`Copy the BOQ from "${cur.name}" into the new project?\n\nOK = duplicate its BOQ rows & columns\nCancel = start with a blank BOQ`)){ sections=JSON.parse(JSON.stringify(data.sections||newSections())); cols=JSON.parse(JSON.stringify(data.cols||[])); }
    const copyBur = cur && burItems.length>0 && confirm(`Copy the BUR rate library (${burItems.length} items) from "${cur.name}" into the new project?\n\nOK = copy all rates as a starting point\nCancel = start with an EMPTY BUR (then use "Paste from Excel" to bring in your own)`);
    const ref=doc(collection(db,"projects"));
    try{
      await setDoc(ref,{name:name.trim(),createdAt:Date.now(),sections,cols,ts:Date.now()});
      if(copyBur){ const src=burItems.slice(); for(let i=0;i<src.length;i+=400){ const batch=writeBatch(db); src.slice(i,i+400).forEach(b=>{ const {id,...rest}=b; batch.set(doc(collection(db,"bur")),{...rest,pid:ref.id}); }); await batch.commit(); } }
      setPid(ref.id); setTab("boq"); addLogEntry(`Created project "${name.trim()}"${copyBur?" (BUR copied)":""}`);
    }catch(e){toast_("⚠️ "+e.message);}
  };
  const renameProject=async()=>{ const cur=projects?.find(p=>p.id===pid); const name=prompt("Rename project:",cur?.name||""); if(!name||!name.trim())return; try{await updateDoc(doc(db,"projects",pid),{name:name.trim()});}catch(e){toast_("⚠️ "+e.message);} };
  const deleteProject=async()=>{ if(!pid)return; const cur=projects?.find(p=>p.id===pid); if(!confirm(`Delete project "${cur?.name}" and all its BOQ items? This cannot be undone.`))return; try{await deleteDoc(doc(db,"projects",pid)); setPid(null);}catch(e){toast_("⚠️ "+e.message);} };

  // ── BOQ writes ──────────────────────────────────────────────────────────────
  const writeProject=nd=>{ if(!pidRef.current)return; setDoc(doc(db,"projects",pidRef.current),{sections:nd.sections,cols:nd.cols||[],ts:Date.now()},{merge:true}).catch(()=>{}); };
  const pushData=useCallback((nd,action)=>{ dirtyRef.current=false; setData({...nd}); writeProject(nd); if(action)addLogEntry(action); },[addLogEntry]);

  const addItem=useCallback(sid=>{ if(!data)return; const nd=JSON.parse(JSON.stringify(data)); const sec=nd.sections.find(s=>s.id===sid); if(sec){sec.items.push({id:uid(),ref:"",desc:"New item",unit:"sum",qty:1,rA:0,rB:0,code:"",status:"Draft",remarks:"",by:uRef.current?.name}); pushData(nd,`Added item in ${sec.name}`);} },[data,pushData]);

  const updItem=useCallback((sid,iid,ch)=>{
    dirtyRef.current=true; let secName;
    setData(prev=>{ if(!prev)return prev; const nd=JSON.parse(JSON.stringify(prev)); const sec=nd.sections.find(s=>s.id===sid); const item=sec?.items.find(i=>i.id===iid); if(item)Object.assign(item,ch,{by:uRef.current?.name}); secName=sec?.name; pendingRef.current=nd; return nd; });
    if(sTimer.current)clearTimeout(sTimer.current);
    sTimer.current=setTimeout(()=>{ const nd=pendingRef.current; if(nd){ writeProject(nd); dirtyRef.current=false; addLogEntry(`Edited item in ${secName||"BOQ"}`);} },700);
  },[addLogEntry]);

  const delItem=useCallback((sid,iid)=>{ if(!data)return; const nd=JSON.parse(JSON.stringify(data)); const sec=nd.sections.find(s=>s.id===sid); const item=sec?.items.find(i=>i.id===iid); if(sec&&item){ sec.items=sec.items.filter(i=>i.id!==iid); pushData(nd,`Deleted item in ${sec.name}`);} },[data,pushData]);

  const blurSave=useCallback((sid,iid,code)=>{ if(code){const m=findBurByCode(burItems,code); if(m&&bTot(m)>0)setRateSugg({secId:sid,iid,rate:bTot(m),code:m.code,desc:m.desc});} },[burItems]);

  // Custom BOQ columns (Excel-style)
  const addCol=useCallback(()=>{
    if(!data)return;
    const label=prompt("New column name (e.g. 'Wastage Amt'):"); if(!label||!label.trim())return;
    const formula=prompt("Optional formula — leave BLANK for a column you type into.\n\nYou can use: Qty, Rate A, Rate B, Amt A, Amt B, and other column names (spaces ok).\nExamples:  Rate A * 0.25      Qty * markup","");
    const nd=JSON.parse(JSON.stringify(data)); nd.cols=[...(nd.cols||[]),{id:uid(),label:label.trim(),formula:(formula||"").trim()}];
    pushData(nd,`Added column "${label.trim()}"`);
  },[data,pushData]);
  const delCol=useCallback(cid=>{ if(!data||!confirm("Delete this column?"))return; const nd=JSON.parse(JSON.stringify(data)); nd.cols=(nd.cols||[]).filter(c=>c.id!==cid); pushData(nd,"Deleted a column"); },[data,pushData]);
  // Set a column's formula (typed as "=expr" in any cell) — applies to ALL rows automatically.
  const setColFormula=useCallback((cid,raw)=>{ if(!data)return; const f=String(raw||"").trim().replace(/^=/,"").trim(); const nd=JSON.parse(JSON.stringify(data)); const col=(nd.cols||[]).find(c=>c.id===cid); if(!col||col.formula===f)return; col.formula=f; pushData(nd,f?`Set formula for "${col.label}"`:`Cleared formula for "${col.label}"`); },[data,pushData]);
  // Export the master BQ: open the bundled template and fill Material/Labour (or U-Rate)
  // by matching each row's column-A CODE to the BUR library — preserving all colours/formats/formulas.
  const exportMasterBQ=useCallback(async()=>{
    try{
      setBqResult(null); toast_("⏳ Building master BQ (a few seconds)…");
      const ExcelJS=(await import("exceljs")).default;
      let buf, tplName="CAG_Dormitory_Master_BQ.xlsx";
      if(bqTplB64){ buf=Uint8Array.from(atob(bqTplB64),c=>c.charCodeAt(0)); tplName=bqTplName||tplName; }
      else { const res=await fetch(import.meta.env.BASE_URL+"boq-template.xlsx"); if(!res.ok)throw new Error("template file not found"); buf=await res.arrayBuffer(); }
      const wb=new ExcelJS.Workbook(); await wb.xlsx.load(buf);
      const codeMap={}; for(const b of burItems){ const k=(b.code||"").trim().toLowerCase(); if(k&&!codeMap[k])codeMap[k]=b; }
      // also match template codes written without (or with) the _unit suffix
      for(const b of burItems){ if(!b.code)continue; for(const k of [codeBase(b.code,b.unit),codeWithUnit(b.code,b.unit),cleanCode(b.code),cleanCode(codeBase(b.code,b.unit))].map(s=>s.toLowerCase())) if(k&&!codeMap[k])codeMap[k]=b; }
      const toStr=v=>{ if(v==null)return ""; if(typeof v==="object"){ if(v.richText)return v.richText.map(t=>t.text).join(""); if(v.text!=null)return String(v.text); if(v.result!=null)return String(v.result); return ""; } return String(v); };
      let filled=0; const unmatched=new Set(); let scanned=0;
      wb.eachSheet(ws=>{
        let hr=0,cCode=0,cUnit=0,cMat=0,cLab=0,cRate=0;
        for(let r=1;r<=Math.min(14,ws.rowCount);r++){
          let found=false;
          ws.getRow(r).eachCell({includeEmpty:false},(cell,col)=>{ const t=toStr(cell.value).trim().toUpperCase(); if(t==="CODE"){cCode=col;found=true;} if(t==="UNIT")cUnit=col; if(t.includes("MATERIAL"))cMat=col; if(t.includes("LABOUR"))cLab=col; if(t.includes("U/RATE"))cRate=col; });
          if(found){hr=r;break;}
        }
        if(!hr||!cCode||!cUnit)return;
        for(let r=hr+1;r<=ws.rowCount;r++){
          const row=ws.getRow(r);
          if(!toStr(row.getCell(cUnit).value).trim())continue;
          const code=toStr(row.getCell(cCode).value).trim(); if(!code)continue;
          scanned++;
          const bur=codeMap[code.toLowerCase()]||codeMap[cleanCode(code).toLowerCase()]; if(!bur){unmatched.add(code);continue;}
          if(cMat&&cLab){ const m=+bur.material||0,l=+bur.labour||0; const mc=row.getCell(cMat); if(!mc.formula)mc.value=m>0?m:null; const lc=row.getCell(cLab); if(!lc.formula)lc.value=l>0?l:null; filled++; }
          else if(cRate){ const rc=row.getCell(cRate); if(!rc.formula){const t=+bTot(bur)||0; rc.value=t>0?t:null; filled++;} }
        }
      });
      const out=await wb.xlsx.writeBuffer();
      const blob=new Blob([out],{type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});
      const url=URL.createObjectURL(blob); const a=document.createElement("a"); a.href=url; a.download=tplName.replace(/\.xlsx$/i,"")+"_priced.xlsx"; a.click(); URL.revokeObjectURL(url);
      setBqResult({filled,scanned,unmatched:[...unmatched].sort(),tpl:tplName});
      toast_(`✅ Exported — ${filled} of ${scanned} rate rows filled`);
    }catch(e){ toast_("⚠️ Export failed: "+(e&&e.message||e)); }
  },[burItems,bqTplB64,bqTplName,toast_]);
  const onTemplateFile=useCallback(async file=>{
    if(!file)return; if(!/\.xlsx$/i.test(file.name)){toast_("⚠️ Use a .xlsx file");return;}
    try{ const bytes=new Uint8Array(await file.arrayBuffer()); let bin=""; const CH=0x8000; for(let i=0;i<bytes.length;i+=CH)bin+=String.fromCharCode.apply(null,bytes.subarray(i,i+CH)); const b=btoa(bin); if(b.length>980000){toast_("⚠️ Template too large (keep under ~700 KB)");return;} await setDoc(doc(db,"boqtemplate",pid),{b64:b,name:file.name}); toast_("✅ Template updated to "+file.name); }
    catch(e){ toast_("⚠️ "+e.message); }
  },[db,pid,toast_]);

  // Export the BOQ (with computed custom columns) to a CSV that opens in Excel.
  const exportBOQ=useCallback(()=>{ if(!data)return; const cols=data.cols||[]; const esc=v=>{const s=String(v??"");return /[",\n]/.test(s)?'"'+s.replace(/"/g,'""')+'"':s;}; const head=["Section","Ref","Description","Unit","Qty","Rate A","Amt A","Rate B","Amt B","Code","Remarks","By",...cols.map(c=>c.label)]; const rows=[]; data.sections.forEach(sec=>(sec.items||[]).forEach(it=>{const cxv=computeCols(it,cols);rows.push([sec.name,it.ref,it.desc,it.unit,it.qty,it.rA,(it.qty||0)*(it.rA||0),it.rB,(it.qty||0)*(it.rB||0),it.code,it.remarks,it.by,...cols.map(c=>cxv[c.id])]);})); const csv="﻿"+[head,...rows].map(r=>r.map(esc).join(",")).join("\r\n"); const blob=new Blob([csv],{type:"text/csv;charset=utf-8;"}); const url=URL.createObjectURL(blob); const a=document.createElement("a"); a.href=url; a.download="BOQ_export.csv"; a.click(); URL.revokeObjectURL(url); toast_(`✅ Exported ${rows.length} BOQ rows to Excel (CSV)`); },[data,toast_]);
  // Read the dropped Excel-grid file and import its rows into the structured BOQ sections (by sheet name).
  const importGridToStructured=useCallback(async()=>{
    try{
      if(!data){toast_("Open a project first");return;}
      const snap=await getDoc(doc(db,"boqfile",pid));
      if(!snap.exists()||!snap.data().b64){toast_("⚠️ No Excel-grid file yet — switch to 'Excel grid' mode and drag your file in first");return;}
      if(!confirm("Import the Excel-grid rows into the Structured BOQ?\n\nThis REPLACES the current structured BOQ items with the items found in your dropped file (matched to sections by sheet name)."))return;
      toast_("⏳ Importing grid → structured…");
      const ExcelJS=(await import("exceljs")).default; const wb=new ExcelJS.Workbook();
      await wb.xlsx.load(Uint8Array.from(atob(snap.data().b64),c=>c.charCodeAt(0)));
      const toStr=v=>{ if(v==null)return ""; if(typeof v==="object"){ if(v.richText)return v.richText.map(t=>t.text).join(""); if(v.text!=null)return String(v.text); if(v.result!=null)return String(v.result); return ""; } return String(v); };
      const num=s=>parseFloat(String(s).replace(/[^0-9.\-]/g,""))||0;
      const sectionFor=name=>{const n=(name||"").toLowerCase(); if(n.includes("prelim"))return"prelim"; if(n.includes("professional")||n.includes("fee"))return"fees"; if(n.includes("building"))return"building"; if(n.includes("external"))return"external"; if(n.includes("m&e")||n.includes("m & e")||n.includes("mech")||n.includes("elec")||n.includes("m and e"))return"mande"; return null;};
      const nd=JSON.parse(JSON.stringify(data)); nd.sections.forEach(s=>{s.items=[];});
      let added=0;
      wb.eachSheet(ws=>{
        const sec=sectionFor(ws.name); if(!sec)return; const target=nd.sections.find(s=>s.id===sec); if(!target)return;
        let hr=0,cCode=0,cDesc=0,cUnit=0,cRate=0,cQty=0;
        for(let r=1;r<=Math.min(14,ws.rowCount);r++){ let f=false; ws.getRow(r).eachCell({includeEmpty:false},(cell,col)=>{const t=toStr(cell.value).trim().toUpperCase(); if(t==="CODE"){cCode=col;f=true;} if(t==="DESCRIPTION")cDesc=col; if(t==="UNIT")cUnit=col; if(t==="U/RATE"||t==="RATE")cRate=col; if(t==="QTY")cQty=col;}); if(f&&cDesc){hr=r;break;} }
        if(!hr||!cDesc||!cUnit)return;
        for(let r=hr+1;r<=ws.rowCount;r++){ const row=ws.getRow(r); const unit=toStr(row.getCell(cUnit).value).trim(); if(!unit)continue; const desc=toStr(row.getCell(cDesc).value).trim(); if(!desc)continue;
          target.items.push({id:uid(),ref:"",desc,unit,qty:cQty?(num(toStr(row.getCell(cQty).value))||1):1,rA:cRate?num(toStr(row.getCell(cRate).value)):0,rB:0,code:cCode?toStr(row.getCell(cCode).value).trim():"",status:"Draft",remarks:"",by:uRef.current?.name}); added++;
        }
      });
      pushData(nd,`Imported ${added} BOQ items from Excel grid`); setBoqMode("structured");
      toast_(`✅ Imported ${added} items into structured BOQ`);
    }catch(e){ toast_("⚠️ Import failed: "+(e&&e.message||e)); }
  },[data,pid,pushData,toast_]);

  // Build the BUR export from LIVE data in the same layout as the master BUR.xlsx:
  // B Description | C Unit | D Code | E Rate (purple) | F Amt | H Quote | K Labour | L Material | M Plant | N OH% | O Profit% | R+ one column per supplier.
  // Each category gets an orange heading row that also carries that category's supplier names (with quote date) above its rate columns.
  const exportBUR=useCallback(async()=>{
    try{
      toast_("⏳ Building BUR Excel…");
      const ExcelJS=(await import("exceljs")).default;
      const wb=new ExcelJS.Workbook(); wb.calcProperties.fullCalcOnLoad=true; const ws=wb.addWorksheet("BUR");
      const F={name:"Arial",size:10}; const NF='_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)';
      const thin={style:"thin"}; const bd={top:thin,left:thin,bottom:thin,right:thin};
      const fill=argb=>({type:"pattern",pattern:"solid",fgColor:{argb}});
      const C_HEAD="FFFFFF00",C_CAT="FFFFC000",C_SUPHDR="FFFFFFCC",C_RATE="FFCC99FF",C_QUOTE="FFCCFFFF",C_SEL="FFA6C9EC",C_OTHER="FFFFCCFF",C_CODE="FFCCFF99";
      const SUP0=18; // column R
      const L=n=>{let s="";while(n>0){const m=(n-1)%26;s=String.fromCharCode(65+m)+s;n=Math.floor((n-1)/26);}return s;};
      const put=(r,c,v,o={})=>{ const cell=ws.getCell(r,c); cell.value=v; cell.font={...F,...(o.font||{})}; cell.alignment={vertical:"top",...(o.al||{})}; if(o.fill)cell.fill=fill(o.fill); if(o.nf)cell.numFmt=o.nf; if(o.bd)cell.border=bd; return cell; };
      const codeCount={}; burItems.forEach(b=>{const k=(b.code||"").trim().toLowerCase(); if(k)codeCount[k]=(codeCount[k]||0)+1;});
      // group by category (same order as the sidebar)
      const groups=[]; const known=new Set();
      const orderedCats=catSort==="none"?cats:[...cats].sort((a,b)=>catSort==="az"?String(a.name).localeCompare(String(b.name)):String(b.name).localeCompare(String(a.name)));
      for(const cat of orderedCats){ known.add(cat.id); const items=orderBySubcat(burItems.filter(b=>b.catId===cat.id).sort((a,b)=>(a.code||"").localeCompare(b.code||""))); if(items.length)groups.push({name:cat.name,items}); }
      const orphan=orderBySubcat(burItems.filter(b=>!known.has(b.catId)).sort((a,b)=>(a.code||"").localeCompare(b.code||""))); if(orphan.length)groups.push({name:"(Uncategorised)",items:orphan});
      // per category: one column per supplier + quote date (case-insensitive), e.g. "Kiong Gay (Jan 24)" and "Kiong Gay (SLA Sep 26)" are separate columns
      const supKey=e=>(String(e.supplier||"").trim()+"|"+String(e.date||"").trim()).toLowerCase();
      let maxSup=0;
      groups.forEach(g=>{ const sups=[]; const idx={};
        g.items.forEach(it=>(it.costData||[]).forEach(e=>{ const nm=String(e.supplier||"").trim(); if(!nm)return; const k=supKey(e); if(!(k in idx)){idx[k]=sups.length;sups.push({name:nm,dates:new Set()});} if(e.date)sups[idx[k]].dates.add(e.date); }));
        g.sups=sups; g.supIdx=idx; maxSup=Math.max(maxSup,sups.length);
        const xn=[]; g.items.forEach(it=>(it.extras||[]).forEach(x=>{ const n=String(x.name||"").trim()||"Other"; if(!xn.some(y=>y.toLowerCase()===n.toLowerCase()))xn.push(n); }));
        g.xcols=xn.length>4?[...xn.slice(0,3),"Others"]:xn; g.xcol=n=>{ const i=g.xcols.findIndex(y=>y.toLowerCase()===n.toLowerCase()); return 14+(i>=0?i:3); }; });
      const notesCol=SUP0+Math.max(maxSup,4), photoCol=notesCol+1; // visible Notes column right after the supplier columns
      // column widths copied from the master sheet
      const W={1:15,2:52,3:6,4:45,5:17.6,6:16.9,7:11,8:15.7,9:8.4,10:15.9,11:12,12:12,13:12,14:13,15:13,16:13,17:13};
      for(let c=1;c<photoCol+12;c++) ws.getColumn(c).width=c>=photoCol?31:(W[c]||18.2);
      put(1,1,"Exported "+new Date().toLocaleDateString("en-SG",{day:"numeric",month:"short",year:"numeric"}),{font:{bold:true}});
      put(3,1,"BUILD UP RATES",{font:{underline:true}});
      const HEAD=[[1,"Description","left"],[3,"Unit","center"],[4,"Code","center"],[5,"Rate","center"],[6,"Amt","center"],[7,"Discount","center"],[8,"Quote","center"],[9,"%","center"],[10,"Waste/ OH","center"],[11,"Labour","center"],[12,"Material","center"],[13,"Plant","center"],[14,"Extra items (S$)","left"],[SUP0,"Subcon / Supplier","center"]];
      HEAD.forEach(([c,t,h])=>put(5,c,t,{font:{bold:true},fill:C_HEAD,bd:1,al:{horizontal:h,wrapText:true}}));
      put(5,notesCol,"Notes",{font:{bold:true},fill:C_HEAD,bd:1,al:{horizontal:"left"}}); ws.getColumn(notesCol).width=45;
      put(5,photoCol,"Photos",{font:{bold:true},fill:C_HEAD,bd:1,al:{horizontal:"center"}});
      let r=7, count=0, photoCount=0; const catRows=new Set(); const codeRow={}, pendingF=[];
      for(const g of groups){
        // category heading row: name in A (orange), component labels, then supplier names over their rate columns
        put(r,1,g.name,{font:{bold:true,underline:true},fill:C_CAT,bd:1}); catRows.add(r);
        [[7,"Discount"],[8,"Quote"],[9,"%"],[10,"Waste/ OH"],[11,"Labour"],[12,"Material"],[13,"Plant"],...g.xcols.map((n,i)=>[14+i,n])].forEach(([c,t])=>put(r,c,t,{font:{bold:true},al:{horizontal:"center",wrapText:true}}));
        g.sups.forEach((s,i)=>{ const d=[...s.dates]; put(r,SUP0+i,s.name+(d.length===1?`\n(${d[0]})`:""),{fill:C_SUPHDR,bd:1,al:{horizontal:"center",wrapText:true}}); });
        ws.getRow(r).height=g.sups.some(s=>s.dates.size===1)?25.5:12.75; r++;
        let curSub="";
        for(const it of g.items){
          const sg=String(it.group||"").trim(); if(sg&&sg!==curSub){ ws.getRow(r).outlineLevel=1; r++; put(r,2,sg,{font:{bold:true,underline:true}}); ws.getRow(r).outlineLevel=1; r++; } curSub=sg;
          ws.getRow(r).outlineLevel=sg?2:1;
          const row=ws.getRow(r);
          const lab=+it.labour||0,mat=+it.material||0,pl=+it.plant||0,sub=+it.subcon||0,wp=+it.waste||0;
          const latest={}; (it.costData||[]).forEach(e=>{ const k=supKey(e); if(!String(e.supplier||"").trim()||!(k in g.supIdx))return; latest[k]=e; });
          // A supplier whose rate equals one of the item's cost figures is the "selected" quote: that cost cell links to it (=+R9) and it turns blue; others stay pink.
          // Prefer a quote filed under the same component, else any supplier with that exact rate.
          const sel={}; const used=new Set();
          [["subcon",sub],["labour",lab],["material",mat],["plant",pl]].forEach(([comp,v])=>{ if(!(v>0))return; const ks=Object.keys(latest).filter(k=>!used.has(k)&&Math.abs((+latest[k].rate||0)-v)<0.005); const k=ks.find(k=>latest[k].component===comp)??ks[0]; if(k!==undefined){sel[comp]=k;used.add(k);} });
          const costCell=(c,comp,v)=>put(r,c,sel[comp]!==undefined?{formula:`+${L(SUP0+g.supIdx[sel[comp]])}${r}`,result:v}:(v||null),{fill:c===8?C_QUOTE:undefined,nf:NF,al:{horizontal:"center"}});
          put(r,2,it.desc||"",{al:{wrapText:true}});
          put(r,3,it.unit||"",{al:{horizontal:"center"}});
          // Code gets "_unit" appended (e.g. Acoustic_feltCeiling_m2) unless it already ends with it
          const codeOut=codeWithUnit(it.code,it.unit);
          const cc=put(r,4,codeOut,{al:{horizontal:"left",wrapText:true}});
          if((codeCount[(it.code||"").trim().toLowerCase()]||0)>1){ cc.fill=fill("FFFECACA"); cc.font={...F,bold:true,color:{argb:"FFB91C1C"}}; }
          const amt=bTot(it);
          put(r,5,{formula:hasRoundUp(it)?`+ROUNDUP(F${r},${+it.roundUp})`:`+F${r}`,result:amt},{font:{bold:true},fill:C_RATE,nf:NF,al:{horizontal:"center"}});
          put(r,6,{formula:`SUM(G${r}:Q${r})-I${r}`,result:bTot({...it,roundUp:null})},{nf:NF});
          costCell(8,"subcon",sub); costCell(11,"labour",lab); costCell(12,"material",mat); costCell(13,"plant",pl);
          // waste: I = % (as a fraction, like the master sheet), J = (Quote+Labour+Material+Plant) x I
          const dp=+it.discount||0; if(dp) put(r,7,{formula:`-(H${r}+K${r}+L${r}+M${r})*${dp/100}`,result:-discAmt(it)},{nf:`"-${+dp}% = "#,##0.00;"-${+dp}% = -"#,##0.00`,al:{horizontal:"center"},font:{color:{argb:"FFDC2626"}}});
          if(wp){ put(r,9,wp/100,{nf:Number.isInteger(wp)?"0%":"0.0#%",al:{horizontal:"center"}}); put(r,10,{formula:`(H${r}+K${r}+L${r}+M${r}+G${r})*I${r}`,result:wasteAmt(it)},{nf:NF,al:{horizontal:"center"}}); }
          // extra items: lump-sum S$ under their named column (N..Q)
          const xsum={}, xnote={}; (it.extras||[]).forEach(x=>{ const c=g.xcol(String(x.name||"").trim()||"Other"); xsum[c]=(xsum[c]||0)+extraAmt(it,x); (xnote[c]=xnote[c]||[]).push(`${x.name||"Other"}: S$ ${extraAmt(it,x).toFixed(2)}`); });
          Object.entries(xsum).forEach(([c,v])=>{ const own=(it.extras||[]).filter(x=>g.xcol(String(x.name||"").trim()||"Other")===+c), xRef=own.length===1&&exprRefs(own[0].expr).length>0, ex=!xRef&&own.length===1&&own[0].type!=="%"&&/[-+*/()]/.test(String(own[0].expr||"").replace(/^\s*[=-]/,""))?String(own[0].expr).replace(/^\s*=/,"").replace(/\s+/g,""):null;
            // a typed calculation goes into Excel as a real formula, e.g. =((1+1)*2)+((3*1)+2)
            const xc=put(r,+c,ex&&isFinite(calcExpr(ex))?{formula:ex,result:+v.toFixed(4)}:(+v.toFixed(4)||null),{nf:NF,al:{horizontal:"center"}}); if(xRef)pendingF.push({cell:xc,expr:own[0].expr}); if(xnote[c].length>1||g.xcols[+c-14]==="Others")xc.note=xnote[c].join("\n"); });
          Object.entries(latest).forEach(([k,e])=>{ const qRef=exprRefs(e.expr).length>0, qx=!qRef&&e.expr&&/[-+*/()]/.test(String(e.expr).replace(/^\s*[=-]/,""))&&isFinite(calcExpr(e.expr))?String(e.expr).replace(/^\s*=/,"").replace(/\s+/g,""):null; const cell=put(r,SUP0+g.supIdx[k],qx?{formula:qx,result:+e.rate||0}:(+e.rate||0),{fill:used.has(k)?C_SEL:C_OTHER,nf:NF,al:{horizontal:"center",wrapText:true}}); if(qRef)pendingF.push({cell,expr:e.expr});
            const note=[e.component&&e.component!=="subcon"&&(CLABEL[e.component]||(ccomps.find(x=>x.key===e.component)||{}).label||e.component),e.location&&`@ ${e.location}`,e.date&&`Date: ${e.date}`,e.note].filter(Boolean).join("\n"); if(note)cell.note=note; });
          // notes typed on cost-data entries, visible in the Notes column (also kept as hover comments on the supplier cells)
          { const nts=(it.costData||[]).filter(e=>String(e.note||"").trim()).map(e=>`${e.supplier?e.supplier+": ":""}${String(e.note).trim()}`); if(nts.length)put(r,notesCol,nts.join("\n"),{al:{wrapText:true}}); }
          // embed this item's photos after the supplier columns, on the same row
          const itemPhotos=(it.costData||[]).flatMap(e=>e.imgs&&e.imgs.length?e.imgs:(e.img?[e.img]:[])).slice(0,12);
          // full stored photo (no re-compression), shown up to 210x160 px keeping its shape; enlarging it in Excel stays sharp
          if(itemPhotos.length){ row.height=124; for(let i=0;i<itemPhotos.length;i++){ try{ const m=/^data:image\/(\w+);base64,(.+)$/.exec(itemPhotos[i]); if(m){ const {w,h}=await imgSize(itemPhotos[i]); const k=Math.min(210/w,160/h,1); const id=wb.addImage({base64:m[2],extension:/png/i.test(m[1])?"png":"jpeg"}); ws.addImage(id,{tl:{col:photoCol-1+i,row:r-1},ext:{width:Math.round(w*k),height:Math.round(h*k)}}); photoCount++; } }catch{} } }
          if(it.code){ codeRow[String(it.code).trim().toLowerCase()]=r; codeRow[codeBase(it.code,it.unit).toLowerCase()]=r; }
          r++; count++;
        }
        r++; // blank row between categories
      }
      // Colour whole rows/columns like the master sheet: header row yellow, category rows orange, and every column keeps its
      // colour all the way down (grey text area, purple Rate, cyan Quote, pink %, pale blue build-up, pink supplier quotes).
      // turn code-linked formulas into Excel links to the linked item's Rate cell, e.g. =rebar+0.13 -> =$E$1903+0.13
      pendingF.forEach(({cell,expr})=>{ let ok=true; const f=String(expr).replace(/^\s*=/,"").replace(/\s+/g,"").replace(/[A-Za-z_][A-Za-z0-9_.&]*/g,m=>{ const row=codeRow[m.toLowerCase()]; if(!row){ok=false;return m;} return "$E$"+row; }); if(ok)cell.value={formula:f,result:typeof cell.value==="number"?cell.value:undefined}; });
      const lastCol=SUP0+Math.max(maxSup,4)-1;
      const colFill=c=>c===5?C_RATE:c===8?C_QUOTE:c===9?"FFE6B8B7":(c>=10&&c<=17)?"FFDDEBF7":(c>=SUP0&&c<=lastCol)?C_OTHER:c===7?"FFFDE2E2":"FFF2F2F2";
      for(let rr=5;rr<r;rr++) for(let c=1;c<=lastCol;c++){ const cell=ws.getCell(rr,c); if(cell.fill&&cell.fill.type==="pattern")continue;
        const f=rr===5?C_HEAD:catRows.has(rr)?C_CAT:colFill(c); if(f)cell.fill=fill(f); if(rr===5||catRows.has(rr))cell.border={top:thin,bottom:thin}; }
      ws.properties.outlineLevelRow=2; ws.properties.outlineProperties={summaryBelow:false,summaryRight:false};
      ws.views=[{state:"frozen",xSplit:9,ySplit:5,topLeftCell:"J6"}];
      const out=await wb.xlsx.writeBuffer(); const blob=new Blob([out],{type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});
      const url=URL.createObjectURL(blob); const a=document.createElement("a"); a.href=url; a.download="BUR.xlsx"; a.click(); URL.revokeObjectURL(url);
      toast_(`✅ Exported ${count} items${photoCount?` + ${photoCount} photos`:""}`);
    }catch(e){ toast_("⚠️ Export failed: "+(e&&e.message||e)); }
  },[burItems,cats,toast_,catSort,ccomps]);

  // ── BUR writes (per-document in shared library) ─────────────────────────────
  // ── BUR undo (this browser session only) ─────────────────────────────────────
  // Each entry remembers how to put things back: {label, kind:"fields", items:[{id,before}]} | {kind:"deleted", docs:[{id,data}]} | {kind:"created", ids:[]}.
  // Rapid typing in the same field of the same item collapses into one entry, so one Undo reverts the whole edit.
  const lastSelRef=useRef(null); // last BUR checkbox clicked (for Shift+click range select)
  const undoRef=useRef([]); const [undoLabel,setUndoLabel]=useState("");
  const syncUndo=()=>setUndoLabel(undoRef.current.length?undoRef.current[undoRef.current.length-1].label:"");
  const pushUndo=useCallback(entry=>{ const st=undoRef.current; st.push({...entry,ts:Date.now()}); if(st.length>50)st.shift(); syncUndo(); },[]);
  const burLabel=id=>{ const b=burItemsRef.current.find(x=>x.id===id); return b?(b.code||b.desc||"item"):"item"; };
  const snapFields=(id,keys)=>{ const b=burItemsRef.current.find(x=>x.id===id)||{}; const o={}; keys.forEach(k=>{ o[k]=b[k]===undefined?null:JSON.parse(JSON.stringify(b[k])); }); return o; };
  const recordFields=useCallback((ids,keys,label,coalesce=false)=>{
    const st=undoRef.current, top=st[st.length-1];
    if(coalesce&&top&&top.kind==="fields"&&top.items.length===1&&ids.length===1&&top.items[0].id===ids[0]&&top.keys===keys.join(",")&&Date.now()-top.ts<4000){ top.ts=Date.now(); return; }
    pushUndo({kind:"fields",keys:keys.join(","),label,items:ids.map(id=>({id,before:snapFields(id,keys)}))});
  },[pushUndo]);
  const undoLast=useCallback(async()=>{
    const e=undoRef.current.pop(); syncUndo(); if(!e)return;
    try{
      if(e.kind==="fields"){
        e.items.forEach(({id})=>{ if(burTimers.current[id]){clearTimeout(burTimers.current[id]);delete burTimers.current[id];} if(burTimers.current["cd_"+id]){clearTimeout(burTimers.current["cd_"+id]);delete burTimers.current["cd_"+id];} });
        setBurItems(prev=>prev.map(b=>{ const it=e.items.find(x=>x.id===b.id); return it?{...b,...it.before}:b; }));
        for(let i=0;i<e.items.length;i+=400){ const batch=writeBatch(db); e.items.slice(i,i+400).forEach(({id,before})=>batch.update(doc(db,"bur",id),before)); await batch.commit(); }
      }else if(e.kind==="deleted"){
        for(let i=0;i<e.docs.length;i+=400){ const batch=writeBatch(db); e.docs.slice(i,i+400).forEach(({id,data})=>batch.set(doc(db,"bur",id),data)); await batch.commit(); }
      }else if(e.kind==="multi"){
        for(const part of e.parts){ if(part.kind==="fields"){ setBurItems(prev=>prev.map(b=>{ const it=part.items.find(x=>x.id===b.id); return it?{...b,...it.before}:b; })); for(let i=0;i<part.items.length;i+=400){ const batch=writeBatch(db); part.items.slice(i,i+400).forEach(({id,before})=>batch.update(doc(db,"bur",id),before)); await batch.commit(); } }
          else if(part.kind==="created"){ for(let i=0;i<part.ids.length;i+=400){ const batch=writeBatch(db); part.ids.slice(i,i+400).forEach(id=>batch.delete(doc(db,"bur",id))); await batch.commit(); } } }
      }else if(e.kind==="created"){
        for(let i=0;i<e.ids.length;i+=400){ const batch=writeBatch(db); e.ids.slice(i,i+400).forEach(id=>batch.delete(doc(db,"bur",id))); await batch.commit(); }
      }
      toast_("↶ Undone: "+e.label);
    }catch(err){ toast_("⚠️ Undo failed: "+(err&&err.message||err)); }
  },[toast_]);
  // Ctrl+Z on the BUR tab (when not typing in a box — inside a box Ctrl+Z undoes the typing as usual)
  useEffect(()=>{ const h=ev=>{ if(tab!=="bur"||!(ev.ctrlKey||ev.metaKey)||ev.key.toLowerCase()!=="z"||ev.shiftKey)return; const t=ev.target, tg=(t&&t.tagName)||""; if(/INPUT|TEXTAREA|SELECT/.test(tg)||(t&&t.isContentEditable))return; ev.preventDefault(); undoLast(); }; window.addEventListener("keydown",h); return ()=>window.removeEventListener("keydown",h); },[tab,undoLast]);
  const addBurItem=useCallback(async catId=>{ const ref=doc(collection(db,"bur")); try{ await setDoc(ref,{pid,catId,code:"",desc:"New Item",unit:"sum",labour:0,material:0,plant:0,subcon:0,oh:0,profit:0,waste:0,extras:[],costData:[],quote:null,group:""}); pushUndo({kind:"created",label:"Add new item",ids:[ref.id]}); setExpBur(ref.id);}catch(e){toast_("⚠️ "+e.message);} },[pid,pushUndo]);
  const updBur=useCallback((id,ch)=>{ recordFields([id],Object.keys(ch),"Edit "+burLabel(id),true); setBurItems(prev=>prev.map(b=>b.id===id?{...b,...ch}:b)); if(burTimers.current[id])clearTimeout(burTimers.current[id]); burTimers.current[id]=setTimeout(()=>{ delete burTimers.current[id]; updateDoc(doc(db,"bur",id),ch).catch(()=>{}); },600); },[recordFields]);
  const delBur=useCallback(id=>{ if(!confirm("Delete this BUR item?"))return; const b=burItemsRef.current.find(x=>x.id===id); if(b){ const {id:_i,...data}=b; pushUndo({kind:"deleted",label:"Delete "+(b.code||b.desc||"item"),docs:[{id,data}]}); } deleteDoc(doc(db,"bur",id)).catch(()=>{}); },[pushUndo]);
  const setBurField=(id,patch)=>{ recordFields([id],Object.keys(patch),"Edit "+burLabel(id)); setBurItems(prev=>prev.map(b=>b.id===id?{...b,...patch}:b)); updateDoc(doc(db,"bur",id),patch).catch(()=>{}); };

  const loadMaster=useCallback(async()=>{
    if(!confirm(`Replace the entire BUR library with the master list?\n\nThis clears the current ${burItems.length} item(s) and loads ${burSeed.items.length} fresh items filed under their ${burSeed.cats.length} categories (from the Excel). Use this to reset / remove duplicates.`))return;
    toast_("⏳ Loading master list (replacing all)…");
    try{
      // 1) clear existing BUR docs (fixes duplicates)
      for(let i=0;i<burItems.length;i+=400){ const batch=writeBatch(db); burItems.slice(i,i+400).forEach(b=>batch.delete(doc(db,"bur",b.id))); await batch.commit(); }
      // 2) set categories from the Excel (column A)
      await setDoc(doc(db,"meta","cats"),{list:burSeed.cats});
      // 3) write the items (scoped to this project)
      for(let i=0;i<burSeed.items.length;i+=400){ const batch=writeBatch(db); burSeed.items.slice(i,i+400).forEach(it=>{ const {id,...rest}=it; batch.set(doc(collection(db,"bur")),{...rest,pid}); }); await batch.commit(); }
      if(burSeed.cats[0])setSelCat(burSeed.cats[0].id);
      addLogEntry(`Loaded master list: ${burSeed.items.length} items in ${burSeed.cats.length} categories`);
      toast_(`✅ Loaded ${burSeed.items.length} items into ${burSeed.cats.length} categories`);
    }catch(e){ toast_("⚠️ Load failed: "+e.message); }
  },[burItems,toast_,addLogEntry,pid]);

  // Auto-sort BUR items into the standard categories by keyword; unmatched → Others. User can move any item after.
  const categorizeMaster=useCallback(async()=>{
    if(!burItems.length){toast_("No BUR items to sort");return;}
    if(!confirm(`Auto-sort ${burItems.length} BUR items into categories?\nItems that don't clearly match go to "Others" — you can move any item afterwards.`))return;
    toast_("⏳ Sorting into categories…");
    try{
      const updates=[];
      for(const b of burItems){
        const text=((b.group||"")+" "+(b.desc||"")+" "+(b.code||"")).toLowerCase();
        let cat="cat13";
        for(const r of CATEGORY_RULES){ if(r.kw.some(k=>text.includes(k))){cat=r.cat;break;} }
        if(b.catId!==cat)updates.push([b.id,cat]);
      }
      if(updates.length)recordFields(updates.map(u=>u[0]),["catId"],`Auto-sort ${updates.length} items`);
      for(let i=0;i<updates.length;i+=400){ const batch=writeBatch(db); updates.slice(i,i+400).forEach(([id,cat])=>batch.update(doc(db,"bur",id),{catId:cat})); await batch.commit(); }
      toast_(`✅ Sorted ${updates.length} items into categories`); addLogEntry(`Auto-sorted ${updates.length} BUR items into categories`);
    }catch(e){toast_("⚠️ "+e.message);}
  },[burItems,toast_,addLogEntry]);

  const importPaste=useCallback(async()=>{
    const lines=pasteText.split(/\r?\n/).filter(l=>l.trim()!=="");
    if(!lines.length){toast_("⚠️ Nothing to paste");return;}
    const first=lines[0].split("\t").map(s=>s.trim().toLowerCase());
    const start=(first[0]==="description"||first.includes("code"))?1:0;
    const rows=[];
    for(let i=start;i<lines.length;i++){ const c=lines[i].split("\t"); const desc=(c[0]||"").trim(),unit=(c[1]||"").trim(),code=(c[2]||"").trim(); const rate=parseFloat(String(c[3]||"").replace(/[^0-9.\-]/g,""))||0; if(!desc&&!code)continue; rows.push({pid,catId:pasteCat,code:codeWithUnit(code,unit||"sum"),desc:desc||"(no description)",unit:unit||"sum",labour:0,material:rate,plant:0,subcon:0,oh:0,profit:0,costData:[],quote:null,group:"Pasted"}); }
    if(!rows.length){toast_("⚠️ No rows parsed — expected: Description, Unit, Code, Rate");return;}
    try{ let batch=writeBatch(db),n=0; const newIds=[]; for(const r of rows){ const ref=doc(collection(db,"bur")); newIds.push(ref.id); batch.set(ref,r); n++; if(n%400===0){await batch.commit();batch=writeBatch(db);} } await batch.commit(); pushUndo({kind:"created",label:`Paste ${rows.length} items`,ids:newIds}); setPasteOpen(false);setPasteText("");setSelCat(pasteCat); toast_(`✅ Imported ${rows.length} items`);}catch(e){toast_("⚠️ "+e.message);}
  },[pasteText,pasteCat,toast_,pid,pushUndo]);

  // ── Cost data / quotes ──────────────────────────────────────────────────────
  const modalItem=costModal?burItems.find(b=>b.id===costModal):null;
  const costEntries=(()=>{ let arr=modalItem?(modalItem.costData||[]).filter(e=>e.component===cType):[]; if(ceSort.by!=="none"){ arr=[...arr].sort((a,b)=>{ if(ceSort.by==="rate")return ((+a.rate||0)-(+b.rate||0))*ceSort.dir; return String(a[ceSort.by]||"").localeCompare(String(b[ceSort.by]||""),undefined,{numeric:true})*ceSort.dir; }); } return arr; })();
  const ceToggleSort=by=>setCeSort(s=>s.by===by?{by,dir:-s.dir}:{by,dir:1});
  const compLabel=c=>CLABEL[c]||(ccomps.find(x=>x.key===c)||{}).label||c;
  // Ticked cost-data entries (e.inRate) make up the item's rate: Labour/Material/Plant/Subcon tabs fill those columns (sum of ticked),
  // every other tab (Mould, Freight, Accessories…) becomes an extra item named after the tab. Manually typed extra items are kept.
  // Before the first tick the old Labour/Material/Plant/Subcon are saved in item.preTick and put back when the last tick is removed.
  const tabExists=c=>COMPS.includes(c)||ccomps.some(x=>x.key===c);
  const tickPatch=(b,cd)=>{ cd=cd.map(e=>e.inRate&&!tabExists(e.component)?{...e,inRate:false}:e); // entries on a removed tab never count
    const manual=(b.extras||[]).filter(x=>!x.fromCD), was=(b.costData||[]).some(e=>e.inRate), now=cd.some(e=>e.inRate);
    if(!now){ if(!was)return {}; const pt=b.preTick||{}; return {labour:+pt.labour||0,material:+pt.material||0,plant:+pt.plant||0,subcon:+pt.subcon||0,preTick:null,extras:manual}; }
    const sum=c=>Math.round(cd.filter(e=>e.inRate&&e.component===c).reduce((t,e)=>t+cdRate(e),0)*10000)/10000;
    const customs=[...new Set(cd.filter(e=>e.inRate&&!COMPS.includes(e.component)).map(e=>e.component))];
    return {...(was?{}:{preTick:{labour:+b.labour||0,material:+b.material||0,plant:+b.plant||0,subcon:+b.subcon||0}}),
      labour:sum("labour"),material:sum("material"),plant:sum("plant"),subcon:sum("subcon"),
      extras:[...manual,...customs.map(c=>({id:"cd_"+c,name:compLabel(c),type:"$",value:sum(c),fromCD:true}))]}; };
  const deleteHiddenEntries=()=>{ const b=burItemsRef.current.find(x=>x.id===costModal); if(!b)return; const gone=(b.costData||[]).filter(e=>!tabExists(e.component)); if(!gone.length)return;
    if(!confirm(`Delete ${gone.length} hidden entr${gone.length>1?"ies":"y"} left over from removed tabs?\n(${gone.map(e=>`${e.supplier||"?"} S$ ${fmt(+e.rate||0)}`).join(", ")})\n\nYou can Undo.`))return;
    const cd=(b.costData||[]).filter(e=>tabExists(e.component)); setBurField(costModal,{costData:cd,...tickPatch(b,cd)}); toast_(`🗑 Deleted ${gone.length} hidden entr${gone.length>1?"ies":"y"}`); };
  const clearHiddenTicks=()=>{ const b=burItemsRef.current.find(x=>x.id===costModal); if(!b)return; const cd=(b.costData||[]).map(e=>e.inRate&&!tabExists(e.component)?{...e,inRate:false}:e); setBurField(costModal,{costData:cd,...tickPatch(b,cd)}); toast_("✅ Removed entries of deleted tabs from the rate"); };
  exprLookup=code=>{ const b=findBurByCode(burItemsRef.current,code); return b?bTot(b):NaN; };
  // When an item's rate changes, every quote / extra item whose formula names its code (e.g. =rebar+0.13) is recalculated and saved.
  // Items in a reference loop (A uses B, B uses A) are left alone.
  useEffect(()=>{ if(!burItems.length)return; const t=setTimeout(async()=>{
      const items=burItemsRef.current; const refsOf=b=>[...(b.costData||[]).map(e=>e.expr),...(b.extras||[]).map(x=>x.expr)].flatMap(exprRefs);
      const inLoop=b=>{ const seen=new Set(), stack=[...refsOf(b)]; while(stack.length){ const tgt=findBurByCode(items,stack.pop()); if(!tgt)continue; if(tgt.id===b.id)return true; if(seen.has(tgt.id))continue; seen.add(tgt.id); stack.push(...refsOf(tgt)); } return false; };
      const ups=[];
      items.forEach(b=>{ if(!refsOf(b).length||inLoop(b))return; let changed=false;
        const cd=(b.costData||[]).map(e=>{ if(!exprRefs(e.expr).length)return e; const n=calcExpr(e.expr); if(isFinite(n)&&Math.abs(n-(+e.rate||0))>0.00005){changed=true;return {...e,rate:Math.round(n*10000)/10000};} return e; });
        const xs=(b.extras||[]).map(x=>{ if(!exprRefs(x.expr).length)return x; const n=calcExpr(x.expr); if(isFinite(n)&&Math.abs(n-(+x.value||0))>0.00005){changed=true;return {...x,value:Math.round(n*10000)/10000};} return x; });
        if(changed)ups.push([b.id,{costData:cd,extras:xs,...(cd.some(e=>e.inRate)?tickPatch({...b,extras:xs},cd):{})}]); });
      if(!ups.length)return;
      try{ for(let i=0;i<ups.length;i+=400){ const batch=writeBatch(db); ups.slice(i,i+400).forEach(([id,pt])=>batch.update(doc(db,"bur",id),pt)); await batch.commit(); } }catch{}
    },1500); return ()=>clearTimeout(t); },[burItems]);
  const toggleInRate=eid=>{ const b=burItemsRef.current.find(x=>x.id===costModal); if(!b)return; const cd=(b.costData||[]).map(e=>e.id===eid?{...e,inRate:!e.inRate}:e); setBurField(costModal,{costData:cd,...tickPatch(b,cd)}); };
  const allComps=[...COMPS,...ccomps.map(c=>c.key)];

  const addCostEntry=useCallback(()=>{ if(!cForm.supplier||!String(cForm.rate).trim()){toast_("⚠️ Enter supplier and rate");return;} const rv=calcExpr(cForm.rate); if(!isFinite(rv)){toast_("⚠️ Rate formula has a mistake — check brackets / symbols");return;} const rExpr=(/[-+*/()]/.test(String(cForm.rate).replace(/^\s*[=-]/,""))||exprRefs(cForm.rate).length>0)?String(cForm.rate).trim():""; const b=burItems.find(x=>x.id===costModal); if(!b)return; const cd=[...(b.costData||[]),{id:uid(),component:cType,supplier:cForm.supplier,rate:Math.round(rv*10000)/10000,...(rExpr?{expr:rExpr}:{}),date:cForm.date,location:cForm.location||"",note:cForm.note,imgs:cForm.imgs||[]}]; setBurField(costModal,{costData:cd}); setCForm({supplier:"",rate:"",date:"",location:"",note:"",imgs:[]}); toast_("✅ Entry added"); },[cType,cForm,costModal,burItems,toast_]);
  const pickCostImg=useCallback(async files=>{ const list=files&&files.length!=null?[...files]:(files?[files]:[]); if(!list.length)return; const out=[]; for(const file of list){ const d=await imgToDataURL(file); if(d&&d.length<=1400000)out.push(d); else if(d)toast_("⚠️ One image too large, skipped"); } if(out.length)setCForm(f=>({...f,imgs:[...(f.imgs||[]),...out]})); },[toast_]);
  const delCostEntry=useCallback(eid=>{ const b=burItems.find(x=>x.id===costModal); if(!b)return; const cd=(b.costData||[]).filter(e=>e.id!==eid); setBurField(costModal,{costData:cd,...tickPatch(b,cd)}); },[costModal,burItems,ccomps]);
  // Edit an existing cost-data entry (optimistic local + debounced save)
  const updCostEntry=useCallback((eid,ch)=>{ recordFields([costModal],["costData","labour","material","plant","subcon","extras","preTick"],"Edit cost data of "+burLabel(costModal),true); setBurItems(prev=>prev.map(b=>{ if(b.id!==costModal)return b; const cd=(b.costData||[]).map(e=>e.id===eid?{...e,...ch}:e); return {...b,costData:cd,...(cd.some(e=>e.inRate)?tickPatch(b,cd):{})}; })); const k="cd_"+costModal; if(burTimers.current[k])clearTimeout(burTimers.current[k]); burTimers.current[k]=setTimeout(()=>{ const b=burItemsRef.current.find(x=>x.id===costModal); if(b)updateDoc(doc(db,"bur",costModal),b.costData.some(e=>e.inRate)?{costData:b.costData,labour:+b.labour||0,material:+b.material||0,plant:+b.plant||0,subcon:+b.subcon||0,extras:b.extras||[],preTick:b.preTick??null}:{costData:b.costData}).catch(()=>{}); },600); },[costModal,ccomps]);
  const entryPhotos=e=>e.imgs&&e.imgs.length?e.imgs:(e.img?[e.img]:[]);
  const addEntryPhotos=useCallback(async(eid,files)=>{ const list=files&&files.length!=null?[...files]:(files?[files]:[]); if(!list.length)return; const out=[]; for(const f of list){ const d=await imgToDataURL(f); if(d&&d.length<=1400000)out.push(d); else if(d)toast_("⚠️ One image too large, skipped"); } if(!out.length)return; const b=burItemsRef.current.find(x=>x.id===costModal); const e=b&&(b.costData||[]).find(x=>x.id===eid); const cur=e?entryPhotos(e):[]; updCostEntry(eid,{imgs:[...cur,...out],img:""}); },[costModal,updCostEntry,toast_]);
  const removeEntryPhoto=useCallback((eid,idx)=>{ const b=burItemsRef.current.find(x=>x.id===costModal); const e=b&&(b.costData||[]).find(x=>x.id===eid); if(!e)return; updCostEntry(eid,{imgs:entryPhotos(e).filter((_,j)=>j!==idx),img:""}); },[costModal,updCostEntry]);
  const useCostEntry=useCallback(entry=>{ setBurField(costModal,{subcon:entry.rate,quote:null}); toast_(`✅ Sub-con rate set: S$ ${fmt(entry.rate)} (${entry.supplier})`); },[costModal]);
  const approveQuote=useCallback(id=>{ if(!uRef.current)return; const b=burItems.find(x=>x.id===id); if(!b||!b.quote)return; setBurField(id,{quote:{...b.quote,status:"approved",approvedBy:uRef.current.name,approvedAt:new Date().toLocaleString("en-SG")}}); toast_(`✅ Approved by ${uRef.current.name}`); },[burItems]);
  const rejectQuote=useCallback(id=>{ setBurField(id,{subcon:0,quote:null}); toast_("🔴 Quote rejected and cleared"); },[]);

  // ── Codes & cats ──────────────────────────────────────────────────────────
  const pushCodes=async nc=>{ setCodes(nc); try{await setDoc(doc(db,"meta","codes"),{list:nc});}catch{} };
  const pushCats=async nc=>{ setCats(nc); try{await setDoc(doc(db,"meta","cats"),{list:nc});}catch{} };
  const saveNewCode=()=>{ if(!codeForm.code.trim())return; pushCodes([...codes,{id:uid(),...codeForm}]); setCodeForm({code:"",desc:"",cat:""}); setShowAddC(false); toast_("✅ Code added"); };
  const saveEditCode=()=>{ pushCodes(codes.map(c=>c.id===editCId?{...c,...codeForm}:c)); setEditCId(null); toast_("✅ Updated"); };
  const delCode=id=>{ if(confirm("Delete?"))pushCodes(codes.filter(c=>c.id!==id)); };

  // ── Calcs ─────────────────────────────────────────────────────────────────
  const aA=i=>(i.qty||0)*(i.rA||0); const aB=i=>(i.qty||0)*(i.rB||0);
  const sTot=id=>{const s=data?.sections.find(x=>x.id===id);return{A:s?.items?.reduce((t,i)=>t+aA(i),0)||0,B:s?.items?.reduce((t,i)=>t+aB(i),0)||0};};
  const gTot=()=>data?.sections.reduce((a,s)=>{const t=sTot(s.id);return{A:a.A+t.A,B:a.B+t.B};},{A:0,B:0})||{A:0,B:0};
  const cSum=()=>{const m={};data?.sections.forEach(s=>s.items?.forEach(i=>{const k=i.code||"Uncoded";if(!m[k])m[k]={A:0,B:0,n:0};m[k].A+=aA(i);m[k].B+=aB(i);m[k].n++;}));return m;};
  const rb=(()=>{const l=+rbu.labour||0,m=+rbu.material||0,p=+rbu.plant||0,s=+rbu.subcon||0,d=l+m+p+s,oh=d*(+rbu.oh||0)/100,sub=d+oh,pr=sub*(+rbu.profit||0)/100;return{d,oh,sub,pr,total:sub+pr};})();
  const toggleCol=id=>{const s=new Set(visCols);s.has(id)?s.delete(id):s.add(id);setVisCols(s);};

  // ── Gates: loading / login / profile / no-project ───────────────────────────
  const splash=msg=>(<div style={{minHeight:"100vh",display:"flex",alignItems:"center",justifyContent:"center",flexDirection:"column",gap:12,color:"#64748b",fontFamily:"system-ui,sans-serif"}}><div style={{fontSize:36}}>🏗️</div><div>{msg}</div></div>);

  if(authUser===undefined)return splash("Loading…");

  if(!authUser)return(
    <div style={{minHeight:"100vh",background:"linear-gradient(135deg,#4f46e5 0%,#9333ea 45%,#ec4899 100%)",display:"flex",alignItems:"center",justifyContent:"center",padding:16,fontFamily:"system-ui,sans-serif"}}>
      <div style={{background:"#fff",borderRadius:20,padding:32,width:"100%",maxWidth:360,boxShadow:"0 20px 60px rgba(0,0,0,.3)"}}>
        <div style={{textAlign:"center",marginBottom:24}}>
          <div style={{fontSize:52,marginBottom:8}}>🏗️</div>
          <h1 style={{fontSize:22,fontWeight:700,color:"#1e293b",margin:0}}>QS Workspace</h1>
          <p style={{fontSize:13,color:"#64748b",marginTop:6}}>Sign in to continue</p>
        </div>
        <div style={{display:"flex",flexDirection:"column",gap:12}}>
          <input type="email" autoComplete="username" style={{border:"1.5px solid #e2e8f0",borderRadius:10,padding:"10px 14px",fontSize:14,outline:"none"}} placeholder="Email" value={email} onChange={e=>setEmail(e.target.value)} onKeyDown={e=>e.key==="Enter"&&doLogin()}/>
          <input type="password" autoComplete="current-password" style={{border:"1.5px solid #e2e8f0",borderRadius:10,padding:"10px 14px",fontSize:14,outline:"none"}} placeholder="Password" value={pw} onChange={e=>setPw(e.target.value)} onKeyDown={e=>e.key==="Enter"&&doLogin()}/>
          {authErr&&<div style={{fontSize:12,color:"#dc2626",background:"#fef2f2",borderRadius:8,padding:"8px 10px"}}>{authErr}</div>}
          <button disabled={authBusy} style={{background:authBusy?"#93c5fd":"#2563eb",color:"#fff",border:"none",borderRadius:10,padding:"12px",fontSize:14,fontWeight:600,cursor:authBusy?"default":"pointer"}} onClick={doLogin}>{authBusy?"Signing in…":"Sign In →"}</button>
          <p style={{fontSize:11,color:"#94a3b8",textAlign:"center",margin:0}}>Accounts are created by your administrator.</p>
        </div>
      </div>
    </div>
  );

  if(!profile)return(
    <div style={{minHeight:"100vh",background:"linear-gradient(135deg,#4f46e5 0%,#9333ea 45%,#ec4899 100%)",display:"flex",alignItems:"center",justifyContent:"center",padding:16,fontFamily:"system-ui,sans-serif"}}>
      <div style={{background:"#fff",borderRadius:20,padding:32,width:"100%",maxWidth:360,boxShadow:"0 20px 60px rgba(0,0,0,.3)"}}>
        <div style={{textAlign:"center",marginBottom:20}}><div style={{fontSize:44,marginBottom:8}}>👤</div><h1 style={{fontSize:20,fontWeight:700,color:"#1e293b",margin:0}}>Set up your profile</h1><p style={{fontSize:13,color:"#64748b",marginTop:6}}>{authUser.email}</p></div>
        <div style={{display:"flex",flexDirection:"column",gap:12}}>
          <input style={{border:"1.5px solid #e2e8f0",borderRadius:10,padding:"10px 14px",fontSize:14,outline:"none"}} placeholder="Your name" value={nameIn} onChange={e=>setNameIn(e.target.value)} onKeyDown={e=>e.key==="Enter"&&saveProfile()}/>
          <select style={{border:"1.5px solid #e2e8f0",borderRadius:10,padding:"10px 14px",fontSize:14,outline:"none",background:"#fff"}} value={roleIn} onChange={e=>setRoleIn(e.target.value)}>{ROLES.map(r=><option key={r}>{r}</option>)}</select>
          <button style={{background:"#2563eb",color:"#fff",border:"none",borderRadius:10,padding:"12px",fontSize:14,fontWeight:600,cursor:"pointer"}} onClick={saveProfile}>Continue →</button>
          <button style={{background:"none",border:"none",color:"#94a3b8",fontSize:12,cursor:"pointer"}} onClick={()=>signOut(auth)}>Sign out</button>
        </div>
      </div>
    </div>
  );

  if(projects===null)return splash("Loading projects…");

  if(projects.length===0)return(
    <div style={{minHeight:"100vh",background:"#f1f5f9",display:"flex",alignItems:"center",justifyContent:"center",padding:16,fontFamily:"system-ui,sans-serif"}}>
      <div style={{background:"#fff",borderRadius:16,padding:40,maxWidth:420,textAlign:"center",boxShadow:"0 1px 4px rgba(0,0,0,.1)"}}>
        <div style={{fontSize:44,marginBottom:10}}>📁</div>
        <h2 style={{fontSize:18,margin:"0 0 6px",color:"#1e293b"}}>No projects yet</h2>
        <p style={{fontSize:13,color:"#64748b",marginTop:0}}>Create your first tender / project to start building the BOQ.</p>
        <button onClick={createProject} style={{marginTop:8,background:"#2563eb",color:"#fff",border:"none",borderRadius:10,padding:"11px 22px",fontSize:14,fontWeight:600,cursor:"pointer"}}>+ Create Project</button>
        <div style={{marginTop:16}}><button onClick={()=>signOut(auth)} style={{background:"none",border:"none",color:"#94a3b8",fontSize:12,cursor:"pointer"}}>Sign out ({user.email})</button></div>
      </div>
    </div>
  );

  if(!data)return splash("Loading project…");

  // ── Main app ────────────────────────────────────────────────────────────────
  const cs=data.sections.find(s=>s.id===selSec);
  const gt=gTot(); const tot=data.sections.reduce((a,s)=>a+(s.items?.length||0),0)||0;
  const vcols=ALL_COLS.filter(c=>visCols.has(c.id));
  const cc=cSum();
  const displayCats=cats;
  const sidebarCats=catSort==="none"?displayCats:[...displayCats].sort((a,b)=>catSort==="az"?String(a.name).localeCompare(String(b.name)):String(b.name).localeCompare(String(a.name)));
  const catName=id=>displayCats.find(c=>c.id===id)?.name||id;
  const BUR_MAX=300; const _q=burSearch.trim().toLowerCase();
  // When searching, look across the WHOLE BUR library (unless "Only in this category" is ticked); otherwise show the selected category.
  // Several words = every word must appear (e.g. "ceiling nrc" in code, description, supplier or note).
  const _base=_q&&!burSearchCat?burItems:burItems.filter(b=>b.catId===selCat); const _words=_q.split(/\s+/).filter(Boolean);
  const _filtered=_q?_base.filter(b=>{ const hay=[b.code,b.desc,...(b.costData||[]).flatMap(e=>[e.supplier,e.note])].map(x=>String(x||"").toLowerCase()).join(" | "); return _words.every(w=>hay.includes(w)); }):_base;
  const _sorted=[..._filtered].sort((a,b)=>{
    let av,bv;
    if(sortBy==="rate"){av=bTot(a);bv=bTot(b);return (av-bv)*sortDir;}
    if(sortBy==="cat"){av=catName(a.catId);bv=catName(b.catId);}
    else{av=a[sortBy]||"";bv=b[sortBy]||"";}
    return String(av).localeCompare(String(bv),undefined,{numeric:true})*sortDir;
  });
  const burGrouped=!(_q&&!burSearchCat); const _ordered=burGrouped?orderBySubcat(_sorted):_sorted;
  // a single category always shows ALL its items; only an all-categories search is capped (to keep the page fast)
  const catTotal=_ordered.length; const catItems=burGrouped?_ordered:_ordered.slice(0,BUR_MAX);
  const subcatsOf=cid=>[...new Set(burItems.filter(b=>b.catId===cid).map(b=>String(b.group||"").trim()).filter(Boolean))].sort((a,b)=>a.localeCompare(b));
  const saveCollapsed=st=>{ try{ localStorage.setItem("burCollapsedSubs",JSON.stringify([...st])); }catch{} return st; };
  const subKey=(cid,g)=>cid+"|"+g;
  const toggleSub=(cid,g)=>setCollapsedSubs(prev=>{ const st=new Set(prev), k=subKey(cid,g); st.has(k)?st.delete(k):st.add(k); return saveCollapsed(st); });
  const shownSubKeys=burGrouped?[...new Set(catItems.filter(b=>String(b.group||"").trim()).map(b=>subKey(b.catId,String(b.group).trim())))]:[];
  const collapseAllSubs=fold=>setCollapsedSubs(prev=>{ const st=new Set(prev); shownSubKeys.forEach(k=>fold?st.add(k):st.delete(k)); return saveCollapsed(st); });
  const toggleSort=f=>{ if(sortBy===f)setSortDir(d=>-d); else{setSortBy(f);setSortDir(1);} };
  const renameCat=cid=>{ const c=cats.find(x=>x.id===cid); if(!c)return; const nm=prompt("Rename category:",c.name); if(nm&&nm.trim())pushCats(cats.map(x=>x.id===cid?{...x,name:nm.trim()}:x)); };
  const delCat=cid=>{ const n=burItems.filter(b=>b.catId===cid).length; if(n>0){alert(`This category has ${n} item(s). Move them to another category first (open an item → CATEGORY), then delete.`);return;} if(confirm("Delete this empty category?")){ pushCats(cats.filter(x=>x.id!==cid)); if(selCat===cid){const o=cats.find(x=>x.id!==cid); setSelCat(o?o.id:"");} } };
  const toggleSel=id=>setSelBur(prev=>{const s=new Set(prev); s.has(id)?s.delete(id):s.add(id); return s;});
  // Checkbox click: plain click toggles one row; Shift+click selects (or unselects) every row between the last clicked one and this one.
  const clickSel=(e,id,list)=>{ const ids=list.map(x=>x.id), a=ids.indexOf(lastSelRef.current), b=ids.indexOf(id);
    if(e.shiftKey&&a>=0&&b>=0){ const on=!selBur.has(id), range=ids.slice(Math.min(a,b),Math.max(a,b)+1); setSelBur(prev=>{const s=new Set(prev); range.forEach(x=>on?s.add(x):s.delete(x)); return s;}); }
    else toggleSel(id);
    lastSelRef.current=id; };
  const bulkSetSubcat=async v=>{ const ids=[...selBur]; if(!ids.length)return; let g=v;
    if(v==="__new"){ g=prompt("New sub-category name (e.g. ALUMINIUM FIN):"); if(g==null)return; g=g.trim(); if(!g)return; }
    if(v==="__none")g="";
    recordFields(ids,["group"],`Sub-category on ${ids.length} item(s)`);
    try{ for(let i=0;i<ids.length;i+=400){const batch=writeBatch(db); ids.slice(i,i+400).forEach(id=>batch.update(doc(db,"bur",id),{group:g})); await batch.commit();} toast_(g?`✅ ${ids.length} item(s) → ${g}`:`✅ Removed sub-category from ${ids.length} item(s)`); clearSel(); }catch(e){ toast_("⚠️ "+e.message); } };
  const renameSubcat=async(cid,oldName)=>{ const nn=prompt(`Rename sub-category "${oldName}" to:\n(leave empty to remove the heading — items stay in the category)`,oldName); if(nn==null)return; const g=nn.trim(); if(g===oldName)return;
    const ids=burItemsRef.current.filter(b=>b.catId===cid&&String(b.group||"").trim()===oldName).map(b=>b.id); if(!ids.length)return;
    recordFields(ids,["group"],`Rename sub-category "${oldName}"`);
    try{ for(let i=0;i<ids.length;i+=400){const batch=writeBatch(db); ids.slice(i,i+400).forEach(id=>batch.update(doc(db,"bur",id),{group:g})); await batch.commit();} toast_("✅ Sub-category updated"); }catch(e){ toast_("⚠️ "+e.message); } };
  // Sync category + sub-category (and Excel row order) from a master BUR.xlsx, matching items by code.
  // Codes, rates, quotes and photos are not touched. Categories named in the Excel but missing in the app are created.
  const subcatsFromExcel=async file=>{ if(!file)return; if(!/\.xlsx$/i.test(file.name)){toast_("⚠️ Use a .xlsx file");return;}
    try{ toast_("⏳ Reading "+file.name+"…");
      const ExcelJS=(await import("exceljs")).default; const wb=new ExcelJS.Workbook(); await wb.xlsx.load(await file.arrayBuffer());
      const parsed=parseBurWorkbook(wb,cats.map(c=>c.name)).map(p=>({...p,code:codeWithUnit(p.code,p.unit)}));
      const findCat=n=>cats.find(c=>String(c.name).trim().toLowerCase()===String(n).trim().toLowerCase());
      const changes=[]; burItemsRef.current.forEach(b=>{ const m=b.code&&findBurByCode(parsed,b.code); if(!m)return; const c=findCat(m.cat);
        const moved=!c||c.id!==b.catId; if(moved||String(b.group||"")!==m.group||b.ord!==m.ord)changes.push({id:b.id,m,moved}); });
      if(!changes.length){toast_("Nothing to update — categories and sub-categories already match");return;}
      const moves=changes.filter(c=>c.moved), named=changes.filter(c=>c.m.group).length;
      const moveSummary=Object.entries(moves.reduce((o,c)=>{o[c.m.cat]=(o[c.m.cat]||0)+1;return o;},{})).map(([k,v])=>`  • ${v} → ${k}`).join("\n");
      if(!confirm(`Sync from ${file.name}?\n\n${changes.length} item(s) matched by code; ${named} get a sub-category heading.${moves.length?`\n\n${moves.length} item(s) will MOVE category to match the Excel:\n${moveSummary}`:""}\n\nCodes, rates, quotes and photos are not changed. You can Undo.`))return;
      let catList=cats; const missing=[...new Set(moves.map(c=>c.m.cat))].filter(n=>!findCat(n));
      if(missing.length){ catList=[...cats,...missing.map(n=>({id:"c"+uid(),name:n}))]; await setDoc(doc(db,"meta","cats"),{list:catList}); setCats(catList); }
      const catIdOf=n=>catList.find(c=>String(c.name).trim().toLowerCase()===String(n).trim().toLowerCase()).id;
      recordFields(changes.map(c=>c.id),["catId","group","ord"],"Sync categories from Excel");
      for(let i=0;i<changes.length;i+=400){const batch=writeBatch(db); changes.slice(i,i+400).forEach(({id,m})=>batch.update(doc(db,"bur",id),{catId:catIdOf(m.cat),group:m.group,ord:m.ord})); await batch.commit();}
      toast_(`✅ Updated ${changes.length} item(s)${moves.length?` (${moves.length} moved category)`:""}`);
    }catch(e){ toast_("⚠️ "+(e&&e.message||e)); } };
  // Duplicate items with everything (rates, waste, extra items, quotes, photos, category & sub-category).
  // The copy gets a free code like "X_copy_m2" (then _copy2, _copy3…) so it never clashes; a single copy opens for editing.
  const copyItems=async ids=>{ const src=ids.map(id=>burItemsRef.current.find(x=>x.id===id)).filter(Boolean); if(!src.length)return;
    const taken=new Set(burItemsRef.current.map(b=>String(b.code||"").trim().toLowerCase())); const newIds=[];
    try{ for(let i=0;i<src.length;i+=400){ const batch=writeBatch(db);
        src.slice(i,i+400).forEach(b=>{ const {id:_id,...data}=JSON.parse(JSON.stringify(b)); let code=data.code||"";
          if(code){ const base=codeBase(code,b.unit); let n=1; do{ code=codeWithUnit(`${base}_copy${n>1?n:""}`,b.unit); n++; }while(taken.has(code.toLowerCase())); taken.add(code.toLowerCase()); }
          data.costData=(data.costData||[]).map(e=>({...e,id:uid()})); data.extras=(data.extras||[]).map(x=>({...x,id:uid()})); if(data.ord!=null)data.ord=+data.ord+0.001;
          const ref=doc(collection(db,"bur")); newIds.push(ref.id); batch.set(ref,{...data,code,pid:b.pid||pid}); });
        await batch.commit(); }
      pushUndo({kind:"created",label:`Copy ${newIds.length} item(s)`,ids:newIds});
      if(newIds.length===1)setExpBur(newIds[0]); else clearSel();
      toast_(newIds.length===1?"⧉ Copied — edit the new item below":`⧉ Copied ${newIds.length} item(s)`);
    }catch(e){ toast_("⚠️ Copy failed: "+(e&&e.message||e)); } };
  // Change one unit to another on every BUR item in this project (e.g. nr -> no.), fixing the "_unit" ending of their codes too.
  // ── Sync from another project ────────────────────────────────────────────
  // Compares this project's BUR with another project's by code and lists what is different there ("changed") or only there ("new").
  // Nothing in this project is deleted. The chosen groups of fields are copied; the whole sync can be undone in one step.
  const SYNC_GROUPS={rate:{label:"Rates & build-up",keys:["labour","material","plant","subcon","waste","discount","extras","roundUp","preTick"]},
    text:{label:"Description & unit",keys:["desc","unit"]}, cd:{label:"Quotes & photos (Cost Data)",keys:["costData"]}, cat:{label:"Category & sub-category",keys:["catId","group","ord"]}};
  const syncNorm=(b,k)=>{ const v=b[k]; if(k==="costData")return JSON.stringify((v||[]).map(e=>[e.component,e.supplier,+e.rate||0,e.expr||"",e.date||"",e.location||"",e.note||"",!!e.inRate,(e.imgs||[]).length,e.per||"",e.cover||""]));
    if(k==="extras")return JSON.stringify((v||[]).map(x=>[x.name,x.type,+x.value||0,x.expr||""])); if(["labour","material","plant","subcon","waste","discount"].includes(k))return +v||0; if(v==null||v==="")return ""; return typeof v==="object"?JSON.stringify(v):String(v); };
  const openSync=()=>setSyncState({src:"",rows:[],sel:new Set(),opts:{rate:true,text:true,cd:true,cat:true}});
  const loadSync=async src=>{ setSyncState(st=>({...st,src,loading:true,rows:[],sel:new Set()})); if(!src)return setSyncState(st=>({...st,loading:false}));
    try{ const snap=await getDocs(query(collection(db,"bur"),where("pid","==",src))); const srcItems=snap.docs.map(d=>({id:d.id,...d.data()}));
      const mine=burItemsRef.current, rows=[];
      srcItems.forEach(sb=>{ if(!sb.code)return; const cb=findBurByCode(mine,sb.code);
        if(!cb){ rows.push({key:sb.id,kind:"new",src:sb}); return; }
        const diffs=Object.entries(SYNC_GROUPS).filter(([,g])=>g.keys.some(k=>syncNorm(cb,k)!==syncNorm(sb,k))).map(([gk])=>gk);
        if(diffs.length)rows.push({key:sb.id,kind:"changed",cur:cb,src:sb,diffs}); });
      rows.sort((a,b)=>(a.kind===b.kind?0:a.kind==="changed"?-1:1)||String(a.src.code).localeCompare(String(b.src.code)));
      setSyncState(st=>st&&st.src===src?{...st,loading:false,rows,sel:new Set(rows.map(r=>r.key))}:st);
    }catch(e){ toast_("⚠️ "+e.message); setSyncState(st=>({...st,loading:false})); } };
  const applySync=async()=>{ const st=syncState; if(!st)return; const keys=Object.entries(st.opts).filter(([,on])=>on).flatMap(([g])=>SYNC_GROUPS[g].keys);
    const pick=st.rows.filter(r=>st.sel.has(r.key)); if(!pick.length||!keys.length){toast_("Nothing selected");return;}
    const chg=pick.filter(r=>r.kind==="changed"), nw=pick.filter(r=>r.kind==="new"); const srcName=(projects||[]).find(p=>p.id===st.src)?.name||"other project";
    if(!confirm(`Sync from "${srcName}"?\n\n• Update ${chg.length} item(s) (${Object.entries(st.opts).filter(([,on])=>on).map(([g])=>SYNC_GROUPS[g].label).join(", ")})\n• Add ${nw.length} new item(s)\n\nNothing is deleted. You can Undo.`))return;
    try{ const parts=[];
      if(chg.length){ parts.push({kind:"fields",items:chg.map(r=>({id:r.cur.id,before:Object.fromEntries(keys.map(k=>[k,r.cur[k]===undefined?null:JSON.parse(JSON.stringify(r.cur[k]))]))}))});
        for(let i=0;i<chg.length;i+=400){ const batch=writeBatch(db); chg.slice(i,i+400).forEach(r=>batch.update(doc(db,"bur",r.cur.id),Object.fromEntries(keys.map(k=>[k,r.src[k]===undefined?null:r.src[k]])))); await batch.commit(); } }
      if(nw.length){ const ids=[]; for(let i=0;i<nw.length;i+=400){ const batch=writeBatch(db); nw.slice(i,i+400).forEach(r=>{ const {id:_i,...data}=r.src; const ref=doc(collection(db,"bur")); ids.push(ref.id); batch.set(ref,{...data,pid}); }); await batch.commit(); } parts.push({kind:"created",ids}); }
      pushUndo({kind:"multi",label:`Sync ${pick.length} item(s) from ${srcName}`,parts});
      toast_(`✅ Synced from ${srcName}: ${chg.length} updated, ${nw.length} added`); setSyncState(null);
    }catch(e){ toast_("⚠️ Sync failed: "+(e&&e.message||e)); } };
  const replaceUnit=async()=>{ const from=prompt("Replace which unit? (exactly as shown, e.g. nr)","nr"); if(from==null||!from.trim())return;
    const to=prompt(`Change every "${from.trim()}" to:`,"no."); if(to==null||!to.trim())return; const f=from.trim(), t=to.trim();
    const list=burItemsRef.current.filter(b=>String(b.unit||"").trim().toLowerCase()===f.toLowerCase()); if(!list.length){toast_(`No BUR items with unit "${f}"`);return;}
    if(!confirm(`Change the unit of ${list.length} BUR item(s) from "${f}" to "${t}"?\nTheir codes change too, e.g. ${list[0].code||"(no code)"} → ${swapUnitTag(list[0].code,list[0].unit,t)||"(no code)"}\n\nYou can Undo.`))return;
    recordFields(list.map(b=>b.id),["unit","code"],`Unit ${f} → ${t} on ${list.length} item(s)`);
    try{ for(let i=0;i<list.length;i+=400){ const batch=writeBatch(db); list.slice(i,i+400).forEach(b=>batch.update(doc(db,"bur",b.id),{unit:t,code:swapUnitTag(b.code,b.unit,t)||""})); await batch.commit(); } toast_(`✅ ${list.length} item(s): ${f} → ${t}`); }catch(e){ toast_("⚠️ "+e.message); } };
  const bulkDelete=async()=>{ const ids=[...selBur]; if(!ids.length)return;
    if(!confirm(`Delete ${ids.length} BUR item(s)?

You can bring them back with the Undo button (until you refresh the page).`))return;
    const docs=ids.map(id=>burItemsRef.current.find(x=>x.id===id)).filter(Boolean).map(({id,...data})=>({id,data}));
    try{ for(let i=0;i<ids.length;i+=400){const batch=writeBatch(db); ids.slice(i,i+400).forEach(id=>batch.delete(doc(db,"bur",id))); await batch.commit();}
      pushUndo({kind:"deleted",label:`Delete ${docs.length} item(s)`,docs}); toast_(`🗑 Deleted ${ids.length} item(s)`); clearSel(); }catch(e){ toast_("⚠️ "+e.message); } };
  const clearSel=()=>setSelBur(new Set());
  const moveItemsToCat=async(ids,cid)=>{ if(!ids.length||!cid)return; recordFields(ids,["catId"],`Move ${ids.length} item(s) to ${catName(cid)}`); for(let i=0;i<ids.length;i+=400){const batch=writeBatch(db); ids.slice(i,i+400).forEach(id=>batch.update(doc(db,"bur",id),{catId:cid})); await batch.commit();} };
  const bulkMoveCat=async cid=>{ const ids=[...selBur]; if(!ids.length||!cid)return; await moveItemsToCat(ids,cid); toast_(`✅ Moved ${ids.length} item(s)`); clearSel(); };
  // Replace this project's whole BUR library with the items in a master BUR workbook (layout of BUR.xlsx).
  // Downloads a JSON backup of the current items first, and only deletes after the user confirms the counts.
  const replaceBurFromExcel=async file=>{
    if(!file||!pid)return; if(!/\.xlsx$/i.test(file.name)){toast_("⚠️ Use a .xlsx file");return;}
    try{
      toast_("⏳ Reading "+file.name+"…");
      const ExcelJS=(await import("exceljs")).default; const wb=new ExcelJS.Workbook(); await wb.xlsx.load(await file.arrayBuffer());
      const parsed=parseBurWorkbook(wb,cats.map(c=>c.name));
      if(!parsed.length){toast_("⚠️ No items with a Code (column D) found in "+file.name);return;}
      const catNames=[...new Set(parsed.map(p=>p.cat))];
      const byName=new Map(cats.map(c=>[String(c.name).trim().toLowerCase(),c]));
      const newCats=catNames.filter(n=>!byName.has(n.toLowerCase()));
      const old=burItemsRef.current.filter(b=>b.pid===pid||!b.pid);
      if(!confirm(`Replace the BUR library for this project?\n\n• DELETE all ${old.length} current BUR items (incl. their cost data & photos)\n• ADD ${parsed.length} items from ${file.name} in ${catNames.length} categories${newCats.length?` (${newCats.length} new: ${newCats.join(", ")})`:""}\n\nA backup of the current items will download first. This cannot be undone in the app.`))return;
      // 1) backup
      const bk=new Blob([JSON.stringify({exported:new Date().toISOString(),pid,cats,items:old},null,1)],{type:"application/json"});
      const a=document.createElement("a"); a.href=URL.createObjectURL(bk); a.download=`BUR_backup_${new Date().toISOString().slice(0,10)}.json`; a.click(); URL.revokeObjectURL(a.href);
      // 2) categories: reuse by name, add the missing ones
      let catList=cats; if(newCats.length){ catList=[...cats,...newCats.map(n=>({id:"c"+uid(),name:n}))]; await setDoc(doc(db,"meta","cats"),{list:catList}); setCats(catList); }
      const catId=n=>catList.find(c=>String(c.name).trim().toLowerCase()===n.toLowerCase()).id;
      // 3) delete old, 4) add new
      toast_(`⏳ Removing ${old.length} old items…`);
      for(let i=0;i<old.length;i+=400){ const batch=writeBatch(db); old.slice(i,i+400).forEach(b=>batch.delete(doc(db,"bur",b.id))); await batch.commit(); }
      toast_(`⏳ Adding ${parsed.length} items…`);
      for(let i=0;i<parsed.length;i+=400){ const batch=writeBatch(db); parsed.slice(i,i+400).forEach(p=>batch.set(doc(collection(db,"bur")),{pid,catId:catId(p.cat),code:codeWithUnit(p.code,p.unit),desc:p.desc,unit:p.unit,labour:0,material:p.rate,plant:0,subcon:0,oh:0,profit:0,
        costData:p.quotes.map(q=>({id:uid(),component:"subcon",supplier:q.supplier,rate:q.rate,date:q.date,location:"",note:"",imgs:[]})),quote:null,group:p.group||"",ord:p.ord})); await batch.commit(); }
      setSelCat(catId(parsed[0].cat));
      toast_(`✅ BUR replaced: ${parsed.length} items from ${file.name}`);
    }catch(e){ toast_("⚠️ Replace failed: "+(e&&e.message||e)); }
  };
  // One-off: append "_unit" to every BUR code in this project that doesn't already end with it.
  const codesMissingUnit=burItems.filter(b=>b.code&&codeWithUnit(b.code,b.unit)!==String(b.code).trim());
  const codesWithSpaces=burItems.filter(b=>b.code&&cleanCode(b.code)!==String(b.code).trim());
  const fixCodeSpaces=async()=>{ const list=codesWithSpaces; if(!list.length)return;
    if(!confirm(`Replace spaces and "/" with "_" in ${list.length} BUR code(s)?\n\n${list.slice(0,6).map(b=>`${b.code}  →  ${cleanCode(b.code)}`).join("\n")}${list.length>6?"\n…":""}\n\nBOQ rows using the old spelling still match. You can Undo.`))return;
    recordFields(list.map(b=>b.id),["code"],`Fix ${list.length} code(s) with spaces / "/"`);
    try{ for(let i=0;i<list.length;i+=400){ const batch=writeBatch(db); list.slice(i,i+400).forEach(b=>batch.update(doc(db,"bur",b.id),{code:cleanCode(b.code)})); await batch.commit(); } toast_(`✅ Fixed ${list.length} code(s)`); }catch(e){ toast_("⚠️ "+e.message); } };
  const addUnitToCodes=async()=>{ const list=codesMissingUnit; if(!list.length)return; const _u=()=>recordFields(list.map(b=>b.id),["code"],`Add _unit to ${list.length} codes`);
    if(!confirm(`Add "_unit" to ${list.length} BUR code(s)?\n\ne.g. ${list[0].code}  →  ${codeWithUnit(list[0].code,list[0].unit)}\n\nBOQ rows and Master BQ templates using the old codes will still match.`))return;
    try{ _u(); for(let i=0;i<list.length;i+=400){ const batch=writeBatch(db); list.slice(i,i+400).forEach(b=>batch.update(doc(db,"bur",b.id),{code:codeWithUnit(b.code,b.unit)})); await batch.commit(); } toast_(`✅ Updated ${list.length} code(s)`); }catch(e){ toast_("⚠️ "+e.message); } };
  // Set Waste/ OH % on all selected items at once (e.g. 10%).
  // Add (or update) one lump-sum extra item on every selected item, e.g. "conc pump" S$ 10. Same name (any case) = update its amount;
  // amount 0 / empty = remove that extra item. The amount may be a formula.
  const bulkAddExtra=async()=>{ const ids=[...selBur]; if(!ids.length)return;
    const name=prompt(`Extra item name for ${ids.length} selected item(s) (e.g. conc pump):`,""); if(name==null||!name.trim())return;
    const amt=prompt(`Amount S$ per unit for "${name.trim()}" (number or formula; 0 = remove it):`,""); if(amt==null)return;
    const n=String(amt).trim()===""?0:calcExpr(amt); if(!isFinite(n)){toast_("⚠️ Amount has a mistake — check the number / formula");return;}
    const nm=name.trim(), key=nm.toLowerCase(), isF=/[-+*/()]/.test(String(amt).replace(/^\s*[=-]/,""));
    recordFields(ids,["extras"],`Extra "${nm}" on ${ids.length} item(s)`);
    try{ for(let i=0;i<ids.length;i+=400){ const batch=writeBatch(db);
        ids.slice(i,i+400).forEach(id=>{ const b=burItemsRef.current.find(x=>x.id===id); if(!b)return; const xs=(b.extras||[]).filter(x=>!(String(x.name||"").trim().toLowerCase()===key&&!x.fromCD));
          batch.update(doc(db,"bur",id),{extras:n?[...xs,{id:uid(),name:nm,type:"$",value:Math.round(n*10000)/10000,...(isF?{expr:String(amt).trim()}:{})}]:xs}); });
        await batch.commit(); }
      toast_(n?`✅ "${nm}" S$ ${fmt(n)} on ${ids.length} item(s)`:`✅ Removed "${nm}" from ${ids.length} item(s)`); clearSel(); }catch(e){ toast_("⚠️ "+e.message); } };
  const bulkSetDiscount=async v=>{ const ids=[...selBur]; if(!ids.length)return; let d=v;
    if(v==="__custom"){ const t=prompt("Discount % for the selected item(s):","5"); if(t==null)return; d=t; }
    const val=+String(d).replace(/%/g,"").trim(); if(!isFinite(val)){toast_("⚠️ Enter a number");return;}
    recordFields(ids,["discount"],`Discount ${val}% on ${ids.length} item(s)`);
    try{ for(let i=0;i<ids.length;i+=400){const batch=writeBatch(db); ids.slice(i,i+400).forEach(id=>batch.update(doc(db,"bur",id),{discount:val})); await batch.commit();} toast_(val?`✅ Discount ${val}% on ${ids.length} item(s)`:`✅ Discount removed from ${ids.length} item(s)`); clearSel(); }catch(e){ toast_("⚠️ "+e.message); } };
  const bulkSetWaste=async v=>{ const ids=[...selBur]; if(!ids.length)return; let w=v;
    if(v==="__custom"){ const t=prompt("Waste/ OH % for the selected item(s):","10"); if(t==null)return; w=t; }
    const val=+String(w).replace(/%/g,"").trim(); if(!isFinite(val)){toast_("⚠️ Enter a number");return;}
    recordFields(ids,["waste"],`Waste/ OH ${val}% on ${ids.length} item(s)`);
    try{ for(let i=0;i<ids.length;i+=400){const batch=writeBatch(db); ids.slice(i,i+400).forEach(id=>batch.update(doc(db,"bur",id),{waste:val})); await batch.commit();} toast_(val?`✅ Waste/ OH ${val}% on ${ids.length} item(s)`:`✅ Waste/ OH removed from ${ids.length} item(s)`); clearSel(); }catch(e){ toast_("⚠️ "+e.message); } };
  const bulkSetRound=async v=>{ const ids=[...selBur]; if(!ids.length)return; const val=v===""?null:+v; recordFields(ids,["roundUp"],`Round-up on ${ids.length} item(s)`); for(let i=0;i<ids.length;i+=400){const batch=writeBatch(db); ids.slice(i,i+400).forEach(id=>batch.update(doc(db,"bur",id),{roundUp:val})); await batch.commit();} toast_(`✅ ${val===null?"Exact rate":`Round up to ${roundLabel(val)}`} on ${ids.length} item(s)`); clearSel(); };
  const bulkSetUnit=async u=>{ const ids=[...selBur]; if(!ids.length||!u)return; recordFields(ids,["unit","code"],`Set unit on ${ids.length} item(s)`); for(let i=0;i<ids.length;i+=400){const batch=writeBatch(db); ids.slice(i,i+400).forEach(id=>{ const b=burItemsRef.current.find(x=>x.id===id); batch.update(doc(db,"bur",id),b&&b.code?{unit:u,code:swapUnitTag(b.code,b.unit,u)}:{unit:u}); }); await batch.commit();} toast_(`✅ Set unit on ${ids.length} item(s)`); clearSel(); };
  const handleDropCat=(cid,e)=>{ e.preventDefault(); const id=e.dataTransfer.getData("text/bur"); if(!id)return; const ids=(selBur.has(id)&&selBur.size>1)?[...selBur]:[id]; moveItemsToCat(ids,cid).then(()=>{toast_(`✅ Moved ${ids.length} item(s)`);clearSel();}); };
  const _codeCount={}; burItems.forEach(b=>{const k=(b.code||"").trim().toLowerCase(); if(k)_codeCount[k]=(_codeCount[k]||0)+1;});
  const isDupCode=code=>{const k=(code||"").trim().toLowerCase(); return !!k&&_codeCount[k]>1;};
  const dupCount=Object.values(_codeCount).filter(n=>n>1).length;
  const burCw=k=>burColW[k]??({desc:360,unit:60,code:150,labour:90,material:95,plant:80,rate:100,cd:90}[k]||80);
  const startBurDrag=(k,e)=>{ e.preventDefault();e.stopPropagation(); burDragRef.current={k,start:e.clientX,startW:burCw(k)}; };
  const burHead=[["desc","Description","left"],["unit","Unit","left"],["code","BUR Code","left"],["labour","Labour","right"],["material","Material","right"],["plant","Plant","right"],["rate","Rate (S$)","right"],["cd","Cost Data","center"]];
  const burEmpty=(<div style={{background:"#fff",borderRadius:12,padding:40,textAlign:"center",color:"#94a3b8"}}><div style={{fontSize:36,marginBottom:8}}>📊</div><div style={{fontWeight:600,fontSize:14,marginBottom:4}}>No items{_q?" match your search":" here"}</div><div style={{fontSize:13}}>{_q?"Try a different search term":'Click "+ Add Item" or "Paste from Excel"'}</div></div>);
  const burTableEl=items=>(
    <div style={{background:"#fff",borderRadius:8,boxShadow:"0 1px 4px rgba(0,0,0,.08)",overflow:"auto"}}>
      <table style={{width:"100%",borderCollapse:"collapse",fontSize:12,tableLayout:"fixed"}}>
        <colgroup><col style={{width:28}}/><col style={{width:24}}/>{burHead.map(([k])=><col key={k} style={{width:burCw(k)}}/>)}<col style={{width:52}}/></colgroup>
        <thead><tr style={{background:"#fef9c3",color:"#92400e",fontSize:11}}>
          <th style={{borderBottom:"1px solid #fde68a",textAlign:"center"}}><input type="checkbox" title="Select all shown" checked={items.length>0&&items.every(it=>selBur.has(it.id))} onChange={e=>{setSelBur(prev=>{const s=new Set(prev); items.forEach(it=>e.target.checked?s.add(it.id):s.delete(it.id)); return s;});}} style={{accentColor:"#7c3aed",cursor:"pointer"}}/></th>
          <th style={{borderBottom:"1px solid #fde68a"}}></th>
          {burHead.map(([k,l,al])=><th key={k} style={{padding:"8px 8px",textAlign:al,fontWeight:700,borderBottom:"1px solid #fde68a",whiteSpace:"nowrap",position:"relative",overflow:"hidden"}}>{l}<div onMouseDown={e=>startBurDrag(k,e)} title="Drag to resize" style={{position:"absolute",top:0,right:0,width:8,height:"100%",cursor:"col-resize"}}/></th>)}
          <th style={{borderBottom:"1px solid #fde68a"}}></th>
        </tr></thead>
        <tbody>
          {items.map((item,ri)=>{
            const isExp=expBur===item.id; const total=bTot(item); const direct=baseDirect(item); const xs=item.extras||[]; const setXs=v=>updBur(item.id,{extras:v}); const cdVis=(item.costData||[]).filter(e=>tabExists(e.component)), cdCount=cdVis.length, cdTip=Object.entries(cdVis.reduce((o,e)=>{const l=compLabel(e.component);o[l]=(o[l]||0)+1;return o;},{})).map(([l,n])=>`${l}: ${n}`).join(" · "); const dup=isDupCode(item.code); const checked=selBur.has(item.id);
            const lbl={fontSize:10,color:"#94a3b8",fontWeight:600,display:"block",marginBottom:2}; const inp={border:"1.5px solid #e2e8f0",borderRadius:6,padding:"5px 8px",fontSize:12,outline:"none"};
            const g=String(item.group||"").trim(), folded=burGrouped&&!!g&&collapsedSubs.has(subKey(item.catId,g)), newGrp=burGrouped&&g&&(ri===0||String(items[ri-1].group||"").trim()!==g), grpItems=newGrp?items.filter(x=>String(x.group||"").trim()===g):[];
            return(<Fragment key={item.id}>
              {newGrp&&<tr style={{background:"#f8fafc"}}><td style={{textAlign:"center",borderBottom:"1px solid #e2e8f0"}}><input type="checkbox" title="Select all in this sub-category" checked={grpItems.every(x=>selBur.has(x.id))} onChange={e=>setSelBur(prev=>{const s=new Set(prev); grpItems.forEach(x=>e.target.checked?s.add(x.id):s.delete(x.id)); return s;})} style={{accentColor:"#7c3aed",cursor:"pointer"}}/></td>
                <td colSpan={burHead.length+2} onClick={e=>{ if(e.target.tagName!=="BUTTON")toggleSub(item.catId,g); }} title={folded?"Click to expand":"Click to collapse"} style={{padding:"7px 8px",borderBottom:"1px solid #e2e8f0",fontWeight:700,color:"#0f172a",fontSize:12,cursor:"pointer",userSelect:"none"}}><span style={{display:"inline-block",width:14,color:"#64748b"}}>{folded?"▸":"▾"}</span><span style={{textDecoration:"underline"}}>{g}</span> <span style={{textDecoration:"none",fontWeight:500,color:"#94a3b8",fontSize:11}}>({grpItems.length})</span>
                  <button onClick={()=>renameSubcat(item.catId,g)} title="Rename this sub-category" style={{marginLeft:8,background:"none",border:"none",cursor:"pointer",color:"#94a3b8",fontSize:12}}>✎</button></td></tr>}
              {!folded&&<>
              <tr style={{borderBottom:"1px solid #f8fafc",background:checked?"#ede9fe":(ri%2===1?"#f5f3ff":"#fff")}}>
                <td style={{padding:"2px 4px",textAlign:"center"}}><input type="checkbox" title="Click to select · Shift+click to select a range" checked={checked} onChange={()=>{}} onClick={e=>clickSel(e,item.id,items)} style={{accentColor:"#7c3aed",cursor:"pointer"}}/></td>
                <td draggable onDragStart={e=>{e.dataTransfer.setData("text/bur",item.id);e.dataTransfer.effectAllowed="move";}} title="Drag me onto a category to move • click to expand" style={{padding:"2px 4px",textAlign:"center",cursor:"grab",color:isExp?"#2563eb":"#94a3b8",userSelect:"none"}} onClick={()=>setExpBur(isExp?null:item.id)}>⠿{isExp?"▼":"▶"}</td>
                <td style={{padding:"3px 6px",overflow:"hidden"}}><input style={{width:"100%",border:"none",fontSize:12,outline:"none",background:"transparent"}} value={item.desc||""} onChange={e=>updBur(item.id,{desc:e.target.value})}/></td>
                <td style={{padding:"3px 6px"}}><select style={{border:"none",fontSize:11,outline:"none",background:"transparent",width:"100%"}} value={item.unit||"sum"} onChange={e=>updBur(item.id,{unit:e.target.value,code:swapUnitTag(item.code,item.unit,e.target.value)})}>{unitOpts(item.unit).map(u=><option key={u}>{u}</option>)}</select></td>
                <td style={{padding:"3px 6px",overflow:"hidden",background:dup?"#fee2e2":undefined}}><input title={dup?"⚠️ Duplicate BUR code":""} style={{width:"100%",border:"none",fontSize:11,outline:"none",background:"transparent",fontFamily:"monospace",fontWeight:700,color:dup?"#b91c1c":"#1d4ed8"}} value={item.code||""} onChange={e=>updBur(item.id,{code:e.target.value})} onBlur={e=>{const c=codeWithUnit(cleanCode(e.target.value),item.unit); if(c!==e.target.value)updBur(item.id,{code:c});}}/></td>
                <td style={{padding:"3px 6px",textAlign:"right"}}><input type="number" style={{width:"100%",border:"none",fontSize:12,outline:"none",background:"transparent",textAlign:"right"}} value={item.labour?item.labour:""} onChange={e=>updBur(item.id,{labour:+e.target.value||0})}/></td>
                <td style={{padding:"3px 6px",textAlign:"right"}}><input type="number" style={{width:"100%",border:"none",fontSize:12,outline:"none",background:"transparent",textAlign:"right"}} value={item.material?item.material:""} onChange={e=>updBur(item.id,{material:+e.target.value||0})}/></td>
                <td style={{padding:"3px 6px",textAlign:"right"}}><input type="number" style={{width:"100%",border:"none",fontSize:12,outline:"none",background:"transparent",textAlign:"right"}} value={item.plant?item.plant:""} onChange={e=>updBur(item.id,{plant:+e.target.value||0})}/></td>
                <td style={{padding:"3px 8px",textAlign:"right",fontWeight:700,color:"#1d4ed8",whiteSpace:"nowrap"}}>{total>0?<span title={hasRoundUp(item)?`Rounded up (${roundLabel(item.roundUp)})`:""}>{hasRoundUp(item)&&<span style={{color:"#7c3aed",fontSize:10,marginRight:2}}>↑</span>}{fmt(total)}</span>:""}</td>
                <td style={{padding:"3px 6px",textAlign:"center"}}><button onClick={()=>{setCostModal(item.id);setCType("subcon");}} title={cdTip||"No cost data yet"} style={{background:"#fef9c3",border:"1px solid #fde68a",borderRadius:6,padding:"3px 8px",fontSize:11,cursor:"pointer",color:"#92400e",fontWeight:600,whiteSpace:"nowrap"}}>📊{cdCount?` ${cdCount}`:""}</button></td>
                <td style={{padding:"3px 4px",textAlign:"center"}}><button onClick={()=>copyItems([item.id])} title="Copy this item (all details, quotes & photos)" style={{border:"none",background:"none",color:"#2563eb",cursor:"pointer",fontSize:13,fontWeight:700,padding:"0 3px"}}>⧉</button><button onClick={()=>delBur(item.id)} title="Delete" style={{border:"none",background:"none",color:"#ef4444",cursor:"pointer",fontSize:13,fontWeight:700,padding:"0 3px"}}>✕</button></td>
              </tr>
              {isExp&&(<tr><td colSpan={11} style={{background:"#fafafa",padding:"12px 16px",borderBottom:"1px solid #eef2f7"}}>
                <div style={{display:"flex",gap:16,flexWrap:"wrap",marginBottom:12}}>
                  <div style={{flex:"2 1 320px"}}><label style={lbl}>DESCRIPTION (full)</label><textarea value={item.desc||""} onChange={e=>updBur(item.id,{desc:e.target.value})} style={{width:"100%",border:"1.5px solid #e2e8f0",borderRadius:6,padding:"7px 10px",fontSize:13,outline:"none",minHeight:64,resize:"vertical",boxSizing:"border-box",lineHeight:1.4}}/></div>
                  <div style={{flex:"1 1 180px"}}><label style={lbl}>BUR CODE</label><input value={item.code||""} onChange={e=>updBur(item.id,{code:e.target.value})} onBlur={e=>{const c=codeWithUnit(cleanCode(e.target.value),item.unit); if(c!==e.target.value)updBur(item.id,{code:c});}} style={{width:"100%",border:"1.5px solid #e2e8f0",borderRadius:6,padding:"7px 10px",fontSize:13,outline:"none",fontFamily:"monospace",fontWeight:700,color:"#1d4ed8",boxSizing:"border-box"}}/></div>
                </div>
                <div style={{display:"flex",gap:16,flexWrap:"wrap",alignItems:"flex-end"}}>
                  <div><label style={lbl}>SUBCON (S$)</label><input type="number" style={{...inp,width:90}} value={item.subcon?item.subcon:""} onChange={e=>updBur(item.id,{subcon:+e.target.value||0})}/></div>
                  <div><label style={lbl}>DISCOUNT %</label><input type="number" style={{...inp,width:70}} placeholder="0" value={item.discount?item.discount:""} onChange={e=>updBur(item.id,{discount:+e.target.value||0})}/></div>
                  <div><label style={lbl}>WASTE/ OH %</label><input type="number" style={{...inp,width:70}} placeholder="0" value={item.waste?item.waste:""} onChange={e=>updBur(item.id,{waste:+e.target.value||0})}/></div>
                  <div><label style={lbl}>ROUND UP RATE</label><select style={{...inp,width:110}} value={hasRoundUp(item)?String(item.roundUp):""} onChange={e=>updBur(item.id,{roundUp:e.target.value===""?null:+e.target.value})}>{ROUND_OPTS.map(([v,l])=><option key={v} value={v}>{l}</option>)}</select></div>
                  <div><label style={lbl}>CATEGORY</label><select style={{...inp,background:"#fff"}} value={item.catId||"cat13"} onChange={e=>updBur(item.id,{catId:e.target.value})}>{cats.map(c=><option key={c.id} value={c.id}>{c.name}</option>)}</select></div>
                  <div><label style={lbl}>SUB-CATEGORY</label><input list={"subcats_"+item.id} style={{...inp,width:200}} placeholder="(none)" value={item.group||""} onChange={e=>updBur(item.id,{group:e.target.value})}/><datalist id={"subcats_"+item.id}>{subcatsOf(item.catId).map(g=><option key={g} value={g}/>)}</datalist></div>
                  <div style={{marginLeft:"auto",fontSize:12,color:"#475569"}}>Direct <b>S$ {fmt(direct)}</b>{+item.discount?<> · Discount ({+item.discount}%) <b style={{color:"#dc2626"}}>−S$ {fmt(discAmt(item))}</b></>:null}{+item.waste?<> · Waste/ OH ({+item.waste}%) <b>S$ {fmt(wasteAmt(item))}</b></>:null}{xs.length?<> · Extras <b>S$ {fmt(extrasTot(item))}</b></>:null} · <span style={{color:"#1d4ed8",fontWeight:800}}>Rate S$ {fmt(total)}/{item.unit}</span></div>
                </div>
                <div style={{marginTop:12,paddingTop:10,borderTop:"1px dashed #e2e8f0"}}>
                  <div style={{display:"flex",alignItems:"center",gap:6,flexWrap:"wrap",marginBottom:xs.length?8:0}}>
                    <span style={lbl}>EXTRA ITEMS (lump sum S$ per {item.unit||"unit"})</span>
                    {EXTRA_PRESETS.map(([n,t])=><button key={n} onClick={()=>setXs([...xs,{id:uid(),name:n,type:t,value:""}])} style={{background:"#f1f5f9",border:"1px solid #e2e8f0",borderRadius:6,padding:"3px 9px",fontSize:11,cursor:"pointer",color:"#334155",fontWeight:600}}>+ {n}</button>)}
                    <button onClick={()=>setXs([...xs,{id:uid(),name:"",type:"$",value:""}])} title="Add your own lump-sum item" style={{background:"#f1f5f9",border:"1px dashed #94a3b8",borderRadius:6,padding:"3px 9px",fontSize:11,cursor:"pointer",color:"#334155",fontWeight:600}}>+ Other…</button>
                  </div>
                  {xs.map(x=><div key={x.id} style={{display:"flex",alignItems:"center",gap:8,marginBottom:6}}>
                    <input list="extraNames" placeholder="Name (e.g. Scaffold)" autoFocus={!x.name} style={{...inp,width:180}} value={x.name} onChange={e=>setXs(xs.map(y=>y.id===x.id?{...y,name:e.target.value}:y))}/>
                    {(()=>{ const txt=x.expr??(x.value===""||x.value==null?"":String(x.value)); const isF=/[-+*/()]/.test(txt.replace(/^\s*-/,"")); const bad=txt.trim()!==""&&!isFinite(calcExpr(txt));
                      return <><input title={isF?`ƒ ${txt}`:"Type a number or a formula, e.g. ((1+1)*2)+((3*1)+2)"} onFocus={()=>setEditExpr(x.id)} onBlur={()=>setEditExpr(null)} style={{...inp,width:isF&&editExpr===x.id?190:80,textAlign:"right",borderColor:bad?"#ef4444":"#e2e8f0"}} value={editExpr===x.id||!isF||bad?txt:String(+(+x.value||0).toFixed(4))} onChange={e=>{ const t=e.target.value, n=calcExpr(t); setXs(xs.map(y=>y.id===x.id?{...y,expr:t,value:t.trim()===""?"":(isFinite(n)?n:(y.value??""))}:y)); }}/>
                        {isF&&<span style={{fontSize:12,color:bad?"#ef4444":editExpr===x.id?"#475569":"#94a3b8",whiteSpace:"nowrap",fontStyle:bad||editExpr===x.id?"normal":"italic"}} title={txt}>{bad?"check formula":editExpr===x.id?`= ${fmt(calcExpr(txt))}`:`ƒ ${txt.length>22?txt.slice(0,22)+"…":txt}`}</span>}</>; })()}
                    {x.type==="%"?<span style={{fontSize:12,color:"#475569",minWidth:90}}>% = S$ {fmt(extraAmt(item,x))}</span>:<span style={{fontSize:12,color:"#94a3b8"}}>S$</span>}
                    <button onClick={()=>setXs(xs.filter(y=>y.id!==x.id))} title="Remove" style={{background:"none",border:"none",color:"#ef4444",cursor:"pointer",fontSize:14}}>✕</button>
                  </div>)}
                </div>
              </td></tr>)}
              </>}
            </Fragment>);
          })}
        </tbody>
      </table>
    </div>
  );
  const projName=projects.find(p=>p.id===pid)?.name||"—";

  return(
    <div style={{minHeight:"100vh",background:"linear-gradient(180deg,#faf5ff 0%,#eff6ff 100%)",display:"flex",flexDirection:"column",fontFamily:"system-ui,sans-serif"}}>

      <header style={{background:"linear-gradient(90deg,#4f46e5 0%,#9333ea 50%,#db2777 100%)",color:"#fff",padding:"10px 16px",display:"flex",alignItems:"center",justifyContent:"space-between",boxShadow:"0 2px 10px rgba(124,58,237,.35)",position:"sticky",top:0,zIndex:200,gap:10,flexWrap:"wrap"}}>
        <div style={{display:"flex",alignItems:"center",gap:10}}>
          <span style={{fontSize:22}}>🏗️</span>
          <div><div style={{fontWeight:700,fontSize:14}}>QS Workspace</div>
          <div style={{fontSize:11,color:"#93c5fd"}}>🟢 Live · {tot} BOQ · {burItems.length} BUR</div></div>
        </div>
        <div style={{display:"flex",alignItems:"center",gap:8,flexWrap:"wrap"}}>
          <select value={pid||""} onChange={e=>setPid(e.target.value)} style={{background:"#16294a",color:"#fff",border:"1px solid #2d4a7a",borderRadius:8,padding:"5px 8px",fontSize:12,fontWeight:600,outline:"none",maxWidth:200}}>
            {projects.map(p=><option key={p.id} value={p.id}>{p.name}</option>)}
          </select>
          <button onClick={createProject} title="New project" style={{background:"#2563eb",color:"#fff",border:"none",borderRadius:8,padding:"5px 9px",fontSize:12,fontWeight:700,cursor:"pointer"}}>+</button>
          <button onClick={renameProject} title="Rename project" style={{background:"#16294a",color:"#cbd5e1",border:"1px solid #2d4a7a",borderRadius:8,padding:"5px 8px",fontSize:12,cursor:"pointer"}}>✎</button>
          <button onClick={deleteProject} title="Delete project" style={{background:"#16294a",color:"#fca5a5",border:"1px solid #2d4a7a",borderRadius:8,padding:"5px 8px",fontSize:12,cursor:"pointer"}}>🗑</button>
          <div style={{background:"#2563eb",borderRadius:20,padding:"4px 12px",fontSize:12,fontWeight:600}}>{user.name} · {user.role}</div>
          <button onClick={()=>signOut(auth)} title="Sign out" style={{background:"#16294a",color:"#cbd5e1",border:"1px solid #2d4a7a",borderRadius:8,padding:"4px 10px",fontSize:12,cursor:"pointer"}}>Sign out</button>
        </div>
      </header>

      {toast&&<div style={{background:"linear-gradient(90deg,#9333ea,#db2777)",color:"#fff",textAlign:"center",padding:"7px",fontSize:13,fontWeight:600}}>{toast}</div>}

      <div style={{background:"#fff",borderBottom:"1px solid #e2e8f0",display:"flex",overflowX:"auto",flexShrink:0,position:"sticky",top:44,zIndex:100}}>
        {TABS.map(t=><button key={t.id} onClick={()=>setTab(t.id)} style={{padding:"10px 14px",fontSize:12,fontWeight:700,border:"none",flexShrink:0,whiteSpace:"nowrap",borderBottom:tab===t.id?"3px solid #9333ea":"3px solid transparent",color:tab===t.id?"#9333ea":"#64748b",background:tab===t.id?"linear-gradient(180deg,rgba(147,51,234,.08),transparent)":"none",cursor:"pointer"}}>{t.label}</button>)}
      </div>

      <datalist id="extraNames">{[...new Set([...EXTRA_PRESETS.map(p=>p[0]),...burItems.flatMap(b=>(b.extras||[]).map(x=>String(x.name||"").trim()))].filter(Boolean))].map(n=><option key={n} value={n}/>)}</datalist>
      <datalist id="burlist">{burItems.filter(b=>b.code).map(b=><option key={b.id} value={b.code}>{b.desc} — S${fmt(bTot(b))}/{b.unit}</option>)}</datalist>

      {rateSugg&&(
        <div style={{background:"#fef9c3",borderBottom:"1px solid #fde68a",padding:"8px 16px",display:"flex",alignItems:"center",gap:12,flexWrap:"wrap",zIndex:90}}>
          <span style={{fontSize:13}}>💡 BUR rate found — <b style={{color:"#1d4ed8"}}>{rateSugg.code}</b>: {rateSugg.desc} = <b>S$ {fmt(rateSugg.rate)}</b></span>
          <div style={{display:"flex",gap:6,marginLeft:"auto"}}>
            <button onClick={()=>{updItem(rateSugg.secId,rateSugg.iid,{rA:rateSugg.rate});setRateSugg(null);toast_("✅ Applied to Rate A");}} style={{background:"#2563eb",color:"#fff",border:"none",borderRadius:6,padding:"4px 10px",fontSize:12,cursor:"pointer",fontWeight:600}}>→ Rate A</button>
            <button onClick={()=>{updItem(rateSugg.secId,rateSugg.iid,{rB:rateSugg.rate});setRateSugg(null);toast_("✅ Applied to Rate B");}} style={{background:"#7c3aed",color:"#fff",border:"none",borderRadius:6,padding:"4px 10px",fontSize:12,cursor:"pointer",fontWeight:600}}>→ Rate B</button>
            <button onClick={()=>setRateSugg(null)} style={{background:"#f1f5f9",border:"none",borderRadius:6,padding:"4px 8px",fontSize:12,cursor:"pointer",color:"#64748b"}}>✕</button>
          </div>
        </div>
      )}

      <div style={{flex:1,overflow:"auto",padding:12}}>

        {/* ══ BOQ — Master BQ rate feed (disabled; kept for export button) ══ */}
        {false&&(
          <div style={{maxWidth:760,margin:"0 auto",display:"flex",flexDirection:"column",gap:12}}>
            <div style={{background:"#fff",borderRadius:12,padding:20,boxShadow:"0 1px 4px rgba(0,0,0,.08)"}}>
              <div style={{fontWeight:700,fontSize:16,color:"#1e293b",marginBottom:6}}>📋 Master BQ — Rate Feed</div>
              <div style={{fontSize:13,color:"#475569",lineHeight:1.6,marginBottom:14}}>
                Keep building your BOQ in your own Excel master BQ (where all the formulas, colours and formatting already work). This app holds your <b>BUR rate library</b> and prices the BQ for you: it opens your template, and for every row it fills <b>Material &amp; Labour</b> (or U-Rate) by matching the <b>CODE in column A</b> to the BUR library — keeping your format and formulas 100% intact.
              </div>
              <div style={{display:"flex",gap:10,flexWrap:"wrap",alignItems:"center"}}>
                <button onClick={exportMasterBQ} style={{background:"#7c3aed",color:"#fff",border:"none",borderRadius:10,padding:"11px 20px",fontSize:14,fontWeight:700,cursor:"pointer"}}>⤓ Export priced Master BQ</button>
                <label style={{background:"#f1f5f9",border:"1px solid #e2e8f0",borderRadius:10,padding:"10px 16px",fontSize:13,fontWeight:600,cursor:"pointer",color:"#475569"}}>Update template…<input type="file" accept=".xlsx" style={{display:"none"}} onChange={e=>onTemplateFile(e.target.files[0])}/></label>
              </div>
              <div style={{fontSize:12,color:"#94a3b8",marginTop:10}}>Template in use: <b style={{color:"#475569"}}>{bqTplName||"CAG_Dormitory_Master_BQ.xlsx (built-in)"}</b> · {burItems.length} BUR codes available</div>
            </div>

            {bqResult&&(
              <div style={{background:"#fff",borderRadius:12,padding:18,boxShadow:"0 1px 4px rgba(0,0,0,.08)"}}>
                <div style={{display:"flex",gap:16,flexWrap:"wrap",marginBottom:bqResult.unmatched.length?12:0}}>
                  <div style={{background:"#dcfce7",borderRadius:10,padding:"12px 18px",textAlign:"center"}}><div style={{fontSize:24,fontWeight:800,color:"#15803d"}}>{bqResult.filled}</div><div style={{fontSize:11,color:"#15803d",fontWeight:600}}>rates filled</div></div>
                  <div style={{background:"#fef9c3",borderRadius:10,padding:"12px 18px",textAlign:"center"}}><div style={{fontSize:24,fontWeight:800,color:"#b45309"}}>{bqResult.unmatched.length}</div><div style={{fontSize:11,color:"#b45309",fontWeight:600}}>codes not in BUR</div></div>
                  <div style={{background:"#eff6ff",borderRadius:10,padding:"12px 18px",textAlign:"center"}}><div style={{fontSize:24,fontWeight:800,color:"#1d4ed8"}}>{bqResult.scanned}</div><div style={{fontSize:11,color:"#1d4ed8",fontWeight:600}}>priced rows scanned</div></div>
                </div>
                {bqResult.unmatched.length>0&&(
                  <div>
                    <div style={{fontSize:12,fontWeight:700,color:"#92400e",marginBottom:6}}>⚠️ Codes in the BQ with no matching BUR rate — add these to the BUR tab, then export again:</div>
                    <div style={{display:"flex",flexWrap:"wrap",gap:6}}>{bqResult.unmatched.map(c=><span key={c} style={{background:"#fffbeb",border:"1px solid #fde68a",borderRadius:6,padding:"3px 8px",fontSize:11,fontFamily:"monospace",color:"#92400e"}}>{c}</span>)}</div>
                  </div>
                )}
              </div>
            )}

            <div style={{background:"#fff",borderRadius:12,padding:16,boxShadow:"0 1px 4px rgba(0,0,0,.08)",fontSize:12,color:"#64748b",lineHeight:1.7}}>
              <b style={{color:"#1e293b"}}>How it works</b><br/>
              1. Maintain rates in the <b>BUR</b> tab (each item has a code, Material &amp; Labour).<br/>
              2. Your master BQ's column A holds those codes.<br/>
              3. Click <b>Export priced Master BQ</b> → download your template, fully formatted, with rates filled in.<br/>
              4. Need a different/updated BQ structure? Click <b>Update template</b> to upload a new .xlsx (used for this project).
            </div>
          </div>
        )}
        {tab==="boq"&&<div style={{display:"flex",gap:6,marginBottom:10,alignItems:"center",flexWrap:"wrap"}}>
          <span style={{fontSize:11,fontWeight:700,color:"#64748b"}}>BOQ mode:</span>
          {[["structured","📋 Structured (code→rate)"],["sheet","📄 Excel grid (drag your file)"]].map(([m,l])=><button key={m} onClick={()=>setBoqMode(m)} style={{border:"1px solid #e2e8f0",background:boqMode===m?"#7c3aed":"#fff",color:boqMode===m?"#fff":"#475569",borderRadius:6,padding:"4px 10px",fontSize:12,cursor:"pointer",fontWeight:600}}>{l}</button>)}
        </div>}
        {tab==="boq"&&boqMode==="sheet"&&<SheetGrid db={db} pid={pid} toast={toast_} baseUrl={import.meta.env.BASE_URL} onToStructured={importGridToStructured} burLookup={code=>{const b=findBurByCode(burItems,code); return b?{material:+b.material||0,labour:+b.labour||0,total:+bTot(b)||0}:null;}}/>}
        {tab==="boq"&&boqMode==="structured"&&(
          <div>
            <div style={{display:"flex",gap:8,marginBottom:12,overflowX:"auto",paddingBottom:4}}>
              {data.sections.map(sec=><button key={sec.id} onClick={()=>setSelSec(sec.id)} style={{padding:"6px 14px",borderRadius:20,fontSize:12,fontWeight:600,border:"none",cursor:"pointer",whiteSpace:"nowrap",flexShrink:0,background:selSec===sec.id?"#2563eb":"#fff",color:selSec===sec.id?"#fff":"#475569",boxShadow:"0 1px 4px rgba(0,0,0,.1)"}}>{sec.name}<span style={{opacity:.7,fontSize:11,marginLeft:4}}>({sec.items?.length||0})</span></button>)}
            </div>
            {cs&&(
              <div style={{background:"#fff",borderRadius:12,boxShadow:"0 1px 4px rgba(0,0,0,.08)",overflow:"hidden"}}>
                <div style={{padding:"10px 14px",borderBottom:"1px solid #f1f5f9",display:"flex",alignItems:"center",justifyContent:"space-between",background:"#f8fafc",flexWrap:"wrap",gap:8}}>
                  <span style={{fontWeight:700,fontSize:14,color:"#1e293b"}}>{cs.name}</span>
                  <div style={{display:"flex",gap:8}}>
                    <div style={{position:"relative"}}>
                      <button onClick={()=>setShowColPick(p=>!p)} style={{background:"#f1f5f9",border:"1px solid #e2e8f0",borderRadius:8,padding:"5px 10px",fontSize:11,fontWeight:600,cursor:"pointer",color:"#475569"}}>⚙ Columns</button>
                      {showColPick&&<div style={{position:"absolute",right:0,top:"110%",background:"#fff",border:"1px solid #e2e8f0",borderRadius:10,padding:10,minWidth:180,boxShadow:"0 4px 16px rgba(0,0,0,.12)",zIndex:300}}>
                        <div style={{fontSize:10,color:"#94a3b8",marginBottom:6,fontWeight:600}}>SHOW / HIDE COLUMNS</div>
                        {ALL_COLS.map(col=><label key={col.id} style={{display:"flex",alignItems:"center",gap:8,padding:"4px 2px",cursor:"pointer",fontSize:12}}><input type="checkbox" checked={visCols.has(col.id)} onChange={()=>toggleCol(col.id)} style={{accentColor:"#2563eb"}}/>{col.label}</label>)}
                      </div>}
                    </div>
                    <button onClick={exportMasterBQ} style={{background:"#7c3aed",color:"#fff",border:"none",borderRadius:8,padding:"6px 12px",fontSize:12,fontWeight:600,cursor:"pointer"}}>⤓ Export Master BQ</button>
                    <button onClick={exportBOQ} style={{background:"#15803d",color:"#fff",border:"none",borderRadius:8,padding:"6px 12px",fontSize:12,fontWeight:600,cursor:"pointer"}}>⤓ CSV</button>
                    <button onClick={addCol} style={{background:"#0d9488",color:"#fff",border:"none",borderRadius:8,padding:"6px 12px",fontSize:12,fontWeight:600,cursor:"pointer"}}>＋ Insert Column</button>
                    <button onClick={()=>addItem(cs.id)} style={{background:"#2563eb",color:"#fff",border:"none",borderRadius:8,padding:"6px 14px",fontSize:12,fontWeight:600,cursor:"pointer"}}>+ Add Row</button>
                  </div>
                </div>
                <div style={{overflowX:"auto"}}>
                  <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                    <thead><tr style={{background:"#f8fafc",color:"#94a3b8",fontSize:11}}>
                      {vcols.map(col=><th key={col.id} style={{padding:"8px 8px",textAlign:col.num||col.id==="by"?"right":col.id==="unit"?"center":"left",fontWeight:600,whiteSpace:"nowrap",borderBottom:"1px solid #e2e8f0",minWidth:col.w}}>{col.label}</th>)}
                      {(data.cols||[]).map(c=><th key={c.id} style={{padding:"8px 8px",textAlign:"right",fontWeight:600,whiteSpace:"nowrap",borderBottom:"1px solid #e2e8f0",minWidth:90,color:"#0f766e"}}><span onClick={()=>{const f=prompt(`Formula for "${c.label}"\n\nUse: Qty, Rate A, Rate B, Amt A, Amt B, and other column names.\nLeave blank for a manual-entry column.\nExample:  Qty * Rate A * markup`,c.formula||"");if(f!==null)setColFormula(c.id,f);}} title={c.formula?`= ${c.formula}  (click to edit)`:"click to set a formula"} style={{cursor:"pointer",textDecoration:"underline dotted"}}>{c.label}{c.formula?" ƒ":""}</span> <button onClick={()=>delCol(c.id)} style={{border:"none",background:"none",color:"#ef4444",cursor:"pointer",fontSize:11,fontWeight:700}}>✕</button></th>)}
                      <th style={{padding:"8px 6px",borderBottom:"1px solid #e2e8f0",width:28}}></th>
                    </tr></thead>
                    <tbody>
                      {!cs.items?.length?<tr><td colSpan={vcols.length+(data.cols?.length||0)+1} style={{padding:"32px",textAlign:"center",color:"#94a3b8",fontSize:13}}>No items — click + Add Row</td></tr>
                      :cs.items.map((item,ri)=>{ const cxv=computeCols(item,data.cols); return (
                        <tr key={item.id} style={{borderBottom:"1px solid #f8fafc"}}>
                          {vcols.map(col=>{
                            const p={padding:"4px 6px"};
                            if(col.id==="ref")return<td key={col.id} style={p}><input style={{width:44,border:"none",fontSize:12,outline:"none",background:"transparent"}} value={item.ref||""} onChange={e=>updItem(cs.id,item.id,{ref:e.target.value})} onBlur={()=>blurSave(cs.id,item.id,item.code)}/></td>;
                            if(col.id==="desc")return<td key={col.id} style={p}><input style={{width:col.w,border:"none",fontSize:12,outline:"none",background:"transparent"}} value={item.desc||""} onChange={e=>updItem(cs.id,item.id,{desc:e.target.value})} onBlur={()=>blurSave(cs.id,item.id,item.code)}/></td>;
                            if(col.id==="unit")return<td key={col.id} style={{...p,textAlign:"center"}}><select style={{border:"none",fontSize:11,outline:"none",background:"transparent"}} value={item.unit||"sum"} onChange={e=>updItem(cs.id,item.id,{unit:e.target.value})}>{unitOpts(item.unit).map(u=><option key={u}>{u}</option>)}</select></td>;
                            if(col.id==="qty")return<td key={col.id} style={{...p,textAlign:"right"}}><input type="number" style={{width:52,border:"none",fontSize:12,outline:"none",background:"transparent",textAlign:"right"}} value={item.qty??""} onChange={e=>updItem(cs.id,item.id,{qty:+e.target.value||0})} onBlur={()=>blurSave(cs.id,item.id,item.code)}/></td>;
                            if(col.id==="rA")return<td key={col.id} style={{...p,textAlign:"right"}}><input type="number" style={{width:68,border:"none",fontSize:12,outline:"none",background:"transparent",textAlign:"right"}} value={item.rA??""} onChange={e=>updItem(cs.id,item.id,{rA:+e.target.value||0})} onBlur={()=>blurSave(cs.id,item.id,item.code)}/></td>;
                            if(col.id==="amtA")return<td key={col.id} style={{...p,textAlign:"right",fontWeight:600,color:"#1d4ed8",whiteSpace:"nowrap"}}>{fmt(aA(item))}</td>;
                            if(col.id==="rB")return<td key={col.id} style={{...p,textAlign:"right"}}><input type="number" style={{width:68,border:"none",fontSize:12,outline:"none",background:"transparent",textAlign:"right"}} value={item.rB??""} onChange={e=>updItem(cs.id,item.id,{rB:+e.target.value||0})} onBlur={()=>blurSave(cs.id,item.id,item.code)}/></td>;
                            if(col.id==="amtB")return<td key={col.id} style={{...p,textAlign:"right",fontWeight:600,color:"#7c3aed",whiteSpace:"nowrap"}}>{fmt(aB(item))}</td>;
                            if(col.id==="code")return<td key={col.id} style={p}><input list="burlist" style={{width:90,border:"none",fontSize:11,outline:"none",background:"transparent"}} placeholder="e.g. CONC40" value={item.code||""} onChange={e=>updItem(cs.id,item.id,{code:e.target.value})} onBlur={e=>blurSave(cs.id,item.id,e.target.value)}/></td>;
                            if(col.id==="status")return<td key={col.id} style={{...p,textAlign:"center"}}><select style={{border:"none",fontSize:11,outline:"none",borderRadius:6,padding:"2px 4px",fontWeight:600,cursor:"pointer",background:SS[item.status||"Draft"].bg,color:SS[item.status||"Draft"].c}} value={item.status||"Draft"} onChange={e=>updItem(cs.id,item.id,{status:e.target.value})}>{STATUSES.map(s=><option key={s}>{s}</option>)}</select></td>;
                            if(col.id==="remarks")return<td key={col.id} style={p}><input style={{width:col.w,border:"none",fontSize:11,outline:"none",background:"transparent",color:"#64748b"}} placeholder="Remark…" value={item.remarks||""} onChange={e=>updItem(cs.id,item.id,{remarks:e.target.value})} onBlur={()=>blurSave(cs.id,item.id,item.code)}/></td>;
                            if(col.id==="by")return<td key={col.id} style={{...p,textAlign:"right",color:"#94a3b8",fontSize:10}}>{item.by||"—"}</td>;
                            return null;
                          })}
                          {(data.cols||[]).map(c=>c.formula
                            ?<td key={c.id} style={{padding:"4px 6px",textAlign:"right",color:"#0f766e",fontWeight:600,whiteSpace:"nowrap"}}>{typeof cxv[c.id]==="number"?fmt(cxv[c.id]):cxv[c.id]}</td>
                            :<td key={c.id} style={{padding:"4px 6px"}}><input style={{width:96,border:"none",fontSize:12,outline:"none",background:"transparent",textAlign:"right"}} placeholder={ri===0?"type value":""} value={item.cx?.[c.id]??""} onChange={e=>updItem(cs.id,item.id,{cx:{...(item.cx||{}),[c.id]:e.target.value}})} onBlur={()=>blurSave(cs.id,item.id,item.code)}/></td>)}
                          <td style={{padding:"4px 4px",textAlign:"center"}}><button onClick={()=>delItem(cs.id,item.id)} style={{border:"none",background:"none",color:"#ef4444",cursor:"pointer",fontSize:13,fontWeight:700}}>✕</button></td>
                        </tr>
                      );})}
                    </tbody>
                    {cs.items?.length>0&&<tfoot><tr style={{background:"#eff6ff",fontWeight:700}}>
                      {vcols.map(col=>{
                        if(col.id==="rA")return<td key={col.id} style={{padding:"8px",textAlign:"right",fontSize:12,color:"#475569"}}>Total →</td>;
                        if(col.id==="amtA")return<td key={col.id} style={{padding:"8px",textAlign:"right",fontSize:13,color:"#1d4ed8"}}>S$ {fmt(sTot(cs.id).A)}</td>;
                        if(col.id==="amtB")return<td key={col.id} style={{padding:"8px",textAlign:"right",fontSize:13,color:"#7c3aed"}}>S$ {fmt(sTot(cs.id).B)}</td>;
                        return<td key={col.id} style={{padding:"8px"}}></td>;
                      })}{(data.cols||[]).map(c=><td key={c.id} style={{padding:"8px"}}></td>)}<td></td>
                    </tr></tfoot>}
                  </table>
                </div>
              </div>
            )}
          </div>
        )}

        {/* ══ BUR ══ */}
        {tab==="bur"&&(
          <div style={{display:"flex",gap:12,alignItems:"flex-start"}}>
            {/* Left vertical category sidebar */}
            <div style={{width:208,flexShrink:0,background:"#fff",borderRadius:12,boxShadow:"0 1px 4px rgba(0,0,0,.08)",maxHeight:"80vh",overflow:"auto",padding:8}}>
              <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:"4px 6px"}}>
                <span style={{fontSize:10,fontWeight:700,color:"#94a3b8",letterSpacing:".5px"}}>CATEGORIES</span>
                <div style={{display:"flex",gap:3}}>
                  <button onClick={()=>setCatSort(s=>s==="az"?"none":"az")} title="Sort A → Z" style={{border:"none",borderRadius:5,padding:"2px 6px",fontSize:10,fontWeight:700,cursor:"pointer",background:catSort==="az"?"#7c3aed":"#f1f5f9",color:catSort==="az"?"#fff":"#64748b"}}>A↓Z</button>
                  <button onClick={()=>setCatSort(s=>s==="za"?"none":"za")} title="Sort Z → A" style={{border:"none",borderRadius:5,padding:"2px 6px",fontSize:10,fontWeight:700,cursor:"pointer",background:catSort==="za"?"#7c3aed":"#f1f5f9",color:catSort==="za"?"#fff":"#64748b"}}>Z↓A</button>
                </div>
              </div>
              {sidebarCats.map(c=>{const n=burItems.filter(b=>b.catId===c.id).length; const sel=selCat===c.id; return(
                <div key={c.id} onDragOver={e=>{e.preventDefault();e.currentTarget.style.outline="2px dashed #7c3aed";}} onDragLeave={e=>{e.currentTarget.style.outline="none";}} onDrop={e=>{e.currentTarget.style.outline="none";handleDropCat(c.id,e);}} style={{display:"flex",alignItems:"center",gap:2,marginBottom:2,borderRadius:8,background:sel?"#7c3aed":"transparent"}}>
                  <button onClick={()=>setSelCat(c.id)} style={{display:"flex",justifyContent:"space-between",alignItems:"center",gap:6,flex:1,minWidth:0,textAlign:"left",padding:"7px 10px",borderRadius:8,fontSize:12,fontWeight:600,border:"none",cursor:"pointer",background:"transparent",color:sel?"#fff":"#475569"}}>
                    <span style={{overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{c.name}</span><span style={{opacity:.7,fontSize:10,flexShrink:0}}>{n}</span>
                  </button>
                  <button onClick={()=>renameCat(c.id)} title="Rename" style={{border:"none",background:"none",cursor:"pointer",fontSize:11,padding:"2px 3px",color:sel?"#e9d5ff":"#94a3b8"}}>✎</button>
                  <button onClick={()=>delCat(c.id)} title="Delete (if empty)" style={{border:"none",background:"none",cursor:"pointer",fontSize:11,padding:"2px 5px 2px 2px",color:sel?"#fecaca":"#cbd5e1"}}>✕</button>
                </div>);})}
              {showNewCat?(
                <div style={{display:"flex",flexDirection:"column",gap:4,padding:"6px 4px"}}>
                  <input style={{border:"1px solid #e2e8f0",borderRadius:7,padding:"5px 9px",fontSize:12,outline:"none"}} placeholder="Category name" value={newCat} onChange={e=>setNewCat(e.target.value)} onKeyDown={e=>{if(e.key==="Enter"&&newCat.trim()){pushCats([...cats,{id:uid(),name:newCat.trim()}]);setNewCat("");setShowNewCat(false);}}}/>
                  <div style={{display:"flex",gap:4}}><button onClick={()=>{if(newCat.trim()){pushCats([...cats,{id:uid(),name:newCat.trim()}]);setNewCat("");setShowNewCat(false);}}} style={{flex:1,background:"#2563eb",color:"#fff",border:"none",borderRadius:7,padding:"5px",fontSize:11,cursor:"pointer",fontWeight:600}}>Add</button><button onClick={()=>setShowNewCat(false)} style={{background:"#f1f5f9",border:"none",borderRadius:7,padding:"5px 8px",fontSize:11,cursor:"pointer",color:"#64748b"}}>✕</button></div>
                </div>
              ):<button onClick={()=>setShowNewCat(true)} style={{width:"100%",marginTop:4,background:"#f1f5f9",border:"1px solid #e2e8f0",borderRadius:8,padding:"6px",fontSize:11,cursor:"pointer",color:"#475569",fontWeight:600}}>+ Category</button>}
            </div>

            {/* Right content */}
            <div style={{flex:1,minWidth:0}}>
              <div style={{display:"flex",gap:8,marginBottom:10,flexWrap:"wrap",alignItems:"center"}}>
                <input value={burSearch} onChange={e=>setBurSearch(e.target.value)} placeholder="🔍 Search code, description, or cost-data supplier/note…" style={{flex:1,minWidth:180,border:"1.5px solid #e2e8f0",borderRadius:8,padding:"7px 12px",fontSize:13,outline:"none"}}/>
                <label title="Limit the search to the category selected on the left" style={{display:"flex",alignItems:"center",gap:4,fontSize:12,color:"#475569",whiteSpace:"nowrap",cursor:"pointer"}}><input type="checkbox" checked={burSearchCat} onChange={e=>setBurSearchCat(e.target.checked)}/>Only in {catName(selCat)}</label>
                {burSearch&&<button onClick={()=>setBurSearch("")} style={{background:"#f1f5f9",border:"none",borderRadius:8,padding:"7px 10px",fontSize:12,cursor:"pointer",color:"#64748b"}}>Clear</button>}
                <button onClick={()=>{setPasteCat(selCat);setPasteOpen(true);}} style={{background:"#0ea5e9",color:"#fff",border:"none",borderRadius:8,padding:"7px 12px",fontSize:12,fontWeight:600,cursor:"pointer",whiteSpace:"nowrap"}}>📋 Paste from Excel</button>
                {codesWithSpaces.length>0&&<button onClick={fixCodeSpaces} title={"Codes with spaces or / can't be used in formulas:\n"+codesWithSpaces.slice(0,8).map(b=>b.code).join("\n")} style={{background:"#fff7ed",color:"#c2410c",border:"1px solid #fdba74",borderRadius:8,padding:"7px 12px",fontSize:12,fontWeight:600,cursor:"pointer",whiteSpace:"nowrap"}}>⚠ Fix {codesWithSpaces.length} code(s) with spaces / "/"</button>}
                {codesMissingUnit.length>0&&<button onClick={addUnitToCodes} title="Append _unit to BUR codes (e.g. Acoustic_feltCeiling → Acoustic_feltCeiling_m2)" style={{background:"#fff7ed",color:"#c2410c",border:"1px solid #fdba74",borderRadius:8,padding:"7px 12px",fontSize:12,fontWeight:600,cursor:"pointer",whiteSpace:"nowrap"}}>+ _unit to {codesMissingUnit.length} codes</button>}
                {shownSubKeys.length>0&&<><button onClick={()=>collapseAllSubs(true)} title="Fold every sub-category in this list" style={{background:"#fff",color:"#334155",border:"1px solid #cbd5e1",borderRadius:8,padding:"7px 10px",fontSize:12,fontWeight:600,cursor:"pointer",whiteSpace:"nowrap"}}>▸ Collapse all</button><button onClick={()=>collapseAllSubs(false)} title="Unfold every sub-category in this list" style={{background:"#fff",color:"#334155",border:"1px solid #cbd5e1",borderRadius:8,padding:"7px 10px",fontSize:12,fontWeight:600,cursor:"pointer",whiteSpace:"nowrap"}}>▾ Expand all</button></>}
                <button onClick={openSync} title="Copy BUR changes / new items from another project into this one" style={{background:"#ecfeff",color:"#0e7490",border:"1px solid #67e8f9",borderRadius:8,padding:"7px 10px",fontSize:12,fontWeight:600,cursor:"pointer",whiteSpace:"nowrap"}}>⇅ Sync from project…</button>
                <button onClick={replaceUnit} title="Change one unit to another on all BUR items (e.g. nr → no.)" style={{background:"#fff",color:"#334155",border:"1px solid #cbd5e1",borderRadius:8,padding:"7px 10px",fontSize:12,fontWeight:600,cursor:"pointer",whiteSpace:"nowrap"}}>⇄ Replace unit…</button>
                {undoLabel&&<button onClick={undoLast} title={"Undo: "+undoLabel+" (Ctrl+Z)"} style={{background:"#fff",color:"#334155",border:"1px solid #cbd5e1",borderRadius:8,padding:"7px 12px",fontSize:12,fontWeight:600,cursor:"pointer",whiteSpace:"nowrap",maxWidth:260,overflow:"hidden",textOverflow:"ellipsis"}}>↶ Undo: {undoLabel}</button>}
                <label title="Match items to a master BUR.xlsx by code and set their category + sub-category (underlined headings in column B). Codes, rates, quotes and photos are not changed." style={{background:"#f5f3ff",color:"#6d28d9",border:"1px solid #c4b5fd",borderRadius:8,padding:"7px 12px",fontSize:12,fontWeight:600,cursor:"pointer",whiteSpace:"nowrap"}}>☰ Sync categories from Excel…<input type="file" accept=".xlsx" style={{display:"none"}} onChange={e=>{const f=e.target.files[0]; e.target.value=""; subcatsFromExcel(f);}}/></label>
                <label title="Clear this project's BUR and load all items from a master BUR.xlsx (a backup downloads first)" style={{background:"#fef2f2",color:"#b91c1c",border:"1px solid #fca5a5",borderRadius:8,padding:"7px 12px",fontSize:12,fontWeight:600,cursor:"pointer",whiteSpace:"nowrap"}}>⟳ Replace BUR from Excel…<input type="file" accept=".xlsx" style={{display:"none"}} onChange={e=>{const f=e.target.files[0]; e.target.value=""; replaceBurFromExcel(f);}}/></label>
                <button onClick={exportBUR} style={{background:"#15803d",color:"#fff",border:"none",borderRadius:8,padding:"7px 12px",fontSize:12,fontWeight:600,cursor:"pointer",whiteSpace:"nowrap"}}>⤓ Export Excel</button>
              </div>

              <div style={{display:"flex",gap:6,marginBottom:8,fontSize:11,color:"#64748b",alignItems:"center",flexWrap:"wrap"}}>
                <span style={{fontWeight:700}}>Sort by:</span>
                {[["code","Code"],["desc","Description"],["unit","Unit"],["rate","Rate"],["cat","Category"]].map(([f,l])=>(
                  <button key={f} onClick={()=>toggleSort(f)} style={{border:"1px solid #e2e8f0",background:sortBy===f?"#1e3a5f":"#fff",color:sortBy===f?"#fff":"#475569",borderRadius:6,padding:"3px 9px",fontSize:11,cursor:"pointer",fontWeight:600}}>{l}{sortBy===f?(sortDir===1?" ▲":" ▼"):""}</button>
                ))}
                <span style={{marginLeft:8,color:"#cbd5e1"}}>· drag column edges to resize</span>
              </div>

              {selBur.size>0&&<div style={{display:"flex",gap:8,alignItems:"center",marginBottom:8,background:"#ede9fe",border:"1px solid #c4b5fd",borderRadius:8,padding:"6px 10px",flexWrap:"wrap"}}>
                <b style={{fontSize:12,color:"#5b21b6"}}>{selBur.size} selected</b>
                <span style={{fontSize:11,color:"#475569"}}>Sub-category:</span>
                <select value="" onChange={e=>{if(e.target.value)bulkSetSubcat(e.target.value);}} style={{border:"1px solid #c4b5fd",borderRadius:6,padding:"4px 8px",fontSize:12,outline:"none",background:"#fff",maxWidth:180}}><option value="">Sub-category…</option><option value="__new">＋ New…</option><option value="__none">(none)</option>{subcatsOf(selCat).map(g=><option key={g} value={g}>{g}</option>)}</select>
                <span style={{fontSize:11,color:"#475569"}}>Move to:</span>
                <select value="" onChange={e=>{if(e.target.value)bulkMoveCat(e.target.value);}} style={{border:"1px solid #c4b5fd",borderRadius:6,padding:"4px 8px",fontSize:12,outline:"none",background:"#fff"}}><option value="">Category…</option>{cats.map(c=><option key={c.id} value={c.id}>{c.name}</option>)}</select>
                <span style={{fontSize:11,color:"#475569"}}>Set unit:</span>
                <select value="" onChange={e=>{if(e.target.value)bulkSetUnit(e.target.value);}} style={{border:"1px solid #c4b5fd",borderRadius:6,padding:"4px 8px",fontSize:12,outline:"none",background:"#fff"}}><option value="">Unit…</option>{UNITS.map(u=><option key={u} value={u}>{u}</option>)}</select>
                <button onClick={bulkAddExtra} title="Add or update a lump-sum extra item (e.g. conc pump S$ 10) on all selected items" style={{background:"#0ea5e9",color:"#fff",border:"none",borderRadius:6,padding:"4px 12px",fontSize:12,cursor:"pointer",fontWeight:700}}>＋ Extra item…</button>
                <select value="none" onChange={e=>{if(e.target.value!=="none")bulkSetDiscount(e.target.value);}} style={{border:"1px solid #c4b5fd",borderRadius:6,padding:"4px 8px",fontSize:12,outline:"none",background:"#fff"}}><option value="none">Discount…</option>{[2,5,10,15].map(m=><option key={m} value={m}>{m}%</option>)}<option value="__custom">Other %…</option><option value="0">None (0%)</option></select>
                <select value="none" onChange={e=>{if(e.target.value!=="none")bulkSetWaste(e.target.value);}} style={{border:"1px solid #c4b5fd",borderRadius:6,padding:"4px 8px",fontSize:12,outline:"none",background:"#fff"}}><option value="none">Waste/ OH…</option>{[5,10,15,20].map(m=><option key={m} value={m}>{m}%</option>)}<option value="__custom">Other %…</option><option value="0">None (0%)</option></select>
                <select value="none" onChange={e=>{if(e.target.value!=="none")bulkSetRound(e.target.value);}} style={{border:"1px solid #c4b5fd",borderRadius:6,padding:"4px 8px",fontSize:12,outline:"none",background:"#fff"}}><option value="none">Round up…</option>{ROUND_OPTS.map(([v,l])=><option key={v} value={v}>{l}</option>)}</select>
                <button onClick={()=>copyItems([...selBur])} style={{background:"#2563eb",color:"#fff",border:"none",borderRadius:6,padding:"4px 12px",fontSize:12,cursor:"pointer",fontWeight:700}}>⧉ Copy {selBur.size}</button>
                <button onClick={bulkDelete} style={{background:"#dc2626",color:"#fff",border:"none",borderRadius:6,padding:"4px 12px",fontSize:12,cursor:"pointer",fontWeight:700}}>🗑 Delete {selBur.size}</button>
                <button onClick={clearSel} style={{marginLeft:"auto",background:"#fff",border:"1px solid #c4b5fd",borderRadius:6,padding:"4px 10px",fontSize:11,cursor:"pointer",color:"#5b21b6",fontWeight:600}}>Clear</button>
              </div>}
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10,flexWrap:"wrap",gap:6}}>
                <span style={{fontSize:13,fontWeight:600,color:"#1e293b"}}>{_q?(burSearchCat?`🔎 Search results in ${catName(selCat)}`:"🔎 Search results (all categories)"):catName(selCat)} <span style={{color:"#64748b",fontWeight:400,fontSize:12}}>— {catTotal} items{catItems.length<catTotal?` (showing first ${catItems.length} — refine search)`:""}</span>{dupCount>0&&<span style={{marginLeft:8,background:"#fee2e2",color:"#b91c1c",borderRadius:6,padding:"2px 8px",fontSize:11,fontWeight:700}}>⚠️ {dupCount} duplicate code{dupCount>1?"s":""}</span>}</span>
                <button onClick={()=>addBurItem(selCat)} style={{background:"#2563eb",color:"#fff",border:"none",borderRadius:8,padding:"6px 14px",fontSize:12,fontWeight:600,cursor:"pointer"}}>+ Add Item</button>
              </div>

              {catItems.length?burTableEl(catItems):burEmpty}
            </div>
          </div>
        )}

        {/* ══ RATES & CODES ══ */}
        {tab==="rates"&&(
          <div style={{display:"grid",gridTemplateColumns:"1fr 1.4fr",gap:12,alignItems:"start"}}>
            <div style={{background:"#fff",borderRadius:12,padding:16,boxShadow:"0 1px 4px rgba(0,0,0,.08)"}}>
              <h3 style={{fontSize:14,fontWeight:700,marginBottom:14,color:"#1e293b",marginTop:0}}>🔢 Rate Build-Up Calculator</h3>
              {[["labour","Labour"],["material","Material"],["plant","Plant & Equipment"],["subcon","Subcontractor"]].map(([k,l])=>(
                <div key={k} style={{display:"flex",alignItems:"center",gap:8,marginBottom:8}}>
                  <label style={{fontSize:12,color:"#475569",width:136,flexShrink:0}}>{l} (S$/unit)</label>
                  <input type="number" style={{flex:1,border:"1.5px solid #e2e8f0",borderRadius:7,padding:"6px 10px",fontSize:12,outline:"none"}} placeholder="0.00" value={rbu[k]} onChange={e=>setRbu(r=>({...r,[k]:e.target.value}))}/>
                </div>
              ))}
              <div style={{borderTop:"1px solid #f1f5f9",paddingTop:8,marginTop:4}}>
                {[["oh","Overhead (%)"],["profit","Profit (%)"]].map(([k,l])=>(
                  <div key={k} style={{display:"flex",alignItems:"center",gap:8,marginBottom:8}}>
                    <label style={{fontSize:12,color:"#475569",width:136,flexShrink:0}}>{l}</label>
                    <input type="number" style={{flex:1,border:"1.5px solid #e2e8f0",borderRadius:7,padding:"6px 10px",fontSize:12,outline:"none"}} value={rbu[k]} onChange={e=>setRbu(r=>({...r,[k]:e.target.value}))}/>
                  </div>
                ))}
              </div>
              <div style={{background:"#eff6ff",borderRadius:10,padding:14,marginTop:10}}>
                {[["Direct Cost",rb.d,false],[`Overhead (${rbu.oh}%)`,rb.oh,false],["Subtotal",rb.sub,true],[`Profit (${rbu.profit}%)`,rb.pr,false]].map(([l,v,b])=>(
                  <div key={l} style={{display:"flex",justifyContent:"space-between",fontSize:12,marginBottom:5,fontWeight:b?700:400,borderTop:b?"1px solid #bfdbfe":"none",paddingTop:b?7:0}}>
                    <span style={{color:"#475569"}}>{l}</span><span>S$ {fmt(v)}</span>
                  </div>
                ))}
                <div style={{display:"flex",justifyContent:"space-between",fontSize:15,fontWeight:800,borderTop:"2px solid #3b82f6",paddingTop:8,marginTop:3,color:"#1d4ed8"}}><span>TOTAL RATE</span><span>S$ {fmt(rb.total)}</span></div>
              </div>
              <button onClick={()=>setRbu({labour:"",material:"",plant:"",subcon:"",oh:15,profit:10})} style={{marginTop:10,width:"100%",background:"#f1f5f9",border:"none",borderRadius:8,padding:"7px",fontSize:12,color:"#64748b",cursor:"pointer",fontWeight:600}}>Clear</button>
            </div>
            <div style={{background:"#fff",borderRadius:12,padding:16,boxShadow:"0 1px 4px rgba(0,0,0,.08)"}}>
              <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:12}}>
                <h3 style={{fontSize:14,fontWeight:700,color:"#1e293b",margin:0}}>🏷️ Cost Codes</h3>
                <button onClick={()=>{setShowAddC(p=>!p);setEditCId(null);setCodeForm({code:"",desc:"",cat:""}); }} style={{background:"#2563eb",color:"#fff",border:"none",borderRadius:7,padding:"5px 10px",fontSize:11,fontWeight:600,cursor:"pointer"}}>+ Add</button>
              </div>
              {showAddC&&!editCId&&(
                <div style={{background:"#f8fafc",borderRadius:8,padding:10,marginBottom:10,border:"1px solid #e2e8f0"}}>
                  <div style={{display:"grid",gridTemplateColumns:"1fr 2fr 1fr",gap:6,marginBottom:8}}>
                    {[["code","Code *"],["desc","Description"],["cat","Category"]].map(([k,l])=>(
                      <div key={k}><div style={{fontSize:10,color:"#94a3b8",marginBottom:2}}>{l}</div><input style={{width:"100%",border:"1px solid #e2e8f0",borderRadius:6,padding:"5px 7px",fontSize:11,outline:"none"}} value={codeForm[k]} onChange={e=>setCodeForm(f=>({...f,[k]:e.target.value}))}/></div>
                    ))}
                  </div>
                  <div style={{display:"flex",gap:6,justifyContent:"flex-end"}}>
                    <button onClick={()=>setShowAddC(false)} style={{background:"#f1f5f9",border:"none",borderRadius:6,padding:"4px 10px",fontSize:11,cursor:"pointer",color:"#64748b"}}>Cancel</button>
                    <button onClick={saveNewCode} style={{background:"#2563eb",color:"#fff",border:"none",borderRadius:6,padding:"4px 10px",fontSize:11,cursor:"pointer",fontWeight:600}}>Save</button>
                  </div>
                </div>
              )}
              <div style={{overflowY:"auto",maxHeight:440}}>
                <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                  <thead><tr style={{color:"#94a3b8",fontSize:10,background:"#f8fafc"}}>{["Code","Description","Category",""].map((h,i)=><th key={i} style={{padding:"6px 8px",textAlign:"left",fontWeight:600,borderBottom:"1px solid #e2e8f0",width:i===3?52:i===0?80:i===2?70:undefined}}>{h}</th>)}</tr></thead>
                  <tbody>
                    {codes.map(c=>(
                      <tr key={c.id} style={{borderBottom:"1px solid #f8fafc"}}>
                        {editCId===c.id?(
                          <>{[["code",70],["desc",null],["cat",70]].map(([k,w])=><td key={k} style={{padding:"3px 4px"}}><input style={{width:w||"100%",border:"1px solid #93c5fd",borderRadius:5,padding:"3px 5px",fontSize:11,outline:"none",boxSizing:"border-box"}} value={codeForm[k]} onChange={e=>setCodeForm(f=>({...f,[k]:e.target.value}))}/></td>)}
                          <td style={{padding:"3px 4px"}}><div style={{display:"flex",gap:3}}><button onClick={saveEditCode} style={{background:"#2563eb",color:"#fff",border:"none",borderRadius:4,padding:"2px 6px",fontSize:10,cursor:"pointer"}}>✓</button><button onClick={()=>setEditCId(null)} style={{background:"#f1f5f9",border:"none",borderRadius:4,padding:"2px 5px",fontSize:10,cursor:"pointer"}}>✕</button></div></td></>
                        ):(
                          <><td style={{padding:"8px 8px",fontWeight:700,color:"#1e293b",fontFamily:"monospace"}}>{c.code}</td><td style={{padding:"8px 8px",color:"#475569",fontSize:11}}>{c.desc}</td><td style={{padding:"8px 8px",color:"#94a3b8",fontSize:11}}>{c.cat}</td>
                          <td style={{padding:"4px"}}><div style={{display:"flex",gap:3,justifyContent:"center"}}><button onClick={()=>{setEditCId(c.id);setCodeForm({code:c.code,desc:c.desc,cat:c.cat});setShowAddC(false);}} style={{background:"none",border:"none",cursor:"pointer",fontSize:12,color:"#3b82f6",padding:"2px"}}>✏</button><button onClick={()=>delCode(c.id)} style={{background:"none",border:"none",cursor:"pointer",fontSize:12,color:"#ef4444",padding:"2px"}}>✕</button></div></td></>
                        )}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        )}

        {/* ══ COMPARE QUOTES ══ */}
        {tab==="compare"&&(
          <div>
            <div style={{display:"flex",gap:6,marginBottom:12,alignItems:"center",flexWrap:"wrap"}}>
              <div style={{display:"flex",gap:6,overflowX:"auto",flex:1,paddingBottom:2}}>
                {comparisons.map(c=><button key={c.id} onClick={()=>setSelCompId(c.id)} style={{padding:"6px 12px",borderRadius:20,fontSize:12,fontWeight:600,border:"none",cursor:"pointer",whiteSpace:"nowrap",flexShrink:0,background:selCompId===c.id?"#1e3a5f":"#fff",color:selCompId===c.id?"#fff":"#475569",boxShadow:"0 1px 3px rgba(0,0,0,.1)"}}>{c.trade||c.project||"Untitled comparison"}</button>)}
              </div>
              <button onClick={addComparison} style={{background:"#f1f5f9",border:"1px solid #e2e8f0",borderRadius:20,padding:"5px 12px",fontSize:11,cursor:"pointer",color:"#475569",fontWeight:600,whiteSpace:"nowrap",flexShrink:0}}>+ New Comparison</button>
            </div>

            {!selComp?(
              <div style={{background:"#fff",borderRadius:12,padding:40,textAlign:"center",color:"#94a3b8"}}>
                <div style={{fontSize:36,marginBottom:8}}>🆚</div>
                <div style={{fontWeight:600,fontSize:14,marginBottom:4}}>No comparison yet</div>
                <div style={{fontSize:13}}>Click "+ New Comparison" to compare subcontractor quotes for a trade</div>
              </div>
            ):(
              <div style={{display:"flex",flexDirection:"column",gap:12}}>

                <div style={{background:"#fff",borderRadius:12,padding:14,boxShadow:"0 1px 4px rgba(0,0,0,.08)",display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(180px,1fr))",gap:10}}>
                  {[["project","Project"],["trade","Trade / Package"],["reference","Reference"],["preparedDate","Prepared Date"]].map(([k,l])=>(
                    <div key={k}><div style={{fontSize:10,color:"#94a3b8",marginBottom:3,fontWeight:600}}>{l.toUpperCase()}</div>
                    <input style={{width:"100%",border:"1.5px solid #e2e8f0",borderRadius:7,padding:"7px 9px",fontSize:13,outline:"none",boxSizing:"border-box"}} value={selComp[k]} onChange={e=>updComp({[k]:e.target.value})}/></div>
                  ))}
                  <div style={{gridColumn:"1/-1",textAlign:"right"}}>
                    <button onClick={()=>delComparison(selComp.id)} style={{background:"none",border:"none",color:"#ef4444",cursor:"pointer",fontSize:12,fontWeight:600}}>🗑 Delete this comparison</button>
                  </div>
                </div>

                <div style={{background:"#fff",borderRadius:12,padding:14,boxShadow:"0 1px 4px rgba(0,0,0,.08)"}}>
                  <div style={{fontSize:12,fontWeight:700,color:"#1e293b",marginBottom:8}}>Subcontractors Bidding</div>
                  <div style={{display:"flex",gap:8,flexWrap:"wrap",alignItems:"center"}}>
                    {selComp.bidders.map(name=>(
                      <span key={name} style={{background:"#eff6ff",color:"#1d4ed8",borderRadius:20,padding:"5px 8px 5px 12px",fontSize:12,fontWeight:600,display:"flex",alignItems:"center",gap:6}}>
                        {name}<button onClick={()=>removeBidder(name)} style={{background:"none",border:"none",color:"#1d4ed8",cursor:"pointer",fontWeight:700,fontSize:12}}>✕</button>
                      </span>
                    ))}
                    <input style={{border:"1.5px solid #e2e8f0",borderRadius:20,padding:"5px 12px",fontSize:12,outline:"none",width:160}} placeholder="e.g. Prefab Tech 3" value={newBidderName} onChange={e=>setNewBidderName(e.target.value)} onKeyDown={e=>e.key==="Enter"&&addBidder()}/>
                    <button onClick={addBidder} style={{background:"#2563eb",color:"#fff",border:"none",borderRadius:20,padding:"6px 14px",fontSize:12,fontWeight:600,cursor:"pointer"}}>+ Add</button>
                  </div>
                  {selComp.bidders.length===0&&<div style={{fontSize:12,color:"#94a3b8",marginTop:8}}>Add at least one subcontractor before adding items.</div>}
                </div>

                {selComp.bidders.length>0&&(()=>{
                  const flagsBySection={}; selComp.sections.forEach(s=>{flagsBySection[s.id]=cmpAutoFlags(s,selComp.bidders);});
                  return (
                  <>
                  {selComp.sections.map(section=>(
                    <div key={section.id} style={{background:"#fff",borderRadius:12,boxShadow:"0 1px 4px rgba(0,0,0,.08)",overflow:"hidden"}}>
                      <div style={{padding:"10px 14px",borderBottom:"1px solid #f1f5f9",background:"#f8fafc",display:"flex",alignItems:"center",gap:8}}>
                        <input style={{flex:1,border:"none",background:"none",fontWeight:700,fontSize:13,color:"#1e293b",outline:"none"}} value={section.name} onChange={e=>renameSection(section.id,e.target.value)}/>
                        <button onClick={()=>addCmpItem(section.id)} style={{background:"#2563eb",color:"#fff",border:"none",borderRadius:7,padding:"5px 12px",fontSize:11,fontWeight:600,cursor:"pointer"}}>+ Item</button>
                        <button onClick={()=>delSection(section.id)} style={{background:"none",border:"none",color:"#ef4444",cursor:"pointer",fontSize:12}}>🗑</button>
                      </div>
                      <div style={{overflowX:"auto"}}>
                        <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                          <thead>
                            <tr style={{background:"#f8fafc",color:"#94a3b8",fontSize:10}}>
                              <th style={{padding:"6px",textAlign:"left",minWidth:36}}>No</th>
                              <th style={{padding:"6px",textAlign:"left",minWidth:200}}>Description</th>
                              <th style={{padding:"6px",textAlign:"center",minWidth:56}}>Unit</th>
                              <th style={{padding:"6px",textAlign:"right",minWidth:70}}>WH Qty</th>
                              {selComp.bidders.map(name=><th key={name} colSpan={4} style={{padding:"6px",textAlign:"center",borderLeft:"2px solid #e2e8f0",minWidth:280}}>{name}</th>)}
                              <th style={{width:28}}></th>
                            </tr>
                            <tr style={{background:"#f8fafc",color:"#cbd5e1",fontSize:9}}>
                              <th colSpan={4}></th>
                              {selComp.bidders.map(name=><Fragment key={name}>
                                <th style={{padding:"2px 6px",borderLeft:"2px solid #e2e8f0",textAlign:"right"}}>Qty</th>
                                <th style={{padding:"2px 6px",textAlign:"right"}}>Rate</th>
                                <th style={{padding:"2px 6px",textAlign:"right"}}>Amount</th>
                                <th style={{padding:"2px 6px",textAlign:"left"}}>Status / Remarks</th>
                              </Fragment>)}
                              <th></th>
                            </tr>
                          </thead>
                          <tbody>
                            {section.items.length===0?<tr><td colSpan={4+selComp.bidders.length*4+1} style={{padding:20,textAlign:"center",color:"#94a3b8"}}>No items yet</td></tr>:
                            section.items.map(item=>{
                              const amounts=selComp.bidders.map(name=>({name,amt:cmpAmount(item,name)}));
                              const priced=amounts.filter(a=>a.amt!=null);
                              const minAmt=priced.length?Math.min(...priced.map(a=>a.amt)):null;
                              const f=flagsBySection[section.id][item.id]||{};
                              return (
                                <tr key={item.id} style={{borderBottom:"1px solid #f8fafc"}}>
                                  <td style={{padding:"4px 6px"}}><input style={{width:28,border:"none",fontSize:11,outline:"none",background:"transparent"}} value={item.no} onChange={e=>updCmpItem(section.id,item.id,{no:e.target.value})}/></td>
                                  <td style={{padding:"4px 6px"}}><input style={{width:"100%",minWidth:190,border:"none",fontSize:12,outline:"none",background:"transparent"}} value={item.desc} onChange={e=>updCmpItem(section.id,item.id,{desc:e.target.value})}/></td>
                                  <td style={{padding:"4px 6px",textAlign:"center"}}><input style={{width:48,border:"none",fontSize:11,outline:"none",background:"transparent",textAlign:"center"}} value={item.unit} onChange={e=>updCmpItem(section.id,item.id,{unit:e.target.value})}/></td>
                                  <td style={{padding:"4px 6px",textAlign:"right"}}><input type="number" style={{width:62,border:"none",fontSize:12,outline:"none",background:"transparent",textAlign:"right"}} value={item.whQty} onChange={e=>updCmpItem(section.id,item.id,{whQty:e.target.value})}/></td>
                                  {selComp.bidders.map(name=>{
                                    const b=cmpBidderCell(item,name); const amt=amounts.find(a=>a.name===name).amt;
                                    const isLowest=amt!=null&&minAmt!=null&&amt===minAmt;
                                    const bf=f[name]||{qty:false,rate:false};
                                    return (
                                      <Fragment key={name}>
                                        <td style={{padding:"4px 6px",borderLeft:"2px solid #f8fafc",background:bf.qty?"#ffe5a3":undefined,textAlign:"right"}}>
                                          <input type="number" style={{width:56,border:"none",fontSize:11,outline:"none",background:"transparent",textAlign:"right"}} value={b.qty} onChange={e=>updCmpBidderCell(section.id,item.id,name,{qty:e.target.value})}/>
                                        </td>
                                        <td style={{padding:"4px 6px",background:bf.rate?"#ffe5a3":undefined,textAlign:"right"}}>
                                          <input type="number" style={{width:66,border:"none",fontSize:11,outline:"none",background:"transparent",textAlign:"right"}} value={b.rate} onChange={e=>updCmpBidderCell(section.id,item.id,name,{rate:e.target.value})}/>
                                        </td>
                                        <td style={{padding:"4px 6px",textAlign:"right",fontWeight:isLowest?800:400,color:isLowest?"#15803d":b.status==="norate"?"#991b1b":"#1e293b",background:isLowest?"#dcfce7":b.status==="norate"?"#fee2e2":undefined,whiteSpace:"nowrap"}}>
                                          {amt!=null?fmt(amt):CMP_STATUS_LABEL[b.status]}
                                        </td>
                                        <td style={{padding:"4px 6px"}}>
                                          <div style={{display:"flex",gap:4,alignItems:"center"}}>
                                            <select style={{border:"none",fontSize:10,outline:"none",background:"transparent",color:"#64748b"}} value={b.status} onChange={e=>updCmpBidderCell(section.id,item.id,name,{status:e.target.value})}>
                                              {CMP_STATUSES.map(s=><option key={s} value={s}>{CMP_STATUS_LABEL[s]}</option>)}
                                            </select>
                                            <input style={{flex:1,minWidth:90,border:"none",fontSize:10,outline:"none",background:"transparent",color:"#94a3b8"}} placeholder="remarks" value={b.remarks} onChange={e=>updCmpBidderCell(section.id,item.id,name,{remarks:e.target.value})}/>
                                          </div>
                                        </td>
                                      </Fragment>
                                    );
                                  })}
                                  <td style={{padding:"4px",textAlign:"center"}}><button onClick={()=>delCmpItem(section.id,item.id)} style={{border:"none",background:"none",color:"#ef4444",cursor:"pointer",fontSize:12,fontWeight:700}}>✕</button></td>
                                </tr>
                              );
                            })}
                          </tbody>
                          {section.items.length>0&&<tfoot><tr style={{background:"#eff6ff",fontWeight:700}}>
                            <td colSpan={4} style={{padding:"8px",textAlign:"right",fontSize:11,color:"#475569"}}>Subtotal →</td>
                            {selComp.bidders.map(name=>{
                              const sub=section.items.reduce((t,i)=>t+(cmpAmount(i,name)||0),0);
                              return <td key={name} colSpan={4} style={{padding:"8px",textAlign:"right",fontSize:12,color:"#1d4ed8",borderLeft:"2px solid #dbeafe"}}>{fmt(sub)}</td>;
                            })}
                            <td></td>
                          </tr></tfoot>}
                        </table>
                      </div>
                    </div>
                  ))}
                  <button onClick={addSection} style={{background:"#f1f5f9",border:"1px dashed #cbd5e1",borderRadius:10,padding:"10px",fontSize:12,fontWeight:600,cursor:"pointer",color:"#475569"}}>+ Add Section</button>
                  </>
                  );
                })()}

                {selComp.bidders.length>0&&selComp.sections.length>0&&(
                  <div style={{background:"#fff",borderRadius:12,boxShadow:"0 1px 4px rgba(0,0,0,.08)",overflow:"hidden"}}>
                    <div style={{padding:"12px 16px",borderBottom:"1px solid #f1f5f9",background:"#f8fafc",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                      <span style={{fontWeight:700,fontSize:14}}>📊 Ranking</span>
                      <button onClick={()=>exportComparisonCSV(selComp)} style={{background:"#2563eb",color:"#fff",border:"none",borderRadius:7,padding:"6px 14px",fontSize:11,fontWeight:600,cursor:"pointer"}}>⬇ Export CSV</button>
                    </div>
                    <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
                      <thead><tr style={{background:"#f8fafc",color:"#94a3b8",fontSize:11}}>{["Rank","Subcontractor","Total Bid Amount (WH Qty)","Unpriced Items"].map((h,i)=><th key={i} style={{padding:"8px 16px",textAlign:i>1?"right":"left",fontWeight:600,borderBottom:"1px solid #e2e8f0"}}>{h}</th>)}</tr></thead>
                      <tbody>
                        {selComp.bidders.map(name=>({name,total:cmpBidderTotal(selComp,name),unpriced:cmpUnpricedCount(selComp,name)}))
                          .sort((a,b)=>a.total-b.total)
                          .map((row,i)=>(
                            <tr key={row.name} style={{borderBottom:"1px solid #f8fafc",background:i===0?"#f0fdf4":undefined}}>
                              <td style={{padding:"10px 16px",fontWeight:800,color:i===0?"#15803d":"#94a3b8"}}>{i+1}</td>
                              <td style={{padding:"10px 16px",fontWeight:600}}>{row.name}</td>
                              <td style={{padding:"10px 16px",textAlign:"right",fontWeight:700,color:i===0?"#15803d":"#1e293b"}}>S$ {fmt(row.total)}</td>
                              <td style={{padding:"10px 16px",textAlign:"right",color:row.unpriced>0?"#b45309":"#94a3b8"}}>{row.unpriced>0?`⚠ ${row.unpriced}`:"—"}</td>
                            </tr>
                          ))}
                      </tbody>
                    </table>
                    <div style={{padding:"8px 16px",fontSize:11,color:"#94a3b8",borderTop:"1px solid #f1f5f9"}}>Amounts are always computed on this project's own required quantity (WH Qty), not the bidder's quoted quantity — that's what makes bids comparable. Lowest total isn't automatically the right choice — check the flags below first.</div>
                  </div>
                )}

                {selComp.bidders.length>0&&(
                  <div style={{background:"#fff",borderRadius:12,boxShadow:"0 1px 4px rgba(0,0,0,.08)",overflow:"hidden"}}>
                    <div style={{padding:"12px 16px",borderBottom:"1px solid #f1f5f9",background:"#f8fafc"}}><span style={{fontWeight:700,fontSize:14}}>🚩 Risk Flags</span></div>

                    {(()=>{
                      const auto=[];
                      selComp.sections.forEach(s=>{const fl=cmpAutoFlags(s,selComp.bidders);s.items.forEach(item=>{selComp.bidders.forEach(name=>{const f=fl[item.id]?.[name];if(f?.qty)auto.push({type:"Quantity discrepancy",item:item.no||item.desc,bidder:name,detail:`Quoted qty differs from WH Qty by more than ${QTY_MISMATCH_PCT*100}%.`,severity:"Medium"});if(f?.rate)auto.push({type:"Outlier rate",item:item.no||item.desc,bidder:name,detail:`Rate differs from this item's median priced rate by more than ${OUTLIER_PCT*100}%.`,severity:"High"});});});});
                      return auto.length>0?(
                        <div style={{padding:"10px 16px",borderBottom:"1px solid #f1f5f9"}}>
                          <div style={{fontSize:11,fontWeight:700,color:"#94a3b8",marginBottom:6}}>AUTO-DETECTED ({auto.length})</div>
                          <div style={{display:"flex",flexDirection:"column",gap:4}}>
                            {auto.map((f,i)=>(<div key={i} style={{display:"flex",gap:8,alignItems:"center",fontSize:12}}>
                              <span style={{background:SEV_COLOR[f.severity].bg,color:SEV_COLOR[f.severity].c,borderRadius:5,padding:"1px 7px",fontWeight:700,fontSize:10,flexShrink:0}}>{f.severity}</span>
                              <span style={{color:"#475569"}}><b>{f.bidder}</b> · Item {f.item} · {f.type} — {f.detail}</span>
                            </div>))}
                          </div>
                        </div>
                      ):null;
                    })()}

                    <div style={{padding:"10px 16px"}}>
                      <div style={{fontSize:11,fontWeight:700,color:"#94a3b8",marginBottom:6}}>MANUAL NOTES ({selComp.riskFlags.length})</div>
                      {selComp.riskFlags.length===0&&<div style={{fontSize:12,color:"#94a3b8",marginBottom:8}}>No manual flags added — use this for things a formula can't catch, like ambiguous scope wording or a weak warranty.</div>}
                      <div style={{display:"flex",flexDirection:"column",gap:6,marginBottom:10}}>
                        {selComp.riskFlags.map(f=>(
                          <div key={f.id} style={{display:"flex",gap:8,alignItems:"flex-start",fontSize:12,background:"#f8fafc",borderRadius:8,padding:"7px 10px"}}>
                            <span style={{background:SEV_COLOR[f.severity].bg,color:SEV_COLOR[f.severity].c,borderRadius:5,padding:"1px 7px",fontWeight:700,fontSize:10,flexShrink:0}}>{f.severity}</span>
                            <span style={{flex:1,color:"#475569"}}><b>{f.type}</b>{(f.item||f.bidder)&&` — ${[f.item&&`Item ${f.item}`,f.bidder].filter(Boolean).join(" / ")}`}: {f.detail}</span>
                            <button onClick={()=>delRiskFlag(f.id)} style={{border:"none",background:"none",color:"#ef4444",cursor:"pointer",fontSize:12,flexShrink:0}}>✕</button>
                          </div>
                        ))}
                      </div>
                      <div style={{display:"grid",gridTemplateColumns:"90px 1fr 1fr 2fr auto",gap:6,alignItems:"start"}}>
                        <select style={{border:"1px solid #e2e8f0",borderRadius:6,padding:"6px",fontSize:11,outline:"none"}} value={newFlag.severity} onChange={e=>setNewFlag(f=>({...f,severity:e.target.value}))}>{["High","Medium","Low"].map(s=><option key={s}>{s}</option>)}</select>
                        <input style={{border:"1px solid #e2e8f0",borderRadius:6,padding:"6px 8px",fontSize:11,outline:"none"}} placeholder="Type e.g. Ambiguous scope" value={newFlag.type} onChange={e=>setNewFlag(f=>({...f,type:e.target.value}))}/>
                        <input style={{border:"1px solid #e2e8f0",borderRadius:6,padding:"6px 8px",fontSize:11,outline:"none"}} placeholder="Item no / bidder" value={newFlag.item} onChange={e=>setNewFlag(f=>({...f,item:e.target.value}))}/>
                        <input style={{border:"1px solid #e2e8f0",borderRadius:6,padding:"6px 8px",fontSize:11,outline:"none"}} placeholder="Detail" value={newFlag.detail} onChange={e=>setNewFlag(f=>({...f,detail:e.target.value}))}/>
                        <button onClick={addRiskFlag} style={{background:"#2563eb",color:"#fff",border:"none",borderRadius:6,padding:"6px 12px",fontSize:11,fontWeight:600,cursor:"pointer"}}>+ Add</button>
                      </div>
                    </div>
                  </div>
                )}

                {selComp.bidders.length>0&&(
                  <div style={{background:"#fff",borderRadius:12,boxShadow:"0 1px 4px rgba(0,0,0,.08)",overflow:"hidden"}}>
                    <div style={{padding:"12px 16px",borderBottom:"1px solid #f1f5f9",background:"#f8fafc"}}><span style={{fontWeight:700,fontSize:14}}>📄 Commercial Terms</span></div>
                    <div style={{overflowX:"auto"}}>
                      <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                        <thead><tr style={{background:"#f8fafc",color:"#94a3b8",fontSize:11}}>{["Subcontractor","Validity","Payment Terms","Lead Time","Warranty","Notes"].map((h,i)=><th key={i} style={{padding:"8px 10px",textAlign:"left",fontWeight:600,borderBottom:"1px solid #e2e8f0",minWidth:i===0?110:130}}>{h}</th>)}</tr></thead>
                        <tbody>
                          {selComp.bidders.map(name=>{const t=selComp.commercialTerms[name]||{};return(
                            <tr key={name} style={{borderBottom:"1px solid #f8fafc"}}>
                              <td style={{padding:"6px 10px",fontWeight:600}}>{name}</td>
                              {["validity","paymentTerms","leadTime","warranty","notes"].map(k=>(
                                <td key={k} style={{padding:"4px 6px"}}><input style={{width:"100%",border:"none",fontSize:11,outline:"none",background:"transparent",boxSizing:"border-box"}} value={t[k]||""} onChange={e=>updCommercialTerm(name,{[k]:e.target.value})}/></td>
                              ))}
                            </tr>
                          );})}
                        </tbody>
                      </table>
                    </div>
                  </div>
                )}
              </div>
            )}
          </div>
        )}

        {/* ══ SUMMARY ══ */}
        {tab==="summary"&&(
          <div style={{maxWidth:680,margin:"0 auto",display:"flex",flexDirection:"column",gap:12}}>
            <div style={{background:"#fff",borderRadius:12,boxShadow:"0 1px 4px rgba(0,0,0,.08)",overflow:"hidden"}}>
              <div style={{padding:"12px 16px",borderBottom:"1px solid #f1f5f9",background:"#f8fafc"}}><span style={{fontWeight:700,fontSize:14}}>📊 {projName} — Tender Summary</span></div>
              <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
                <thead><tr style={{background:"#f8fafc",color:"#94a3b8",fontSize:11}}>{["Section","Items","Phase A (S$)","Phase B (S$)","Combined (S$)"].map((h,i)=><th key={i} style={{padding:"8px 16px",textAlign:i>1?"right":"left",fontWeight:600,borderBottom:"1px solid #e2e8f0"}}>{h}</th>)}</tr></thead>
                <tbody>{data.sections.map(sec=>{const t=sTot(sec.id);return(<tr key={sec.id} style={{borderBottom:"1px solid #f8fafc"}}><td style={{padding:"11px 16px",fontWeight:600}}>{sec.name}</td><td style={{padding:"11px 16px",color:"#64748b"}}>{sec.items?.length||0}</td><td style={{padding:"11px 16px",textAlign:"right",color:"#1d4ed8"}}>{fmt(t.A)}</td><td style={{padding:"11px 16px",textAlign:"right",color:"#7c3aed"}}>{fmt(t.B)}</td><td style={{padding:"11px 16px",textAlign:"right",fontWeight:700}}>{fmt(t.A+t.B)}</td></tr>);})}</tbody>
                <tfoot><tr style={{background:"#7c3aed",color:"#fff",fontWeight:700,fontSize:14}}><td style={{padding:"12px 16px"}}>GRAND TOTAL</td><td style={{padding:"12px 16px"}}>{tot}</td><td style={{padding:"12px 16px",textAlign:"right"}}>S$ {fmt(gt.A)}</td><td style={{padding:"12px 16px",textAlign:"right"}}>S$ {fmt(gt.B)}</td><td style={{padding:"12px 16px",textAlign:"right"}}>S$ {fmt(gt.A+gt.B)}</td></tr></tfoot>
              </table>
            </div>
            <div style={{background:"#fff",borderRadius:12,boxShadow:"0 1px 4px rgba(0,0,0,.08)",overflow:"hidden"}}>
              <div style={{padding:"12px 16px",borderBottom:"1px solid #f1f5f9",background:"#f8fafc"}}><span style={{fontWeight:700,fontSize:14}}>🏷️ By Cost Code</span></div>
              <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
                <thead><tr style={{background:"#f8fafc",color:"#94a3b8",fontSize:11}}>{["Code","Items","Phase A","Phase B","Total"].map((h,i)=><th key={i} style={{padding:"8px 16px",textAlign:i<2?"left":"right",fontWeight:600,borderBottom:"1px solid #e2e8f0"}}>{h}</th>)}</tr></thead>
                <tbody>{Object.keys(cc).length===0?<tr><td colSpan={5} style={{padding:24,textAlign:"center",color:"#94a3b8"}}>No BUR codes assigned yet</td></tr>:Object.entries(cc).sort((a,b)=>a[0].localeCompare(b[0])).map(([code,v])=><tr key={code} style={{borderBottom:"1px solid #f8fafc"}}><td style={{padding:"9px 16px",fontWeight:700,fontFamily:"monospace",color:"#1d4ed8"}}>{code}</td><td style={{padding:"9px 16px",color:"#64748b"}}>{v.n}</td><td style={{padding:"9px 16px",textAlign:"right",color:"#1d4ed8"}}>{fmt(v.A)}</td><td style={{padding:"9px 16px",textAlign:"right",color:"#7c3aed"}}>{fmt(v.B)}</td><td style={{padding:"9px 16px",textAlign:"right",fontWeight:700}}>{fmt(v.A+v.B)}</td></tr>)}</tbody>
              </table>
            </div>
          </div>
        )}

        {/* ══ ACTIVITY ══ */}
        {tab==="log"&&(
          <div style={{background:"#fff",borderRadius:12,boxShadow:"0 1px 4px rgba(0,0,0,.08)",overflow:"hidden",maxWidth:600,margin:"0 auto"}}>
            <div style={{padding:"12px 16px",borderBottom:"1px solid #f1f5f9",background:"#f8fafc",display:"flex",justifyContent:"space-between",alignItems:"center"}}><span style={{fontWeight:700,fontSize:14}}>📝 Activity Log</span><span style={{fontSize:11,color:"#94a3b8"}}>{log.length} entries</span></div>
            <div style={{maxHeight:500,overflowY:"auto"}}>
              {log.length===0?<div style={{padding:32,textAlign:"center",color:"#94a3b8",fontSize:13}}>No activity yet</div>
              :log.map(e=><div key={e.id} style={{display:"flex",alignItems:"flex-start",gap:10,padding:"10px 16px",borderBottom:"1px solid #f8fafc"}}>
                <div style={{width:32,height:32,borderRadius:"50%",background:"#dbeafe",color:"#1d4ed8",display:"flex",alignItems:"center",justifyContent:"center",fontSize:13,fontWeight:700,flexShrink:0}}>{(e.user||"?")[0].toUpperCase()}</div>
                <div style={{flex:1}}><span style={{fontWeight:600,fontSize:12}}>{e.user}</span><span style={{fontSize:11,color:"#94a3b8"}}> · {e.role}</span><div style={{fontSize:12,color:"#475569",marginTop:2}}>{e.action}</div></div>
                <div style={{fontSize:11,color:"#94a3b8",whiteSpace:"nowrap"}}>{e.time}</div>
              </div>)}
            </div>
          </div>
        )}

      </div>

      {/* ══ SYNC-FROM-PROJECT MODAL ══ */}
      {syncState&&(()=>{ const st=syncState, others=(projects||[]).filter(p=>p.id!==pid), allOn=st.rows.length>0&&st.rows.every(r=>st.sel.has(r.key));
        const setSel=fn=>setSyncState(x=>({...x,sel:fn(new Set(x.sel))})); const nChg=st.rows.filter(r=>r.kind==="changed").length, nNew=st.rows.length-nChg;
        const th={padding:"7px 8px",textAlign:"left",fontSize:11,color:"#64748b",fontWeight:700,borderBottom:"1px solid #e2e8f0",position:"sticky",top:0,background:"#f8fafc"};
        return (
        <div style={{position:"fixed",inset:0,background:"rgba(15,23,42,.6)",zIndex:500,display:"flex",alignItems:"center",justifyContent:"center",padding:16}}>
          <div style={{background:"#fff",borderRadius:16,width:"100%",maxWidth:980,maxHeight:"90vh",display:"flex",flexDirection:"column",boxShadow:"0 25px 60px rgba(0,0,0,.3)"}}>
            <div style={{padding:"14px 20px",borderBottom:"1px solid #f1f5f9",display:"flex",alignItems:"center",gap:10,flexWrap:"wrap"}}>
              <b style={{fontSize:15,color:"#1e293b"}}>⇅ Sync BUR from another project</b>
              <select value={st.src} onChange={e=>loadSync(e.target.value)} style={{border:"1.5px solid #e2e8f0",borderRadius:8,padding:"6px 10px",fontSize:13,outline:"none",minWidth:220}}>
                <option value="">Choose project…</option>{others.map(p=><option key={p.id} value={p.id}>{p.name}</option>)}
              </select>
              <button onClick={()=>setSyncState(null)} style={{marginLeft:"auto",background:"#f1f5f9",border:"none",borderRadius:8,padding:"6px 12px",fontSize:12,cursor:"pointer",color:"#64748b",fontWeight:600}}>Close ✕</button>
            </div>
            <div style={{padding:"10px 20px",display:"flex",gap:14,flexWrap:"wrap",alignItems:"center",fontSize:12,color:"#334155",borderBottom:"1px solid #f1f5f9"}}>
              <span style={{fontWeight:700,color:"#64748b"}}>Copy:</span>
              {Object.entries(SYNC_GROUPS).map(([g,def])=><label key={g} style={{display:"flex",gap:4,alignItems:"center",cursor:"pointer"}}><input type="checkbox" checked={!!st.opts[g]} onChange={e=>setSyncState(x=>({...x,opts:{...x.opts,[g]:e.target.checked}}))}/>{def.label}</label>)}
            </div>
            <div style={{flex:1,overflow:"auto",padding:"0 20px"}}>
              {!st.src?<div style={{padding:"40px 0",textAlign:"center",color:"#94a3b8",fontSize:13}}>Pick the project to copy from. Items are matched by BUR code; nothing in this project is deleted.</div>
              :st.loading?<div style={{padding:"40px 0",textAlign:"center",color:"#94a3b8",fontSize:13}}>⏳ Comparing…</div>
              :!st.rows.length?<div style={{padding:"40px 0",textAlign:"center",color:"#059669",fontSize:13}}>✅ Nothing to sync — this project already matches.</div>
              :<table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                <thead><tr><th style={{...th,width:30}}><input type="checkbox" checked={allOn} onChange={e=>setSel(()=>e.target.checked?new Set(st.rows.map(r=>r.key)):new Set())}/></th><th style={th}>What</th><th style={th}>Code</th><th style={th}>Description</th><th style={{...th,textAlign:"right"}}>Rate here</th><th style={{...th,textAlign:"right"}}>Rate there</th><th style={th}>Different</th></tr></thead>
                <tbody>{st.rows.map(r=>{ const on=st.sel.has(r.key); return (
                  <tr key={r.key} style={{borderBottom:"1px solid #f1f5f9",background:on?"#f0fdfa":"#fff"}}>
                    <td style={{padding:"5px 8px"}}><input type="checkbox" checked={on} onChange={()=>setSel(x=>{ x.has(r.key)?x.delete(r.key):x.add(r.key); return x; })}/></td>
                    <td style={{padding:"5px 8px"}}><span style={{background:r.kind==="new"?"#dcfce7":"#fef3c7",color:r.kind==="new"?"#166534":"#92400e",borderRadius:4,padding:"1px 6px",fontSize:10,fontWeight:700}}>{r.kind==="new"?"NEW":"CHANGED"}</span></td>
                    <td style={{padding:"5px 8px",fontFamily:"monospace",color:"#1d4ed8",fontWeight:700}}>{r.src.code}</td>
                    <td style={{padding:"5px 8px",color:"#334155",maxWidth:300,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}} title={r.src.desc}>{r.src.desc}</td>
                    <td style={{padding:"5px 8px",textAlign:"right",color:"#64748b"}}>{r.cur?fmt(bTot(r.cur)):"—"}</td>
                    <td style={{padding:"5px 8px",textAlign:"right",fontWeight:700,color:r.cur&&Math.abs(bTot(r.cur)-bTot(r.src))>0.005?"#b45309":"#1e293b"}}>{fmt(bTot(r.src))}</td>
                    <td style={{padding:"5px 8px",fontSize:11,color:"#64748b"}}>{r.kind==="new"?"only in that project":r.diffs.map(d=>SYNC_GROUPS[d].label).join(", ")}</td>
                  </tr>); })}</tbody>
              </table>}
            </div>
            <div style={{padding:"12px 20px",borderTop:"1px solid #f1f5f9",display:"flex",alignItems:"center",gap:10,flexWrap:"wrap"}}>
              {st.rows.length>0&&<span style={{fontSize:12,color:"#64748b"}}>{nChg} changed · {nNew} new · <b>{st.sel.size} selected</b></span>}
              <button onClick={applySync} disabled={!st.sel.size} style={{marginLeft:"auto",background:st.sel.size?"#0e7490":"#cbd5e1",color:"#fff",border:"none",borderRadius:8,padding:"8px 18px",fontSize:13,fontWeight:700,cursor:st.sel.size?"pointer":"default"}}>⇅ Sync {st.sel.size} item(s)</button>
            </div>
          </div>
        </div>); })()}
      {/* ══ PASTE-FROM-EXCEL MODAL ══ */}
      {pasteOpen&&(
        <div style={{position:"fixed",inset:0,background:"rgba(15,23,42,.6)",zIndex:500,display:"flex",alignItems:"center",justifyContent:"center",padding:16}}>
          <div style={{background:"#fff",borderRadius:16,width:"100%",maxWidth:680,maxHeight:"88vh",display:"flex",flexDirection:"column",boxShadow:"0 25px 60px rgba(0,0,0,.3)"}}>
            <div style={{padding:"16px 20px",borderBottom:"1px solid #f1f5f9",display:"flex",alignItems:"center",justifyContent:"space-between"}}>
              <div style={{fontWeight:700,fontSize:15,color:"#1e293b"}}>📋 Paste from Excel</div>
              <button onClick={()=>setPasteOpen(false)} style={{background:"#f1f5f9",border:"none",borderRadius:8,padding:"6px 12px",fontSize:12,cursor:"pointer",color:"#64748b",fontWeight:600}}>Close ✕</button>
            </div>
            <div style={{padding:"16px 20px",overflow:"auto"}}>
              <div style={{fontSize:12,color:"#475569",marginBottom:8}}>In Excel, select cells in <b>4 columns in this order: Description · Unit · Code · Rate</b>, copy, then paste below. A header row is auto-detected and skipped.</div>
              <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:8,flexWrap:"wrap"}}>
                <label style={{fontSize:12,color:"#475569"}}>Import into category:</label>
                <select value={pasteCat} onChange={e=>setPasteCat(e.target.value)} style={{border:"1.5px solid #e2e8f0",borderRadius:7,padding:"6px 10px",fontSize:12,outline:"none",background:"#fff"}}>
                  {displayCats.map(c=><option key={c.id} value={c.id}>{c.name}</option>)}
                </select>
              </div>
              <textarea value={pasteText} onChange={e=>setPasteText(e.target.value)} placeholder={"Acoustic Ceiling Panel AC1\tm2\tAcousticCeilingPanel_C1\t130.50\n…"} style={{width:"100%",height:220,border:"1.5px solid #e2e8f0",borderRadius:8,padding:"10px 12px",fontSize:12,fontFamily:"monospace",outline:"none",boxSizing:"border-box",resize:"vertical"}}/>
              <div style={{display:"flex",gap:8,justifyContent:"flex-end",marginTop:12}}>
                <button onClick={()=>setPasteText("")} style={{background:"#f1f5f9",border:"none",borderRadius:7,padding:"8px 14px",fontSize:12,cursor:"pointer",color:"#64748b"}}>Clear</button>
                <button onClick={importPaste} style={{background:"#2563eb",color:"#fff",border:"none",borderRadius:7,padding:"8px 20px",fontSize:12,cursor:"pointer",fontWeight:600}}>Import</button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* ══ COST DATA MODAL ══ */}
      {costModal&&modalItem&&(
        <div style={{position:"fixed",inset:0,background:"rgba(15,23,42,.6)",zIndex:500,display:"flex",alignItems:"center",justifyContent:"center",padding:16}}>
          <div onPaste={async e=>{ const it=[...(e.clipboardData?.items||[])].find(x=>x.type&&x.type.startsWith("image/")); if(it){const f=it.getAsFile(); if(f){await pickCostImg(f); toast_("📷 Photo pasted — fill details, then Add Entry");}} }} style={{background:"#fff",borderRadius:16,width:"96%",maxWidth:1080,maxHeight:"92vh",display:"flex",flexDirection:"column",boxShadow:"0 25px 60px rgba(0,0,0,.3)"}}>
            <div style={{padding:"16px 20px",borderBottom:"1px solid #f1f5f9",display:"flex",alignItems:"flex-start",justifyContent:"space-between"}}>
              <div>
                <div style={{fontWeight:700,fontSize:15,color:"#1e293b",display:"flex",alignItems:"center",gap:8}}>
                  📊 Cost Data
                  <span style={{background:"#eff6ff",color:"#1d4ed8",fontFamily:"monospace",fontWeight:800,padding:"2px 8px",borderRadius:5,fontSize:13}}>{modalItem.code||"—"}</span>
                </div>
                <div style={{fontSize:12,color:"#64748b",marginTop:3}}>{modalItem.desc} · {modalItem.unit}</div>
              </div>
              <button onClick={()=>setCostModal(null)} style={{background:"#f1f5f9",border:"none",borderRadius:8,padding:"6px 12px",fontSize:12,cursor:"pointer",color:"#64748b",fontWeight:600,flexShrink:0}}>Close ✕</button>
            </div>
            <div style={{display:"flex",borderBottom:"1px solid #e2e8f0",padding:"0 20px",flexShrink:0,overflowX:"auto"}}>
              {allComps.map(c=>{const cnt=(modalItem.costData||[]).filter(e=>e.component===c).length; const custom=!COMPS.includes(c); return(
                <span key={c} style={{display:"inline-flex",alignItems:"center",borderBottom:cType===c?"2.5px solid #2563eb":"2.5px solid transparent"}}>
                  <button onClick={()=>setCType(c)} style={{padding:"9px 10px 9px 14px",fontSize:12,fontWeight:600,border:"none",color:cType===c?"#2563eb":"#64748b",background:"none",cursor:"pointer",whiteSpace:"nowrap"}}>
                    {compLabel(c)}{cnt>0&&<span style={{marginLeft:4,background:cType===c?"#dbeafe":"#f1f5f9",color:cType===c?"#1d4ed8":"#64748b",borderRadius:10,padding:"0 5px",fontSize:10}}>{cnt}</span>}
                  </button>
                  {custom&&<button onClick={()=>delComponentTab(c)} title="Remove tab" style={{border:"none",background:"none",color:"#cbd5e1",cursor:"pointer",fontSize:11,padding:"0 8px 0 0"}}>✕</button>}
                </span>
              );})}
              <button onClick={addComponentTab} title="Add a tab" style={{padding:"9px 12px",fontSize:12,fontWeight:700,border:"none",color:"#7c3aed",background:"none",cursor:"pointer",whiteSpace:"nowrap"}}>＋ Tab</button>
            </div>
            <div style={{flex:1,overflow:"auto",padding:"16px 20px"}}>
              {(()=>{ const hid=(modalItem.costData||[]).filter(e=>e.inRate&&!tabExists(e.component));
                const gone=(modalItem.costData||[]).filter(e=>!tabExists(e.component));
                const warn=gone.length>0&&<div style={{background:hid.length?"#fef2f2":"#f8fafc",border:"1px solid "+(hid.length?"#fecaca":"#e2e8f0"),borderRadius:8,padding:"8px 12px",marginBottom:8,fontSize:12,color:hid.length?"#991b1b":"#475569",display:"flex",gap:8,alignItems:"center",flexWrap:"wrap"}}>
                  {hid.length?<>⚠️ {hid.length} ticked entr{hid.length>1?"ies are":"y is"} on a removed tab (S$ {fmt(hid.reduce((t,e)=>t+cdRate(e),0))}) and still in this item's rate.</>:<>ℹ️ {gone.length} entr{gone.length>1?"ies":"y"} from removed tabs {gone.length>1?"are":"is"} hidden (not counted).</>}
                  {hid.length>0&&<button onClick={clearHiddenTicks} style={{background:"#dc2626",color:"#fff",border:"none",borderRadius:6,padding:"4px 10px",fontSize:12,fontWeight:700,cursor:"pointer"}}>Remove from rate</button>}
                  <button onClick={deleteHiddenEntries} style={{background:"#fff",color:"#b91c1c",border:"1px solid #fca5a5",borderRadius:6,padding:"4px 10px",fontSize:12,fontWeight:700,cursor:"pointer"}}>🗑 Delete hidden entries</button></div>;
                const t=(modalItem.costData||[]).filter(e=>e.inRate&&tabExists(e.component)); if(!t.length)return warn|| <div style={{fontSize:11,color:"#94a3b8",marginBottom:10}}>💡 Tick ✓ entries on any tab (Labour, Material, Mould, Freight…) to add them up into this item's rate per {modalItem.unit}.</div>;
                const by={}; t.forEach(e=>{ by[e.component]=(by[e.component]||0)+cdRate(e); });
                return <>{warn}<div style={{background:"#ecfdf5",border:"1px solid #a7f3d0",borderRadius:8,padding:"8px 12px",marginBottom:12,fontSize:12,color:"#065f46",display:"flex",flexWrap:"wrap",gap:6,alignItems:"center"}}>
                  <b>Rate from ticked ({t.length}):</b>
                  {Object.entries(by).map(([c,v],i)=><span key={c}>{i>0?"+ ":""}{compLabel(c)} <b>{fmt(v)}</b></span>)}
                  {+modalItem.waste?<span>+ Waste/ OH {+modalItem.waste}%</span>:null}
                  {(modalItem.extras||[]).some(x=>!x.fromCD)?<span>+ other extra items</span>:null}
                  <span style={{marginLeft:"auto",fontWeight:800,color:"#1d4ed8",fontSize:13}}>= S$ {fmt(bTot(modalItem))} / {modalItem.unit}</span>
                </div></>; })()}
              {costEntries.length>0&&<div style={{display:"flex",gap:6,alignItems:"center",marginBottom:8,fontSize:11,color:"#64748b",flexWrap:"wrap"}}>
                <span style={{fontWeight:700}}>Sort:</span>
                {[["supplier","Supplier"],["location","Location"],["rate","Rate"],["date","Date"]].map(([k,l])=><button key={k} onClick={()=>ceToggleSort(k)} style={{border:"1px solid #e2e8f0",background:ceSort.by===k?"#7c3aed":"#fff",color:ceSort.by===k?"#fff":"#475569",borderRadius:6,padding:"3px 9px",fontSize:11,cursor:"pointer",fontWeight:600}}>{l}{ceSort.by===k?(ceSort.dir===1?" ▲":" ▼"):""}</button>)}
                <button onClick={()=>setShowGraph(g=>!g)} style={{marginLeft:8,border:"1px solid #e2e8f0",background:showGraph?"#16a34a":"#fff",color:showGraph?"#fff":"#475569",borderRadius:6,padding:"3px 10px",fontSize:11,cursor:"pointer",fontWeight:700}}>📈 Graph</button>
              </div>}
              {showGraph&&costEntries.length>0&&(()=>{ const data=[...costEntries].sort((a,b)=>String(a.date||"").localeCompare(String(b.date||""))); const max=Math.max(...data.map(e=>+e.rate||0),1); const H=180,pad=34,step=Math.max(56,Math.min(90,520/data.length)),W=Math.max(300,data.length*step+30),bw=Math.min(46,step-14);
                return <div style={{overflowX:"auto",marginBottom:14,border:"1px solid #e2e8f0",borderRadius:8,padding:10,background:"#fff"}}>
                  <div style={{fontSize:11,fontWeight:700,color:"#475569",marginBottom:4}}>{compLabel(cType)} rate trend (by date)</div>
                  <svg width={W} height={H}>
                    <line x1={18} y1={H-pad} x2={W-6} y2={H-pad} stroke="#e2e8f0"/>
                    {data.map((e,i)=>{ const x=24+i*step; const bh=(H-pad-22)*((+e.rate||0)/max); const y=H-pad-bh; return <g key={i}>
                      <rect x={x} y={y} width={bw} height={bh} rx={3} fill="#7c3aed"/>
                      <text x={x+bw/2} y={y-3} fontSize="10" textAnchor="middle" fill="#1e293b" fontWeight="700">{fmt(+e.rate||0)}</text>
                      <text x={x+bw/2} y={H-pad+13} fontSize="9" textAnchor="middle" fill="#64748b">{(e.supplier||"").slice(0,9)}</text>
                      <text x={x+bw/2} y={H-pad+24} fontSize="8" textAnchor="middle" fill="#94a3b8">{e.date||""}</text>
                    </g>;})}
                  </svg>
                </div>;
              })()}
              {costEntries.length===0?(
                <div style={{textAlign:"center",padding:"28px 0",color:"#94a3b8",fontSize:13}}>No {compLabel(cType)} entries yet — add one below</div>
              ):(
                <table style={{width:"100%",borderCollapse:"collapse",fontSize:12,marginBottom:16}}>
                  <thead><tr style={{background:"#f8fafc",color:"#94a3b8",fontSize:11}}>
                    <th title="Tick to include in the item's rate" style={{padding:"8px 6px",borderBottom:"1px solid #e2e8f0",width:30,textAlign:"center"}}>✓</th>{["Supplier / Vendor","Rate (S$)",`Per → ${modalItem.unit||"unit"}`,"Date","Location","Note","Photo",""].map((h,i)=><th key={i} style={{padding:"8px 10px",textAlign:i===1?"right":"left",fontWeight:600,borderBottom:"1px solid #e2e8f0"}}>{h}</th>)}
                  </tr></thead>
                  <tbody>
                    {costEntries.map(e=>(
                      <tr key={e.id} style={{borderBottom:"1px solid #f8fafc",background:e.inRate?"#ecfdf5":undefined}}>
                        <td style={{padding:"6px 6px",textAlign:"center"}}><input type="checkbox" checked={!!e.inRate} onChange={()=>toggleInRate(e.id)} title="Include this in the item's rate" style={{accentColor:"#059669",cursor:"pointer",width:15,height:15}}/></td>
                        <td style={{padding:"6px 8px"}}><input value={e.supplier||""} onChange={ev=>updCostEntry(e.id,{supplier:ev.target.value})} style={{width:"100%",border:"1px solid transparent",borderRadius:5,padding:"4px 6px",fontSize:12,fontWeight:600,color:"#1e293b",outline:"none",background:"#f8fafc"}}/></td>
                        <td style={{padding:"6px 8px",textAlign:"right"}}><input title={e.expr?`ƒ ${e.expr}`:"Type a number or a formula, e.g. =(100+20)/2 or =rebar_kg+0.13"} onFocus={()=>setEditExpr(e.id)} onBlur={()=>setEditExpr(null)} value={editExpr===e.id||!e.expr||!isFinite(calcExpr(e.expr))?(e.expr??(e.rate??"")):String(+(+e.rate||0).toFixed(4))} onChange={ev=>{ const t=ev.target.value, n=calcExpr(t); updCostEntry(e.id,{expr:(/[-+*/()]/.test(String(t).replace(/^\s*[=-]/,""))||String(n)!==t.trim())?t:null, /* keep the typed text (e.g. "2239." while typing a decimal) unless it is exactly the plain number */...(isFinite(n)?{rate:Math.round(n*10000)/10000}:t.trim()===""?{rate:0}:{})}); }} style={{width:e.expr&&editExpr===e.id?150:90,border:"1px solid transparent",borderRadius:5,padding:"4px 6px",fontSize:13,fontWeight:700,color:"#1d4ed8",outline:"none",textAlign:"right",background:"#f8fafc"}}/>{e.expr&&(()=>{ const refs=exprRefs(e.expr), miss=refs.filter(c=>!findBurByCode(burItemsRef.current,c)), self=refs.filter(c=>{const t=findBurByCode(burItemsRef.current,c);return t&&t.id===modalItem.id;}), n=calcExpr(e.expr);
                          return <div style={{fontSize:10,marginTop:2,whiteSpace:"nowrap",color:miss.length||self.length||!isFinite(n)?"#dc2626":"#059669"}}>{miss.length?`code not found: ${miss.join(", ")}`:self.length?"can't use this item's own code":!isFinite(n)?"check formula":<>{editExpr===e.id?<>= S$ {fmt(n)}</>:<span style={{color:"#94a3b8",fontStyle:"italic"}} title={e.expr}>ƒ {String(e.expr).length>22?String(e.expr).slice(0,22)+"…":e.expr}</span>}{!e.inRate&&<span style={{color:"#b45309"}}> · tick ✓ to use</span>}</>}</div>; })()}</td>
                        <td style={{padding:"6px 8px",whiteSpace:"nowrap",fontSize:11,color:"#64748b"}} title={`Rate is per what? And how many ${modalItem.unit} does ONE of it cover? e.g. per no., 1 no. = 50 ${modalItem.unit}`}>
                          per <input value={e.per||""} placeholder={modalItem.unit} onChange={ev=>updCostEntry(e.id,{per:ev.target.value})} style={{width:44,border:"1px solid #e2e8f0",borderRadius:5,padding:"3px 5px",fontSize:11,outline:"none"}}/>
                          {e.per&&String(e.per).trim().toLowerCase()!==String(modalItem.unit||"").trim().toLowerCase()&&<> · 1 = <input value={e.cover??""} placeholder="?" onChange={ev=>updCostEntry(e.id,{cover:ev.target.value})} style={{width:60,border:"1px solid "+(String(e.cover??"").trim()!==""&&!isFinite(calcExpr(e.cover))?"#ef4444":"#e2e8f0"),borderRadius:5,padding:"3px 5px",fontSize:11,outline:"none",textAlign:"right"}}/> {modalItem.unit}
                            {cdCover(e)!==1&&<b style={{color:"#1d4ed8",marginLeft:4}}>= S$ {fmt(cdRate(e))}/{modalItem.unit}</b>}</>}
                        </td>
                        <td style={{padding:"6px 8px"}}><input value={e.date||""} placeholder="—" onChange={ev=>updCostEntry(e.id,{date:ev.target.value})} style={{width:80,border:"1px solid transparent",borderRadius:5,padding:"4px 6px",fontSize:12,color:"#64748b",outline:"none",background:"#f8fafc"}}/></td>
                        <td style={{padding:"6px 8px"}}><input value={e.location||""} placeholder="—" onChange={ev=>updCostEntry(e.id,{location:ev.target.value})} style={{width:100,border:"1px solid transparent",borderRadius:5,padding:"4px 6px",fontSize:12,color:"#64748b",outline:"none",background:"#f8fafc"}}/></td>
                        <td style={{padding:"6px 8px"}}><input value={e.note||""} placeholder="—" onChange={ev=>updCostEntry(e.id,{note:ev.target.value})} style={{width:"100%",minWidth:90,border:"1px solid transparent",borderRadius:5,padding:"4px 6px",fontSize:11,color:"#64748b",outline:"none",background:"#f8fafc"}}/></td>
                        <td style={{padding:"6px 10px"}}><div style={{display:"flex",gap:5,flexWrap:"wrap",alignItems:"center",maxWidth:170}}>
                          {entryPhotos(e).map((p,i)=><span key={i} style={{position:"relative",display:"inline-flex"}}><img src={p} alt="" onClick={()=>setImgView(p)} style={{height:30,width:40,objectFit:"cover",borderRadius:3,cursor:"zoom-in",border:"1px solid #e2e8f0"}}/><button onClick={()=>removeEntryPhoto(e.id,i)} title="Remove photo" style={{position:"absolute",top:-6,right:-6,background:"#ef4444",color:"#fff",border:"none",borderRadius:"50%",width:15,height:15,fontSize:9,lineHeight:"15px",padding:0,cursor:"pointer"}}>✕</button></span>)}
                          <label title="Add photo" style={{height:30,width:30,display:"flex",alignItems:"center",justifyContent:"center",border:"1px dashed #cbd5e1",borderRadius:4,cursor:"pointer",color:"#64748b",fontSize:15}}>＋<input type="file" accept="image/*" multiple style={{display:"none"}} onChange={ev=>addEntryPhotos(e.id,ev.target.files)}/></label>
                        </div></td>
                        <td style={{padding:"10px 8px"}}>
                          <div style={{display:"flex",gap:6,justifyContent:"flex-end"}}>
                            {cType==="subcon"&&<button onClick={()=>useCostEntry(e)} style={{background:"#f59e0b",color:"#fff",border:"none",borderRadius:6,padding:"4px 10px",fontSize:11,cursor:"pointer",fontWeight:700,whiteSpace:"nowrap"}}>→ Use as Quote</button>}
                            <button onClick={()=>delCostEntry(e.id)} style={{background:"none",border:"none",color:"#ef4444",cursor:"pointer",fontSize:13,fontWeight:700,padding:"2px 4px"}}>✕</button>
                          </div>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              )}
              <div style={{background:"#f8fafc",borderRadius:10,padding:14,border:"1px solid #e2e8f0"}}>
                <div style={{fontSize:12,fontWeight:600,color:"#475569",marginBottom:10}}>+ Add {compLabel(cType)} Quote / Price</div>
                <div style={{display:"grid",gridTemplateColumns:"2fr 1fr 1fr 1.4fr 2fr",gap:8,marginBottom:10}}>
                  {[["supplier","Supplier / Vendor","text","e.g. ABC Contractors Pte Ltd"],["rate","Rate (S$)","text","0.00, =(100+20)/2 or =rebar+0.13"],["date","Date","text","e.g. Jun 2025"],["location","Location","text","e.g. Level 3 / Block A"],["note","Note","text","e.g. FOB, ex-GST"]].map(([k,l,t,ph])=>(
                    <div key={k}><div style={{fontSize:10,color:"#94a3b8",marginBottom:3,fontWeight:600}}>{l}</div>
                    <input type={t} style={{width:"100%",border:"1.5px solid #e2e8f0",borderRadius:7,padding:"7px 9px",fontSize:12,outline:"none",boxSizing:"border-box"}} placeholder={ph} value={cForm[k]} onChange={e=>setCForm(f=>({...f,[k]:e.target.value}))}/>{k==="rate"&&(/[-+*/()]/.test(String(cForm.rate).replace(/^\s*[=-]/,""))||exprRefs(cForm.rate).length>0)&&(()=>{ const miss=exprRefs(cForm.rate).filter(c=>!findBurByCode(burItemsRef.current,c)), n=calcExpr(cForm.rate);
                        return <div style={{fontSize:11,marginTop:2,color:!miss.length&&isFinite(n)?"#059669":"#ef4444"}}>{miss.length?`code not found: ${miss.join(", ")}`:isFinite(n)?`= S$ ${fmt(n)}`:"check formula"}</div>; })()}</div>
                  ))}
                </div>
                <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:10,flexWrap:"wrap"}}>
                  <label style={{background:"#fff",border:"1px solid #e2e8f0",borderRadius:7,padding:"6px 12px",fontSize:12,cursor:"pointer",color:"#475569",fontWeight:600}}>📷 Attach photo(s)<input type="file" accept="image/*" multiple style={{display:"none"}} onChange={e=>pickCostImg(e.target.files)}/></label>
                  <span style={{fontSize:11,color:"#94a3b8"}}>…or paste screenshots (Ctrl+V) — you can add many</span>
                  {(cForm.imgs||[]).map((p,i)=><span key={i} style={{display:"inline-flex",alignItems:"center",gap:3}}><img src={p} alt="" style={{height:38,width:50,objectFit:"cover",borderRadius:4,border:"1px solid #e2e8f0"}}/><button onClick={()=>setCForm(f=>({...f,imgs:f.imgs.filter((_,j)=>j!==i)}))} style={{border:"none",background:"none",color:"#ef4444",cursor:"pointer",fontSize:12,fontWeight:700}}>✕</button></span>)}
                </div>
                <div style={{display:"flex",gap:8,justifyContent:"flex-end"}}>
                  <button onClick={()=>setCForm({supplier:"",rate:"",date:"",location:"",note:"",imgs:[]})} style={{background:"#f1f5f9",border:"none",borderRadius:7,padding:"7px 14px",fontSize:12,cursor:"pointer",color:"#64748b"}}>Clear</button>
                  <button onClick={addCostEntry} style={{background:"#2563eb",color:"#fff",border:"none",borderRadius:7,padding:"7px 18px",fontSize:12,cursor:"pointer",fontWeight:600}}>Add Entry</button>
                </div>
              </div>
              {cType==="subcon"&&<div style={{marginTop:10,background:"#eff6ff",borderRadius:8,padding:"8px 12px",fontSize:11,color:"#1d4ed8"}}>💡 Click <b>"→ Use as Quote"</b> to set this supplier's rate as the item's Sub-Con rate.</div>}
            </div>
          </div>
        </div>
      )}

      {/* ══ IMAGE LIGHTBOX ══ */}
      {imgView&&(
        <div onClick={()=>setImgView(null)} style={{position:"fixed",inset:0,background:"rgba(0,0,0,.82)",zIndex:600,display:"flex",alignItems:"center",justifyContent:"center",padding:20,cursor:"zoom-out"}}>
          <img src={imgView} alt="" style={{maxWidth:"95%",maxHeight:"95%",borderRadius:8,boxShadow:"0 10px 40px rgba(0,0,0,.5)"}}/>
        </div>
      )}
    </div>
  );
}
