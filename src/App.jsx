import { useState, useEffect, useRef, useCallback, Fragment } from "react";
import { auth, db } from "./firebase";
import { onAuthStateChanged, signInWithEmailAndPassword, signOut } from "firebase/auth";
import { collection, doc, onSnapshot, setDoc, updateDoc, deleteDoc, writeBatch, getDoc } from "firebase/firestore";
// NOTE: BUR items are scoped per project via a `pid` field.
import burSeed from "./burSeed.json"; // master BUR list imported from BUR.xlsx
import SheetGrid from "./SheetGrid.jsx";

const SECTIONS=[{id:"prelim",name:"Preliminaries"},{id:"building",name:"Building Works"},{id:"external",name:"External Works"},{id:"mande",name:"M&E Works"},{id:"fees",name:"Professional Fees"}];
const UNITS=["m²","m³","m","nr","sum","lot","kg","t","m run","%","item"];
const STATUSES=["Draft","Under Review","Confirmed"];
const SS={Draft:{bg:"#f1f5f9",c:"#475569"},"Under Review":{bg:"#fef9c3",c:"#b45309"},Confirmed:{bg:"#dcfce7",c:"#15803d"}};
const ROLES=["Lead QS","Estimator","PM","Client"];
const TABS=[{id:"boq",label:"📋 BOQ"},{id:"bur",label:"📚 BUR"},{id:"summary",label:"📊 Summary"},{id:"log",label:"📝 Activity"}];
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

const newSections=()=>SECTIONS.map(s=>({...s,items:[]}));
const fmt=n=>(n||0).toLocaleString("en-SG",{minimumFractionDigits:2,maximumFractionDigits:2});
// Resize/compress an image File or Blob to a small JPEG data URL (for storing quote photos).
// Aggressively shrink a data-URL image for embedding into Excel (small dimensions + low quality).
const shrinkDataURL=(dataUrl,max=460,q=0.45)=>new Promise(res=>{ if(!dataUrl){res("");return;} const im=new Image(); im.onload=()=>{ let w=im.width,h=im.height; if(w>max){h=h*max/w;w=max;} if(h>max){w=w*max/h;h=max;} const cv=document.createElement("canvas"); cv.width=w; cv.height=h; cv.getContext("2d").drawImage(im,0,0,w,h); try{res(cv.toDataURL("image/jpeg",q));}catch{res(dataUrl);} }; im.onerror=()=>res(dataUrl); im.src=dataUrl; });
const imgToDataURL=blob=>new Promise(res=>{ if(!blob){res("");return;} const fr=new FileReader(); fr.onload=()=>{ const im=new Image(); im.onload=()=>{ const max=1100; let w=im.width,h=im.height; if(w>max){h=h*max/w;w=max;} if(h>max){w=w*max/h;h=max;} const cv=document.createElement("canvas"); cv.width=w; cv.height=h; cv.getContext("2d").drawImage(im,0,0,w,h); try{res(cv.toDataURL("image/jpeg",0.7));}catch{res("");} }; im.onerror=()=>res(""); im.src=fr.result; }; fr.onerror=()=>res(""); fr.readAsDataURL(blob); });
const uid=()=>"_"+Math.random().toString(36).slice(2,10);
const bTot=b=>{const d=(+b.labour||0)+(+b.material||0)+(+b.plant||0)+(+b.subcon||0),oh=d*(+b.oh||0)/100,s=d+oh;return s*(1+(+b.profit||0)/100);};
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
  const [burSearch,setBurSearch]=useState("");
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

  const sTimer=useRef(null); const uRef=useRef(null); const pidRef=useRef(null); const dirtyRef=useRef(false); const pendingRef=useRef(null); const logRef=useRef([]); const burTimers=useRef({});
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

  useEffect(()=>{ if(!ready)return; const ref=doc(db,"meta","codes"); return onSnapshot(ref,s=>{ if(s.exists())setCodes(s.data().list||[]); else{setCodes(DEF_CODES);setDoc(ref,{list:DEF_CODES}).catch(()=>{});} }); },[ready]);
  useEffect(()=>{ if(!ready)return; const ref=doc(db,"meta","cats"); return onSnapshot(ref,s=>{ if(s.exists())setCats(s.data().list||[]); else{setCats(DEF_CATS);setDoc(ref,{list:DEF_CATS}).catch(()=>{});} }); },[ready]);
  // Keep the selected category valid when the category list changes (e.g. after loading the master list).
  useEffect(()=>{ if(cats.length&&!cats.some(c=>c.id===selCat))setSelCat(cats[0].id); },[cats]); // eslint-disable-line
  // Per-project master BQ template override (uploaded). Falls back to the bundled CAG template.
  useEffect(()=>{ if(!ready||!pid){setBqTplB64(null);setBqTplName("");return;} return onSnapshot(doc(db,"boqtemplate",pid),s=>{ if(s.exists()){setBqTplB64(s.data().b64||null);setBqTplName(s.data().name||"");}else{setBqTplB64(null);setBqTplName("");} }); },[ready,pid]);
  useEffect(()=>{ if(!ready)return; return onSnapshot(doc(db,"meta","log"),s=>{ const e=s.exists()?(s.data().entries||[]):[]; logRef.current=e; setLog(e); }); },[ready]);
  useEffect(()=>{ if(!ready)return; return onSnapshot(doc(db,"meta","ccomps"),s=>setCcomps(s.exists()?(s.data().list||[]):[])); },[ready]);
  const addComponentTab=async()=>{ const name=prompt("New cost-data tab name (e.g. Equipment, Transport, Specialist):"); if(!name||!name.trim())return; const key="cc_"+Math.random().toString(36).slice(2,8); try{await setDoc(doc(db,"meta","ccomps"),{list:[...ccomps,{key,label:name.trim()}]});setCType(key);}catch(e){toast_("⚠️ "+e.message);} };
  const delComponentTab=async key=>{ if(!confirm("Remove this tab? Existing entries stay in the data but the tab is hidden."))return; try{await setDoc(doc(db,"meta","ccomps"),{list:ccomps.filter(c=>c.key!==key)});}catch{} if(cType===key)setCType("subcon"); };

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

  const blurSave=useCallback((sid,iid,code)=>{ if(code){const m=burItems.find(b=>b.code&&b.code.toLowerCase()===code.toLowerCase()); if(m&&bTot(m)>0)setRateSugg({secId:sid,iid,rate:bTot(m),code:m.code,desc:m.desc});} },[burItems]);

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
          const bur=codeMap[code.toLowerCase()]; if(!bur){unmatched.add(code);continue;}
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

  // Build the BUR export from LIVE data, grouped by category, formatted like the master build-up sheet.
  const exportBUR=useCallback(async()=>{
    try{
      toast_("⏳ Building BUR Excel…");
      const ExcelJS=(await import("exceljs")).default;
      const wb=new ExcelJS.Workbook(); const ws=wb.addWorksheet("BUILD UP RATES");
      const bd={top:{style:"thin",color:{argb:"FFBFBFBF"}},left:{style:"thin",color:{argb:"FFBFBFBF"}},bottom:{style:"thin",color:{argb:"FFBFBFBF"}},right:{style:"thin",color:{argb:"FFBFBFBF"}}};
      const headers=["Description","Unit","Code","Labour","Material","Plant","Subcon","Rate","Sub-Con Quotes"];
      ws.columns=[{width:50},{width:8},{width:22},{width:10},{width:11},{width:9},{width:11},{width:12},{width:55}];
      const moneyCols=new Set([4,5,6,7,8]);
      const codeCount={}; burItems.forEach(b=>{const k=(b.code||"").trim().toLowerCase(); if(k)codeCount[k]=(codeCount[k]||0)+1;});
      // group by category
      const groups=[]; const known=new Set();
      const orderedCats=catSort==="none"?cats:[...cats].sort((a,b)=>catSort==="az"?String(a.name).localeCompare(String(b.name)):String(b.name).localeCompare(String(a.name)));
      for(const cat of orderedCats){ known.add(cat.id); const items=burItems.filter(b=>b.catId===cat.id).sort((a,b)=>(a.code||"").localeCompare(b.code||"")); if(items.length)groups.push({name:cat.name,items}); }
      const orphan=burItems.filter(b=>!known.has(b.catId)).sort((a,b)=>(a.code||"").localeCompare(b.code||"")); if(orphan.length)groups.push({name:"(Uncategorised)",items:orphan});
      ws.mergeCells(1,1,1,headers.length); const t=ws.getCell(1,1); t.value="BUILD UP RATES"; t.font={bold:true,size:14,color:{argb:"FF1E3A5F"}};
      const hr=ws.getRow(2); headers.forEach((h,i)=>{ const c=hr.getCell(i+1); c.value=h; c.font={bold:true}; c.fill={type:"pattern",pattern:"solid",fgColor:{argb:"FFFFF200"}}; c.border=bd; c.alignment={horizontal:(i>=3&&i<=7)?"right":"left",vertical:"middle",wrapText:true}; }); hr.height=22;
      { const pc=ws.getCell(2,headers.length+1); pc.value="Photos"; pc.font={bold:true}; pc.fill={type:"pattern",pattern:"solid",fgColor:{argb:"FFFFF200"}}; pc.border=bd; }
      for(let c=headers.length+1;c<=headers.length+12;c++)ws.getColumn(c).width=18;
      const PAL=["FFFDE68A","FFBFDBFE","FFBBF7D0","FFFBCFE8","FFDDD6FE","FFFED7AA","FFA5F3FC","FFE9D5FF","FFFEF08A","FFD9F99D","FFFECACA","FFC7D2FE","FFFCD34D","FF99F6E4"];
      const tint=hex=>{ const r=parseInt(hex.slice(2,4),16),g=parseInt(hex.slice(4,6),16),b=parseInt(hex.slice(6,8),16); const m=x=>Math.round(x+(255-x)*0.62).toString(16).padStart(2,"0"); return "FF"+m(r)+m(g)+m(b); };
      let r=3, count=0, photoCount=0;
      for(let gi=0;gi<groups.length;gi++){ const g=groups[gi]; const band=PAL[gi%PAL.length]; const rowTint=tint(band);
        ws.mergeCells(r,1,r,headers.length+12); const cc=ws.getCell(r,1); cc.value=g.name+"   ("+g.items.length+")"; cc.font={bold:true,size:12,color:{argb:"FF1E293B"}}; cc.fill={type:"pattern",pattern:"solid",fgColor:{argb:band}}; cc.border=bd; ws.getRow(r).height=20; r++;
        for(let j=0;j<g.items.length;j++){ const it=g.items[j]; const fillArgb=(j%2===1)?rowTint:"FFFFFFFF";
          const q=(it.costData||[]).filter(e=>e.component==="subcon").map(e=>`${e.supplier}: ${e.rate}${e.location?` @${e.location}`:""}${e.date?` (${e.date})`:""}`).join("  |  ");
          const total=+bTot(it)||0;
          const vals=[it.desc||"",it.unit||"",it.code||"",+it.labour||0,+it.material||0,+it.plant||0,+it.subcon||0,total,q];
          const row=ws.getRow(r);
          vals.forEach((v,i)=>{ const c=row.getCell(i+1); c.value=(moneyCols.has(i+1)&&!(+v>0))?null:v; c.border=bd; c.fill={type:"pattern",pattern:"solid",fgColor:{argb:fillArgb}}; c.alignment={vertical:"top",horizontal:(i>=3&&i<=7)?"right":"left",wrapText:(i===0||i===8)}; if(moneyCols.has(i+1))c.numFmt="#,##0.00"; if(i===2)c.font={name:"Consolas",color:{argb:"FF1D4ED8"}}; if(i===7)c.font={bold:true,color:{argb:"FF1D4ED8"}}; });
          if((codeCount[(it.code||"").trim().toLowerCase()]||0)>1){ const cc2=row.getCell(3); cc2.fill={type:"pattern",pattern:"solid",fgColor:{argb:"FFFECACA"}}; cc2.font={name:"Consolas",bold:true,color:{argb:"FFB91C1C"}}; }
          // embed this item's photos in the columns after the table, on the same row
          const itemPhotos=(it.costData||[]).flatMap(e=>e.imgs&&e.imgs.length?e.imgs:(e.img?[e.img]:[])).slice(0,12);
          if(itemPhotos.length){ row.height=98; for(let i=0;i<itemPhotos.length;i++){ try{ const small=await shrinkDataURL(itemPhotos[i]); const m=/^data:image\/(\w+);base64,(.+)$/.exec(small); if(m){ const id=wb.addImage({base64:m[2],extension:"jpeg"}); ws.addImage(id,{tl:{col:headers.length+i,row:r-1},ext:{width:124,height:92}}); photoCount++; } }catch{} } }
          r++; count++;
        }
      }
      ws.views=[{state:"frozen",ySplit:2}];
      const out=await wb.xlsx.writeBuffer(); const blob=new Blob([out],{type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});
      const url=URL.createObjectURL(blob); const a=document.createElement("a"); a.href=url; a.download="BUR_BuildUpRates.xlsx"; a.click(); URL.revokeObjectURL(url);
      toast_(`✅ Exported ${count} items${photoCount?` + ${photoCount} photos`:""}`);
    }catch(e){ toast_("⚠️ Export failed: "+(e&&e.message||e)); }
  },[burItems,cats,toast_,catSort]);

  // ── BUR writes (per-document in shared library) ─────────────────────────────
  const addBurItem=useCallback(async catId=>{ const ref=doc(collection(db,"bur")); try{ await setDoc(ref,{pid,catId,code:"",desc:"New Item",unit:"sum",labour:0,material:0,plant:0,subcon:0,oh:15,profit:10,costData:[],quote:null,group:""}); setExpBur(ref.id);}catch(e){toast_("⚠️ "+e.message);} },[pid]);
  const updBur=useCallback((id,ch)=>{ setBurItems(prev=>prev.map(b=>b.id===id?{...b,...ch}:b)); if(burTimers.current[id])clearTimeout(burTimers.current[id]); burTimers.current[id]=setTimeout(()=>updateDoc(doc(db,"bur",id),ch).catch(()=>{}),600); },[]);
  const delBur=useCallback(id=>{ if(confirm("Delete this BUR item?"))deleteDoc(doc(db,"bur",id)).catch(()=>{}); },[]);
  const setBurField=(id,patch)=>{ setBurItems(prev=>prev.map(b=>b.id===id?{...b,...patch}:b)); updateDoc(doc(db,"bur",id),patch).catch(()=>{}); };

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
    for(let i=start;i<lines.length;i++){ const c=lines[i].split("\t"); const desc=(c[0]||"").trim(),unit=(c[1]||"").trim(),code=(c[2]||"").trim(); const rate=parseFloat(String(c[3]||"").replace(/[^0-9.\-]/g,""))||0; if(!desc&&!code)continue; rows.push({pid,catId:pasteCat,code,desc:desc||"(no description)",unit:unit||"sum",labour:0,material:rate,plant:0,subcon:0,oh:0,profit:0,costData:[],quote:null,group:"Pasted"}); }
    if(!rows.length){toast_("⚠️ No rows parsed — expected: Description, Unit, Code, Rate");return;}
    try{ let batch=writeBatch(db),n=0; for(const r of rows){ batch.set(doc(collection(db,"bur")),r); n++; if(n%400===0){await batch.commit();batch=writeBatch(db);} } await batch.commit(); setPasteOpen(false);setPasteText("");setSelCat(pasteCat); toast_(`✅ Imported ${rows.length} items`);}catch(e){toast_("⚠️ "+e.message);}
  },[pasteText,pasteCat,toast_,pid]);

  // ── Cost data / quotes ──────────────────────────────────────────────────────
  const modalItem=costModal?burItems.find(b=>b.id===costModal):null;
  const costEntries=(()=>{ let arr=modalItem?(modalItem.costData||[]).filter(e=>e.component===cType):[]; if(ceSort.by!=="none"){ arr=[...arr].sort((a,b)=>{ if(ceSort.by==="rate")return ((+a.rate||0)-(+b.rate||0))*ceSort.dir; return String(a[ceSort.by]||"").localeCompare(String(b[ceSort.by]||""),undefined,{numeric:true})*ceSort.dir; }); } return arr; })();
  const ceToggleSort=by=>setCeSort(s=>s.by===by?{by,dir:-s.dir}:{by,dir:1});
  const compLabel=c=>CLABEL[c]||(ccomps.find(x=>x.key===c)||{}).label||c;
  const allComps=[...COMPS,...ccomps.map(c=>c.key)];

  const addCostEntry=useCallback(()=>{ if(!cForm.supplier||!cForm.rate){toast_("⚠️ Enter supplier and rate");return;} const b=burItems.find(x=>x.id===costModal); if(!b)return; const cd=[...(b.costData||[]),{id:uid(),component:cType,supplier:cForm.supplier,rate:+cForm.rate,date:cForm.date,location:cForm.location||"",note:cForm.note,imgs:cForm.imgs||[]}]; setBurField(costModal,{costData:cd}); setCForm({supplier:"",rate:"",date:"",location:"",note:"",imgs:[]}); toast_("✅ Entry added"); },[cType,cForm,costModal,burItems,toast_]);
  const pickCostImg=useCallback(async files=>{ const list=files&&files.length!=null?[...files]:(files?[files]:[]); if(!list.length)return; const out=[]; for(const file of list){ const d=await imgToDataURL(file); if(d&&d.length<=1400000)out.push(d); else if(d)toast_("⚠️ One image too large, skipped"); } if(out.length)setCForm(f=>({...f,imgs:[...(f.imgs||[]),...out]})); },[toast_]);
  const delCostEntry=useCallback(eid=>{ const b=burItems.find(x=>x.id===costModal); if(!b)return; setBurField(costModal,{costData:(b.costData||[]).filter(e=>e.id!==eid)}); },[costModal,burItems]);
  // Edit an existing cost-data entry (optimistic local + debounced save)
  const updCostEntry=useCallback((eid,ch)=>{ setBurItems(prev=>prev.map(b=>b.id===costModal?{...b,costData:(b.costData||[]).map(e=>e.id===eid?{...e,...ch}:e)}:b)); const k="cd_"+costModal; if(burTimers.current[k])clearTimeout(burTimers.current[k]); burTimers.current[k]=setTimeout(()=>{ const b=burItemsRef.current.find(x=>x.id===costModal); if(b)updateDoc(doc(db,"bur",costModal),{costData:b.costData}).catch(()=>{}); },600); },[costModal]);
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
  const BUR_MAX=200; const _q=burSearch.trim().toLowerCase();
  // When searching, look across the WHOLE BUR library; otherwise show the selected category.
  const _base=_q?burItems:burItems.filter(b=>b.catId===selCat);
  const _filtered=_q?_base.filter(b=>(b.code||"").toLowerCase().includes(_q)||(b.desc||"").toLowerCase().includes(_q)||(b.costData||[]).some(e=>(e.supplier||"").toLowerCase().includes(_q)||(e.note||"").toLowerCase().includes(_q))):_base;
  const _sorted=[..._filtered].sort((a,b)=>{
    let av,bv;
    if(sortBy==="rate"){av=bTot(a);bv=bTot(b);return (av-bv)*sortDir;}
    if(sortBy==="cat"){av=catName(a.catId);bv=catName(b.catId);}
    else{av=a[sortBy]||"";bv=b[sortBy]||"";}
    return String(av).localeCompare(String(bv),undefined,{numeric:true})*sortDir;
  });
  const catTotal=_sorted.length; const catItems=_sorted.slice(0,BUR_MAX);
  const toggleSort=f=>{ if(sortBy===f)setSortDir(d=>-d); else{setSortBy(f);setSortDir(1);} };
  const renameCat=cid=>{ const c=cats.find(x=>x.id===cid); if(!c)return; const nm=prompt("Rename category:",c.name); if(nm&&nm.trim())pushCats(cats.map(x=>x.id===cid?{...x,name:nm.trim()}:x)); };
  const delCat=cid=>{ const n=burItems.filter(b=>b.catId===cid).length; if(n>0){alert(`This category has ${n} item(s). Move them to another category first (open an item → CATEGORY), then delete.`);return;} if(confirm("Delete this empty category?")){ pushCats(cats.filter(x=>x.id!==cid)); if(selCat===cid){const o=cats.find(x=>x.id!==cid); setSelCat(o?o.id:"");} } };
  const toggleSel=id=>setSelBur(prev=>{const s=new Set(prev); s.has(id)?s.delete(id):s.add(id); return s;});
  const clearSel=()=>setSelBur(new Set());
  const moveItemsToCat=async(ids,cid)=>{ if(!ids.length||!cid)return; for(let i=0;i<ids.length;i+=400){const batch=writeBatch(db); ids.slice(i,i+400).forEach(id=>batch.update(doc(db,"bur",id),{catId:cid})); await batch.commit();} };
  const bulkMoveCat=async cid=>{ const ids=[...selBur]; if(!ids.length||!cid)return; await moveItemsToCat(ids,cid); toast_(`✅ Moved ${ids.length} item(s)`); clearSel(); };
  const bulkSetUnit=async u=>{ const ids=[...selBur]; if(!ids.length||!u)return; for(let i=0;i<ids.length;i+=400){const batch=writeBatch(db); ids.slice(i,i+400).forEach(id=>batch.update(doc(db,"bur",id),{unit:u})); await batch.commit();} toast_(`✅ Set unit on ${ids.length} item(s)`); clearSel(); };
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
        <colgroup><col style={{width:28}}/><col style={{width:24}}/>{burHead.map(([k])=><col key={k} style={{width:burCw(k)}}/>)}<col style={{width:34}}/></colgroup>
        <thead><tr style={{background:"#fef9c3",color:"#92400e",fontSize:11}}>
          <th style={{borderBottom:"1px solid #fde68a",textAlign:"center"}}><input type="checkbox" title="Select all shown" checked={items.length>0&&items.every(it=>selBur.has(it.id))} onChange={e=>{setSelBur(prev=>{const s=new Set(prev); items.forEach(it=>e.target.checked?s.add(it.id):s.delete(it.id)); return s;});}} style={{accentColor:"#7c3aed",cursor:"pointer"}}/></th>
          <th style={{borderBottom:"1px solid #fde68a"}}></th>
          {burHead.map(([k,l,al])=><th key={k} style={{padding:"8px 8px",textAlign:al,fontWeight:700,borderBottom:"1px solid #fde68a",whiteSpace:"nowrap",position:"relative",overflow:"hidden"}}>{l}<div onMouseDown={e=>startBurDrag(k,e)} title="Drag to resize" style={{position:"absolute",top:0,right:0,width:8,height:"100%",cursor:"col-resize"}}/></th>)}
          <th style={{borderBottom:"1px solid #fde68a"}}></th>
        </tr></thead>
        <tbody>
          {items.map((item,ri)=>{
            const isExp=expBur===item.id; const total=bTot(item); const direct=(+item.labour||0)+(+item.material||0)+(+item.plant||0)+(+item.subcon||0); const cdCount=(item.costData||[]).length; const dup=isDupCode(item.code); const checked=selBur.has(item.id);
            const lbl={fontSize:10,color:"#94a3b8",fontWeight:600,display:"block",marginBottom:2}; const inp={border:"1.5px solid #e2e8f0",borderRadius:6,padding:"5px 8px",fontSize:12,outline:"none"};
            return(<Fragment key={item.id}>
              <tr style={{borderBottom:"1px solid #f8fafc",background:checked?"#ede9fe":(ri%2===1?"#f5f3ff":"#fff")}}>
                <td style={{padding:"2px 4px",textAlign:"center"}}><input type="checkbox" checked={checked} onChange={()=>toggleSel(item.id)} style={{accentColor:"#7c3aed",cursor:"pointer"}}/></td>
                <td draggable onDragStart={e=>{e.dataTransfer.setData("text/bur",item.id);e.dataTransfer.effectAllowed="move";}} title="Drag me onto a category to move • click to expand" style={{padding:"2px 4px",textAlign:"center",cursor:"grab",color:isExp?"#2563eb":"#94a3b8",userSelect:"none"}} onClick={()=>setExpBur(isExp?null:item.id)}>⠿{isExp?"▼":"▶"}</td>
                <td style={{padding:"3px 6px",overflow:"hidden"}}><input style={{width:"100%",border:"none",fontSize:12,outline:"none",background:"transparent"}} value={item.desc||""} onChange={e=>updBur(item.id,{desc:e.target.value})}/></td>
                <td style={{padding:"3px 6px"}}><select style={{border:"none",fontSize:11,outline:"none",background:"transparent",width:"100%"}} value={item.unit||"sum"} onChange={e=>updBur(item.id,{unit:e.target.value})}>{UNITS.map(u=><option key={u}>{u}</option>)}</select></td>
                <td style={{padding:"3px 6px",overflow:"hidden",background:dup?"#fee2e2":undefined}}><input title={dup?"⚠️ Duplicate BUR code":""} style={{width:"100%",border:"none",fontSize:11,outline:"none",background:"transparent",fontFamily:"monospace",fontWeight:700,color:dup?"#b91c1c":"#1d4ed8"}} value={item.code||""} onChange={e=>updBur(item.id,{code:e.target.value})}/></td>
                <td style={{padding:"3px 6px",textAlign:"right"}}><input type="number" style={{width:"100%",border:"none",fontSize:12,outline:"none",background:"transparent",textAlign:"right"}} value={item.labour?item.labour:""} onChange={e=>updBur(item.id,{labour:+e.target.value||0})}/></td>
                <td style={{padding:"3px 6px",textAlign:"right"}}><input type="number" style={{width:"100%",border:"none",fontSize:12,outline:"none",background:"transparent",textAlign:"right"}} value={item.material?item.material:""} onChange={e=>updBur(item.id,{material:+e.target.value||0})}/></td>
                <td style={{padding:"3px 6px",textAlign:"right"}}><input type="number" style={{width:"100%",border:"none",fontSize:12,outline:"none",background:"transparent",textAlign:"right"}} value={item.plant?item.plant:""} onChange={e=>updBur(item.id,{plant:+e.target.value||0})}/></td>
                <td style={{padding:"3px 8px",textAlign:"right",fontWeight:700,color:"#1d4ed8",whiteSpace:"nowrap"}}>{total>0?fmt(total):""}</td>
                <td style={{padding:"3px 6px",textAlign:"center"}}><button onClick={()=>{setCostModal(item.id);setCType("subcon");}} style={{background:"#fef9c3",border:"1px solid #fde68a",borderRadius:6,padding:"3px 8px",fontSize:11,cursor:"pointer",color:"#92400e",fontWeight:600,whiteSpace:"nowrap"}}>📊{cdCount?` ${cdCount}`:""}</button></td>
                <td style={{padding:"3px 4px",textAlign:"center"}}><button onClick={()=>delBur(item.id)} style={{border:"none",background:"none",color:"#ef4444",cursor:"pointer",fontSize:13,fontWeight:700}}>✕</button></td>
              </tr>
              {isExp&&(<tr><td colSpan={11} style={{background:"#fafafa",padding:"12px 16px",borderBottom:"1px solid #eef2f7"}}>
                <div style={{display:"flex",gap:16,flexWrap:"wrap",marginBottom:12}}>
                  <div style={{flex:"2 1 320px"}}><label style={lbl}>DESCRIPTION (full)</label><textarea value={item.desc||""} onChange={e=>updBur(item.id,{desc:e.target.value})} style={{width:"100%",border:"1.5px solid #e2e8f0",borderRadius:6,padding:"7px 10px",fontSize:13,outline:"none",minHeight:64,resize:"vertical",boxSizing:"border-box",lineHeight:1.4}}/></div>
                  <div style={{flex:"1 1 180px"}}><label style={lbl}>BUR CODE</label><input value={item.code||""} onChange={e=>updBur(item.id,{code:e.target.value})} style={{width:"100%",border:"1.5px solid #e2e8f0",borderRadius:6,padding:"7px 10px",fontSize:13,outline:"none",fontFamily:"monospace",fontWeight:700,color:"#1d4ed8",boxSizing:"border-box"}}/></div>
                </div>
                <div style={{display:"flex",gap:16,flexWrap:"wrap",alignItems:"flex-end"}}>
                  <div><label style={lbl}>SUBCON (S$)</label><input type="number" style={{...inp,width:90}} value={item.subcon?item.subcon:""} onChange={e=>updBur(item.id,{subcon:+e.target.value||0})}/></div>
                  <div><label style={lbl}>OH %</label><input type="number" style={{...inp,width:70}} value={item.oh??15} onChange={e=>updBur(item.id,{oh:+e.target.value||0})}/></div>
                  <div><label style={lbl}>PROFIT %</label><input type="number" style={{...inp,width:70}} value={item.profit??10} onChange={e=>updBur(item.id,{profit:+e.target.value||0})}/></div>
                  <div><label style={lbl}>CATEGORY</label><select style={{...inp,background:"#fff"}} value={item.catId||"cat13"} onChange={e=>updBur(item.id,{catId:e.target.value})}>{cats.map(c=><option key={c.id} value={c.id}>{c.name}</option>)}</select></div>
                  <div style={{marginLeft:"auto",fontSize:12,color:"#475569"}}>Direct <b>S$ {fmt(direct)}</b> · OH({item.oh}%) <b>S$ {fmt(direct*(+item.oh||0)/100)}</b> · <span style={{color:"#1d4ed8",fontWeight:800}}>Total S$ {fmt(total)}/{item.unit}</span></div>
                </div>
              </td></tr>)}
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
        {tab==="boq"&&boqMode==="sheet"&&<SheetGrid db={db} pid={pid} toast={toast_} baseUrl={import.meta.env.BASE_URL} onToStructured={importGridToStructured} burLookup={code=>{const b=burItems.find(x=>(x.code||"").trim().toLowerCase()===String(code).trim().toLowerCase()); return b?{material:+b.material||0,labour:+b.labour||0,total:+bTot(b)||0}:null;}}/>}
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
                            if(col.id==="unit")return<td key={col.id} style={{...p,textAlign:"center"}}><select style={{border:"none",fontSize:11,outline:"none",background:"transparent"}} value={item.unit||"sum"} onChange={e=>updItem(cs.id,item.id,{unit:e.target.value})}>{UNITS.map(u=><option key={u}>{u}</option>)}</select></td>;
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
                {burSearch&&<button onClick={()=>setBurSearch("")} style={{background:"#f1f5f9",border:"none",borderRadius:8,padding:"7px 10px",fontSize:12,cursor:"pointer",color:"#64748b"}}>Clear</button>}
                <button onClick={()=>{setPasteCat(selCat);setPasteOpen(true);}} style={{background:"#0ea5e9",color:"#fff",border:"none",borderRadius:8,padding:"7px 12px",fontSize:12,fontWeight:600,cursor:"pointer",whiteSpace:"nowrap"}}>📋 Paste from Excel</button>
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
                <span style={{fontSize:11,color:"#475569"}}>Move to:</span>
                <select value="" onChange={e=>{if(e.target.value)bulkMoveCat(e.target.value);}} style={{border:"1px solid #c4b5fd",borderRadius:6,padding:"4px 8px",fontSize:12,outline:"none",background:"#fff"}}><option value="">Category…</option>{cats.map(c=><option key={c.id} value={c.id}>{c.name}</option>)}</select>
                <span style={{fontSize:11,color:"#475569"}}>Set unit:</span>
                <select value="" onChange={e=>{if(e.target.value)bulkSetUnit(e.target.value);}} style={{border:"1px solid #c4b5fd",borderRadius:6,padding:"4px 8px",fontSize:12,outline:"none",background:"#fff"}}><option value="">Unit…</option>{UNITS.map(u=><option key={u} value={u}>{u}</option>)}</select>
                <button onClick={clearSel} style={{marginLeft:"auto",background:"#fff",border:"1px solid #c4b5fd",borderRadius:6,padding:"4px 10px",fontSize:11,cursor:"pointer",color:"#5b21b6",fontWeight:600}}>Clear</button>
              </div>}
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10,flexWrap:"wrap",gap:6}}>
                <span style={{fontSize:13,fontWeight:600,color:"#1e293b"}}>{_q?"🔎 Search results (all categories)":catName(selCat)} <span style={{color:"#64748b",fontWeight:400,fontSize:12}}>— {catTotal} items{catItems.length<catTotal?` (showing first ${catItems.length} — refine search)`:""}</span>{dupCount>0&&<span style={{marginLeft:8,background:"#fee2e2",color:"#b91c1c",borderRadius:6,padding:"2px 8px",fontSize:11,fontWeight:700}}>⚠️ {dupCount} duplicate code{dupCount>1?"s":""}</span>}</span>
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
                    {["Supplier / Vendor","Rate (S$)","Date","Location","Note","Photo",""].map((h,i)=><th key={i} style={{padding:"8px 10px",textAlign:i===1?"right":"left",fontWeight:600,borderBottom:"1px solid #e2e8f0"}}>{h}</th>)}
                  </tr></thead>
                  <tbody>
                    {costEntries.map(e=>(
                      <tr key={e.id} style={{borderBottom:"1px solid #f8fafc"}}>
                        <td style={{padding:"6px 8px"}}><input value={e.supplier||""} onChange={ev=>updCostEntry(e.id,{supplier:ev.target.value})} style={{width:"100%",border:"1px solid transparent",borderRadius:5,padding:"4px 6px",fontSize:12,fontWeight:600,color:"#1e293b",outline:"none",background:"#f8fafc"}}/></td>
                        <td style={{padding:"6px 8px",textAlign:"right"}}><input type="number" value={e.rate??""} onChange={ev=>updCostEntry(e.id,{rate:+ev.target.value||0})} style={{width:90,border:"1px solid transparent",borderRadius:5,padding:"4px 6px",fontSize:13,fontWeight:700,color:"#1d4ed8",outline:"none",textAlign:"right",background:"#f8fafc"}}/></td>
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
                  {[["supplier","Supplier / Vendor","text","e.g. ABC Contractors Pte Ltd"],["rate","Rate (S$)","number","0.00"],["date","Date","text","e.g. Jun 2025"],["location","Location","text","e.g. Level 3 / Block A"],["note","Note","text","e.g. FOB, ex-GST"]].map(([k,l,t,ph])=>(
                    <div key={k}><div style={{fontSize:10,color:"#94a3b8",marginBottom:3,fontWeight:600}}>{l}</div>
                    <input type={t} style={{width:"100%",border:"1.5px solid #e2e8f0",borderRadius:7,padding:"7px 9px",fontSize:12,outline:"none",boxSizing:"border-box"}} placeholder={ph} value={cForm[k]} onChange={e=>setCForm(f=>({...f,[k]:e.target.value}))}/></div>
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
