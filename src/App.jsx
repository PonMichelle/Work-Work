import { useState, useEffect, useRef, useCallback } from "react";
import { auth, db } from "./firebase";
import { onAuthStateChanged, signInWithEmailAndPassword, signOut } from "firebase/auth";
import { collection, doc, onSnapshot, setDoc, updateDoc, deleteDoc, writeBatch } from "firebase/firestore";
import burSeed from "./burSeed.json"; // master BUR list imported from BUR.xlsx

const SECTIONS=[{id:"prelim",name:"Preliminaries"},{id:"building",name:"Building Works"},{id:"external",name:"External Works"},{id:"mande",name:"M&E Works"},{id:"fees",name:"Professional Fees"}];
const UNITS=["m²","m³","m","nr","sum","lot","kg","t","m run","%","item"];
const STATUSES=["Draft","Under Review","Confirmed"];
const SS={Draft:{bg:"#f1f5f9",c:"#475569"},"Under Review":{bg:"#fef9c3",c:"#b45309"},Confirmed:{bg:"#dcfce7",c:"#15803d"}};
const ROLES=["Lead QS","Estimator","PM","Client"];
const TABS=[{id:"boq",label:"📋 BOQ"},{id:"bur",label:"📚 BUR"},{id:"rates",label:"🔢 Rates & Codes"},{id:"summary",label:"📊 Summary"},{id:"log",label:"📝 Activity"}];
const ALL_COLS=[{id:"ref",label:"Ref",w:48},{id:"desc",label:"Description",w:190},{id:"unit",label:"Unit",w:64},{id:"qty",label:"Qty",w:56,num:1},{id:"rA",label:"Rate A",w:74,num:1},{id:"amtA",label:"Amt A (S$)",w:100,num:1},{id:"rB",label:"Rate B",w:74,num:1},{id:"amtB",label:"Amt B (S$)",w:100,num:1},{id:"code",label:"BUR Code",w:90},{id:"status",label:"Status",w:110},{id:"remarks",label:"Remarks",w:130},{id:"by",label:"By",w:56}];
const DEF_COLS=new Set(["ref","desc","unit","qty","rA","amtA","rB","amtB","code","status","remarks","by"]);
const DEF_CODES=[{id:"dc1",code:"PRELIM",desc:"Preliminaries",cat:"Prelim"},{id:"dc2",code:"BLDG-A",desc:"Building Works Phase A",cat:"Building"},{id:"dc3",code:"BLDG-B",desc:"Building Works Phase B",cat:"Building"},{id:"dc4",code:"EXT-A",desc:"External Works Phase A",cat:"External"},{id:"dc5",code:"EXT-B",desc:"External Works Phase B",cat:"External"},{id:"dc6",code:"ME-ACMV",desc:"Air Conditioning & Mech Ventilation",cat:"M&E"},{id:"dc7",code:"ME-ELV",desc:"Electrical & Low Voltage",cat:"M&E"},{id:"dc8",code:"ME-FP",desc:"Fire Protection",cat:"M&E"},{id:"dc9",code:"ME-STP",desc:"Sewage Treatment Plant",cat:"M&E"},{id:"dc10",code:"FEES",desc:"Professional Fees",cat:"Fees"}];
const DEF_CATS=["Concrete Works","Formwork","Reinforcement","Brickwork & Blockwork","Waterproofing","Structural Steelwork","Floor & Wall Finishes","Roofing","Doors, Windows & Glazing","External Works","M&E Works","Preliminaries","Others"].map((name,i)=>({id:`cat${i+1}`,name}));
const MASTER_CAT={id:"cat_master",name:"Master BUR"};
const COMPS=["labour","material","plant","subcon"];
const CLABEL={labour:"Labour",material:"Material",plant:"Plant",subcon:"Subcon / Supplier"};

const newSections=()=>SECTIONS.map(s=>({...s,items:[]}));
const fmt=n=>(n||0).toLocaleString("en-SG",{minimumFractionDigits:2,maximumFractionDigits:2});
const uid=()=>"_"+Math.random().toString(36).slice(2,10);
const bTot=b=>{const d=(+b.labour||0)+(+b.material||0)+(+b.plant||0)+(+b.subcon||0),oh=d*(+b.oh||0)/100,s=d+oh;return s*(1+(+b.profit||0)/100);};
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
  const [pasteOpen,setPasteOpen]=useState(false); const [pasteText,setPasteText]=useState(""); const [pasteCat,setPasteCat]=useState("cat_master");
  const [costModal,setCostModal]=useState(null);
  const [cType,setCType]=useState("subcon");
  const [cForm,setCForm]=useState({supplier:"",rate:"",date:"",note:""});
  const [newCat,setNewCat]=useState(""); const [showNewCat,setShowNewCat]=useState(false);
  const [codes,setCodes]=useState(DEF_CODES);
  const [editCId,setEditCId]=useState(null); const [codeForm,setCodeForm]=useState({code:"",desc:"",cat:""}); const [showAddC,setShowAddC]=useState(false);
  const [rbu,setRbu]=useState({labour:"",material:"",plant:"",subcon:"",oh:15,profit:10});
  const [log,setLog]=useState([]);
  const [toast,setToast]=useState(null);

  const sTimer=useRef(null); const uRef=useRef(null); const pidRef=useRef(null); const dirtyRef=useRef(false); const pendingRef=useRef(null); const logRef=useRef([]); const burTimers=useRef({});
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

  useEffect(()=>{ if(!ready)return; return onSnapshot(collection(db,"bur"),snap=>setBurItems(snap.docs.map(d=>({id:d.id,...d.data()})))); },[ready]);

  useEffect(()=>{ if(!ready)return; const ref=doc(db,"meta","codes"); return onSnapshot(ref,s=>{ if(s.exists())setCodes(s.data().list||[]); else{setCodes(DEF_CODES);setDoc(ref,{list:DEF_CODES}).catch(()=>{});} }); },[ready]);
  useEffect(()=>{ if(!ready)return; const ref=doc(db,"meta","cats"); return onSnapshot(ref,s=>{ if(s.exists())setCats(s.data().list||[]); else{setCats(DEF_CATS);setDoc(ref,{list:DEF_CATS}).catch(()=>{});} }); },[ready]);
  useEffect(()=>{ if(!ready)return; return onSnapshot(doc(db,"meta","log"),s=>{ const e=s.exists()?(s.data().entries||[]):[]; logRef.current=e; setLog(e); }); },[ready]);

  // Current project's BOQ
  useEffect(()=>{ if(!ready||!pid){setData(null);return;} return onSnapshot(doc(db,"projects",pid),s=>{ if(!s.exists())return; if(dirtyRef.current)return; const d=s.data(); setData({sections:d.sections||newSections(),ts:d.ts||0}); }); },[ready,pid]);

  // ── Project ops ─────────────────────────────────────────────────────────────
  const createProject=async()=>{ const name=prompt("New project / tender name:"); if(!name||!name.trim())return; const ref=doc(collection(db,"projects")); try{ await setDoc(ref,{name:name.trim(),createdAt:Date.now(),sections:newSections(),ts:Date.now()}); setPid(ref.id); setTab("boq"); addLogEntry(`Created project "${name.trim()}"`);}catch(e){toast_("⚠️ "+e.message);} };
  const renameProject=async()=>{ const cur=projects?.find(p=>p.id===pid); const name=prompt("Rename project:",cur?.name||""); if(!name||!name.trim())return; try{await updateDoc(doc(db,"projects",pid),{name:name.trim()});}catch(e){toast_("⚠️ "+e.message);} };
  const deleteProject=async()=>{ if(!pid)return; const cur=projects?.find(p=>p.id===pid); if(!confirm(`Delete project "${cur?.name}" and all its BOQ items? This cannot be undone.`))return; try{await deleteDoc(doc(db,"projects",pid)); setPid(null);}catch(e){toast_("⚠️ "+e.message);} };

  // ── BOQ writes ──────────────────────────────────────────────────────────────
  const writeProject=nd=>{ if(!pidRef.current)return; setDoc(doc(db,"projects",pidRef.current),{sections:nd.sections,ts:Date.now()},{merge:true}).catch(()=>{}); };
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

  // ── BUR writes (per-document in shared library) ─────────────────────────────
  const addBurItem=useCallback(async catId=>{ const ref=doc(collection(db,"bur")); try{ await setDoc(ref,{catId,code:"",desc:"New Item",unit:"sum",labour:0,material:0,plant:0,subcon:0,oh:15,profit:10,costData:[],quote:null,group:""}); setExpBur(ref.id);}catch(e){toast_("⚠️ "+e.message);} },[]);
  const updBur=useCallback((id,ch)=>{ setBurItems(prev=>prev.map(b=>b.id===id?{...b,...ch}:b)); if(burTimers.current[id])clearTimeout(burTimers.current[id]); burTimers.current[id]=setTimeout(()=>updateDoc(doc(db,"bur",id),ch).catch(()=>{}),600); },[]);
  const delBur=useCallback(id=>{ if(confirm("Delete this BUR item?"))deleteDoc(doc(db,"bur",id)).catch(()=>{}); },[]);
  const setBurField=(id,patch)=>{ setBurItems(prev=>prev.map(b=>b.id===id?{...b,...patch}:b)); updateDoc(doc(db,"bur",id),patch).catch(()=>{}); };

  const loadMaster=useCallback(async()=>{
    if(burItems.some(b=>b.catId==="cat_master")&&!confirm(`This adds the ${burSeed.length}-item master list to the SHARED library for everyone. Continue?`))return;
    toast_("⏳ Loading master list to shared library…");
    try{ let batch=writeBatch(db),n=0; for(const it of burSeed){ const ref=doc(collection(db,"bur")); const {id,...rest}=it; batch.set(ref,rest); n++; if(n%400===0){await batch.commit();batch=writeBatch(db);} } await batch.commit(); setSelCat("cat_master"); addLogEntry(`Loaded ${burSeed.length} master BUR items`); toast_(`✅ Loaded ${burSeed.length} master items (shared)`); }
    catch(e){ toast_("⚠️ Load failed: "+e.message); }
  },[burItems,toast_,addLogEntry]);

  const importPaste=useCallback(async()=>{
    const lines=pasteText.split(/\r?\n/).filter(l=>l.trim()!=="");
    if(!lines.length){toast_("⚠️ Nothing to paste");return;}
    const first=lines[0].split("\t").map(s=>s.trim().toLowerCase());
    const start=(first[0]==="description"||first.includes("code"))?1:0;
    const rows=[];
    for(let i=start;i<lines.length;i++){ const c=lines[i].split("\t"); const desc=(c[0]||"").trim(),unit=(c[1]||"").trim(),code=(c[2]||"").trim(); const rate=parseFloat(String(c[3]||"").replace(/[^0-9.\-]/g,""))||0; if(!desc&&!code)continue; rows.push({catId:pasteCat,code,desc:desc||"(no description)",unit:unit||"sum",labour:0,material:rate,plant:0,subcon:0,oh:0,profit:0,costData:[],quote:null,group:"Pasted"}); }
    if(!rows.length){toast_("⚠️ No rows parsed — expected: Description, Unit, Code, Rate");return;}
    try{ let batch=writeBatch(db),n=0; for(const r of rows){ batch.set(doc(collection(db,"bur")),r); n++; if(n%400===0){await batch.commit();batch=writeBatch(db);} } await batch.commit(); setPasteOpen(false);setPasteText("");setSelCat(pasteCat); toast_(`✅ Imported ${rows.length} items`);}catch(e){toast_("⚠️ "+e.message);}
  },[pasteText,pasteCat,toast_]);

  // ── Cost data / quotes ──────────────────────────────────────────────────────
  const modalItem=costModal?burItems.find(b=>b.id===costModal):null;
  const costEntries=modalItem?(modalItem.costData||[]).filter(e=>e.component===cType):[];

  const addCostEntry=useCallback(()=>{ if(!cForm.supplier||!cForm.rate){toast_("⚠️ Enter supplier and rate");return;} const b=burItems.find(x=>x.id===costModal); if(!b)return; const cd=[...(b.costData||[]),{id:uid(),component:cType,supplier:cForm.supplier,rate:+cForm.rate,date:cForm.date,note:cForm.note}]; setBurField(costModal,{costData:cd}); setCForm({supplier:"",rate:"",date:"",note:""}); toast_("✅ Entry added"); },[cType,cForm,costModal,burItems,toast_]);
  const delCostEntry=useCallback(eid=>{ const b=burItems.find(x=>x.id===costModal); if(!b)return; setBurField(costModal,{costData:(b.costData||[]).filter(e=>e.id!==eid)}); },[costModal,burItems]);
  const useCostEntry=useCallback(entry=>{ setBurField(costModal,{subcon:entry.rate,quote:{supplier:entry.supplier,rate:entry.rate,date:entry.date||"",note:entry.note||"",status:"pending",approvedBy:null,approvedAt:null}}); toast_("🟡 Rate set as Sub-Con Quotation — pending Lead QS approval"); },[costModal]);
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
    <div style={{minHeight:"100vh",background:"linear-gradient(135deg,#1e3a5f,#2563eb)",display:"flex",alignItems:"center",justifyContent:"center",padding:16,fontFamily:"system-ui,sans-serif"}}>
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
    <div style={{minHeight:"100vh",background:"linear-gradient(135deg,#1e3a5f,#2563eb)",display:"flex",alignItems:"center",justifyContent:"center",padding:16,fontFamily:"system-ui,sans-serif"}}>
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
  const displayCats=cats.some(c=>c.id==="cat_master")?cats:[MASTER_CAT,...cats];
  const catName=id=>displayCats.find(c=>c.id===id)?.name||id;
  const BUR_MAX=200; const _q=burSearch.trim().toLowerCase();
  const _allCat=burItems.filter(b=>b.catId===selCat);
  const _filtered=_q?_allCat.filter(b=>(b.code||"").toLowerCase().includes(_q)||(b.desc||"").toLowerCase().includes(_q)):_allCat;
  const catTotal=_filtered.length; const catItems=_filtered.slice(0,BUR_MAX);
  const projName=projects.find(p=>p.id===pid)?.name||"—";

  return(
    <div style={{minHeight:"100vh",background:"#f1f5f9",display:"flex",flexDirection:"column",fontFamily:"system-ui,sans-serif"}}>

      <header style={{background:"#1e3a5f",color:"#fff",padding:"10px 16px",display:"flex",alignItems:"center",justifyContent:"space-between",boxShadow:"0 2px 8px rgba(0,0,0,.2)",position:"sticky",top:0,zIndex:200,gap:10,flexWrap:"wrap"}}>
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

      {toast&&<div style={{background:"#1d4ed8",color:"#fff",textAlign:"center",padding:"7px",fontSize:13,fontWeight:500}}>{toast}</div>}

      <div style={{background:"#fff",borderBottom:"1px solid #e2e8f0",display:"flex",overflowX:"auto",flexShrink:0,position:"sticky",top:44,zIndex:100}}>
        {TABS.map(t=><button key={t.id} onClick={()=>setTab(t.id)} style={{padding:"10px 13px",fontSize:12,fontWeight:600,border:"none",flexShrink:0,whiteSpace:"nowrap",borderBottom:tab===t.id?"2.5px solid #2563eb":"2.5px solid transparent",color:tab===t.id?"#2563eb":"#64748b",background:"none",cursor:"pointer"}}>{t.label}</button>)}
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

        {/* ══ BOQ ══ */}
        {tab==="boq"&&(
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
                    <button onClick={()=>addItem(cs.id)} style={{background:"#2563eb",color:"#fff",border:"none",borderRadius:8,padding:"6px 14px",fontSize:12,fontWeight:600,cursor:"pointer"}}>+ Add Row</button>
                  </div>
                </div>
                <div style={{overflowX:"auto"}}>
                  <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                    <thead><tr style={{background:"#f8fafc",color:"#94a3b8",fontSize:11}}>
                      {vcols.map(col=><th key={col.id} style={{padding:"8px 8px",textAlign:col.num||col.id==="by"?"right":col.id==="status"||col.id==="unit"?"center":"left",fontWeight:600,whiteSpace:"nowrap",borderBottom:"1px solid #e2e8f0",minWidth:col.w}}>{col.label}</th>)}
                      <th style={{padding:"8px 6px",borderBottom:"1px solid #e2e8f0",width:28}}></th>
                    </tr></thead>
                    <tbody>
                      {!cs.items?.length?<tr><td colSpan={vcols.length+1} style={{padding:"32px",textAlign:"center",color:"#94a3b8",fontSize:13}}>No items — click + Add Row</td></tr>
                      :cs.items.map(item=>(
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
                          <td style={{padding:"4px 4px",textAlign:"center"}}><button onClick={()=>delItem(cs.id,item.id)} style={{border:"none",background:"none",color:"#ef4444",cursor:"pointer",fontSize:13,fontWeight:700}}>✕</button></td>
                        </tr>
                      ))}
                    </tbody>
                    {cs.items?.length>0&&<tfoot><tr style={{background:"#eff6ff",fontWeight:700}}>
                      {vcols.map(col=>{
                        if(col.id==="rA")return<td key={col.id} style={{padding:"8px",textAlign:"right",fontSize:12,color:"#475569"}}>Total →</td>;
                        if(col.id==="amtA")return<td key={col.id} style={{padding:"8px",textAlign:"right",fontSize:13,color:"#1d4ed8"}}>S$ {fmt(sTot(cs.id).A)}</td>;
                        if(col.id==="amtB")return<td key={col.id} style={{padding:"8px",textAlign:"right",fontSize:13,color:"#7c3aed"}}>S$ {fmt(sTot(cs.id).B)}</td>;
                        return<td key={col.id} style={{padding:"8px"}}></td>;
                      })}<td></td>
                    </tr></tfoot>}
                  </table>
                </div>
              </div>
            )}
          </div>
        )}

        {/* ══ BUR ══ */}
        {tab==="bur"&&(
          <div>
            <div style={{display:"flex",gap:6,marginBottom:12,alignItems:"center",flexWrap:"wrap"}}>
              <div style={{display:"flex",gap:6,overflowX:"auto",flex:1,paddingBottom:2}}>
                {displayCats.map(c=><button key={c.id} onClick={()=>setSelCat(c.id)} style={{padding:"6px 12px",borderRadius:20,fontSize:12,fontWeight:600,border:"none",cursor:"pointer",whiteSpace:"nowrap",flexShrink:0,background:selCat===c.id?"#1e3a5f":"#fff",color:selCat===c.id?"#fff":"#475569",boxShadow:"0 1px 3px rgba(0,0,0,.1)"}}>{c.name}<span style={{opacity:.6,fontSize:10,marginLeft:4}}>({burItems.filter(b=>b.catId===c.id).length})</span></button>)}
              </div>
              {showNewCat?(
                <div style={{display:"flex",gap:4,flexShrink:0}}>
                  <input style={{border:"1px solid #e2e8f0",borderRadius:7,padding:"5px 9px",fontSize:12,outline:"none",width:160}} placeholder="Category name" value={newCat} onChange={e=>setNewCat(e.target.value)} onKeyDown={e=>{if(e.key==="Enter"&&newCat.trim()){pushCats([...cats,{id:uid(),name:newCat.trim()}]);setNewCat("");setShowNewCat(false);}}}/>
                  <button onClick={()=>{if(newCat.trim()){pushCats([...cats,{id:uid(),name:newCat.trim()}]);setNewCat("");setShowNewCat(false);}}} style={{background:"#2563eb",color:"#fff",border:"none",borderRadius:7,padding:"5px 10px",fontSize:11,cursor:"pointer",fontWeight:600}}>Add</button>
                  <button onClick={()=>setShowNewCat(false)} style={{background:"#f1f5f9",border:"none",borderRadius:7,padding:"5px 8px",fontSize:11,cursor:"pointer",color:"#64748b"}}>✕</button>
                </div>
              ):<button onClick={()=>setShowNewCat(true)} style={{background:"#f1f5f9",border:"1px solid #e2e8f0",borderRadius:20,padding:"5px 12px",fontSize:11,cursor:"pointer",color:"#475569",fontWeight:600,whiteSpace:"nowrap",flexShrink:0}}>+ Category</button>}
            </div>

            <div style={{display:"flex",gap:8,marginBottom:10,flexWrap:"wrap",alignItems:"center"}}>
              <input value={burSearch} onChange={e=>setBurSearch(e.target.value)} placeholder="🔍 Search code or description…" style={{flex:1,minWidth:200,border:"1.5px solid #e2e8f0",borderRadius:8,padding:"7px 12px",fontSize:13,outline:"none"}}/>
              {burSearch&&<button onClick={()=>setBurSearch("")} style={{background:"#f1f5f9",border:"none",borderRadius:8,padding:"7px 10px",fontSize:12,cursor:"pointer",color:"#64748b"}}>Clear</button>}
              <button onClick={()=>{setPasteCat(selCat);setPasteOpen(true);}} style={{background:"#0ea5e9",color:"#fff",border:"none",borderRadius:8,padding:"7px 12px",fontSize:12,fontWeight:600,cursor:"pointer",whiteSpace:"nowrap"}}>📋 Paste from Excel</button>
              <button onClick={loadMaster} style={{background:"#1e3a5f",color:"#fff",border:"none",borderRadius:8,padding:"7px 12px",fontSize:12,fontWeight:600,cursor:"pointer",whiteSpace:"nowrap"}}>⬇ Load master list ({burSeed.length})</button>
            </div>

            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10,flexWrap:"wrap",gap:6}}>
              <span style={{fontSize:13,fontWeight:600,color:"#1e293b"}}>{catName(selCat)} <span style={{color:"#64748b",fontWeight:400,fontSize:12}}>— {catTotal} items{catItems.length<catTotal?` (showing first ${catItems.length} — refine search)`:""}</span></span>
              <button onClick={()=>addBurItem(selCat)} style={{background:"#2563eb",color:"#fff",border:"none",borderRadius:8,padding:"6px 14px",fontSize:12,fontWeight:600,cursor:"pointer"}}>+ Add Item</button>
            </div>

            {catItems.length===0?(
              <div style={{background:"#fff",borderRadius:12,padding:40,textAlign:"center",color:"#94a3b8"}}>
                <div style={{fontSize:36,marginBottom:8}}>📊</div>
                <div style={{fontWeight:600,fontSize:14,marginBottom:4}}>No items{_q?" match your search":" in this category"}</div>
                <div style={{fontSize:13}}>{_q?"Try a different search term":'Click "+ Add Item" or "Load master list"'}</div>
              </div>
            ):(
              <div style={{display:"flex",flexDirection:"column",gap:8}}>
                {catItems.map(item=>{
                  const isExp=expBur===item.id;
                  const isPending=item.quote?.status==="pending";
                  const isApproved=item.quote?.status==="approved";
                  const total=bTot(item);
                  const direct=(+item.labour||0)+(+item.material||0)+(+item.plant||0)+(+item.subcon||0);
                  return(
                    <div key={item.id} style={{background:"#fff",borderRadius:12,boxShadow:"0 1px 4px rgba(0,0,0,.08)",overflow:"hidden",border:isPending?"1.5px solid #fcd34d":isApproved?"1.5px solid #86efac":"none"}}>
                      <div style={{display:"flex",alignItems:"center",gap:10,padding:"11px 14px",cursor:"pointer"}} onClick={()=>setExpBur(isExp?null:item.id)}>
                        <span style={{fontSize:12,color:isExp?"#2563eb":"#cbd5e1",flexShrink:0}}>{isExp?"▼":"▶"}</span>
                        <div style={{flex:1,display:"flex",alignItems:"center",gap:10,flexWrap:"wrap",minWidth:0}}>
                          <span style={{fontWeight:800,fontSize:12,background:"#eff6ff",color:"#1d4ed8",padding:"2px 8px",borderRadius:5,flexShrink:0,fontFamily:"monospace"}}>{item.code||"—"}</span>
                          {item.group&&<span style={{fontSize:10,color:"#94a3b8",flexShrink:0,fontStyle:"italic"}}>{item.group}</span>}
                          <span style={{fontSize:13,fontWeight:500,color:"#1e293b",flex:1,minWidth:100}}>{item.desc}</span>
                          <span style={{fontSize:11,color:"#94a3b8",flexShrink:0}}>{item.unit}</span>
                          <div style={{display:"flex",gap:8,fontSize:11,flexShrink:0,flexWrap:"wrap"}}>
                            {COMPS.map(k=>+item[k]>0&&k!=="subcon"&&<span key={k} style={{color:"#64748b"}}>{k[0].toUpperCase()}:{fmt(item[k])}</span>)}
                            {isPending&&<span style={{background:"#fef08a",padding:"1px 6px",borderRadius:4,fontWeight:700,color:"#b45309"}}>🟡 SC Pending</span>}
                            {isApproved&&<span style={{background:"#bbf7d0",padding:"1px 6px",borderRadius:4,fontWeight:700,color:"#15803d"}}>✅ SC Approved</span>}
                            {!isPending&&!isApproved&&item.subcon>0&&<span style={{color:"#7c3aed"}}>SC:{fmt(item.subcon)}</span>}
                          </div>
                          <span style={{fontWeight:700,fontSize:13,color:"#1d4ed8",flexShrink:0}}>S$ {fmt(total)}/{item.unit}</span>
                        </div>
                        <div style={{display:"flex",gap:6,flexShrink:0}} onClick={e=>e.stopPropagation()}>
                          <button onClick={()=>{setCostModal(item.id);setCType("subcon");}} style={{background:"#fef9c3",border:"1px solid #fde68a",borderRadius:6,padding:"3px 8px",fontSize:11,cursor:"pointer",color:"#92400e",fontWeight:600,whiteSpace:"nowrap"}}>📊 Cost Data</button>
                          <button onClick={()=>delBur(item.id)} style={{border:"none",background:"none",color:"#ef4444",cursor:"pointer",fontSize:13,fontWeight:700,padding:"2px 4px"}}>✕</button>
                        </div>
                      </div>

                      {isExp&&(
                        <div style={{borderTop:"1px solid #f1f5f9",padding:"16px",background:"#fafafa"}}>
                          <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(190px,1fr))",gap:10,marginBottom:12}}>
                            <div><label style={{fontSize:11,color:"#94a3b8",display:"block",marginBottom:3,fontWeight:600}}>BUR CODE</label><input style={{width:"100%",border:"1.5px solid #e2e8f0",borderRadius:7,padding:"7px 10px",fontSize:13,outline:"none",fontFamily:"monospace",fontWeight:700,color:"#1d4ed8"}} value={item.code||""} onChange={e=>updBur(item.id,{code:e.target.value})}/></div>
                            <div style={{gridColumn:"span 2"}}><label style={{fontSize:11,color:"#94a3b8",display:"block",marginBottom:3,fontWeight:600}}>DESCRIPTION</label><input style={{width:"100%",border:"1.5px solid #e2e8f0",borderRadius:7,padding:"7px 10px",fontSize:13,outline:"none"}} value={item.desc||""} onChange={e=>updBur(item.id,{desc:e.target.value})}/></div>
                            <div><label style={{fontSize:11,color:"#94a3b8",display:"block",marginBottom:3,fontWeight:600}}>UNIT</label><select style={{width:"100%",border:"1.5px solid #e2e8f0",borderRadius:7,padding:"7px 10px",fontSize:13,outline:"none",background:"#fff"}} value={item.unit||"sum"} onChange={e=>updBur(item.id,{unit:e.target.value})}>{UNITS.map(u=><option key={u}>{u}</option>)}</select></div>
                          </div>
                          <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(160px,1fr))",gap:10,marginBottom:12}}>
                            {COMPS.map(k=>{
                              const hasCostData=(item.costData||[]).filter(e=>e.component===k).length;
                              return(
                                <div key={k}>
                                  <label style={{fontSize:11,color:"#94a3b8",display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:3,fontWeight:600}}>
                                    {CLABEL[k].toUpperCase()}
                                    <button onClick={e=>{e.stopPropagation();setCostModal(item.id);setCType(k);}} style={{fontSize:9,background:hasCostData?"#fef9c3":"#f1f5f9",border:"none",borderRadius:3,padding:"1px 5px",cursor:"pointer",color:hasCostData?"#92400e":"#64748b",fontWeight:600}}>📊{hasCostData?` ${hasCostData}`:""}</button>
                                  </label>
                                  <input type="number" style={{width:"100%",border:k==="subcon"&&isPending?"1.5px solid #f59e0b":k==="subcon"&&isApproved?"1.5px solid #16a34a":"1.5px solid #e2e8f0",borderRadius:7,padding:"7px 10px",fontSize:13,outline:"none",background:k==="subcon"&&isPending?"#fffbeb":k==="subcon"&&isApproved?"#f0fdf4":"#fff"}} value={item[k]??""} onChange={e=>updBur(item.id,{[k]:+e.target.value||0})}/>
                                  {k==="subcon"&&isPending&&<div style={{fontSize:9,color:"#b45309",marginTop:2,fontWeight:600}}>🟡 {item.quote.supplier} · {item.quote.date}</div>}
                                  {k==="subcon"&&isApproved&&<div style={{fontSize:9,color:"#15803d",marginTop:2,fontWeight:600}}>✅ Approved · {item.quote.approvedBy}</div>}
                                </div>
                              );
                            })}
                            <div><label style={{fontSize:11,color:"#94a3b8",display:"block",marginBottom:3,fontWeight:600}}>OH%</label><input type="number" style={{width:"100%",border:"1.5px solid #e2e8f0",borderRadius:7,padding:"7px 10px",fontSize:13,outline:"none"}} value={item.oh??15} onChange={e=>updBur(item.id,{oh:+e.target.value||0})}/></div>
                            <div><label style={{fontSize:11,color:"#94a3b8",display:"block",marginBottom:3,fontWeight:600}}>PROFIT%</label><input type="number" style={{width:"100%",border:"1.5px solid #e2e8f0",borderRadius:7,padding:"7px 10px",fontSize:13,outline:"none"}} value={item.profit??10} onChange={e=>updBur(item.id,{profit:+e.target.value||0})}/></div>
                          </div>
                          <div style={{background:"#eff6ff",borderRadius:8,padding:"10px 14px",display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:8}}>
                            <div style={{fontSize:12,color:"#475569",display:"flex",gap:16,flexWrap:"wrap"}}>
                              <span>Direct: <b>S$ {fmt(direct)}</b></span>
                              <span>OH ({item.oh}%): <b>S$ {fmt(direct*(+item.oh||0)/100)}</b></span>
                            </div>
                            <span style={{fontSize:15,fontWeight:800,color:"#1d4ed8"}}>Total Rate: S$ {fmt(total)} / {item.unit}</span>
                          </div>
                          {isPending&&(
                            <div style={{marginTop:10,background:"#fffbeb",borderRadius:8,padding:"10px 14px",border:"1px solid #fde68a",display:"flex",alignItems:"center",justifyContent:"space-between",flexWrap:"wrap",gap:8}}>
                              <div><span style={{fontSize:12,fontWeight:700,color:"#92400e"}}>🟡 Sub-Con Quotation Pending: </span><span style={{fontSize:12,color:"#92400e"}}>S$ {fmt(item.quote.rate)} · {item.quote.supplier} · {item.quote.date}</span></div>
                              {user.role==="Lead QS"?(
                                <div style={{display:"flex",gap:6}}>
                                  <button onClick={()=>approveQuote(item.id)} style={{background:"#16a34a",color:"#fff",border:"none",borderRadius:6,padding:"5px 12px",fontSize:12,cursor:"pointer",fontWeight:700}}>✓ Approve</button>
                                  <button onClick={()=>rejectQuote(item.id)} style={{background:"#ef4444",color:"#fff",border:"none",borderRadius:6,padding:"5px 10px",fontSize:12,cursor:"pointer"}}>✕ Reject</button>
                                </div>
                              ):<span style={{fontSize:11,color:"#92400e",fontStyle:"italic"}}>Awaiting Lead QS approval</span>}
                            </div>
                          )}
                          {isApproved&&<div style={{marginTop:10,background:"#f0fdf4",borderRadius:8,padding:"9px 14px",border:"1px solid #86efac",fontSize:12,color:"#15803d",fontWeight:600}}>✅ Sub-Con Quote Approved: S$ {fmt(item.quote.rate)} · {item.quote.supplier} · Approved by {item.quote.approvedBy} on {item.quote.approvedAt}</div>}
                        </div>
                      )}
                    </div>
                  );
                })}
              </div>
            )}
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
                <tfoot><tr style={{background:"#1e3a5f",color:"#fff",fontWeight:700,fontSize:14}}><td style={{padding:"12px 16px"}}>GRAND TOTAL</td><td style={{padding:"12px 16px"}}>{tot}</td><td style={{padding:"12px 16px",textAlign:"right"}}>S$ {fmt(gt.A)}</td><td style={{padding:"12px 16px",textAlign:"right"}}>S$ {fmt(gt.B)}</td><td style={{padding:"12px 16px",textAlign:"right"}}>S$ {fmt(gt.A+gt.B)}</td></tr></tfoot>
              </table>
            </div>
            <div style={{background:"#fff",borderRadius:12,boxShadow:"0 1px 4px rgba(0,0,0,.08)",overflow:"hidden"}}>
              <div style={{padding:"12px 16px",borderBottom:"1px solid #f1f5f9",background:"#f8fafc"}}><span style={{fontWeight:700,fontSize:14}}>🏷️ By Cost Code</span></div>
              <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
                <thead><tr style={{background:"#f8fafc",color:"#94a3b8",fontSize:11}}>{["Code","Items","Phase A","Phase B","Total"].map((h,i)=><th key={i} style={{padding:"8px 16px",textAlign:i<2?"left":"right",fontWeight:600,borderBottom:"1px solid #e2e8f0"}}>{h}</th>)}</tr></thead>
                <tbody>{Object.keys(cc).length===0?<tr><td colSpan={5} style={{padding:24,textAlign:"center",color:"#94a3b8"}}>No BUR codes assigned yet</td></tr>:Object.entries(cc).sort((a,b)=>a[0].localeCompare(b[0])).map(([code,v])=><tr key={code} style={{borderBottom:"1px solid #f8fafc"}}><td style={{padding:"9px 16px",fontWeight:700,fontFamily:"monospace",color:"#1d4ed8"}}>{code}</td><td style={{padding:"9px 16px",color:"#64748b"}}>{v.n}</td><td style={{padding:"9px 16px",textAlign:"right",color:"#1d4ed8"}}>{fmt(v.A)}</td><td style={{padding:"9px 16px",textAlign:"right",color:"#7c3aed"}}>{fmt(v.B)}</td><td style={{padding:"9px 16px",textAlign:"right",fontWeight:700}}>{fmt(v.A+v.B)}</td></tr>)}</tbody>
              </table>
            </div>
            {burItems.some(b=>b.quote?.status==="pending")&&(
              <div style={{background:"#fffbeb",borderRadius:12,boxShadow:"0 1px 4px rgba(0,0,0,.08)",overflow:"hidden",border:"1.5px solid #fcd34d"}}>
                <div style={{padding:"12px 16px",borderBottom:"1px solid #fde68a",display:"flex",justifyContent:"space-between"}}><span style={{fontWeight:700,fontSize:14}}>🟡 Pending Sub-Con Approvals</span><span style={{fontSize:12,color:"#92400e",fontWeight:600}}>{burItems.filter(b=>b.quote?.status==="pending").length} items</span></div>
                <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
                  <thead><tr style={{background:"#fef9c3",color:"#92400e",fontSize:11}}>{["Code","Description","Supplier","Date","Rate (S$)",""].map((h,i)=><th key={i} style={{padding:"8px 12px",textAlign:i>=4?"right":"left",fontWeight:600,borderBottom:"1px solid #fde68a"}}>{h}</th>)}</tr></thead>
                  <tbody>{burItems.filter(b=>b.quote?.status==="pending").map(b=><tr key={b.id} style={{borderBottom:"1px solid #fef9c3"}}><td style={{padding:"9px 12px",fontWeight:700,fontFamily:"monospace",color:"#1d4ed8"}}>{b.code||"—"}</td><td style={{padding:"9px 12px"}}>{b.desc}</td><td style={{padding:"9px 12px"}}>{b.quote.supplier}</td><td style={{padding:"9px 12px",color:"#64748b"}}>{b.quote.date}</td><td style={{padding:"9px 12px",textAlign:"right",fontWeight:700,color:"#b45309"}}>S$ {fmt(b.quote.rate)}</td><td style={{padding:"9px 8px",textAlign:"right"}}>{user.role==="Lead QS"&&<div style={{display:"flex",gap:4,justifyContent:"flex-end"}}><button onClick={()=>approveQuote(b.id)} style={{background:"#16a34a",color:"#fff",border:"none",borderRadius:5,padding:"3px 10px",fontSize:11,cursor:"pointer",fontWeight:700}}>✓ Approve</button><button onClick={()=>rejectQuote(b.id)} style={{background:"#ef4444",color:"#fff",border:"none",borderRadius:5,padding:"3px 8px",fontSize:11,cursor:"pointer"}}>✕</button></div>}</td></tr>)}</tbody>
                </table>
              </div>
            )}
            <div style={{background:"#fff",borderRadius:12,padding:16,boxShadow:"0 1px 4px rgba(0,0,0,.08)"}}>
              <div style={{fontWeight:700,fontSize:13,marginBottom:12}}>Item Status</div>
              <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:10}}>
                {STATUSES.map(s=>{const n=data.sections.reduce((a,sec)=>a+(sec.items?.filter(i=>i.status===s).length||0),0),st=SS[s];return(<div key={s} style={{background:st.bg,borderRadius:10,padding:"14px 10px",textAlign:"center"}}><div style={{fontSize:26,fontWeight:800,color:st.c}}>{n}</div><div style={{fontSize:11,fontWeight:600,color:st.c,marginTop:2}}>{s}</div></div>);})}
              </div>
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
          <div style={{background:"#fff",borderRadius:16,width:"100%",maxWidth:720,maxHeight:"88vh",display:"flex",flexDirection:"column",boxShadow:"0 25px 60px rgba(0,0,0,.3)"}}>
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
            <div style={{display:"flex",borderBottom:"1px solid #e2e8f0",padding:"0 20px",flexShrink:0}}>
              {COMPS.map(c=>{const cnt=(modalItem.costData||[]).filter(e=>e.component===c).length;return(
                <button key={c} onClick={()=>setCType(c)} style={{padding:"9px 14px",fontSize:12,fontWeight:600,border:"none",borderBottom:cType===c?"2.5px solid #2563eb":"2.5px solid transparent",color:cType===c?"#2563eb":"#64748b",background:"none",cursor:"pointer",whiteSpace:"nowrap"}}>
                  {CLABEL[c]}{cnt>0&&<span style={{marginLeft:4,background:cType===c?"#dbeafe":"#f1f5f9",color:cType===c?"#1d4ed8":"#64748b",borderRadius:10,padding:"0 5px",fontSize:10}}>{cnt}</span>}
                </button>
              );})}
            </div>
            <div style={{flex:1,overflow:"auto",padding:"16px 20px"}}>
              {costEntries.length===0?(
                <div style={{textAlign:"center",padding:"28px 0",color:"#94a3b8",fontSize:13}}>No {CLABEL[cType]} entries yet — add one below</div>
              ):(
                <table style={{width:"100%",borderCollapse:"collapse",fontSize:12,marginBottom:16}}>
                  <thead><tr style={{background:"#f8fafc",color:"#94a3b8",fontSize:11}}>
                    {["Supplier / Vendor","Rate (S$)","Date","Note",""].map((h,i)=><th key={i} style={{padding:"8px 10px",textAlign:i===1?"right":"left",fontWeight:600,borderBottom:"1px solid #e2e8f0"}}>{h}</th>)}
                  </tr></thead>
                  <tbody>
                    {costEntries.map(e=>(
                      <tr key={e.id} style={{borderBottom:"1px solid #f8fafc"}}>
                        <td style={{padding:"10px 10px",fontWeight:600,color:"#1e293b"}}>{e.supplier}</td>
                        <td style={{padding:"10px 10px",textAlign:"right",color:"#1d4ed8",fontWeight:700,fontSize:14}}>S$ {fmt(e.rate)}</td>
                        <td style={{padding:"10px 10px",color:"#64748b"}}>{e.date||"—"}</td>
                        <td style={{padding:"10px 10px",color:"#94a3b8",fontSize:11}}>{e.note||"—"}</td>
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
                <div style={{fontSize:12,fontWeight:600,color:"#475569",marginBottom:10}}>+ Add {CLABEL[cType]} Quote / Price</div>
                <div style={{display:"grid",gridTemplateColumns:"2fr 1fr 1fr 2fr",gap:8,marginBottom:10}}>
                  {[["supplier","Supplier / Vendor","text","e.g. ABC Contractors Pte Ltd"],["rate","Rate (S$)","number","0.00"],["date","Date","text","e.g. Jun 2025"],["note","Note","text","e.g. FOB, ex-GST"]].map(([k,l,t,ph])=>(
                    <div key={k}><div style={{fontSize:10,color:"#94a3b8",marginBottom:3,fontWeight:600}}>{l}</div>
                    <input type={t} style={{width:"100%",border:"1.5px solid #e2e8f0",borderRadius:7,padding:"7px 9px",fontSize:12,outline:"none",boxSizing:"border-box"}} placeholder={ph} value={cForm[k]} onChange={e=>setCForm(f=>({...f,[k]:e.target.value}))}/></div>
                  ))}
                </div>
                <div style={{display:"flex",gap:8,justifyContent:"flex-end"}}>
                  <button onClick={()=>setCForm({supplier:"",rate:"",date:"",note:""})} style={{background:"#f1f5f9",border:"none",borderRadius:7,padding:"7px 14px",fontSize:12,cursor:"pointer",color:"#64748b"}}>Clear</button>
                  <button onClick={addCostEntry} style={{background:"#2563eb",color:"#fff",border:"none",borderRadius:7,padding:"7px 18px",fontSize:12,cursor:"pointer",fontWeight:600}}>Add Entry</button>
                </div>
              </div>
              {cType==="subcon"&&<div style={{marginTop:10,background:"#fef9c3",borderRadius:8,padding:"8px 12px",fontSize:11,color:"#92400e"}}>💡 Click <b>"→ Use as Quote"</b> to set this supplier rate as the Sub-Con Quotation. It highlights 🟡 until the Lead QS approves it.</div>}
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
