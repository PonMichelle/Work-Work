import { useState, useEffect, useRef, useCallback } from "react";
import { doc, onSnapshot, setDoc } from "firebase/firestore";

const argbToCss=argb=>{ if(!argb)return null; const s=String(argb); if(s.length===8)return "#"+s.slice(2); if(s.length===6)return "#"+s; return null; };
const colLetter=n=>{ let s=""; while(n>0){const m=(n-1)%26; s=String.fromCharCode(65+m)+s; n=Math.floor((n-1)/26);} return s; };
const a1=addr=>{ const m=/^([A-Z]+)(\d+)$/.exec(addr); if(!m)return null; let c=0; for(const ch of m[1])c=c*26+(ch.charCodeAt(0)-64); return {r:+m[2],c}; };
const bufToB64=buf=>{ const b=new Uint8Array(buf); let s=""; const CH=0x8000; for(let i=0;i<b.length;i+=CH)s+=String.fromCharCode.apply(null,b.subarray(i,i+CH)); return btoa(s); };
const b64ToBytes=b64=>Uint8Array.from(atob(b64),c=>c.charCodeAt(0));
const ck=(si,r,c)=>`S${si}!${r}:${c}`, cwk=(si,c)=>`S${si}!C${c}`, rhk=(si,r)=>`S${si}!R${r}`;
const COLW_DEF=9, MAXR=1200, MAXC=60;

// Evaluate an "=" formula referencing other cells (A1), with SUM/ROUND/ROUNDUP/AVERAGE/MIN/MAX/ABS/IF.
function evalFormula(expr,getVal,depth){
  if(depth>25)return "#REF";
  let e=String(expr).trim().slice(1);
  e=e.replace(/\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)/g,(m,c1,r1,c2,r2)=>{ const A=a1(c1+r1),B=a1(c2+r2); if(!A||!B)return "0"; const v=[]; for(let r=Math.min(A.r,B.r);r<=Math.max(A.r,B.r);r++)for(let c=Math.min(A.c,B.c);c<=Math.max(A.c,B.c);c++)v.push(getVal(r,c,depth+1)); return "["+v.join(",")+"]"; });
  e=e.replace(/\$?([A-Z]+)\$?(\d+)/g,(m,col,row)=>{ const A=a1(col+row); return A?String(getVal(A.r,A.c,depth+1)):"0"; });
  const SUM=(...a)=>a.flat(9).reduce((s,x)=>s+(+x||0),0);
  const ROUND=(x,n=0)=>{const f=Math.pow(10,n);return Math.round((+x||0)*f)/f;};
  const ROUNDUP=(x,n=0)=>{const f=Math.pow(10,n);return Math.ceil((+x||0)*f)/f;};
  const AVERAGE=(...a)=>{const f=a.flat(9);return f.length?SUM(...f)/f.length:0;};
  const MIN=(...a)=>Math.min(...a.flat(9).map(Number)); const MAX=(...a)=>Math.max(...a.flat(9).map(Number));
  const ABS=x=>Math.abs(+x||0); const IF=(c,t,f)=>c?t:f;
  try{ const fn=new Function("SUM","ROUND","ROUNDUP","AVERAGE","MIN","MAX","ABS","IF",`return (${e});`); const r=fn(SUM,ROUND,ROUNDUP,AVERAGE,MIN,MAX,ABS,IF); return (typeof r==="number"&&isFinite(r))?r:r; }catch{ return "#ERR"; }
}
const fmtNum=r=>{ if(typeof r!=="number")return String(r); if(!isFinite(r))return "#ERR"; return Number.isInteger(r)?String(r):String(Math.round(r*100)/100); };

export default function SheetGrid({ db, pid, toast, baseUrl, burLookup }){
  const [b64,setB64]=useState(null); const [fileName,setFileName]=useState("");
  const [edits,setEdits]=useState({cells:{},colW:{},rowH:{},active:0});
  const [grid,setGrid]=useState(null); const [sheets,setSheets]=useState([]);
  const [busy,setBusy]=useState(false); const [dragLive,setDragLive]=useState(null);
  const wbRef=useRef(null); const saveT=useRef(null);
  const editsRef=useRef(edits); useEffect(()=>{editsRef.current=edits;},[edits]);
  const dragLiveRef=useRef(dragLive); useEffect(()=>{dragLiveRef.current=dragLive;},[dragLive]);
  const dragRef=useRef(null);

  useEffect(()=>{ if(!pid)return; return onSnapshot(doc(db,"boqfile",pid),s=>{ if(s.exists()){setB64(s.data().b64||null);setFileName(s.data().name||"");}else{setB64(null);setFileName("");} }); },[pid,db]);
  useEffect(()=>{ if(!pid)return; return onSnapshot(doc(db,"boqedits",pid),s=>{ if(s.exists())setEdits({cells:{},colW:{},rowH:{},active:0,...s.data()}); }); },[pid,db]);

  const buildGrid=useCallback((wb,idx)=>{
    const ws=wb.worksheets[idx]; if(!ws){setGrid(null);return;}
    const right=Math.min(ws.columnCount||1,MAXC), bottom=Math.min(ws.rowCount||1,MAXR);
    const cols=[]; for(let c=1;c<=right;c++){const w=ws.getColumn(c).width; cols.push(Math.round((w||COLW_DEF)*7+6));}
    let merges=[]; try{merges=ws.model&&ws.model.merges?ws.model.merges.slice():[];}catch{}
    const rows=[];
    for(let r=1;r<=bottom;r++){
      const row=ws.getRow(r); const h=row.height?Math.round(row.height*1.34):21; const cells=[];
      for(let c=1;c<=right;c++){
        const cell=row.getCell(c); const st=cell.style||{}; const font=st.font||{};
        let bg=null; try{ if(st.fill&&st.fill.fgColor)bg=argbToCss(st.fill.fgColor.argb); }catch{}
        let txt=""; try{txt=cell.text==null?"":String(cell.text);}catch{txt="";}
        cells.push({v:txt,bg,bold:!!font.bold,italic:!!font.italic,size:font.size||10,color:font.color?argbToCss(font.color.argb):null,
          align:(st.alignment&&st.alignment.horizontal)||null,valign:(st.alignment&&st.alignment.vertical)||"bottom",wrap:!!(st.alignment&&st.alignment.wrapText),
          bL:!!(st.border&&st.border.left),bR:!!(st.border&&st.border.right),bT:!!(st.border&&st.border.top),bB:!!(st.border&&st.border.bottom)});
      }
      rows.push({h,cells});
    }
    // detect header columns (CODE / UNIT / Material / Labour / U/RATE) for the BUR rate link
    let meta=null,hr=0,cCode=0,cUnit=0,cMat=0,cLab=0,cRate=0;
    for(let r=1;r<=Math.min(14,bottom)&&!hr;r++){ let found=false; for(let c=1;c<=right;c++){ const t=(rows[r-1].cells[c-1].v||"").trim().toUpperCase(); if(t==="CODE"){cCode=c;found=true;} if(t==="UNIT")cUnit=c; if(t.includes("MATERIAL"))cMat=c; if(t.includes("LABOUR"))cLab=c; if(t.includes("U/RATE"))cRate=c; } if(found)hr=r; }
    if(hr&&cCode)meta={hr,cCode,cUnit,cMat,cLab,cRate};
    setGrid({right,bottom,cols,rows,merges,name:ws.name,meta});
  },[]);

  useEffect(()=>{ let cancel=false; (async()=>{
    if(!b64){setGrid(null);setSheets([]);wbRef.current=null;return;}
    setBusy(true);
    try{ const ExcelJS=(await import("exceljs")).default; const wb=new ExcelJS.Workbook(); await wb.xlsx.load(b64ToBytes(b64)); if(cancel)return;
      wbRef.current=wb; const names=wb.worksheets.map(w=>w.name); setSheets(names); buildGrid(wb,Math.min(editsRef.current.active||0,names.length-1));
    }catch(e){ toast("⚠️ Could not read workbook: "+(e&&e.message||e)); }
    if(!cancel)setBusy(false);
  })(); return ()=>{cancel=true;}; },[b64,buildGrid,toast]);

  const persist=useCallback(ne=>{ if(saveT.current)clearTimeout(saveT.current); saveT.current=setTimeout(()=>setDoc(doc(db,"boqedits",pid),ne).catch(()=>{}),600); },[db,pid]);
  const persistNow=useCallback(ne=>setDoc(doc(db,"boqedits",pid),ne).catch(()=>{}),[db,pid]);
  const setActive=idx=>{ const ne={...editsRef.current,active:idx}; setEdits(ne); persistNow(ne); if(wbRef.current)buildGrid(wbRef.current,idx); };

  const commitCell=(r,c,val)=>{
    const orig=grid?.rows[r-1]?.cells[c-1]?.v??""; const key=ck(edits.active,r,c);
    const cells={...edits.cells};
    if(val===orig)delete cells[key]; else cells[key]=val;
    const mt=grid?.meta;
    if(mt&&c===mt.cCode&&val&&val.trim()&&burLookup){ const info=burLookup(val); if(info){ if(mt.cMat&&mt.cLab){cells[ck(edits.active,r,mt.cMat)]=String(info.material);cells[ck(edits.active,r,mt.cLab)]=String(info.labour);} else if(mt.cRate){cells[ck(edits.active,r,mt.cRate)]=String(info.total);} toast(`✅ Rate from BUR "${val.trim()}" filled`); } }
    const ne={...edits,cells}; setEdits(ne); persist(ne);
  };

  useEffect(()=>{
    const mm=e=>{ const d=dragRef.current; if(!d)return; if(d.kind==="col")setDragLive({kind:"col",idx:d.idx,px:Math.max(24,d.startPx+(e.clientX-d.start))}); else setDragLive({kind:"row",idx:d.idx,px:Math.max(14,d.startPx+(e.clientY-d.start))}); };
    const mu=()=>{ const d=dragRef.current,live=dragLiveRef.current; if(d&&live){ const ne={...editsRef.current}; if(d.kind==="col")ne.colW={...ne.colW,[cwk(editsRef.current.active,d.idx)]:live.px}; else ne.rowH={...ne.rowH,[rhk(editsRef.current.active,d.idx)]:live.px}; setEdits(ne); persistNow(ne); } dragRef.current=null; setDragLive(null); };
    window.addEventListener("mousemove",mm); window.addEventListener("mouseup",mu);
    return ()=>{window.removeEventListener("mousemove",mm);window.removeEventListener("mouseup",mu);};
  },[persistNow]);
  const startDrag=(kind,idx,startPx,e)=>{ e.preventDefault();e.stopPropagation(); dragRef.current={kind,idx,start:kind==="col"?e.clientX:e.clientY,startPx}; setDragLive({kind,idx,px:startPx}); };

  const loadBytes=useCallback(async(buf,name)=>{ const b=bufToB64(buf); if(b.length>980000){toast("⚠️ File too large (keep .xlsx under ~700 KB)");return;} try{ await setDoc(doc(db,"boqfile",pid),{b64:b,name:name||"sheet.xlsx"}); await setDoc(doc(db,"boqedits",pid),{cells:{},colW:{},rowH:{},active:0}); toast("✅ Loaded "+(name||"workbook")); }catch(e){toast("⚠️ Save failed: "+e.message);} },[db,pid,toast]);
  const onFile=async file=>{ if(!file)return; if(!/\.xlsx$/i.test(file.name)){toast("⚠️ Please use a .xlsx file");return;} setBusy(true); try{await loadBytes(await file.arrayBuffer(),file.name);}finally{setBusy(false);} };
  const onDrop=e=>{ e.preventDefault(); const f=e.dataTransfer.files&&e.dataTransfer.files[0]; if(f)onFile(f); };
  const useTemplate=async()=>{ setBusy(true); try{ const r=await fetch(baseUrl+"boq-template.xlsx"); await loadBytes(await r.arrayBuffer(),"CAG_Dormitory_Master_BQ.xlsx"); }catch(e){toast("⚠️ "+e.message);} finally{setBusy(false);} };

  const exportXlsx=useCallback(async()=>{
    if(!b64){toast("Nothing to export");return;} setBusy(true);
    try{ const ExcelJS=(await import("exceljs")).default; const wb=new ExcelJS.Workbook(); await wb.xlsx.load(b64ToBytes(b64));
      for(const [key,val] of Object.entries(edits.cells||{})){ const m=/^S(\d+)!(\d+):(\d+)$/.exec(key); if(!m)continue; const ws=wb.worksheets[+m[1]]; if(!ws)continue; const cell=ws.getRow(+m[2]).getCell(+m[3]); const s=String(val); if(s.trim().startsWith("=")){cell.value={formula:s.trim().slice(1)};} else { const num=Number(s.replace(/,/g,"")); cell.value=(s!==""&&/^-?[\d.,]+$/.test(s)&&!isNaN(num))?num:s; } }
      for(const [key,px] of Object.entries(edits.colW||{})){ const m=/^S(\d+)!C(\d+)$/.exec(key); if(!m)continue; const ws=wb.worksheets[+m[1]]; if(ws)ws.getColumn(+m[2]).width=Math.max(2,(px-6)/7); }
      for(const [key,px] of Object.entries(edits.rowH||{})){ const m=/^S(\d+)!R(\d+)$/.exec(key); if(!m)continue; const ws=wb.worksheets[+m[1]]; if(ws)ws.getRow(+m[2]).height=Math.max(6,px/1.34); }
      const out=await wb.xlsx.writeBuffer(); const blob=new Blob([out],{type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});
      const url=URL.createObjectURL(blob); const a=document.createElement("a"); a.href=url; a.download=(fileName||"BOQ").replace(/\.xlsx$/i,"")+"_export.xlsx"; a.click(); URL.revokeObjectURL(url); toast("✅ Exported "+a.download);
    }catch(e){toast("⚠️ Export failed: "+(e&&e.message||e));} finally{setBusy(false);}
  },[b64,edits,fileName,toast]);

  if(!b64)return(
    <div onDragOver={e=>e.preventDefault()} onDrop={onDrop} style={{background:"#fff",border:"2px dashed #cbd5e1",borderRadius:14,padding:48,textAlign:"center",color:"#64748b"}}>
      <div style={{fontSize:42,marginBottom:10}}>📄⬇️</div>
      <div style={{fontWeight:700,fontSize:16,color:"#1e293b",marginBottom:6}}>Drag your Excel (.xlsx) here</div>
      <div style={{fontSize:13,marginBottom:16}}>Loads with the same columns, rows, colours & formatting — editable, resizable, exportable. Type a BUR code in the CODE column to auto-fill its rate.</div>
      <div style={{display:"flex",gap:10,justifyContent:"center",flexWrap:"wrap"}}>
        <label style={{background:"#2563eb",color:"#fff",borderRadius:10,padding:"10px 18px",fontSize:13,fontWeight:600,cursor:"pointer"}}>Choose file<input type="file" accept=".xlsx" style={{display:"none"}} onChange={e=>onFile(e.target.files[0])}/></label>
        <button onClick={useTemplate} disabled={busy} style={{background:"#7c3aed",color:"#fff",border:"none",borderRadius:10,padding:"10px 18px",fontSize:13,fontWeight:600,cursor:"pointer"}}>{busy?"Loading…":"Use CAG master template"}</button>
      </div>
    </div>
  );

  const cw=c=>(dragLive&&dragLive.kind==="col"&&dragLive.idx===c)?dragLive.px:(edits.colW?.[cwk(edits.active,c)]??grid?.cols[c-1]??70);
  const rh=r=>(dragLive&&dragLive.kind==="row"&&dragLive.idx===r)?dragLive.px:(edits.rowH?.[rhk(edits.active,r)]??grid?.rows[r-1]?.h??21);

  const getVal=(r,c,depth=0)=>{ if(depth>25)return 0; const raw=editsRef.current.cells?.[ck(editsRef.current.active,r,c)] ?? grid?.rows[r-1]?.cells[c-1]?.v ?? ""; if(typeof raw==="string"&&raw.trim().startsWith("=")){const v=evalFormula(raw.trim(),getVal,depth+1);return typeof v==="number"?v:(parseFloat(v)||0);} const n=parseFloat(String(raw).replace(/,/g,"")); return isNaN(n)?0:n; };

  let covered=new Set(), spanMap={};
  if(grid){ for(const rng of grid.merges){ const [a,b]=String(rng).split(":"); const A=a1(a),B=a1(b); if(!A||!B)continue; spanMap[`${A.r}_${A.c}`]={rs:B.r-A.r+1,cs:B.c-A.c+1}; for(let r=A.r;r<=B.r;r++)for(let c=A.c;c<=B.c;c++){if(!(r===A.r&&c===A.c))covered.add(`${r}_${c}`);} } }

  return(
    <div onDragOver={e=>e.preventDefault()} onDrop={onDrop}>
      <div style={{display:"flex",gap:8,alignItems:"center",marginBottom:8,flexWrap:"wrap"}}>
        <span style={{fontWeight:700,fontSize:13,color:"#1e293b"}}>📄 {fileName||"Sheet"}</span>
        {busy&&<span style={{fontSize:12,color:"#2563eb"}}>working…</span>}
        {grid&&grid.meta&&<span style={{fontSize:11,color:"#15803d",background:"#dcfce7",borderRadius:6,padding:"2px 8px"}}>🔗 BUR code link active (col {colLetter(grid.meta.cCode)})</span>}
        <div style={{marginLeft:"auto",display:"flex",gap:8,flexWrap:"wrap"}}>
          <button onClick={exportXlsx} style={{background:"#15803d",color:"#fff",border:"none",borderRadius:8,padding:"6px 12px",fontSize:12,fontWeight:600,cursor:"pointer"}}>⤓ Export Excel</button>
          <label style={{background:"#2563eb",color:"#fff",borderRadius:8,padding:"6px 12px",fontSize:12,fontWeight:600,cursor:"pointer"}}>Replace file<input type="file" accept=".xlsx" style={{display:"none"}} onChange={e=>onFile(e.target.files[0])}/></label>
        </div>
      </div>
      <div style={{display:"flex",gap:4,marginBottom:8,overflowX:"auto"}}>
        {sheets.map((nm,i)=><button key={i} onClick={()=>setActive(i)} style={{padding:"4px 10px",borderRadius:"6px 6px 0 0",fontSize:12,fontWeight:600,border:"1px solid #e2e8f0",borderBottom:"none",cursor:"pointer",whiteSpace:"nowrap",background:edits.active===i?"#1e3a5f":"#f8fafc",color:edits.active===i?"#fff":"#475569"}}>{nm}</button>)}
      </div>

      {!grid?<div style={{padding:40,textAlign:"center",color:"#94a3b8"}}>Loading sheet…</div>:(
        <div style={{overflow:"auto",border:"1px solid #e2e8f0",borderRadius:8,background:"#fff",maxHeight:"72vh"}}>
          <table style={{borderCollapse:"collapse",tableLayout:"fixed",fontSize:11}}>
            <colgroup><col style={{width:44}}/>{Array.from({length:grid.right},(_,i)=><col key={i} style={{width:cw(i+1)}}/>)}</colgroup>
            <thead>
              <tr>
                <th style={{position:"sticky",top:0,left:0,zIndex:4,background:"#e2e8f0",border:"1px solid #cbd5e1",height:22}}></th>
                {Array.from({length:grid.right},(_,i)=>(
                  <th key={i} style={{position:"sticky",top:0,zIndex:3,background:"#e2e8f0",border:"1px solid #cbd5e1",fontSize:10,color:"#475569",fontWeight:700,padding:0,height:22,userSelect:"none"}}>
                    <div style={{padding:"2px 4px",textAlign:"center"}}>{colLetter(i+1)}</div>
                    <div onMouseDown={e=>startDrag("col",i+1,cw(i+1),e)} title="Drag to resize column" style={{position:"absolute",top:0,right:0,width:8,height:"100%",cursor:"col-resize",zIndex:6,background:dragLive&&dragLive.kind==="col"&&dragLive.idx===i+1?"#2563eb":"transparent"}}/>
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {grid.rows.map((row,ri)=>{ const r=ri+1; return(
                <tr key={r} style={{height:rh(r)}}>
                  <td style={{position:"sticky",left:0,zIndex:2,background:"#e2e8f0",border:"1px solid #cbd5e1",textAlign:"center",fontSize:10,color:"#475569",fontWeight:700,userSelect:"none",padding:0}}>
                    <div style={{height:rh(r),display:"flex",alignItems:"center",justifyContent:"center"}}>{r}</div>
                    <div onMouseDown={e=>startDrag("row",r,rh(r),e)} title="Drag to resize row" style={{position:"absolute",left:0,bottom:0,height:8,width:"100%",cursor:"row-resize",zIndex:3,background:dragLive&&dragLive.kind==="row"&&dragLive.idx===r?"#2563eb":"transparent"}}/>
                  </td>
                  {row.cells.map((cell,ci)=>{ const c=ci+1; if(covered.has(`${r}_${c}`))return null; const sp=spanMap[`${r}_${c}`]; const key=ck(edits.active,r,c); const raw=edits.cells?.[key]??cell.v; const disp=(typeof raw==="string"&&raw.trim().startsWith("="))?fmtNum(evalFormula(raw.trim(),getVal,0)):raw;
                    return(
                      <td key={c} colSpan={sp?sp.cs:1} rowSpan={sp?sp.rs:1} contentEditable suppressContentEditableWarning data-rc={`${r}_${c}`}
                        onKeyDown={e=>{ if(e.key==="Enter"){e.preventDefault();e.currentTarget.blur();const t=document.querySelector(`[data-rc="${r+1}_${c}"]`);if(t){t.focus();const sel=window.getSelection();sel.selectAllChildren(t);sel.collapseToEnd();}} else if(e.key==="Tab"){e.preventDefault();e.currentTarget.blur();const t=document.querySelector(`[data-rc="${r}_${c+(e.shiftKey?-1:1)}"]`);if(t)t.focus();} else if(e.key==="Escape"){e.currentTarget.blur();} }}
                        onBlur={e=>commitCell(r,c,e.currentTarget.innerText)}
                        style={{border:"1px solid #eef2f7",borderLeft:cell.bL?"1px solid #94a3b8":undefined,borderTop:cell.bT?"1px solid #94a3b8":undefined,borderRight:cell.bR?"1px solid #94a3b8":undefined,borderBottom:cell.bB?"1px solid #94a3b8":undefined,
                          background:cell.bg||"transparent",color:cell.color||"#1e293b",fontWeight:cell.bold?700:400,fontStyle:cell.italic?"italic":"normal",fontSize:Math.max(9,Math.min(16,cell.size)),
                          textAlign:cell.align||"left",verticalAlign:cell.valign==="middle"||cell.valign==="center"?"middle":cell.valign==="top"?"top":"bottom",whiteSpace:cell.wrap?"normal":"nowrap",overflow:"hidden",padding:"1px 4px",outline:"none"}}>
                        {disp}
                      </td>
                    );
                  })}
                </tr>
              );})}
            </tbody>
          </table>
        </div>
      )}
      <div style={{fontSize:11,color:"#94a3b8",marginTop:6}}>Resize: drag the right edge of a column letter / bottom edge of a row number. Type <b>=A1*B1</b> style formulas (SUM, ROUND supported) for live totals. Type a BUR code in the CODE column to pull its rate. {grid&&(grid.bottom>=MAXR||grid.right>=MAXC)?`(First ${grid.bottom}×${grid.right} shown.)`:""}</div>
    </div>
  );
}
