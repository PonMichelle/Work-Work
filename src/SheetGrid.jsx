import { useState, useEffect, useRef, useCallback } from "react";
import { doc, onSnapshot, setDoc } from "firebase/firestore";

// ── helpers ───────────────────────────────────────────────────────────────
const argbToCss=argb=>{ if(!argb)return null; const s=String(argb); if(s.length===8)return "#"+s.slice(2); if(s.length===6)return "#"+s; return null; };
const colLetter=n=>{ let s=""; while(n>0){const m=(n-1)%26; s=String.fromCharCode(65+m)+s; n=Math.floor((n-1)/26);} return s; };
const a1=addr=>{ const m=/^([A-Z]+)(\d+)$/.exec(addr); if(!m)return null; let c=0; for(const ch of m[1])c=c*26+(ch.charCodeAt(0)-64); return {r:+m[2],c}; };
const bufToB64=buf=>{ const b=new Uint8Array(buf); let s=""; const CH=0x8000; for(let i=0;i<b.length;i+=CH)s+=String.fromCharCode.apply(null,b.subarray(i,i+CH)); return btoa(s); };
const b64ToBytes=b64=>Uint8Array.from(atob(b64),c=>c.charCodeAt(0));
const ck=(si,r,c)=>`S${si}!${r}:${c}`, cwk=(si,c)=>`S${si}!C${c}`, rhk=(si,r)=>`S${si}!R${r}`;
const COLW_DEF=9, MAXR=1200, MAXC=60;

export default function SheetGrid({ db, pid, toast, baseUrl }){
  const [b64,setB64]=useState(null); const [fileName,setFileName]=useState("");
  const [edits,setEdits]=useState({cells:{},colW:{},rowH:{},active:0});
  const [grid,setGrid]=useState(null); const [sheets,setSheets]=useState([]);
  const [busy,setBusy]=useState(false); const [dragLive,setDragLive]=useState(null);
  const wbRef=useRef(null); const saveT=useRef(null); const editsRef=useRef(edits); const dragRef=useRef(null);
  useEffect(()=>{editsRef.current=edits;},[edits]);

  // subscribe to the source file + edit overlay (per project)
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
    setGrid({right,bottom,cols,rows,merges,name:ws.name});
  },[]);

  // parse workbook whenever the source bytes change
  useEffect(()=>{ let cancel=false; (async()=>{
    if(!b64){setGrid(null);setSheets([]);wbRef.current=null;return;}
    setBusy(true);
    try{
      const ExcelJS=(await import("exceljs")).default; const wb=new ExcelJS.Workbook();
      await wb.xlsx.load(b64ToBytes(b64)); if(cancel)return;
      wbRef.current=wb; const names=wb.worksheets.map(w=>w.name); setSheets(names);
      buildGrid(wb,Math.min(editsRef.current.active||0,names.length-1));
    }catch(e){ toast("⚠️ Could not read workbook: "+(e&&e.message||e)); }
    if(!cancel)setBusy(false);
  })(); return ()=>{cancel=true;}; },[b64,buildGrid,toast]);

  const persist=useCallback(ne=>{ if(saveT.current)clearTimeout(saveT.current); saveT.current=setTimeout(()=>{ setDoc(doc(db,"boqedits",pid),ne).catch(()=>{}); },600); },[db,pid]);
  const persistNow=useCallback(ne=>{ setDoc(doc(db,"boqedits",pid),ne).catch(()=>{}); },[db,pid]);

  const setActive=idx=>{ const ne={...editsRef.current,active:idx}; setEdits(ne); persistNow(ne); if(wbRef.current)buildGrid(wbRef.current,idx); };
  const commitCell=(r,c,val)=>{ const key=ck(edits.active,r,c); const cur=edits.cells?.[key]; const orig=grid?.rows[r-1]?.cells[c-1]?.v??""; const norm=val===orig?undefined:val; if((cur??undefined)===(norm??undefined))return; const cells={...edits.cells}; if(norm===undefined)delete cells[key]; else cells[key]=val; const ne={...edits,cells}; setEdits(ne); persist(ne); };

  // resize via drag
  useEffect(()=>{
    const mm=e=>{ const d=dragRef.current; if(!d)return; if(d.kind==="col"){const px=Math.max(24,d.startPx+(e.clientX-d.start));setDragLive({kind:"col",idx:d.idx,px});}else{const px=Math.max(14,d.startPx+(e.clientY-d.start));setDragLive({kind:"row",idx:d.idx,px});} };
    const mu=()=>{ const d=dragRef.current, live=dragLiveRef.current; if(d&&live){ const ne={...editsRef.current}; if(d.kind==="col"){ne.colW={...ne.colW,[cwk(editsRef.current.active,d.idx)]:live.px};}else{ne.rowH={...ne.rowH,[rhk(editsRef.current.active,d.idx)]:live.px};} setEdits(ne); persistNow(ne); } dragRef.current=null; setDragLive(null); };
    window.addEventListener("mousemove",mm); window.addEventListener("mouseup",mu);
    return ()=>{window.removeEventListener("mousemove",mm);window.removeEventListener("mouseup",mu);};
  },[persistNow]);
  const dragLiveRef=useRef(dragLive); useEffect(()=>{dragLiveRef.current=dragLive;},[dragLive]);
  const startDrag=(kind,idx,startPx,e)=>{ e.preventDefault();e.stopPropagation(); dragRef.current={kind,idx,start:kind==="col"?e.clientX:e.clientY,startPx}; };

  // load a workbook (dropped/picked file, or the bundled CAG template)
  const loadBytes=useCallback(async(buf,name)=>{
    const b=bufToB64(buf);
    if(b.length>980000){ toast("⚠️ File too large to store (keep the .xlsx under ~700 KB)"); return; }
    try{ await setDoc(doc(db,"boqfile",pid),{b64:b,name:name||"sheet.xlsx"}); await setDoc(doc(db,"boqedits",pid),{cells:{},colW:{},rowH:{},active:0}); toast("✅ Loaded "+(name||"workbook")); }
    catch(e){ toast("⚠️ Save failed: "+e.message); }
  },[db,pid,toast]);
  const onFile=async file=>{ if(!file)return; if(!/\.xlsx$/i.test(file.name)){toast("⚠️ Please use a .xlsx file");return;} setBusy(true); try{await loadBytes(await file.arrayBuffer(),file.name);}finally{setBusy(false);} };
  const onDrop=e=>{ e.preventDefault(); const f=e.dataTransfer.files&&e.dataTransfer.files[0]; if(f)onFile(f); };
  const useTemplate=async()=>{ setBusy(true); try{ const r=await fetch(baseUrl+"boq-template.xlsx"); await loadBytes(await r.arrayBuffer(),"CAG_Dormitory_Master_BQ.xlsx"); }catch(e){toast("⚠️ "+e.message);} finally{setBusy(false);} };

  const exportXlsx=useCallback(async()=>{
    if(!b64){toast("Nothing to export");return;} setBusy(true);
    try{
      const ExcelJS=(await import("exceljs")).default; const wb=new ExcelJS.Workbook(); await wb.xlsx.load(b64ToBytes(b64));
      for(const [key,val] of Object.entries(edits.cells||{})){ const m=/^S(\d+)!(\d+):(\d+)$/.exec(key); if(!m)continue; const ws=wb.worksheets[+m[1]]; if(!ws)continue; const cell=ws.getRow(+m[2]).getCell(+m[3]); const num=Number(String(val).replace(/,/g,"")); cell.value=(val!==""&&val!=null&&/^-?[\d.,]+%?$/.test(String(val))&&!isNaN(num))?num:val; }
      for(const [key,px] of Object.entries(edits.colW||{})){ const m=/^S(\d+)!C(\d+)$/.exec(key); if(!m)continue; const ws=wb.worksheets[+m[1]]; if(ws)ws.getColumn(+m[2]).width=Math.max(2,(px-6)/7); }
      for(const [key,px] of Object.entries(edits.rowH||{})){ const m=/^S(\d+)!R(\d+)$/.exec(key); if(!m)continue; const ws=wb.worksheets[+m[1]]; if(ws)ws.getRow(+m[2]).height=Math.max(6,px/1.34); }
      const out=await wb.xlsx.writeBuffer(); const blob=new Blob([out],{type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});
      const url=URL.createObjectURL(blob); const a=document.createElement("a"); a.href=url; a.download=(fileName||"BOQ").replace(/\.xlsx$/i,"")+"_export.xlsx"; a.click(); URL.revokeObjectURL(url);
      toast("✅ Exported "+a.download);
    }catch(e){ toast("⚠️ Export failed: "+(e&&e.message||e)); } finally{setBusy(false);}
  },[b64,edits,fileName,toast]);

  // ── empty state ───────────────────────────────────────────────────────────
  if(!b64)return(
    <div onDragOver={e=>e.preventDefault()} onDrop={onDrop} style={{background:"#fff",border:"2px dashed #cbd5e1",borderRadius:14,padding:48,textAlign:"center",color:"#64748b"}}>
      <div style={{fontSize:42,marginBottom:10}}>📄⬇️</div>
      <div style={{fontWeight:700,fontSize:16,color:"#1e293b",marginBottom:6}}>Drag your Excel (.xlsx) here</div>
      <div style={{fontSize:13,marginBottom:16}}>It loads with the same columns, rows, colours & formatting — editable, resizable, and exportable.</div>
      <div style={{display:"flex",gap:10,justifyContent:"center",flexWrap:"wrap"}}>
        <label style={{background:"#2563eb",color:"#fff",borderRadius:10,padding:"10px 18px",fontSize:13,fontWeight:600,cursor:"pointer"}}>Choose file<input type="file" accept=".xlsx" style={{display:"none"}} onChange={e=>onFile(e.target.files[0])}/></label>
        <button onClick={useTemplate} disabled={busy} style={{background:"#7c3aed",color:"#fff",border:"none",borderRadius:10,padding:"10px 18px",fontSize:13,fontWeight:600,cursor:"pointer"}}>{busy?"Loading…":"Use CAG master template"}</button>
      </div>
      <div style={{fontSize:11,color:"#94a3b8",marginTop:14}}>Keep files under ~700 KB. Formulas are preserved on export (they show their saved value while in the app).</div>
    </div>
  );

  const cw=c=> (dragLive&&dragLive.kind==="col"&&dragLive.idx===c)?dragLive.px : (edits.colW?.[cwk(edits.active,c)] ?? grid?.cols[c-1] ?? 70);
  const rh=r=> (dragLive&&dragLive.kind==="row"&&dragLive.idx===r)?dragLive.px : (edits.rowH?.[rhk(edits.active,r)] ?? grid?.rows[r-1]?.h ?? 21);

  // merge maps
  let covered=new Set(), spanMap={};
  if(grid){ for(const rng of grid.merges){ const [a,b]=String(rng).split(":"); const A=a1(a),B=a1(b); if(!A||!B)continue; spanMap[`${A.r}_${A.c}`]={rs:B.r-A.r+1,cs:B.c-A.c+1}; for(let r=A.r;r<=B.r;r++)for(let c=A.c;c<=B.c;c++){ if(!(r===A.r&&c===A.c))covered.add(`${r}_${c}`);} } }

  return(
    <div onDragOver={e=>e.preventDefault()} onDrop={onDrop}>
      {/* toolbar */}
      <div style={{display:"flex",gap:8,alignItems:"center",marginBottom:8,flexWrap:"wrap"}}>
        <span style={{fontWeight:700,fontSize:13,color:"#1e293b"}}>📄 {fileName||"Sheet"}</span>
        {busy&&<span style={{fontSize:12,color:"#2563eb"}}>working…</span>}
        <div style={{marginLeft:"auto",display:"flex",gap:8,flexWrap:"wrap"}}>
          <button onClick={exportXlsx} style={{background:"#15803d",color:"#fff",border:"none",borderRadius:8,padding:"6px 12px",fontSize:12,fontWeight:600,cursor:"pointer"}}>⤓ Export Excel</button>
          <label style={{background:"#2563eb",color:"#fff",borderRadius:8,padding:"6px 12px",fontSize:12,fontWeight:600,cursor:"pointer"}}>Replace file<input type="file" accept=".xlsx" style={{display:"none"}} onChange={e=>onFile(e.target.files[0])}/></label>
        </div>
      </div>
      {/* sheet tabs */}
      <div style={{display:"flex",gap:4,marginBottom:8,overflowX:"auto"}}>
        {sheets.map((nm,i)=><button key={i} onClick={()=>setActive(i)} style={{padding:"4px 10px",borderRadius:"6px 6px 0 0",fontSize:12,fontWeight:600,border:"1px solid #e2e8f0",borderBottom:"none",cursor:"pointer",whiteSpace:"nowrap",background:edits.active===i?"#1e3a5f":"#f8fafc",color:edits.active===i?"#fff":"#475569"}}>{nm}</button>)}
      </div>

      {!grid?<div style={{padding:40,textAlign:"center",color:"#94a3b8"}}>Loading sheet…</div>:(
        <div style={{overflow:"auto",border:"1px solid #e2e8f0",borderRadius:8,background:"#fff",maxHeight:"72vh"}}>
          <table style={{borderCollapse:"collapse",tableLayout:"fixed",fontSize:11}}>
            <colgroup><col style={{width:42}}/>{Array.from({length:grid.right},(_,i)=><col key={i} style={{width:cw(i+1)}}/>)}</colgroup>
            <thead>
              <tr>
                <th style={{position:"sticky",top:0,left:0,zIndex:3,background:"#f1f5f9",border:"1px solid #cbd5e1",height:22}}></th>
                {Array.from({length:grid.right},(_,i)=>(
                  <th key={i} style={{position:"sticky",top:0,zIndex:2,background:"#f1f5f9",border:"1px solid #cbd5e1",fontSize:10,color:"#64748b",fontWeight:600,padding:0,height:22,userSelect:"none"}}>
                    <div style={{position:"relative",padding:"2px 4px"}}>{colLetter(i+1)}
                      <div onMouseDown={e=>startDrag("col",i+1,cw(i+1),e)} style={{position:"absolute",top:0,right:-3,width:6,height:"100%",cursor:"col-resize"}}/>
                    </div>
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {grid.rows.map((row,ri)=>{ const r=ri+1; return(
                <tr key={r} style={{height:rh(r)}}>
                  <td style={{position:"sticky",left:0,zIndex:1,background:"#f1f5f9",border:"1px solid #cbd5e1",textAlign:"center",fontSize:10,color:"#64748b",fontWeight:600,userSelect:"none",padding:0}}>
                    <div style={{position:"relative",height:rh(r),display:"flex",alignItems:"center",justifyContent:"center"}}>{r}
                      <div onMouseDown={e=>startDrag("row",r,rh(r),e)} style={{position:"absolute",left:0,bottom:-3,height:6,width:"100%",cursor:"row-resize"}}/>
                    </div>
                  </td>
                  {row.cells.map((cell,ci)=>{ const c=ci+1; if(covered.has(`${r}_${c}`))return null; const sp=spanMap[`${r}_${c}`]; const key=ck(edits.active,r,c); const disp=edits.cells?.[key]??cell.v;
                    return(
                      <td key={c} colSpan={sp?sp.cs:1} rowSpan={sp?sp.rs:1}
                        contentEditable suppressContentEditableWarning
                        onBlur={e=>commitCell(r,c,e.currentTarget.innerText)}
                        style={{border:"1px solid #e2e8f0",borderLeft:cell.bL?"1px solid #475569":"1px solid #eef2f7",borderTop:cell.bT?"1px solid #475569":"1px solid #eef2f7",borderRight:cell.bR?"1px solid #475569":"1px solid #eef2f7",borderBottom:cell.bB?"1px solid #475569":"1px solid #eef2f7",
                          background:cell.bg||"transparent",color:cell.color||"#1e293b",fontWeight:cell.bold?700:400,fontStyle:cell.italic?"italic":"normal",fontSize:Math.max(9,Math.min(16,cell.size)),
                          textAlign:cell.align||"left",verticalAlign:cell.valign==="center"?"middle":cell.valign==="top"?"top":"bottom",whiteSpace:cell.wrap?"normal":"nowrap",overflow:"hidden",padding:"1px 4px",outline:"none"}}>
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
      <div style={{fontSize:11,color:"#94a3b8",marginTop:6}}>Tip: click a cell to edit · drag a column-letter's right edge / a row-number's bottom edge to resize · changes auto-save & sync. {grid&&(grid.bottom>=MAXR||grid.right>=MAXC)?`(Showing first ${grid.bottom} rows × ${grid.right} cols.)`:""}</div>
    </div>
  );
}
