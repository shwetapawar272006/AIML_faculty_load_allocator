import { useState, useRef } from 'react'
import axios from 'axios'

/* ─── types ─────────────────────────────────────────────────────────────── */
interface FRow { idx:number; name:string; designation:string; theory:number|null; lab:number|null; tu:number; r:number|null; c:number|null; p:number|null; i:number|null; e:number|null; admin:number|null; total:number|null }
interface CRow { idx:number; name:string; designation:string; cores:Array<{type:string;detail:string}|null> }
interface Data { faculty:FRow[]; cores:CRow[]; totalTU:number; totalR:number; totalC:number; totalP:number; totalI:number; totalE:number; totalAdm:number }

const n = (v:any) => v==null?0:Number(v)||0
const f = (v:any) => v==null?'—':Number.isInteger(Number(v))?String(Number(v)):parseFloat(v).toFixed(4).replace(/\.?0+$/,'')
const fh = (v:number) => (v*6).toFixed(2).replace(/\.?0+$/,'')

function parse(d:any): Data {
  const faculty = (d.subject_allocation||[]).map((r:any)=>({
    idx:r.idx, name:r.name, designation:r.designation,
    theory:r.theory??null, lab:r.lab??null, tu:r.tu??0,
    r:r.r??null, c:r.c??null, p:r.p??null, i:r.i??null, e:r.e??null, admin:r.admin??null, total:r.total??null
  }))
  return {
    faculty, cores:(d.core_breakup||[]).map((r:any)=>({idx:r.idx,name:r.name,designation:r.designation,cores:r.cores||[]})),
    totalTU:faculty.reduce((s:number,r:FRow)=>s+n(r.tu),0),
    totalR :faculty.reduce((s:number,r:FRow)=>s+n(r.r),0),
    totalC :faculty.reduce((s:number,r:FRow)=>s+n(r.c),0),
    totalP :faculty.reduce((s:number,r:FRow)=>s+n(r.p),0),
    totalI :faculty.reduce((s:number,r:FRow)=>s+n(r.i),0),
    totalE :faculty.reduce((s:number,r:FRow)=>s+n(r.e),0),
    totalAdm:faculty.reduce((s:number,r:FRow)=>s+n(r.admin),0),
  }
}

/* ─── main component ─────────────────────────────────────────────────────── */
export default function App() {
  const [file,setFile]=useState<File|null>(null)
  const [isDrag,setIsDrag]=useState(false)
  const [loading,setLoading]=useState(false)
  const [err,setErr]=useState('')
  const [data,setData]=useState<Data|null>(null)
  const [blob,setBlob]=useState<Blob|null>(null)
  const [tab,setTab]=useState(0)
  const [search,setSearch]=useState('')
  const inputRef=useRef<HTMLInputElement>(null)

  const onFile=(f:File|null)=>{
    if(!f)return
    if(!f.name.match(/\.xlsx?$/i)){setErr('Please select an .xlsx file');return}
    setFile(f);setErr('');setData(null);setBlob(null)
  }

  const process=async()=>{
    if(!file)return
    setLoading(true);setErr('')
    const fd=new FormData();fd.append('file',file)
    try{
      const [r1,r2]=await Promise.all([
        axios.post('http://localhost:8000/generate',fd),
        axios.post('http://localhost:8000/download',fd,{responseType:'blob'})
      ])
      setData(parse(r1.data));setBlob(r2.data)
    }catch(e:any){setErr(e.response?.data?.detail||e.message||'Error')}
    finally{setLoading(false)}
  }

  const download=()=>{
    if(!blob)return
    const url=URL.createObjectURL(blob)
    const a=document.createElement('a');a.href=url
    a.download='CMRIT_TRCPIE_Faculty_Core_Load_AIML_2025-26_Even_Sem.xlsx'
    a.click();URL.revokeObjectURL(url)
  }

  const filteredFaculty = data?.faculty.filter(r=>search===''||r.name.toLowerCase().includes(search.toLowerCase())) || []
  const filteredCores   = data?.cores.filter(r=>search===''||r.name.toLowerCase().includes(search.toLowerCase())) || []

  /* ── Upload screen ── */
  if(!data) return (
    <div style={{minHeight:'100vh',background:'linear-gradient(135deg,#0f172a 0%,#1e3a8a 50%,#0f172a 100%)',display:'flex',alignItems:'center',justifyContent:'center',padding:24}}>
      <div style={{width:'100%',maxWidth:520}}>
        <div style={{textAlign:'center',marginBottom:32,color:'#fff'}}>
          <div style={{fontSize:48,marginBottom:12}}>🎓</div>
          <h1 style={{fontSize:26,fontWeight:700,margin:'0 0 8px'}}>Faculty Load Allocator</h1>
          <p style={{fontSize:13,opacity:.75,margin:0}}>CMR Institute of Technology · Dept. of AI &amp; ML · 2025-26 Even Semester</p>
        </div>
        <div style={{background:'rgba(255,255,255,.07)',backdropFilter:'blur(12px)',border:'1px solid rgba(255,255,255,.15)',borderRadius:16,padding:28}}>
          <div onClick={()=>inputRef.current?.click()}
            onDragOver={e=>{e.preventDefault();setIsDrag(true)}}
            onDragLeave={()=>setIsDrag(false)}
            onDrop={e=>{e.preventDefault();setIsDrag(false);onFile(e.dataTransfer.files[0])}}
            style={{border:`2px dashed ${isDrag?'#60a5fa':'rgba(255,255,255,.3)'}`,borderRadius:12,padding:'32px 20px',
              textAlign:'center',cursor:'pointer',background:isDrag?'rgba(96,165,250,.1)':'rgba(255,255,255,.03)',marginBottom:16,transition:'all .2s'}}>
            <div style={{fontSize:36,marginBottom:10}}>📊</div>
            <div style={{fontSize:15,fontWeight:600,color:'#fff',marginBottom:4}}>{file?file.name:'Drop Subject Allotment Excel here'}</div>
            <div style={{fontSize:12,color:'rgba(255,255,255,.6)'}}>{file?`${(file.size/1024).toFixed(1)} KB · ready`:'or click to browse · .xlsx'}</div>
            <input ref={inputRef} type="file" accept=".xlsx,.xls" style={{display:'none'}} onChange={e=>onFile(e.target.files?.[0]??null)}/>
          </div>
          {err&&<div style={{background:'rgba(239,68,68,.2)',border:'1px solid rgba(239,68,68,.4)',color:'#fca5a5',borderRadius:8,padding:'10px 14px',marginBottom:14,fontSize:13}}>❌ {err}</div>}
          <button onClick={process} disabled={!file||loading}
            style={{width:'100%',padding:14,borderRadius:10,border:'none',cursor:!file||loading?'not-allowed':'pointer',
              background:!file||loading?'rgba(255,255,255,.2)':'linear-gradient(135deg,#2563eb,#7c3aed)',
              color:'#fff',fontSize:15,fontWeight:700,letterSpacing:.3}}>
            {loading?'⏳  Processing…':'▶  Process & Preview Output'}
          </button>
        </div>
      </div>
    </div>
  )

  /* ── Preview screen ── */
  const TABS=['📋 Subject Allocation','🔲 Core 1–4 Break-Up','📊 Summary Stats']
  const TRCPIE_CARDS=[
    {label:'Total Core Load',val:data.totalTU+data.totalR+data.totalC+data.totalP+data.totalI+data.totalE+data.totalAdm,color:'#2563eb',bg:'#dbeafe'},
    {label:'Teaching (T)',val:data.totalTU,color:'#1d4ed8',bg:'#dbeafe'},
    {label:'Research (R)',val:data.totalR,color:'#15803d',bg:'#dcfce7'},
    {label:'Consultancy (C)',val:data.totalC,color:'#b45309',bg:'#fef3c7'},
    {label:'Sponsored Project (P)',val:data.totalP,color:'#7e22ce',bg:'#f3e8ff'},
    {label:'Innovation (I)',val:data.totalI,color:'#0e7490',bg:'#cffafe'},
    {label:'Entrepreneurship (E)',val:data.totalE,color:'#be123c',bg:'#ffe4e6'},
    {label:'Admin (A)',val:data.totalAdm,color:'#374151',bg:'#f3f4f6'},
  ]

  return (
    <div style={{minHeight:'100vh',background:'#f0f4f8',fontFamily:'system-ui,sans-serif'}}>
      {/* ── top bar ── */}
      <div style={{background:'linear-gradient(90deg,#1e3a8a,#2563eb)',padding:'14px 24px',display:'flex',alignItems:'center',justifyContent:'space-between',boxShadow:'0 2px 12px rgba(0,0,0,.2)'}}>
        <div style={{color:'#fff'}}>
          <div style={{fontSize:17,fontWeight:700}}>🎓 AIML Faculty Core Load — 2025-26 Even Semester</div>
          <div style={{fontSize:12,opacity:.75,marginTop:2}}>CMR Institute of Technology · {data.faculty.length} Faculty · {data.totalTU.toFixed(2)} Total Teaching Units</div>
        </div>
        <div style={{display:'flex',gap:8,alignItems:'center'}}>
          <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="🔍 Search faculty…"
            style={{padding:'7px 12px',borderRadius:8,border:'1px solid rgba(255,255,255,.3)',background:'rgba(255,255,255,.15)',color:'#fff',fontSize:13,outline:'none',width:200}}/>
          <button onClick={()=>{setData(null);setBlob(null);setFile(null)}}
            style={{padding:'7px 14px',borderRadius:8,border:'1px solid rgba(255,255,255,.3)',background:'rgba(255,255,255,.1)',color:'#fff',fontSize:13,cursor:'pointer'}}>↩ New</button>
          <button onClick={download}
            style={{padding:'7px 16px',borderRadius:8,border:'none',background:'#fff',color:'#1e3a8a',fontSize:13,fontWeight:700,cursor:'pointer'}}>⬇ Download Excel</button>
        </div>
      </div>

      {/* ── tab bar ── */}
      <div style={{background:'#fff',borderBottom:'1px solid #e2e8f0',padding:'0 24px',display:'flex',gap:4}}>
        {TABS.map((t,i)=>(
          <button key={i} onClick={()=>setTab(i)}
            style={{padding:'12px 18px',border:'none',borderBottom:`3px solid ${tab===i?'#2563eb':'transparent'}`,background:'transparent',
              color:tab===i?'#1e3a8a':'#64748b',fontSize:13,fontWeight:tab===i?600:400,cursor:'pointer',transition:'all .15s'}}>
            {t}
          </button>
        ))}
      </div>

      <div style={{padding:24}}>

        {/* ════ TAB 0: Subject Allocation ════ */}
        {tab===0&&(
          <div style={{background:'#fff',borderRadius:12,boxShadow:'0 1px 4px rgba(0,0,0,.08)',overflow:'hidden'}}>
            <div style={{background:'#f8fafc',padding:'12px 16px',borderBottom:'1px solid #e2e8f0',display:'flex',alignItems:'center',justifyContent:'space-between'}}>
              <span style={{fontSize:13,fontWeight:600,color:'#1e293b'}}>Subject Allocation — {filteredFaculty.length} faculty</span>
              <span style={{fontSize:11,color:'#94a3b8'}}>Theory + Lab breakdown with TRCPIE components</span>
            </div>
            <div style={{overflowX:'auto'}}>
              <table style={{borderCollapse:'collapse',width:'100%',fontSize:12}}>
                <thead>
                  <tr style={{background:'#1e3a8a'}}>
                    {['#','Name','Designation','Theory','Lab','Teaching Units','R','C','P','I','E','Admin','Total'].map(h=>(
                      <th key={h} style={{padding:'9px 12px',color:'#fff',fontWeight:600,fontSize:11,whiteSpace:'nowrap',textAlign:'left'}}>{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {filteredFaculty.map((r,ri)=>{
                    const tuBg=r.tu>=4?'#dcfce7':r.tu>=3?'#dbeafe':'#fef3c7'
                    return(
                    <tr key={r.idx} style={{background:ri%2===0?'#fff':'#f8fafc'}}>
                      <td style={td}>{r.idx}</td>
                      <td style={{...td,fontWeight:500,maxWidth:200}}>{r.name}</td>
                      <td style={td}>
                        <span style={{fontSize:10,padding:'2px 7px',borderRadius:4,fontWeight:600,
                          background:r.designation==='Associate Professor'?'#dbeafe':'#f1f5f9',
                          color:r.designation==='Associate Professor'?'#1e40af':'#475569'}}>
                          {r.designation==='Associate Professor'?'Assoc Prof':'Asst Prof'}
                        </span>
                      </td>
                      <td style={{...td,color:'#64748b'}}>{f(r.theory)}</td>
                      <td style={{...td,color:'#64748b'}}>{f(r.lab)}</td>
                      <td style={{...td,padding:'6px 8px'}}>
                        <div style={{display:'flex',alignItems:'center',gap:6}}>
                          <div style={{width:44,height:7,background:'#e2e8f0',borderRadius:4}}>
                            <div style={{width:`${Math.min(100,r.tu/4*100)}%`,height:'100%',background:'#2563eb',borderRadius:4}}/>
                          </div>
                          <span style={{fontWeight:700,fontSize:12,background:tuBg,padding:'1px 6px',borderRadius:4}}>{f(r.tu)}</span>
                        </div>
                      </td>
                      {[r.r,r.c,r.p,r.i,r.e,r.admin].map((v,vi)=>(
                        <td key={vi} style={{...td,color:'#64748b'}}>{f(v)}</td>
                      ))}
                      <td style={{...td,fontWeight:700}}>{r.total}</td>
                    </tr>
                  )})}
                  <tr style={{background:'#1e3a8a'}}>
                    <td colSpan={5} style={{...td,color:'#93c5fd',textAlign:'right',fontSize:11}}>Total →</td>
                    <td style={{...td,color:'#fff',fontWeight:700}}>{data.totalTU.toFixed(4)}</td>
                    <td colSpan={7} style={td}/>
                  </tr>
                </tbody>
              </table>
            </div>
          </div>
        )}

        {/* ════ TAB 1: Core 1-4 ════ */}
        {tab===1&&(
          <div style={{background:'#fff',borderRadius:12,boxShadow:'0 1px 4px rgba(0,0,0,.08)',overflow:'hidden'}}>
            <div style={{background:'#f8fafc',padding:'12px 16px',borderBottom:'1px solid #e2e8f0',display:'flex',alignItems:'center',justifyContent:'space-between'}}>
              <span style={{fontSize:13,fontWeight:600,color:'#1e293b'}}>Core 1–4 Break-Up — {filteredCores.length} faculty</span>
              <div style={{display:'flex',gap:10}}>
                {[['#dbeafe','1T / 1 T  Theory'],['#fef3c7','TR  Fractional'],['#dcfce7','R/C/P/I/E  Research'],['#f1f5f9','1A  Admin']].map(([bg,lbl])=>(
                  <span key={lbl} style={{fontSize:10,background:bg,padding:'2px 8px',borderRadius:4}}>{lbl}</span>
                ))}
              </div>
            </div>
            <div style={{overflowX:'auto'}}>
              <table style={{borderCollapse:'collapse',width:'100%',fontSize:12}}>
                <thead>
                  <tr style={{background:'#1e3a8a'}}>
                    {['#','Name','Core 1','Core 2','Core 3','Core 4'].map(h=>(
                      <th key={h} style={{padding:'9px 12px',color:'#fff',fontWeight:600,fontSize:11,textAlign:'left',whiteSpace:'nowrap'}}>{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {filteredCores.map((r,ri)=>(
                    <tr key={r.idx} style={{background:ri%2===0?'#fff':'#f8fafc',verticalAlign:'top'}}>
                      <td style={td}>{r.idx}</td>
                      <td style={{...td,fontWeight:500,maxWidth:180,whiteSpace:'normal'}}>{r.name}</td>
                      {r.cores.map((c,ci)=>(
                        <td key={ci} style={{...td,maxWidth:220,whiteSpace:'normal'}}>
                          {c?(
                            <>
                              <span style={{display:'inline-block',fontSize:10,padding:'1px 7px',borderRadius:3,fontWeight:700,marginBottom:4,
                                background:c.type==='1 T'||c.type==='1T'?'#dbeafe':c.type==='TR'?'#fef3c7':c.type==='1A'?'#f1f5f9':'#dcfce7',
                                color:c.type==='1 T'||c.type==='1T'?'#1e40af':c.type==='TR'?'#92400e':c.type==='1A'?'#475569':'#166534'}}>
                                {c.type}
                              </span>
                              <div style={{fontSize:11,color:'#475569',lineHeight:1.5,whiteSpace:'pre-wrap'}}>{c.detail}</div>
                            </>
                          ):<span style={{color:'#e2e8f0'}}>—</span>}
                        </td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {/* ════ TAB 2: Summary Stats ════ */}
        {tab===2&&(
          <div>
            {/* metric cards */}
            <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(155px,1fr))',gap:12,marginBottom:20}}>
              {[
                {icon:'👥',label:'Total Faculty',val:data.faculty.length,sub:''},
                {icon:'📚',label:'Teaching Units',val:data.totalTU.toFixed(2),sub:fh(data.totalTU)+' hrs'},
                {icon:'📊',label:'Avg Units/Faculty',val:(data.totalTU/data.faculty.length).toFixed(2),sub:'per faculty'},
                {icon:'✅',label:'Full Load (4 units)',val:data.faculty.filter(r=>r.tu>=4).length,sub:'faculty'},
                {icon:'🎓',label:'Associate Profs',val:data.faculty.filter(r=>r.designation==='Associate Professor').length,sub:''},
                {icon:'👨‍🏫',label:'Assistant Profs',val:data.faculty.filter(r=>r.designation==='Assistant Professor').length,sub:''},
              ].map(c=>(
                <div key={c.label} style={{background:'#fff',borderRadius:12,padding:'16px 18px',boxShadow:'0 1px 4px rgba(0,0,0,.08)',border:'1px solid #e2e8f0'}}>
                  <div style={{fontSize:24,marginBottom:8}}>{c.icon}</div>
                  <div style={{fontSize:24,fontWeight:700,color:'#1e293b',lineHeight:1}}>{c.val}</div>
                  <div style={{fontSize:11,color:'#64748b',marginTop:5}}>{c.label}</div>
                  {c.sub&&<div style={{fontSize:10,color:'#94a3b8',marginTop:2}}>{c.sub}</div>}
                </div>
              ))}
            </div>

            {/* TRCPIE totals */}
            <div style={{background:'#fff',borderRadius:12,boxShadow:'0 1px 4px rgba(0,0,0,.08)',padding:20,marginBottom:20}}>
              <div style={{fontSize:14,fontWeight:700,color:'#1e293b',marginBottom:14}}>
                TRCPIE Core Load Totals
                <span style={{fontSize:11,fontWeight:400,color:'#94a3b8',marginLeft:8}}>1 unit = 6 hours of core activity</span>
              </div>
              <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(150px,1fr))',gap:10}}>
                {TRCPIE_CARDS.map(c=>(
                  <div key={c.label} style={{background:c.bg,borderRadius:10,padding:'14px 16px'}}>
                    <div style={{fontSize:22,fontWeight:700,color:c.color}}>{n(c.val).toFixed(2)}</div>
                    <div style={{fontSize:11,fontWeight:600,color:c.color,marginTop:3}}>{c.label}</div>
                    <div style={{fontSize:10,color:c.color,opacity:.75,marginTop:2}}>{fh(n(c.val))} core hrs</div>
                  </div>
                ))}
              </div>
            </div>

            {/* distribution bar chart */}
            <div style={{background:'#fff',borderRadius:12,boxShadow:'0 1px 4px rgba(0,0,0,.08)',padding:20}}>
              <div style={{fontSize:14,fontWeight:700,color:'#1e293b',marginBottom:14}}>Teaching Load per Faculty</div>
              {data.faculty.filter(r=>search===''||r.name.toLowerCase().includes(search.toLowerCase())).map(r=>{
                const col=r.tu>=4?'#16a34a':r.tu>=3?'#2563eb':'#f59e0b'
                const pct=Math.min(100,r.tu/4*100)
                return(
                  <div key={r.idx} style={{display:'flex',alignItems:'center',gap:10,marginBottom:7}}>
                    <div style={{fontSize:11,color:'#64748b',width:180,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap',flexShrink:0}}>
                      {r.name.replace(/^(Mr\.|Ms\.|Dr\.)\s*/,'')}
                    </div>
                    <div style={{flex:1,background:'#f1f5f9',borderRadius:5,height:18,overflow:'hidden'}}>
                      <div style={{width:`${pct}%`,height:'100%',background:col,borderRadius:5,transition:'width .4s',
                        display:'flex',alignItems:'center',paddingLeft:6}}>
                        {pct>20&&<span style={{fontSize:10,color:'#fff',fontWeight:600}}>{f(r.tu)}</span>}
                      </div>
                    </div>
                    {pct<=20&&<span style={{fontSize:11,fontWeight:600,width:32,textAlign:'right',color:col}}>{f(r.tu)}</span>}
                    <div style={{fontSize:10,color:'#94a3b8',width:60,textAlign:'right',flexShrink:0}}>{fh(r.tu)} hrs</div>
                  </div>
                )
              })}
              <div style={{display:'flex',gap:16,marginTop:12,paddingTop:12,borderTop:'1px solid #f1f5f9'}}>
                {[['#16a34a','≥ 4 units — Full load'],['#2563eb','3–3.9 units'],['#f59e0b','< 3 units']].map(([c,l])=>(
                  <div key={l} style={{display:'flex',alignItems:'center',gap:5}}>
                    <div style={{width:10,height:10,borderRadius:2,background:c}}/>
                    <span style={{fontSize:11,color:'#64748b'}}>{l}</span>
                  </div>
                ))}
              </div>
            </div>
          </div>
        )}
      </div>
    </div>
  )
}

const td:React.CSSProperties={padding:'7px 12px',borderBottom:'1px solid #f1f5f9',whiteSpace:'nowrap',fontSize:12}
