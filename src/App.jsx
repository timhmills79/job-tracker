import { useState, useCallback, useRef, useEffect } from 'react';
import { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, LevelFormat } from 'docx';

// ── Constants ─────────────────────────────────────────────────────────────────
const SENIORITY_OPTIONS = [
  { label: 'Any Level', value: '' }, { label: 'Entry Level', value: 'NO_EXPERIENCE' },
  { label: 'Junior', value: 'JUNIOR' }, { label: 'Mid-Level', value: 'MID' }, { label: 'Senior', value: 'SENIOR' },
];
const LOCATION_TYPE_OPTIONS = [
  { label: 'Any', value: '' }, { label: 'Remote', value: 'REMOTE' },
  { label: 'Hybrid', value: 'HYBRID' }, { label: 'On-site', value: 'PHYSICAL' },
];
const SOURCES = [
  { id: 'ziprecruiter',   label: 'ZipRecruiter',      emoji: '🔍', color: '#4a90d9' },
  { id: 'indeed',         label: 'Indeed',             emoji: '💼', color: '#003a9b' },
  { id: 'linkedin',       label: 'LinkedIn',           emoji: '🔗', color: '#0077b5' },
  { id: 'glassdoor',      label: 'Glassdoor',          emoji: '🟢', color: '#0caa41' },
  { id: 'usajobs',        label: 'USAJobs',            emoji: '🏛️', color: '#1a3e6f' },
  { id: 'weworkremotely', label: 'We Work Remotely',   emoji: '🌎', color: '#1f9e6e' },
];

// ── API helpers ───────────────────────────────────────────────────────────────
async function apiClaude(messages, mcpServers = [], maxTokens = 1000, tools = []) {
  const body = { model: 'claude-sonnet-4-20250514', max_tokens: maxTokens, messages };
  if (mcpServers.length) body.mcp_servers = mcpServers;
  if (tools.length) body.tools = tools;
  const res = await fetch('/api/claude', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(body),
  });
  if (res.status === 401) throw new Error('SESSION_EXPIRED');
  return res.json();
}

async function apiSheets(method, body) {
  const res = await fetch('/api/sheets', {
    method,
    headers: { 'Content-Type': 'application/json' },
    body: method !== 'GET' ? JSON.stringify(body) : undefined,
  });
  if (res.status === 401) throw new Error('SESSION_EXPIRED');
  return res.json();
}

function getTextBlock(data) { return data.content?.find(b => b.type === 'text')?.text || ''; }
function parseJSON(raw) {
  const clean = (raw || '').replace(/```json|```/g, '').trim();
  const s = clean.search(/[\[{]/), e = Math.max(clean.lastIndexOf(']'), clean.lastIndexOf('}'));
  if (s === -1 || e === -1) return null;
  try { return JSON.parse(clean.slice(s, e + 1)); } catch { return null; }
}

const ZIP_MCP   = [{ type: 'url', url: 'https://api.ziprecruiter.com/mcp', name: 'ziprecruiter-mcp' }];
const INDEED_MCP= [{ type: 'url', url: 'https://mcp.indeed.com/claude/mcp', name: 'indeed-mcp' }];
const WEB_TOOLS = [{ type: 'web_search_20250305', name: 'web_search' }];

// ── Duplicate detection ───────────────────────────────────────────────────────
function normT(t) { return (t||'').toLowerCase().replace(/\s+/g,' ').trim(); }
function normC(c) { return (c||'').toLowerCase().replace(/[^a-z0-9]/g,''); }
function isDuplicate(job, seen) {
  const nt=normT(job.title),nc=normC(job.company),url=job.url||'';
  return seen.some(s=>{
    if(s.id===job.id) return false;
    if(s.url&&url&&s.url===url) return true;
    if(normC(s.company)===nc&&normT(s.title)===nt) return true;
    if(normC(s.company)===nc){
      const tw=new Set(nt.split(' ').filter(w=>w.length>2));
      const sw=new Set(normT(s.title).split(' ').filter(w=>w.length>2));
      const ov=[...tw].filter(w=>sw.has(w)).length;
      const mx=Math.max(tw.size,sw.size);
      if(mx>0&&ov/mx>=0.6) return true;
    }
    return false;
  });
}

// ── Source search functions ───────────────────────────────────────────────────
async function searchZip(title,loc,locType,sen,salMin){
  const senStr=sen?` experience level ${sen}`:''; const salStr=salMin?` minimum salary $${salMin}`:'';
  const d=await apiClaude([{role:'user',content:`Search ZipRecruiter.com for current "${title}" job postings${loc?` in ${loc}`:''}${locType==='REMOTE'?' that are remote':''}${senStr}${salStr}. Find real listings from ziprecruiter.com. Return ONLY JSON array of up to 6:[{id,title,company,location,salary,url,snippet,posted}]. Unique string IDs. salary=Not specified if unknown. No markdown.`}],[],1000,WEB_TOOLS);
  const p=parseJSON(getTextBlock(d));
  return Array.isArray(p)?p.map((j,i)=>({...j,id:`zip-${title}-${j.id||i}`,source:'ziprecruiter',searchTitle:title})):[];
}
async function searchIndeed(title,loc,locType,sen,salMin){
  const senStr=sen?` experience level ${sen}`:''; const salStr=salMin?` minimum salary $${salMin}`:'';
  const d=await apiClaude([{role:'user',content:`Search Indeed.com for current "${title}" job postings${loc?` in ${loc}`:''}${locType==='REMOTE'?' that are remote':''}${senStr}${salStr}. Find real listings from indeed.com. Return ONLY JSON array of up to 6:[{id,title,company,location,salary,url,snippet,posted}]. Unique string IDs. salary=Not specified if unknown. No markdown.`}],[],1000,WEB_TOOLS);
  const p=parseJSON(getTextBlock(d));
  return Array.isArray(p)?p.map((j,i)=>({...j,id:`indeed-${title}-${j.id||i}`,source:'indeed',searchTitle:title})):[];
}
async function searchLinkedIn(title,loc,locType){
  const d=await apiClaude([{role:'user',content:`Search LinkedIn Jobs for "${title}"${loc?` in ${loc}`:''}${locType==='REMOTE'?' remote':''}.Return up to 6 current postings as ONLY JSON array:[{id,title,company,location,salary,url,snippet,posted}].salary="Not specified" if unknown.No markdown.`}],[],1000,WEB_TOOLS);
  const p=parseJSON(getTextBlock(d));
  return Array.isArray(p)?p.map((j,i)=>({...j,id:`linkedin-${title}-${j.id||i}`,source:'linkedin',searchTitle:title})):[];
}
async function searchGlassdoor(title,loc,locType){
  const d=await apiClaude([{role:'user',content:`Search Glassdoor Jobs for "${title}"${loc?` in ${loc}`:''}${locType==='REMOTE'?' remote':''}.Return up to 6 as ONLY JSON array:[{id,title,company,location,salary,url,snippet,posted}].No markdown.`}],[],1000,WEB_TOOLS);
  const p=parseJSON(getTextBlock(d));
  return Array.isArray(p)?p.map((j,i)=>({...j,id:`glassdoor-${title}-${j.id||i}`,source:'glassdoor',searchTitle:title})):[];
}
async function searchUSAJobs(title,loc,locType){
  const d=await apiClaude([{role:'user',content:`Search usajobs.gov for "${title}" federal jobs${loc?` in ${loc}`:''}${locType==='REMOTE'?' telework eligible':''}.Return up to 6 as ONLY JSON:[{id,title,company,location,salary,url,snippet,posted}].Use agency as company.No markdown.`}],[],1000,WEB_TOOLS);
  const p=parseJSON(getTextBlock(d));
  return Array.isArray(p)?p.map((j,i)=>({...j,id:`usa-${title}-${j.id||i}`,source:'usajobs',searchTitle:title})):[];
}
async function searchWWR(title){
  const d=await apiClaude([{role:'user',content:`Search weworkremotely.com for "${title}" remote jobs.Return up to 6 as ONLY JSON:[{id,title,company,location,salary,url,snippet,posted}].location="Remote".No markdown.`}],[],1000,WEB_TOOLS);
  const p=parseJSON(getTextBlock(d));
  return Array.isArray(p)?p.map((j,i)=>({...j,id:`wwr-${title}-${j.id||i}`,source:'weworkremotely',searchTitle:title,location:'Remote'})):[];
}

// ── DOCX download ─────────────────────────────────────────────────────────────
async function downloadDocx(text, title, company) {
  const lines=text.split('\n'), children=[];
  const numCfg=[{reference:'bullets',levels:[{level:0,format:LevelFormat.BULLET,text:'•',alignment:AlignmentType.LEFT,style:{paragraph:{indent:{left:720,hanging:360}}}}]}];
  for(const line of lines){
    const t=line.trim();
    if(!t){children.push(new Paragraph({children:[new TextRun('')]}));continue;}
    if(t.startsWith('# '))   children.push(new Paragraph({heading:HeadingLevel.HEADING_1,children:[new TextRun({text:t.slice(2),bold:true,font:'Arial',size:32})]}));
    else if(t.startsWith('## ')) children.push(new Paragraph({heading:HeadingLevel.HEADING_2,children:[new TextRun({text:t.slice(3),bold:true,font:'Arial',size:26})]}));
    else if(t.startsWith('- ')||t.startsWith('• ')) children.push(new Paragraph({numbering:{reference:'bullets',level:0},children:[new TextRun({text:t.replace(/^[-•]\s*/,''),font:'Arial',size:22})]}));
    else{const parts=t.split(/(\*\*[^*]+\*\*)/g);children.push(new Paragraph({children:parts.map(p=>p.startsWith('**')&&p.endsWith('**')?new TextRun({text:p.slice(2,-2),bold:true,font:'Arial',size:22}):new TextRun({text:p,font:'Arial',size:22}))}));}
  }
  const doc=new Document({numbering:{config:numCfg},styles:{default:{document:{run:{font:'Arial',size:22}}},paragraphStyles:[{id:'Heading1',name:'Heading 1',basedOn:'Normal',next:'Normal',quickFormat:true,run:{size:32,bold:true,font:'Arial'},paragraph:{spacing:{before:240,after:120},outlineLevel:0}},{id:'Heading2',name:'Heading 2',basedOn:'Normal',next:'Normal',quickFormat:true,run:{size:26,bold:true,font:'Arial'},paragraph:{spacing:{before:200,after:80},outlineLevel:1}}]},sections:[{properties:{page:{size:{width:12240,height:15840},margin:{top:1440,right:1440,bottom:1440,left:1440}}},children}]});
  const buffer=await Packer.toBuffer(doc);
  const blob=new Blob([buffer],{type:'application/vnd.openxmlformats-officedocument.wordprocessingml.document'});
  const url=URL.createObjectURL(blob),a=document.createElement('a');
  a.href=url;a.download=`Resume_${title.replace(/[^a-z0-9]/gi,'_')}_${company.replace(/[^a-z0-9]/gi,'_')}.docx`;a.click();
  URL.revokeObjectURL(url);
}

// ── Shared UI bits ────────────────────────────────────────────────────────────
function fitColor(s){if(s>=75)return{bg:'#edf7ed',border:'#a5d6a7',text:'#2e7d32'};if(s>=50)return{bg:'#fff8e1',border:'#ffe082',text:'#f57f17'};return{bg:'#fce4ec',border:'#f48fb1',text:'#c62828'};}
function FitBadge({score}){const c=fitColor(score);return<span style={{fontSize:12,fontWeight:700,fontFamily:"'DM Mono',monospace",padding:'2px 10px',borderRadius:20,background:c.bg,border:`1px solid ${c.border}`,color:c.text}}>{score}% fit</span>;}
function StatusBadge({status}){const map={New:'#e8f4fd;#1a6fa8;#b3d8f0',Saved:'#edf7ed;#2e7d32;#a5d6a7',Skipped:'#fafafa;#999;#ddd',Applied:'#fff3e0;#e65100;#ffcc80'};const[bg,color,border]=(map[status]||map.New).split(';');return<span style={{fontSize:11,fontWeight:700,letterSpacing:'0.04em',padding:'2px 9px',borderRadius:20,fontFamily:"'DM Mono',monospace",background:bg,color,border:`1px solid ${border}`}}>{status}</span>;}
function SourceBadge({sourceId}){const src=SOURCES.find(s=>s.id===sourceId);if(!src)return null;return<span style={{fontSize:10,fontWeight:700,fontFamily:"'DM Mono',monospace",padding:'1px 7px',borderRadius:20,background:src.color+'18',border:`1px solid ${src.color}44`,color:src.color}}>{src.emoji} {src.label}</span>;}
function TitleTag({title,onRemove}){return<div style={{display:'inline-flex',alignItems:'center',gap:6,background:'#1a1a2e',color:'#f0ece2',padding:'4px 10px 4px 12px',borderRadius:20,fontSize:13,fontFamily:"'DM Mono',monospace"}}>{title}<button onClick={()=>onRemove(title)} style={{background:'none',border:'none',color:'#c9a84c',cursor:'pointer',fontSize:15,lineHeight:1,padding:0}}>×</button></div>;}
function fileToBase64(file){return new Promise((r,j)=>{const rd=new FileReader();rd.onload=()=>r(rd.result.split(',')[1]);rd.onerror=j;rd.readAsDataURL(file);});}

function SalaryResearch({jobId,listedSalary,salaryData,onResearch,loading}){
  if(!salaryData&&!loading)return<button onClick={onResearch} style={{fontSize:11,color:'#1a6fa8',background:'none',border:'none',cursor:'pointer',fontFamily:"'DM Mono',monospace",textDecoration:'underline',padding:0}}>💰 Research salary</button>;
  if(loading)return<span style={{fontSize:11,color:'#c9a84c',fontFamily:"'DM Mono',monospace"}}>💰 Researching…</span>;
  const{low,mid,high,currency,verdict,verdictColor}=salaryData;
  return<div style={{marginTop:7,padding:'7px 11px',background:'#f8f7f3',border:'1px solid #e8e4db',borderRadius:7,display:'flex',alignItems:'center',gap:10,flexWrap:'wrap'}}>
    <div><span style={{fontSize:10,fontFamily:"'DM Mono',monospace",color:'#aaa',textTransform:'uppercase',letterSpacing:'0.08em'}}>Market Range</span><div style={{fontSize:13,fontWeight:700,color:'#1a1a2e',fontFamily:"'DM Mono',monospace",marginTop:1}}>{currency}{low}–{currency}{high} <span style={{fontWeight:400,color:'#888',fontSize:12}}>· med {currency}{mid}</span></div></div>
    {listedSalary&&listedSalary!=='Not specified'&&<div style={{borderLeft:'1px solid #e0dbd0',paddingLeft:10}}><span style={{fontSize:10,fontFamily:"'DM Mono',monospace",color:'#aaa',textTransform:'uppercase',letterSpacing:'0.08em'}}>Listed</span><div style={{fontSize:13,fontFamily:"'DM Mono',monospace",color:'#333',marginTop:1}}>{listedSalary}</div></div>}
    {verdict&&<span style={{fontSize:11,fontWeight:700,fontFamily:"'DM Mono',monospace",padding:'2px 9px',borderRadius:20,background:verdictColor?.bg||'#f0f0f0',color:verdictColor?.text||'#555',border:`1px solid ${verdictColor?.border||'#ddd'}`}}>{verdict}</span>}
  </div>;
}

// ── Login screen ──────────────────────────────────────────────────────────────
function LoginScreen() {
  return (
    <div style={{minHeight:'100vh',background:'#f7f6f2',display:'flex',alignItems:'center',justifyContent:'center'}}>
      <div style={{textAlign:'center',maxWidth:420,padding:40}}>
        <div style={{background:'#1a1a2e',borderRadius:16,padding:'40px 36px',color:'#f0ece2',boxShadow:'0 8px 40px rgba(0,0,0,0.18)'}}>
          <div style={{fontSize:11,letterSpacing:'0.18em',color:'#c9a84c',fontFamily:"'DM Mono',monospace",textTransform:'uppercase',marginBottom:10}}>Career Intelligence</div>
          <h1 style={{fontSize:28,fontWeight:400,marginBottom:8}}>Job Search Tracker</h1>
          <p style={{color:'#9a9070',fontSize:14,fontStyle:'italic',marginBottom:32,lineHeight:1.6}}>Search 6 job boards · Score against your resume · Sync to your own Google Sheets</p>
          <a href="/api/auth/login" style={{display:'inline-flex',alignItems:'center',gap:10,padding:'13px 24px',background:'#fff',color:'#1a1a2e',borderRadius:8,textDecoration:'none',fontSize:15,fontWeight:600,fontFamily:"'DM Mono',monospace",letterSpacing:'0.03em'}}>
            <svg width="20" height="20" viewBox="0 0 48 48"><path fill="#EA4335" d="M24 9.5c3.54 0 6.71 1.22 9.21 3.6l6.85-6.85C35.9 2.38 30.47 0 24 0 14.62 0 6.51 5.38 2.56 13.22l7.98 6.19C12.43 13.72 17.74 9.5 24 9.5z"/><path fill="#4285F4" d="M46.98 24.55c0-1.57-.15-3.09-.38-4.55H24v9.02h12.94c-.58 2.96-2.26 5.48-4.78 7.18l7.73 6c4.51-4.18 7.09-10.36 7.09-17.65z"/><path fill="#FBBC05" d="M10.53 28.59c-.48-1.45-.76-2.99-.76-4.59s.27-3.14.76-4.59l-7.98-6.19C.92 16.46 0 20.12 0 24c0 3.88.92 7.54 2.56 10.78l7.97-6.19z"/><path fill="#34A853" d="M24 48c6.48 0 11.93-2.13 15.89-5.81l-7.73-6c-2.15 1.45-4.92 2.3-8.16 2.3-6.26 0-11.57-4.22-13.47-9.91l-7.98 6.19C6.51 42.62 14.62 48 24 48z"/></svg>
            Sign in with Google
          </a>
          <p style={{marginTop:20,fontSize:12,color:'#666',lineHeight:1.5}}>Your data stays in your own Google Drive. No data is shared between users.</p>
        </div>
      </div>
    </div>
  );
}

// ── Resume preview ────────────────────────────────────────────────────────────
function ResumePreview({text}){
  return<div style={{fontFamily:"'Georgia',serif",lineHeight:1.65,color:'#1a1a2e'}}>{text.split('\n').map((line,i)=>{const t=line.trim();if(!t)return<div key={i} style={{height:6}}/>;if(t.startsWith('# '))return<h1 key={i} style={{margin:'0 0 4px',fontSize:20,fontWeight:700,borderBottom:'2px solid #c9a84c',paddingBottom:5}}>{t.slice(2)}</h1>;if(t.startsWith('## '))return<h2 key={i} style={{margin:'13px 0 3px',fontSize:13,fontWeight:700,borderBottom:'1px solid #e5e0d8',paddingBottom:3,letterSpacing:'0.05em',textTransform:'uppercase'}}>{t.slice(3)}</h2>;if(t.startsWith('- ')||t.startsWith('• '))return<div key={i} style={{display:'flex',gap:6,marginLeft:8,marginBottom:2}}><span style={{color:'#c9a84c',flexShrink:0}}>•</span><span style={{fontSize:13}} dangerouslySetInnerHTML={{__html:t.replace(/^[-•]\s*/,'').replace(/\*\*([^*]+)\*\*/g,'<strong>$1</strong>')}}/></div>;return<p key={i} style={{margin:'0 0 3px',fontSize:13}} dangerouslySetInnerHTML={{__html:t.replace(/\*\*([^*]+)\*\*/g,'<strong>$1</strong>')}}/>;})}</div>;
}

// ── Job card ──────────────────────────────────────────────────────────────────
function JobCard({job,status,fit,isScoring,isDuplicate,salaryData,onResearchSalary,onStatusChange,onScore,resumeReady,onTailor,tailored,showSource}){
  return<div style={{background:isDuplicate?'#fafaf8':'#fff',border:`1.5px solid ${isDuplicate?'#e0dbd0':'#e8e4db'}`,borderRadius:10,padding:'13px 16px',display:'grid',gridTemplateColumns:'1fr auto',gap:11,alignItems:'start',opacity:isDuplicate?0.82:1}}>
    <div>
      <div style={{display:'flex',alignItems:'center',gap:6,flexWrap:'wrap',marginBottom:3}}>
        <span style={{fontSize:15,fontWeight:600,color:'#1a1a2e'}}>{job.title}</span>
        <StatusBadge status={status}/>
        {isDuplicate&&<span style={{fontSize:10,fontWeight:700,fontFamily:"'DM Mono',monospace",padding:'1px 7px',borderRadius:20,background:'#fff8e1',border:'1px solid #ffe082',color:'#f57f17'}}>⚠ dupe</span>}
        {isScoring&&<span style={{fontSize:11,fontFamily:"'DM Mono',monospace",color:'#c9a84c'}}>scoring…</span>}
        {fit&&!isScoring&&<FitBadge score={fit.score}/>}
        {tailored&&<span style={{fontSize:11,fontFamily:"'DM Mono',monospace",color:'#7b5ea7',background:'#f3eeff',border:'1px solid #d4b8ff',padding:'1px 8px',borderRadius:20}}>📝 tailored</span>}
        {showSource&&job.source&&<SourceBadge sourceId={job.source}/>}
      </div>
      <div style={{fontSize:13,color:'#555',marginBottom:3}}><strong style={{color:'#333'}}>{job.company}</strong>{job.location&&<span> · {job.location}</span>}{job.salary&&job.salary!=='Not specified'&&<span style={{color:'#2e7d32',marginLeft:8,fontFamily:"'DM Mono',monospace",fontSize:12}}>{job.salary}</span>}</div>
      {job.snippet&&<p style={{margin:'0 0 4px',fontSize:13,color:'#777',lineHeight:1.5,fontStyle:'italic'}}>{job.snippet}</p>}
      {fit?.highlights&&<div style={{marginBottom:4}}>{fit.highlights.map((h,i)=><div key={i} style={{fontSize:12,color:'#555',fontFamily:"'DM Mono',monospace"}}>{h}</div>)}</div>}
      <SalaryResearch jobId={job.id} listedSalary={job.salary} salaryData={salaryData?.data} loading={salaryData?.loading} onResearch={onResearchSalary}/>
      <div style={{display:'flex',gap:7,alignItems:'center',flexWrap:'wrap',marginTop:6}}>
        {job.posted&&<span style={{fontSize:11,color:'#bbb',fontFamily:"'DM Mono',monospace"}}>{job.posted}</span>}
        {job.dateAdded&&<span style={{fontSize:11,color:'#bbb',fontFamily:"'DM Mono',monospace"}}>Added {job.dateAdded}</span>}
        {job.url&&<a href={job.url} target="_blank" rel="noreferrer" style={{fontSize:12,color:'#1a6fa8',textDecoration:'underline'}}>View →</a>}
        {resumeReady&&!fit&&!isScoring&&<button onClick={onScore} style={{fontSize:11,color:'#c9a84c',background:'none',border:'none',cursor:'pointer',fontFamily:"'DM Mono',monospace",textDecoration:'underline',padding:0}}>Score</button>}
        {resumeReady&&<button onClick={onTailor} style={{fontSize:11,color:'#7b5ea7',background:'none',border:'none',cursor:'pointer',fontFamily:"'DM Mono',monospace",textDecoration:'underline',padding:0}}>{tailored?'Re-tailor':'✨ Tailor Resume'}</button>}
      </div>
    </div>
    <div style={{display:'flex',flexDirection:'column',gap:5}}>
      {['Saved','Applied','Skipped'].map(s=><button key={s} onClick={()=>onStatusChange(s)} style={{padding:'4px 9px',borderRadius:5,border:'1px solid #e0dbd0',background:status===s?'#1a1a2e':'#f7f6f2',color:status===s?'#f0ece2':'#555',fontSize:11,fontFamily:"'DM Mono',monospace",cursor:'pointer',fontWeight:status===s?700:400}}>{s}</button>)}
    </div>
  </div>;
}

// ── Main App ──────────────────────────────────────────────────────────────────
export default function App() {
  const [user, setUser] = useState(null);
  const [authChecked, setAuthChecked] = useState(false);

  // Check auth on mount
  useEffect(() => {
    fetch('/api/auth/me')
      .then(r => r.ok ? r.json() : null)
      .then(u => { setUser(u); setAuthChecked(true); })
      .catch(() => setAuthChecked(true));
  }, []);

  if (!authChecked) return (
    <div style={{minHeight:'100vh',background:'#f7f6f2',display:'flex',alignItems:'center',justifyContent:'center'}}>
      <span style={{fontFamily:"'DM Mono',monospace",fontSize:13,color:'#aaa'}}>Loading…</span>
    </div>
  );

  if (!user) return <LoginScreen />;
  return <Tracker user={user} />;
}

// ── Tracker (authenticated) ───────────────────────────────────────────────────
function Tracker({ user }) {
  const [titles, setTitles] = useState(['Software Engineer']);
  const [titleInput, setTitleInput] = useState('');
  const [location, setLocation] = useState('');
  const [locationType, setLocationType] = useState('REMOTE');
  const [seniority, setSeniority] = useState('');
  const [salaryMin, setSalaryMin] = useState('');
  const [enabledSources, setEnabledSources] = useState(new Set(SOURCES.map(s => s.id)));

  const [resumeText, setResumeText] = useState('');
  const [resumeBase64, setResumeBase64] = useState(null);
  const [resumeFileName, setResumeFileName] = useState('');
  const [resumeMode, setResumeMode] = useState('paste');
  const [resumeReady, setResumeReady] = useState(false);
  const fileRef = useRef();

  const [jobs, setJobs] = useState([]);
  const [sheetJobs, setSheetJobs] = useState([]);
  const [sheetId, setSheetId] = useState(null);
  const [jobStatus, setJobStatusState] = useState({});
  const [fitScores, setFitScores] = useState({});
  const [scoringId, setScoringId] = useState(null);
  const [duplicates, setDuplicates] = useState(new Set());
  const [hideDuplicates, setHideDuplicates] = useState(false);
  const [salaryResearch, setSalaryResearch] = useState({});
  const [tailoredResumes, setTailoredResumes] = useState({});
  const [selectedTailorJob, setSelectedTailorJob] = useState(null);
  const [downloadingId, setDownloadingId] = useState(null);

  const [activeTab, setActiveTab] = useState('search');
  const [loading, setLoading] = useState(false);
  const [sourceStatus, setSourceStatus] = useState({});
  const [syncing, setSyncing] = useState(false);
  const [syncMsg, setSyncMsg] = useState('');
  const [syncError, setSyncError] = useState('');
  const [error, setError] = useState('');
  const [searched, setSearched] = useState(false);

  // ── Load from user's own Sheet ─────────────────────────────────────────────
  const loadFromSheet = useCallback(async () => {
    setSyncing(true); setSyncError('');
    try {
      const data = await apiSheets('GET');
      if (data.error) throw new Error(data.error);
      setSheetId(data.sheetId);
      setSheetJobs(data.rows || []);
      const sm={},fm={};
      (data.rows||[]).forEach(j=>{if(j.id){sm[j.id]=j.status||'Saved';if(j.fitScore)fm[j.id]={score:parseInt(j.fitScore),highlights:j.fitNotes?j.fitNotes.split(';'):[]}}});
      setJobStatusState(p=>({...p,...sm})); setFitScores(p=>({...p,...fm}));
      setSyncMsg(`Loaded ${data.rows?.length||0} saved jobs from your Google Sheets`);
    } catch(e){ if(e.message!=='SESSION_EXPIRED') setSyncError('Could not load: '+e.message); }
    finally { setSyncing(false); }
  }, []);

  useEffect(() => { loadFromSheet(); }, []);

  const saveJob = useCallback(async (job, status, fit) => {
    const payload = { ...job, status, fitScore: fit?.score||'', fitNotes: fit?.highlights?.join('; ')||'', dateAdded: new Date().toLocaleDateString() };
    try {
      const data = await apiSheets('POST', { job: payload, action: 'save', sheetId });
      if (data.sheetId && !sheetId) setSheetId(data.sheetId);
    } catch(e){ console.warn('Save failed:', e.message); }
  }, [sheetId]);

  const updateStatus = useCallback(async (jobId, newStatus) => {
    try { await apiSheets('POST', { job: { id: jobId, status: newStatus }, action: 'updateStatus', sheetId }); }
    catch(e){ console.warn('Update failed:', e.message); }
  }, [sheetId]);

  // ── Salary research ────────────────────────────────────────────────────────
  const researchSalary = useCallback(async (job) => {
    setSalaryResearch(p=>({...p,[job.id]:{data:null,loading:true}}));
    try {
      const locHint=job.location||location||'United States';
      const d=await apiClaude([{role:'user',content:`Search for ${new Date().getFullYear()} salary data for "${job.title}" jobs${locHint!=='United States'?` in ${locHint}`:' in the US'}. Use Glassdoor,LinkedIn Salary,Indeed,Levels.fyi,or BLS. Return ONLY JSON:{"low":<25pct int>,"mid":<median int>,"high":<75pct int>,"currency":"$","verdict":"<Below Market|At Market|Above Market|Market Rate>","verdictColor":{"bg":"<hex>","border":"<hex>","text":"<hex>"}}. verdictColors: Below Market=(#fce4ec,#f48fb1,#c62828),At Market=(#e8f4fd,#b3d8f0,#1a6fa8),Above Market=(#edf7ed,#a5d6a7,#2e7d32),Market Rate=(#f0f0f0,#ddd,#666). Listed salary:"${job.salary||'Not specified'}". No other text.`}],[],1000,WEB_TOOLS);
      const parsed=parseJSON(getTextBlock(d));
      if(parsed?.low){const fmt=n=>{const num=parseInt(String(n).replace(/[^0-9]/g,''));return num>999?`${Math.round(num/1000)}k`:String(num);};setSalaryResearch(p=>({...p,[job.id]:{data:{...parsed,low:fmt(parsed.low),mid:fmt(parsed.mid),high:fmt(parsed.high)},loading:false}}));}
      else{setSalaryResearch(p=>({...p,[job.id]:{data:{low:'N/A',mid:'N/A',high:'N/A',currency:'$',verdict:'No Data',verdictColor:{bg:'#fafafa',border:'#ddd',text:'#999'}},loading:false}}));}
    } catch{setSalaryResearch(p=>({...p,[job.id]:{data:null,loading:false}}));}
  }, [location]);

  // ── File upload ────────────────────────────────────────────────────────────
  const handleFileUpload = async (e) => {
    const file=e.target.files[0]; if(!file) return;
    setResumeFileName(file.name); setResumeBase64(await fileToBase64(file)); setResumeReady(true);
  };

  // ── Fit scoring ────────────────────────────────────────────────────────────
  const scoreJob = useCallback(async (job) => {
    if(!resumeReady) return; setScoringId(job.id);
    try {
      const prompt=`Analyze resume vs job. Return ONLY JSON:{"score":0-100,"highlights":["✅ or ⚠️ phrase max 10 words","...","..."]}. Job:${job.title} at ${job.company}. ${job.snippet||''}`;
      let messages;
      if(resumeMode==='upload'&&resumeBase64) messages=[{role:'user',content:[{type:'document',source:{type:'base64',media_type:'application/pdf',data:resumeBase64}},{type:'text',text:prompt}]}];
      else messages=[{role:'user',content:`Resume:\n${resumeText}\n---\n${prompt}`}];
      const data=await apiClaude(messages);
      const parsed=parseJSON(getTextBlock(data));
      if(parsed){setFitScores(p=>({...p,[job.id]:parsed}));return parsed;}
    } catch(e){console.warn('Scoring failed',e);}finally{setScoringId(null);}
  }, [resumeReady,resumeMode,resumeBase64,resumeText]);

  // ── Resume tailoring ───────────────────────────────────────────────────────
  const tailorResume = useCallback(async (job) => {
    if(!resumeReady){alert('Please add your resume first.');return;}
    setTailoredResumes(p=>({...p,[job.id]:{text:'',loading:true,error:null}}));
    setSelectedTailorJob(job); setActiveTab('tailored');
    try {
      const prompt=`You are an expert resume writer. Tailor the provided resume for this job. Keep facts accurate. Mirror keywords. Reorder for relevance. Strengthen bullets. Adjust summary. Format in clean Markdown. Return complete resume only.\n\nJob:${job.title} at ${job.company}\nDescription:${job.snippet||''}`;
      let messages;
      if(resumeMode==='upload'&&resumeBase64) messages=[{role:'user',content:[{type:'document',source:{type:'base64',media_type:'application/pdf',data:resumeBase64}},{type:'text',text:prompt}]}];
      else messages=[{role:'user',content:`Original Resume:\n\n${resumeText}\n\n---\n\n${prompt}`}];
      const res=await fetch('/api/claude',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({model:'claude-sonnet-4-20250514',max_tokens:4000,messages})});
      const data=await res.json();
      const text=getTextBlock(data);
      if(!text) throw new Error('No response');
      setTailoredResumes(p=>({...p,[job.id]:{text,loading:false,error:null}}));
    } catch(e){setTailoredResumes(p=>({...p,[job.id]:{text:'',loading:false,error:e.message}}));}
  }, [resumeReady,resumeMode,resumeBase64,resumeText]);

  // ── Search ─────────────────────────────────────────────────────────────────
  const searchJobs = useCallback(async () => {
    if(titles.length===0){setError('Add at least one job title.');return;}
    setLoading(true); setError(''); setJobs([]); setDuplicates(new Set()); setSearched(true);
    const ss={}; SOURCES.forEach(s=>{ss[s.id]=enabledSources.has(s.id)?'searching':'idle';}); setSourceStatus(ss);
    const allJobs=[],dupSet=new Set(),initStatus={};

    const addResults=(results)=>{
      const seen=[...allJobs,...sheetJobs];
      results.forEach(job=>{if(isDuplicate(job,seen))dupSet.add(job.id);allJobs.push(job);initStatus[job.id]=jobStatus[job.id]||'New';});
      setJobs([...allJobs]); setDuplicates(new Set(dupSet)); setJobStatusState(p=>({...p,...initStatus}));
    };

    try {
      for(const title of titles){
        const searches=[];
        if(enabledSources.has('ziprecruiter')) searches.push(searchZip(title,location,locationType,seniority,salaryMin).then(r=>{addResults(r);setSourceStatus(p=>({...p,ziprecruiter:'done'}));}).catch(()=>setSourceStatus(p=>({...p,ziprecruiter:'error'}))));
        if(enabledSources.has('indeed'))       searches.push(searchIndeed(title,location,locationType,seniority,salaryMin).then(r=>{addResults(r);setSourceStatus(p=>({...p,indeed:'done'}));}).catch(()=>setSourceStatus(p=>({...p,indeed:'error'}))));
        if(enabledSources.has('linkedin'))     searches.push(searchLinkedIn(title,location,locationType).then(r=>{addResults(r);setSourceStatus(p=>({...p,linkedin:'done'}));}).catch(()=>setSourceStatus(p=>({...p,linkedin:'error'}))));
        if(enabledSources.has('glassdoor'))    searches.push(searchGlassdoor(title,location,locationType).then(r=>{addResults(r);setSourceStatus(p=>({...p,glassdoor:'done'}));}).catch(()=>setSourceStatus(p=>({...p,glassdoor:'error'}))));
        if(enabledSources.has('usajobs'))      searches.push(searchUSAJobs(title,location,locationType).then(r=>{addResults(r);setSourceStatus(p=>({...p,usajobs:'done'}));}).catch(()=>setSourceStatus(p=>({...p,usajobs:'error'}))));
        if(enabledSources.has('weworkremotely'))searches.push(searchWWR(title).then(r=>{addResults(r);setSourceStatus(p=>({...p,weworkremotely:'done'}));}).catch(()=>setSourceStatus(p=>({...p,weworkremotely:'error'}))));
        await Promise.all(searches);
      }
      if(allJobs.length===0) throw new Error('No results found from any source.');
      if(resumeReady){for(const job of allJobs)if(!dupSet.has(job.id))await scoreJob(job);}
    } catch(e){if(allJobs.length===0)setError(e.message);}
    finally{setLoading(false);}
  }, [titles,location,locationType,seniority,salaryMin,enabledSources,resumeReady,scoreJob,sheetJobs]);

  const setJobStatus = useCallback((job, newStatus) => {
    setJobStatusState(p=>({...p,[job.id]:newStatus}));
    const fit=fitScores[job.id];
    if(newStatus==='Saved'||newStatus==='Applied'){
      saveJob(job,newStatus,fit).then(()=>setSheetJobs(p=>{const ex=p.find(j=>j.id===job.id);return ex?p.map(j=>j.id===job.id?{...j,status:newStatus}:j):[...p,{...job,status:newStatus,fitScore:fit?.score,fitNotes:fit?.highlights?.join('; ')}];}));
    } else if(newStatus==='Skipped'){
      updateStatus(job.id,newStatus);
      setSheetJobs(p=>p.map(j=>j.id===job.id?{...j,status:newStatus}:j));
    }
  }, [fitScores,saveJob,updateStatus]);

  const addTitle=()=>{const t=titleInput.trim();if(t&&!titles.includes(t))setTitles(p=>[...p,t]);setTitleInput('');};
  const removeTitle=t=>setTitles(p=>p.filter(x=>x!==t));
  const handleTitleKey=e=>{if(e.key==='Enter'||e.key===','){e.preventDefault();addTitle();}};
  const toggleSource=id=>setEnabledSources(prev=>{const n=new Set(prev);n.has(id)?n.delete(id):n.add(id);return n;});

  const visibleJobs=hideDuplicates?jobs.filter(j=>!duplicates.has(j.id)):jobs;
  const sourceGroups=SOURCES.filter(s=>enabledSources.has(s.id)).map(src=>({source:src,titleGroups:titles.map(t=>({title:t,jobs:visibleJobs.filter(j=>j.source===src.id&&j.searchTitle===t)})).filter(tg=>tg.jobs.length>0)})).filter(sg=>sg.titleGroups.length>0);
  const dupCount=duplicates.size, scoredCount=Object.keys(fitScores).length, tailoredCount=Object.keys(tailoredResumes).length;
  const allJobsForTailoring=[...jobs,...sheetJobs.filter(sj=>!jobs.find(j=>j.id===sj.id))];
  const isScoringAll=scoringId!==null;

  const TABS=[['search','🔍 Search'],['saved',`📋 Saved (${sheetJobs.filter(j=>j.status!=='Skipped').length})`],['tailored',`📝 Tailored (${tailoredCount})`]];

  return (
    <div style={{minHeight:'100vh',background:'#f7f6f2',fontFamily:"'Georgia',serif"}}>
      {/* Header */}
      <div style={{background:'#1a1a2e',color:'#f0ece2',padding:'20px 36px 16px',borderBottom:'3px solid #c9a84c'}}>
        <div style={{maxWidth:1020,margin:'0 auto',display:'flex',alignItems:'center',gap:16}}>
          <div style={{flex:1}}>
            <span style={{fontSize:11,letterSpacing:'0.18em',color:'#c9a84c',fontFamily:"'DM Mono',monospace",textTransform:'uppercase'}}>Career Intelligence</span>
            <h1 style={{margin:'4px 0 2px',fontSize:24,fontWeight:400}}>Job Search Tracker</h1>
          </div>
          <div style={{display:'flex',alignItems:'center',gap:10}}>
            {user.picture&&<img src={user.picture} alt="" style={{width:32,height:32,borderRadius:'50%',border:'2px solid #c9a84c'}}/>}
            <div>
              <div style={{fontSize:13,color:'#f0ece2'}}>{user.name}</div>
              <div style={{fontSize:11,color:'#9a9070'}}>{user.email}</div>
            </div>
            <a href="/api/auth/logout" style={{fontSize:11,fontFamily:"'DM Mono',monospace",color:'#9a9070',textDecoration:'underline',marginLeft:8}}>Sign out</a>
          </div>
        </div>
      </div>

      {/* Sync bar */}
      <div style={{background:syncing?'#fff8e1':syncError?'#fff5f5':'#f0faf0',borderBottom:'1px solid #e5e0d8',padding:'6px 36px'}}>
        <div style={{maxWidth:1020,margin:'0 auto',display:'flex',alignItems:'center',gap:10}}>
          {syncing&&<span style={{fontSize:12,fontFamily:"'DM Mono',monospace",color:'#888'}}>⟳ Syncing with your Google Sheets…</span>}
          {!syncing&&syncMsg&&<span style={{fontSize:12,fontFamily:"'DM Mono',monospace",color:'#2e7d32'}}>✓ {syncMsg}</span>}
          {!syncing&&syncError&&<span style={{fontSize:12,fontFamily:"'DM Mono',monospace",color:'#c0392b'}}>⚠ {syncError}</span>}
          <button onClick={loadFromSheet} disabled={syncing} style={{marginLeft:'auto',fontSize:11,fontFamily:"'DM Mono',monospace",background:'none',border:'1px solid #ddd',borderRadius:5,padding:'3px 10px',cursor:'pointer',color:'#666'}}>↻ Refresh</button>
        </div>
      </div>

      {/* Tabs */}
      <div style={{background:'#fff',borderBottom:'1px solid #e5e0d8'}}>
        <div style={{maxWidth:1020,margin:'0 auto',padding:'0 36px',display:'flex'}}>
          {TABS.map(([tab,label])=><button key={tab} onClick={()=>setActiveTab(tab)} style={{padding:'12px 18px',border:'none',background:'none',cursor:'pointer',fontSize:13,fontFamily:"'DM Mono',monospace",letterSpacing:'0.04em',color:activeTab===tab?'#1a1a2e':'#999',borderBottom:activeTab===tab?'2px solid #c9a84c':'2px solid transparent',fontWeight:activeTab===tab?700:400}}>{label}</button>)}
        </div>
      </div>

      <div style={{maxWidth:1020,margin:'0 auto',padding:'0 36px'}}>

        {/* ── SEARCH TAB ── */}
        {activeTab==='search'&&<>
          {/* Resume */}
          <div style={{background:'#fff',border:'1.5px solid #e5e0d8',borderRadius:10,margin:'16px 0 0',overflow:'hidden'}}>
            <div style={{padding:'11px 18px',borderBottom:'1px solid #f0ece2',display:'flex',alignItems:'center',gap:10}}>
              <span style={{fontSize:11,fontFamily:"'DM Mono',monospace",letterSpacing:'0.12em',color:'#c9a84c',textTransform:'uppercase',fontWeight:700}}>Resume</span>
              {resumeReady&&<span style={{fontSize:11,fontFamily:"'DM Mono',monospace",color:'#2e7d32',background:'#edf7ed',border:'1px solid #a5d6a7',padding:'1px 8px',borderRadius:20}}>✓ Ready</span>}
            </div>
            <div style={{padding:'10px 18px 0'}}><div style={{display:'inline-flex',border:'1.5px solid #e0dbd0',borderRadius:7,overflow:'hidden'}}>{[['paste','Paste Text'],['upload','Upload PDF']].map(([mode,label])=><button key={mode} onClick={()=>setResumeMode(mode)} style={{padding:'5px 14px',border:'none',cursor:'pointer',fontSize:12,fontFamily:"'DM Mono',monospace",background:resumeMode===mode?'#1a1a2e':'#faf9f6',color:resumeMode===mode?'#f0ece2':'#666'}}>{label}</button>)}</div></div>
            <div style={{padding:'10px 18px 13px'}}>
              {resumeMode==='paste'?<textarea value={resumeText} onChange={e=>{setResumeText(e.target.value);setResumeReady(e.target.value.trim().length>50);}} placeholder="Paste your resume text here…" rows={4} style={{width:'100%',padding:'9px 11px',border:'1.5px solid #e0dbd0',borderRadius:7,fontSize:13,fontFamily:"'Georgia',serif",color:'#1a1a2e',background:'#faf9f6',resize:'vertical',outline:'none',boxSizing:'border-box',lineHeight:1.6}}/>:<div><input ref={fileRef} type="file" accept=".pdf" onChange={handleFileUpload} style={{display:'none'}}/><button onClick={()=>fileRef.current.click()} style={{padding:'8px 16px',border:'1.5px dashed #c9a84c',borderRadius:8,background:'#fffdf4',color:'#1a1a2e',fontSize:13,fontFamily:"'DM Mono',monospace",cursor:'pointer'}}>{resumeFileName?`📄 ${resumeFileName}`:'📁 Choose PDF…'}</button></div>}
            </div>
          </div>

          {/* Sources */}
          <div style={{background:'#fff',border:'1.5px solid #e5e0d8',borderRadius:10,margin:'11px 0 0'}}>
            <div style={{padding:'11px 18px',borderBottom:'1px solid #f0ece2',display:'flex',alignItems:'center',gap:10}}><span style={{fontSize:11,fontFamily:"'DM Mono',monospace",letterSpacing:'0.12em',color:'#c9a84c',textTransform:'uppercase',fontWeight:700}}>Sources</span><span style={{fontSize:11,color:'#bbb',fontFamily:"'DM Mono',monospace"}}>{enabledSources.size} of {SOURCES.length} enabled</span></div>
            <div style={{padding:'11px 18px 13px',display:'flex',gap:8,flexWrap:'wrap'}}>
              {SOURCES.map(src=>{const on=enabledSources.has(src.id);const ss=sourceStatus[src.id];return<button key={src.id} onClick={()=>toggleSource(src.id)} style={{display:'flex',alignItems:'center',gap:6,padding:'6px 13px',borderRadius:20,border:`1.5px solid ${on?src.color:'#e0dbd0'}`,background:on?src.color+'12':'#faf9f6',color:on?src.color:'#aaa',fontSize:12,fontFamily:"'DM Mono',monospace",fontWeight:on?700:400,cursor:'pointer'}}>{src.emoji} {src.label}{ss==='searching'&&<span style={{fontSize:9,marginLeft:2}}>⟳</span>}{ss==='done'&&<span style={{fontSize:9,marginLeft:2,color:'#2e7d32'}}>✓</span>}{ss==='error'&&<span style={{fontSize:9,marginLeft:2,color:'#c62828'}}>✗</span>}</button>;})}
            </div>
          </div>

          {/* Search controls */}
          <div style={{background:'#fff',border:'1.5px solid #e5e0d8',borderRadius:10,margin:'11px 0 0'}}>
            <div style={{padding:'11px 18px',borderBottom:'1px solid #f0ece2'}}><span style={{fontSize:11,fontFamily:"'DM Mono',monospace",letterSpacing:'0.12em',color:'#c9a84c',textTransform:'uppercase',fontWeight:700}}>Search</span></div>
            <div style={{padding:'12px 18px 16px'}}>
              <div style={{marginBottom:10}}>
                <label style={labelStyle}>Job Titles — Enter or comma to add</label>
                <div style={{display:'flex',flexWrap:'wrap',gap:7,alignItems:'center',padding:'8px 10px',border:'1.5px solid #e0dbd0',borderRadius:8,background:'#faf9f6',minHeight:42}}>
                  {titles.map(t=><TitleTag key={t} title={t} onRemove={removeTitle}/>)}
                  <input value={titleInput} onChange={e=>setTitleInput(e.target.value)} onKeyDown={handleTitleKey} onBlur={addTitle} placeholder={titles.length===0?'e.g. Software Engineer':'Add another…'} style={{border:'none',outline:'none',background:'transparent',fontSize:14,fontFamily:"'Georgia',serif",color:'#1a1a2e',flex:1,minWidth:130}}/>
                </div>
              </div>
              <div style={{display:'grid',gridTemplateColumns:'1.5fr 1fr 1fr 1fr',gap:9}}>
                <div><label style={labelStyle}>Location</label><input value={location} onChange={e=>setLocation(e.target.value)} placeholder="City, State" style={inputStyle}/></div>
                <div><label style={labelStyle}>Work Type</label><select value={locationType} onChange={e=>setLocationType(e.target.value)} style={inputStyle}>{LOCATION_TYPE_OPTIONS.map(o=><option key={o.value} value={o.value}>{o.label}</option>)}</select></div>
                <div><label style={labelStyle}>Level</label><select value={seniority} onChange={e=>setSeniority(e.target.value)} style={inputStyle}>{SENIORITY_OPTIONS.map(o=><option key={o.value} value={o.value}>{o.label}</option>)}</select></div>
                <div><label style={labelStyle}>Min Salary ($)</label><input value={salaryMin} onChange={e=>setSalaryMin(e.target.value)} placeholder="80000" style={inputStyle} type="number"/></div>
              </div>
              <div style={{marginTop:11,display:'flex',gap:8,alignItems:'center',flexWrap:'wrap'}}>
                <button onClick={searchJobs} disabled={loading||isScoringAll||enabledSources.size===0} style={primaryBtn}>{loading?'Searching…':isScoringAll?'Scoring…':`Search ${enabledSources.size} Source${enabledSources.size!==1?'s':''}${resumeReady?' + Score':''}`}</button>
                {dupCount>0&&<button onClick={()=>setHideDuplicates(h=>!h)} style={{...ghostBtn,fontSize:12,color:hideDuplicates?'#c9a84c':'#888',borderColor:hideDuplicates?'#c9a84c':'#e0dbd0'}}>{hideDuplicates?`Show ${dupCount} dupe${dupCount!==1?'s':''}`:`Hide ${dupCount} dupe${dupCount!==1?'s':''}`}</button>}
                {jobs.length>0&&<span style={{marginLeft:'auto',fontSize:12,color:'#888',fontFamily:"'DM Mono',monospace"}}>{jobs.length} total{dupCount>0?` · ${dupCount} dupes`:''}{resumeReady&&scoredCount>0?` · ${scoredCount} scored`:''}</span>}
              </div>
              {error&&<div style={{marginTop:8,padding:'8px 12px',background:'#fff5f5',border:'1px solid #ffc9c9',borderRadius:6,color:'#c0392b',fontSize:13}}>{error}</div>}
            </div>
          </div>

          {/* Results */}
          <div style={{padding:'16px 0 40px'}}>
            {loading&&jobs.length===0&&<div style={{textAlign:'center',padding:'40px 0',color:'#aaa'}}>
              <div style={{fontSize:13,fontFamily:"'DM Mono',monospace",marginBottom:10}}>Searching all sources in parallel…</div>
              <div style={{display:'flex',gap:8,justifyContent:'center',flexWrap:'wrap'}}>
                {SOURCES.filter(s=>enabledSources.has(s.id)).map(src=><span key={src.id} style={{fontSize:11,fontFamily:"'DM Mono',monospace",padding:'3px 10px',borderRadius:20,background:src.color+'18',border:`1px solid ${src.color}44`,color:src.color}}>{src.emoji} {src.label} {sourceStatus[src.id]==='done'?'✓':sourceStatus[src.id]==='error'?'✗':'⟳'}</span>)}
              </div>
            </div>}
            {!loading&&searched&&visibleJobs.length===0&&!error&&<div style={{textAlign:'center',padding:'40px 0',color:'#aaa',fontStyle:'italic'}}>No results found.</div>}
            {sourceGroups.map(({source:src,titleGroups})=>(
              <div key={src.id} style={{marginBottom:26}}>
                <div style={{display:'flex',alignItems:'center',gap:10,marginBottom:12,padding:'8px 14px',background:src.color+'10',border:`1px solid ${src.color}30`,borderRadius:8}}>
                  <span style={{fontSize:14}}>{src.emoji}</span>
                  <span style={{fontSize:13,fontWeight:700,fontFamily:"'DM Mono',monospace",color:src.color}}>{src.label}</span>
                  <span style={{fontSize:11,fontFamily:"'DM Mono',monospace",color:'#aaa'}}>{titleGroups.reduce((a,tg)=>a+tg.jobs.length,0)} results</span>
                  {sourceStatus[src.id]==='searching'&&<span style={{fontSize:11,fontFamily:"'DM Mono',monospace",color:src.color}}>⟳ searching…</span>}
                  {sourceStatus[src.id]==='done'&&<span style={{fontSize:11,fontFamily:"'DM Mono',monospace",color:'#2e7d32'}}>✓</span>}
                  {sourceStatus[src.id]==='error'&&<span style={{fontSize:11,fontFamily:"'DM Mono',monospace",color:'#c62828'}}>✗ error</span>}
                </div>
                {titleGroups.map(({title,jobs:tJobs})=>(
                  <div key={title} style={{marginBottom:12,marginLeft:8}}>
                    {titles.length>1&&<div style={{display:'flex',alignItems:'center',gap:8,marginBottom:7}}><span style={{fontSize:10,fontFamily:"'DM Mono',monospace",letterSpacing:'0.1em',color:'#c9a84c',textTransform:'uppercase',fontWeight:700}}>{title}</span><div style={{flex:1,height:1,background:'#ede9e1'}}/><span style={{fontSize:10,fontFamily:"'DM Mono',monospace",color:'#ccc'}}>{tJobs.length}</span></div>}
                    <div style={{display:'flex',flexDirection:'column',gap:8}}>
                      {tJobs.map(job=><JobCard key={job.id} job={job} status={jobStatus[job.id]||'New'} fit={fitScores[job.id]} isScoring={scoringId===job.id} isDuplicate={duplicates.has(job.id)} salaryData={salaryResearch[job.id]} onResearchSalary={()=>researchSalary(job)} onStatusChange={s=>setJobStatus(job,s)} onScore={()=>scoreJob(job)} resumeReady={resumeReady} onTailor={()=>tailorResume(job)} tailored={!!tailoredResumes[job.id]}/>)}
                    </div>
                  </div>
                ))}
              </div>
            ))}
          </div>
        </>}

        {/* ── SAVED TAB ── */}
        {activeTab==='saved'&&<div style={{padding:'16px 0 40px'}}>
          {sheetJobs.length===0&&!syncing&&<div style={{textAlign:'center',padding:'50px 0',color:'#aaa',fontStyle:'italic'}}>No saved jobs yet.</div>}
          {syncing&&<div style={{textAlign:'center',padding:'40px 0',color:'#aaa',fontFamily:"'DM Mono',monospace",fontSize:13}}>Loading from your Google Sheets…</div>}
          {['Saved','Applied','Skipped'].map(status=>{
            const group=sheetJobs.filter(j=>(jobStatus[j.id]||j.status)===status);
            if(!group.length) return null;
            return<div key={status} style={{marginBottom:22}}>
              <div style={{display:'flex',alignItems:'center',gap:10,marginBottom:10}}><span style={{fontSize:11,fontFamily:"'DM Mono',monospace",letterSpacing:'0.12em',color:'#c9a84c',textTransform:'uppercase',fontWeight:700}}>{status}</span><div style={{flex:1,height:1,background:'#e5e0d8'}}/><span style={{fontSize:11,fontFamily:"'DM Mono',monospace",color:'#bbb'}}>{group.length}</span></div>
              <div style={{display:'flex',flexDirection:'column',gap:8}}>
                {group.map(job=><JobCard key={job.id} job={job} status={jobStatus[job.id]||job.status||'Saved'} fit={fitScores[job.id]||(job.fitScore?{score:parseInt(job.fitScore),highlights:job.fitNotes?job.fitNotes.split(';'):[]}:null)} isScoring={scoringId===job.id} isDuplicate={false} salaryData={salaryResearch[job.id]} onResearchSalary={()=>researchSalary(job)} onStatusChange={s=>setJobStatus(job,s)} onScore={()=>scoreJob(job)} resumeReady={resumeReady} onTailor={()=>tailorResume(job)} tailored={!!tailoredResumes[job.id]} showSource/>)}
              </div>
            </div>;
          })}
        </div>}

        {/* ── TAILORED TAB ── */}
        {activeTab==='tailored'&&<div style={{padding:'16px 0 40px'}}>
          {tailoredCount===0&&<div style={{textAlign:'center',padding:'60px 0',color:'#aaa'}}><div style={{fontSize:30,marginBottom:10}}>📝</div><div style={{fontStyle:'italic',marginBottom:6}}>No tailored resumes yet.</div><div style={{fontSize:13,color:'#bbb'}}>Click "✨ Tailor Resume" on any job card.</div></div>}
          {tailoredCount>0&&<div style={{display:'grid',gridTemplateColumns:'250px 1fr',gap:16,alignItems:'start'}}>
            <div style={{background:'#fff',border:'1.5px solid #e5e0d8',borderRadius:10,overflow:'hidden',position:'sticky',top:16}}>
              <div style={{padding:'10px 14px',borderBottom:'1px solid #f0ece2'}}><span style={{fontSize:11,fontFamily:"'DM Mono',monospace",letterSpacing:'0.1em',color:'#c9a84c',textTransform:'uppercase',fontWeight:700}}>Tailored Versions</span></div>
              {Object.entries(tailoredResumes).map(([jobId,tr])=>{const job=allJobsForTailoring.find(j=>j.id===jobId)||{id:jobId,title:'Unknown',company:''};const isSel=selectedTailorJob?.id===jobId;return<div key={jobId} onClick={()=>setSelectedTailorJob(job)} style={{padding:'10px 14px',borderBottom:'1px solid #f5f2ed',cursor:'pointer',background:isSel?'#fffdf4':'transparent',borderLeft:isSel?'3px solid #c9a84c':'3px solid transparent'}}><div style={{fontSize:13,fontWeight:600,color:'#1a1a2e',marginBottom:1}}>{job.title}</div><div style={{fontSize:12,color:'#888'}}>{job.company}</div>{tr.loading&&<div style={{fontSize:11,color:'#c9a84c',fontFamily:"'DM Mono',monospace",marginTop:2}}>Generating…</div>}{!tr.loading&&tr.text&&<div style={{fontSize:11,color:'#2e7d32',fontFamily:"'DM Mono',monospace",marginTop:2}}>✓ Ready</div>}{tr.error&&<div style={{fontSize:11,color:'#c0392b',fontFamily:"'DM Mono',monospace",marginTop:2}}>⚠ Error</div>}</div>;})}
            </div>
            <div>
              {!selectedTailorJob&&<div style={{padding:40,textAlign:'center',color:'#aaa',fontStyle:'italic'}}>Select a version</div>}
              {selectedTailorJob&&(()=>{const tr=tailoredResumes[selectedTailorJob.id];if(!tr)return null;return<div style={{background:'#fff',border:'1.5px solid #e5e0d8',borderRadius:10,overflow:'hidden'}}>
                <div style={{padding:'12px 18px',borderBottom:'1px solid #f0ece2',display:'flex',alignItems:'center',gap:10,flexWrap:'wrap'}}>
                  <div><div style={{fontSize:14,fontWeight:600,color:'#1a1a2e'}}>Tailored for: {selectedTailorJob.title}</div><div style={{fontSize:12,color:'#888'}}>{selectedTailorJob.company}</div></div>
                  <div style={{marginLeft:'auto',display:'flex',gap:7}}>
                    {tr.text&&<><button onClick={()=>navigator.clipboard.writeText(tr.text)} style={{...ghostBtn,fontSize:12}}>📋 Copy</button><button onClick={()=>downloadDocx(tr.text,selectedTailorJob.title,selectedTailorJob.company)} disabled={downloadingId===selectedTailorJob.id} style={{...primaryBtn,fontSize:12,background:'#c9a84c',color:'#1a1a2e'}}>{downloadingId===selectedTailorJob.id?'Generating…':'⬇ .docx'}</button></>}
                    <button onClick={()=>tailorResume(selectedTailorJob)} disabled={tr.loading} style={{...ghostBtn,fontSize:12}}>↻ Redo</button>
                  </div>
                </div>
                <div style={{padding:'20px 24px',minHeight:300}}>
                  {tr.loading&&<div style={{textAlign:'center',padding:'40px 0',color:'#aaa'}}><div style={{fontSize:13,fontFamily:"'DM Mono',monospace"}}>✨ Tailoring resume…</div><div style={{fontSize:12,color:'#ccc',marginTop:4}}>~15–20 seconds</div></div>}
                  {tr.error&&<div style={{color:'#c0392b',fontSize:13}}>Error: {tr.error}</div>}
                  {!tr.loading&&!tr.error&&tr.text&&<ResumePreview text={tr.text}/>}
                </div>
              </div>;})()}
            </div>
          </div>}
        </div>}
      </div>
    </div>
  );
}

const labelStyle={display:'block',fontSize:11,fontFamily:"'DM Mono',monospace",letterSpacing:'0.1em',color:'#888',textTransform:'uppercase',marginBottom:4};
const inputStyle={width:'100%',padding:'8px 10px',border:'1.5px solid #e0dbd0',borderRadius:6,fontSize:14,fontFamily:"'Georgia',serif",color:'#1a1a2e',background:'#faf9f6',outline:'none',boxSizing:'border-box'};
const primaryBtn={padding:'9px 18px',background:'#1a1a2e',color:'#f0ece2',border:'none',borderRadius:6,fontSize:13,fontFamily:"'DM Mono',monospace",letterSpacing:'0.05em',cursor:'pointer',fontWeight:600};
const ghostBtn={padding:'7px 13px',background:'transparent',color:'#555',border:'1.5px solid #e0dbd0',borderRadius:6,fontSize:13,fontFamily:"'DM Mono',monospace",cursor:'pointer'};
