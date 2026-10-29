import {initializeApp} from 'https://www.gstatic.com/firebasejs/10.12.2/firebase-app.js';
import {getAuth,onAuthStateChanged,signInWithEmailAndPassword,createUserWithEmailAndPassword,sendPasswordResetEmail,sendEmailVerification,updateProfile,signOut,GoogleAuthProvider,signInWithPopup} from 'https://www.gstatic.com/firebasejs/10.12.2/firebase-auth.js';
import {initializeFirestore,persistentLocalCache,persistentMultipleTabManager,doc as fdoc,collection as fcol,getDoc,getDocs,setDoc,updateDoc,deleteDoc,onSnapshot as fsnap,query as fquery,orderBy as forder,limit as flimit} from 'https://www.gstatic.com/firebasejs/10.12.2/firebase-firestore.js';
import {firebaseConfig,OWNER_EMAIL} from './firebase-config.js';


const NS='http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const RNS='http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PKG='http://schemas.openxmlformats.org/package/2006/relationships';
const XMLNS_XML='http://www.w3.org/XML/1998/namespace';
const STATUSES=['DONE','ONGOING','NOT DONE'];
const BLOCKERS=[
  ['Scaffold not ready','Scaffolding'],
  ['Permit / isolation','Permits'],
  ['Crane / lifting not available','Lifting'],
  ['Lifting plan / AP','Lifting'],
  ['QC hold / inspection','QC'],
  ['Material not on site','Materials'],
  ['Access / area clash','Area coordination'],
  ['Manpower','Resourcing'],
  ['Drawing / engineering query','Engineering'],
  ['Time ran out','Planning'],
  ['Other','Other']
];
const STORE_KEY='dabsNight.v1';

let doc=null;      // parsed DABS {fileName, sections:[{title,trades,items}]}
let bytes=null;    // original file Uint8Array
let st={};         // per item state {status:{col:val}, blockers:[], note:'', extra:[cols], open:bool}
let prepBy='';
let filter={view:'all',level:null};
let added=[];      // activities added on the night, not on the DABS
const itemsOf=si=>[...doc.sections[si].items,...added.filter(a=>a.sec===si)];
const allItems=()=>doc.sections.flatMap((s,i)=>itemsOf(i));

/* ---------- utils ---------- */
const $=s=>document.querySelector(s);
const esc=s=>String(s??'').replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
function colNum(c){let n=0;for(const ch of c)n=n*26+ch.charCodeAt(0)-64;return n}
function colName(n){let s='';while(n>0){const m=(n-1)%26;s=String.fromCharCode(65+m)+s;n=Math.floor((n-1)/26)}return s}
function splitRef(r){const m=/^([A-Z]+)(\d+)$/.exec(r);return [colNum(m[1]),+m[2]]}
const parseXml=s=>new DOMParser().parseFromString(s,'application/xml');
const serXml=d=>new XMLSerializer().serializeToString(d);
const kids=(el,name)=>Array.from(el.children).filter(c=>c.localName===name);
const kid=(el,name)=>kids(el,name)[0]||null;
function toast(msg){const t=$('#toast');t.textContent=msg;t.hidden=false;clearTimeout(toast._t);toast._t=setTimeout(()=>t.hidden=true,2200)}
function b64enc(u8){let s='';for(let i=0;i<u8.length;i+=0x8000)s+=String.fromCharCode.apply(null,u8.subarray(i,i+0x8000));return btoa(s)}
function b64dec(b){const s=atob(b);const u=new Uint8Array(s.length);for(let i=0;i<s.length;i++)u[i]=s.charCodeAt(i);return u}
function fmtLevel(v){if(typeof v==='number'){const s=v.toFixed(2);return v>0?'+'+s:s}const t=String(v??'').trim();return t||'—'}
function rowStatus(it){
  const s=st[it.id];const vals=allCols(it).map(c=>s.status[c]||null);
  if(!vals.length)return null;
  if(vals.every(v=>v==='DONE'))return 'DONE';
  if(vals.every(v=>v==='NOT DONE'))return 'NOT DONE';
  if(vals.some(v=>v==='DONE'||v==='ONGOING'))return 'ONGOING';
  if(vals.every(v=>v===null))return null;
  return 'NOT DONE';
}
const allCols=it=>[...it.trades.map(t=>t.col),...(st[it.id]?.extra||[])];

/* ---------- xlsx reading ---------- */
async function sheetPaths(zip){
  const wb=parseXml(await zip.file('xl/workbook.xml').async('string'));
  const sheet=wb.getElementsByTagNameNS(NS,'sheet')[0];
  const rid=sheet.getAttributeNS(RNS,'id');
  const rels=parseXml(await zip.file('xl/_rels/workbook.xml.rels').async('string'));
  let target=null;
  for(const r of rels.getElementsByTagNameNS(PKG,'Relationship'))if(r.getAttribute('Id')===rid)target=r.getAttribute('Target');
  if(!target)throw new Error('Could not find the first worksheet in this file.');
  const path=target.startsWith('/')?target.slice(1):'xl/'+target.replace(/^\.\//,'');
  return {path,name:sheet.getAttribute('name')};
}
async function readShared(zip){
  const f=zip.file('xl/sharedStrings.xml');if(!f)return [];
  const d=parseXml(await f.async('string'));
  return Array.from(d.getElementsByTagNameNS(NS,'si')).map(si=>{
    let s='';for(const t of si.getElementsByTagNameNS(NS,'t')){if(t.parentNode.localName==='rPh')continue;s+=t.textContent}return s});
}
function cellValue(c,ss){
  const t=c.getAttribute('t');
  if(t==='inlineStr'){let s='';for(const x of c.getElementsByTagNameNS(NS,'t'))s+=x.textContent;return s}
  const v=kid(c,'v');if(!v)return null;
  if(t==='s')return ss[+v.textContent]??'';
  if(t==='str'||t==='e')return v.textContent;
  if(t==='b')return v.textContent==='1';
  const n=Number(v.textContent);return isNaN(n)?v.textContent:n;
}
async function parseDabs(u8,fileName){
  const zip=await JSZip.loadAsync(u8);
  const {path}=await sheetPaths(zip);
  const ss=await readShared(zip);
  const sx=parseXml(await zip.file(path).async('string'));
  const grid=new Map();let maxRow=0;
  for(const c of sx.getElementsByTagNameNS(NS,'c')){
    const [col,row]=splitRef(c.getAttribute('r'));const v=cellValue(c,ss);
    if(v!==null&&v!=='')grid.set(row+','+col,v);if(row>maxRow)maxRow=row;
  }
  const own=(r,c)=>{const v=grid.get(r+','+c);return v==null?'':v};
  const merged=new Map(grid);
  for(const m of sx.getElementsByTagNameNS(NS,'mergeCell')){
    const [a,b]=m.getAttribute('ref').split(':');if(!b)continue;
    const [c1,r1]=splitRef(a),[c2,r2]=splitRef(b);const v=grid.get(r1+','+c1);
    if(v==null)continue;for(let r=r1;r<=r2;r++)for(let c=c1;c<=c2;c++)merged.set(r+','+c,v);
  }
  const mv=(r,c)=>{const v=merged.get(r+','+c);return v==null?'':v};
  const txt=v=>String(v??'').trim();
  // header rows
  const headers=[];
  for(let r=1;r<=maxRow;r++){
    for(let c=1;c<=60;c++){if(txt(own(r,c)).toLowerCase()==='activity description'){headers.push({row:r,actCol:c});break}}
  }
  if(!headers.length)throw new Error('No "Activity Description" header found. Is this a DABS handover sheet?');
  const sections=[];
  headers.forEach((h,i)=>{
    const cols={};const hdr=[];
    for(let c=1;c<=60;c++){const t=txt(own(h.row,c));if(t)hdr.push([c,t])}
    for(const [c,t] of hdr){const l=t.toLowerCase();
      if(l==='level')cols.level=c;else if(l==='shift')cols.shift=c;else if(l.startsWith('black'))cols.bh=c;
      else if(l.includes('remark')||l.includes('blocker'))cols.remark=c;}
    cols.act=h.actCol;
    const before=hdr.filter(([c])=>c<h.actCol&&![cols.level,cols.shift,cols.bh].includes(c));
    cols.tag=before.length?before[before.length-1][0]:h.actCol-1;
    const tagLabel=before.length?before[before.length-1][1]:'Tag';
    const trades=hdr.filter(([c])=>c>h.actCol&&c!==cols.remark&&(!cols.remark||c<cols.remark)).map(([c,t])=>({col:c,name:t}));
    if(!cols.remark)cols.remark=(trades.length?trades[trades.length-1].col:h.actCol)+1;
    let title='';let titleRow=h.row;
    for(let r=h.row-1;r>=Math.max(1,h.row-4);r--){const vals=[];for(let c=1;c<=20;c++){const t=txt(own(r,c));if(t)vals.push(t)}
      if(vals.length){title=vals[0];titleRow=r;for(let c=1;c<=20;c++)if(txt(own(r,c))){h.titleCol=c;break}break}}
    h.titleRow=titleRow;
    sections.push({title:title||('Section '+(i+1)),header:h.row,titleRow,titleCol:h.titleCol||cols.level||1,hdr,cols,trades,tagLabel,items:[]});
  });
  sections.forEach((s,i)=>{
    const end=i+1<headers.length?headers[i+1].titleRow-1:maxRow;
    for(let r=s.header+1;r<=end;r++){
      const act=txt(own(r,s.cols.act)),tag=txt(own(r,s.cols.tag));
      if(!act&&!tag)continue;
      const trades=[];const pre={};
      for(const t of s.trades){const v=txt(own(r,t.col)).toUpperCase();
        if(v==='X'||STATUSES.includes(v)){trades.push(t);if(STATUSES.includes(v))pre[t.col]=v}}
      s.items.push({id:i+'-'+r,row:r,sec:i,level:fmtLevel(s.cols.level?own(r,s.cols.level):''),
        shift:txt(own(r,s.cols.shift)),bh:txt(mv(r,s.cols.bh)),tag,act,dayRemark:txt(own(r,s.cols.remark)),trades,pre});
    }
  });
  const shortName=t=>t.replace(/^dabs\s*/i,'').replace(/\s+\d.*$/,'').trim()||t;
  sections.forEach(s=>s.short=shortName(s.title));
  return {fileName,sections};
}

/* ---------- state ---------- */
function blankState(it){return {status:{...(it.pre||{})},blockers:[],note:'',extra:[]}}
function initState(){for(const it of allItems())if(!st[it.id])st[it.id]=blankState(it)}

/* ---------- shared store (Firebase, works offline and syncs when signal returns) ---------- */
let db=null,authUser=null,canWrite=false,isAdmin=false,ready=false;
let nights=[],cur=null,unsubItems=null,unsubNights=null,opening=false;
const meta={currentId:null,prepBy:''};
const LOCAL='dabsNight.pwa';
function loadLocal(){try{Object.assign(meta,JSON.parse(localStorage.getItem(LOCAL)||'{}'))}catch(e){}}
function saveLocal(){try{localStorage.setItem(LOCAL,JSON.stringify(meta))}catch(e){}}
const myName=()=>(authUser&&(authUser.displayName||(authUser.email||'').split('@')[0]))||'';
const byDate=()=>nights.slice().sort((a,b)=>b.loadedAt-a.loadedAt);
const latest=()=>byDate()[0]||null;
function nightLabel(d,loadedAt){
  const m=/(\d{1,2})(?:st|nd|rd|th)?[\s\-.]*([A-Za-z]{3})/.exec(d.sections.map(s=>s.title).join(' '));
  if(m)return m[1]+' '+m[2][0].toUpperCase()+m[2].slice(1,3).toLowerCase();
  return new Date(loadedAt).toLocaleDateString('en-GB',{day:'numeric',month:'short'});
}
function hashStr(s){let h=0x811c9dc5;for(let i=0;i<s.length;i++){h^=s.charCodeAt(i);h=Math.imul(h,16777619)}return (h>>>0).toString(36)}
function compactItems(){return allItems().map(it=>{const s=st[it.id]||{};return {t:it.tag,a:it.act,l:it.level,s:rowStatus(it),b:s.blockers||[],x:!!it.added}})}

let localPending=false;
function setSync(s){const el=$('#sync');if(!el)return;el.dataset.s=s;
  el.textContent={ok:'Synced',saving:'Sending…',offline:'Saved on this phone · sends when signal returns',ro:'View only',load:'Connecting…'}[s]||''}
function updateSync(){if(!authUser)return;if(!canWrite){setSync('ro');return}setSync(localPending?(navigator.onLine?'saving':'offline'):'ok')}
window.addEventListener('online',updateSync);window.addEventListener('offline',updateSync);

const timers={},dirty=new Set();
function writeItem(id,delay){
  if(!cur||!canWrite)return;const nid=cur.id;
  dirty.add(id);clearTimeout(timers[id]);
  timers[id]=setTimeout(()=>{
    dirty.delete(id);const s=st[id];const it=findItem(id);if(!s||!cur||cur.id!==nid)return;
    const body={status:s.status,blockers:s.blockers,note:s.note,extra:s.extra,open:!!s.open,by:authUser.uid,byName:myName(),at:Date.now()};
    if(it&&it.added)body.added={sec:it.sec,level:it.level,tag:it.tag,act:it.act,trades:it.trades,created:it.created||Date.now(),createdByName:it.createdByName||myName()};
    s.by=body.by;s.byName=body.byName;s.at=body.at;
    db.doc('nights/'+nid+'/items/'+id).set(body).catch(e=>toast(e&&e.code==='permission-denied'?'Not saved: your access was changed. Ask the night shift lead.':'Not saved ('+((e&&e.code)||'error')+')'));
    localPending=true;updateSync();scheduleSummary();
  },delay||0);
}
let sumT=null;
function scheduleSummary(){clearTimeout(sumT);sumT=setTimeout(()=>{
  if(!cur||!canWrite)return;
  db.doc('nights/'+cur.id).update({summary:counts(),items:compactItems(),updatedAt:Date.now()}).catch(()=>{})},1500)}

function applyItemDoc(id,d){
  if(d.added){
    const info={id,sec:d.added.sec,added:true,level:d.added.level,shift:'Night',bh:'',tag:d.added.tag||'',act:d.added.act||'',dayRemark:'',trades:d.added.trades||[],pre:{},created:d.added.created,createdByName:d.added.createdByName};
    const a=added.find(x=>x.id===id);if(a)Object.assign(a,info);else added.push(info);
  }
  if(dirty.has(id))return false;      // this phone has a newer edit waiting to be written
  const focused=document.activeElement&&document.activeElement.id==='note-'+id;
  st[id]={status:{...(d.status||{})},blockers:[...(d.blockers||[])],note:focused&&st[id]?st[id].note:(d.note||''),extra:[...(d.extra||[])],open:!!d.open,by:d.by||null,byName:d.byName||'',at:d.at||0};
  return true;
}
function subscribeItems(){
  if(unsubItems)unsubItems();
  const nid=cur.id;let first=true;
  unsubItems=db.collection('nights/'+nid+'/items').onSnapshot(snap=>{
    if(!cur||cur.id!==nid)return;
    localPending=!!(snap.metadata&&snap.metadata.hasPendingWrites);updateSync();
    let structural=first;const changed=[];
    for(const ch of snap.docChanges()){
      const id=ch.doc.id;
      if(ch.type==='removed'){
        if(added.some(a=>a.id===id)){added=added.filter(a=>a.id!==id);delete st[id];structural=true}
        else{const it=findItem(id);if(it&&!dirty.has(id)){st[id]=blankState(it);changed.push(id)}}
        continue;
      }
      const d=ch.doc.data()||{};
      if(d.added&&!added.some(a=>a.id===id))structural=true;
      if(applyItemDoc(id,d))changed.push(id);
    }
    first=false;initState();
    if(structural)renderAll();else{changed.forEach(id=>rerenderCard(id,true));renderTop()}
  },e=>{toast('Live updates stopped ('+(e.code||'error')+'). Reopen the app.')},true);
}
function subscribeNights(){
  if(unsubNights)return;
  unsubNights=db.collection('nights').orderBy('loadedAt','desc').limit(300).onSnapshot(async snap=>{
    nights=snap.docs.map(d=>Object.assign({id:d.id},d.data()));histIdx=null;ready=true;
    if(opening)return;
    if(!cur){const want=nights.find(n=>n.id===meta.currentId)||latest();
      if(want){try{await openNight(want)}catch(e){renderAll();toast(e.message)}}else renderAll();return}
    const c=nights.find(n=>n.id===cur.id);
    if(!c){if(unsubItems)unsubItems();unsubItems=null;cur=null;doc=null;bytes=null;st={};added=[];
      const l=latest();if(l)await openNight(l);else renderAll();toast('That night was deleted by an admin');return}
    Object.assign(cur,c);renderTop();renderBanner();if(archiveOpen)openArchive();
  },e=>{if(e.code==='permission-denied')return;$('#main').innerHTML=`<p class="err">Could not load the team's DABS (${esc(e.code||'error')}). Reopen the app.</p>`});
}
function stopListening(){if(unsubItems)unsubItems();if(unsubNights)unsubNights();unsubItems=unsubNights=null;nights=[];cur=null;doc=null;bytes=null;st={};added=[];ready=false}
async function openNight(n){
  opening=true;
  try{
    const f=await db.doc('files/'+n.id).get();
    if(!f.exists)throw new Error('The DABS file for '+(n.label||'this night')+' is not on this phone yet. Open it once with signal.');
    const u8=b64dec(f.data().b64);const d=await parseDabs(u8,n.fileName);
    if(unsubItems){unsubItems();unsubItems=null}
    bytes=u8;doc=d;added=[];st={};dirty.clear();initState();
    cur=n;meta.currentId=n.id;saveLocal();filter={view:'all',level:null};histIdx=null;
    renderAll();subscribeItems();
  }finally{opening=false}
}
async function uploadDabs(f){
  const u8=new Uint8Array(await f.arrayBuffer());const d=await parseDabs(u8,f.name);const b64=b64enc(u8);
  if(b64.length>900000)throw new Error('This file is too large to share (over 650 KB). Remove pictures or extra sheets and try again.');
  const id='d'+hashStr(b64)+u8.length.toString(36);
  const ex=nights.find(n=>n.id===id);
  if(ex){await openNight(ex);toast('Already shared. Opened '+ex.label);return}
  if(!canWrite)throw new Error('You can view the DABS but not add files. Ask the night shift lead for Contributor access.');
  if(!navigator.onLine)throw new Error('No signal. Sharing a new DABS needs a connection; try again near the cabins or on Wi-Fi.');
  const now=Date.now();
  const m={fileName:f.name,label:nightLabel(d,now),loadedAt:now,updatedAt:now,createdBy:authUser.uid,createdByName:myName(),summary:null,items:[]};
  await db.doc('files/'+id).set({b64,fileName:f.name});
  await db.doc('nights/'+id).set(m);
  await openNight(Object.assign({id},m));
  scheduleSummary();
  toast(d.sections.reduce((k,s)=>k+s.items.length,0)+' activities shared with the team');
}
function resolveNames(){}

/* ---------- history of a tag across earlier nights ---------- */
let histIdx=null;
function tokens(tag){
  const out=[];let prefix='';
  for(let p of String(tag||'').toUpperCase().split(/&|,|\+|\bAND\b/)){
    p=p.replace(/\s+/g,'');if(!p||!/\d/.test(p))continue;
    const m=/^(.*?--)(.+)$/.exec(p);
    if(m){prefix=m[1];out.push(p)}else if(prefix&&/^[A-Z]{1,3}\d+$/.test(p))out.push(prefix+p);else out.push(p);
  }
  return out;
}
function buildHist(){
  histIdx=new Map();if(!cur)return;
  for(const n of nights){if(n.id===cur.id||n.loadedAt>=cur.loadedAt)continue;
    for(const it of n.items||[])for(const tk of tokens(it.t)){
      const arr=histIdx.get(tk)||[];arr.push({id:n.id,label:n.label||'',at:n.loadedAt,a:it.a,s:it.s,b:it.b||[]});histIdx.set(tk,arr)}}
}
function historyOf(tag){
  if(!histIdx)buildHist();const seen=new Map();
  for(const tk of tokens(tag))for(const h of histIdx.get(tk)||[]){const k=h.id+'|'+h.a;if(!seen.has(k))seen.set(k,h)}
  return [...seen.values()].sort((a,b)=>b.at-a.at);
}
const nightsNotDone=h=>new Set(h.filter(x=>x.s!=='DONE').map(x=>x.id)).size;

/* ---------- rendering ---------- */
function counts(){
  const c={DONE:0,ONGOING:0,'NOT DONE':0,unset:0,total:0};
  for(const it of allItems()){const r=rowStatus(it);c.total++;if(r)c[r]++;else c.unset++}
  return c;
}
function renderTop(){
  $('#fileName').textContent=doc?(cur&&cur.label?cur.label+' · ':'')+doc.fileName:(authUser?'No DABS loaded':'');
  $('#btnArchive').hidden=!authUser||!ready;
  $('#btnNew').textContent=doc?'New DABS':'Load DABS';$('#btnNew').hidden=!authUser||!canWrite;
  $('#progressBox').hidden=!doc;$('#filters').hidden=!doc;$('#actions').hidden=!doc;
  if(!doc)return;
  const c=counts();const pct=n=>c.total?(n/c.total*100)+'%':'0';
  $('#bar').innerHTML=`<span style="width:${pct(c.DONE)};background:var(--ok)"></span><span style="width:${pct(c.ONGOING)};background:var(--go)"></span><span style="width:${pct(c['NOT DONE'])};background:var(--no)"></span>`;
  $('#counts').innerHTML=`<span><i class="dot" style="background:var(--ok)"></i>DONE <b>${c.DONE}</b></span><span><i class="dot" style="background:var(--go)"></i>ONGOING <b>${c.ONGOING}</b></span><span><i class="dot" style="background:var(--no)"></i>NOT DONE <b>${c['NOT DONE']}</b></span><span><i class="dot" style="background:var(--unset)"></i>Not updated <b>${c.unset}</b></span><span>of <b>${c.total}</b> activities</span>`;
}
function renderFilters(){
  const levels=[...new Set(allItems().map(i=>i.level))].sort((a,b)=>parseFloat(a)-parseFloat(b));
  const f=[['all','All'],['open','Open items'],...doc.sections.map((s,i)=>['sec'+i,s.short])];
  $('#filters').innerHTML=f.map(([k,l])=>`<button class="chip" type="button" data-view="${k}" aria-pressed="${filter.view===k}">${esc(l)}</button>`).join('')
   +levels.map(l=>`<button class="chip" type="button" data-level="${esc(l)}" aria-pressed="${filter.level===l}">Lvl ${esc(l)}</button>`).join('');
}
function visible(it){
  if(filter.level!==null&&it.level!==filter.level)return false;
  if(filter.view==='open'){const r=rowStatus(it);return r!=='DONE'}
  if(filter.view.startsWith('sec'))return it.sec===+filter.view.slice(3);
  return true;
}
function cardHtml(it){
  const s=st[it.id];const sec=doc.sections[it.sec];const rs=rowStatus(it);
  const cls=rs?'s-'+rs.replace(' ',''):'';
  const cols=allCols(it);
  const trades=cols.map(col=>{const t=sec.trades.find(x=>x.col===col);const v=s.status[col]||null;
    return `<div class="trade"><span class="tn">${esc(t?t.name:colName(col))}</span><div class="seg" role="group" aria-label="${esc(t?t.name:'')} status">${
      STATUSES.map(x=>`<button type="button" data-act="st" data-id="${it.id}" data-col="${col}" data-v="${x}" class="${v===x?'on':''}" aria-pressed="${v===x}">${x}</button>`).join('')}</div></div>`}).join('');
  const missing=sec.trades.filter(t=>!cols.includes(t.col));
  const needBlk=s.open||s.blockers.length||s.note||rs==='NOT DONE'||rs==='ONGOING';
  const blk=needBlk?`<div class="blk"><span class="lbl">Blocker / reason</span><div class="bchips">${
    BLOCKERS.map(([b])=>`<button type="button" class="bchip" data-act="blk" data-id="${it.id}" data-b="${esc(b)}" aria-pressed="${s.blockers.includes(b)}">${esc(b)}</button>`).join('')
    }</div><textarea id="note-${it.id}" data-act="note" data-id="${it.id}" placeholder="Details for the day shift (who, what is needed)">${esc(s.note)}</textarea></div>`:'';
  return `<article class="card ${cls}" id="card-${it.id}">
    <div class="head"><div class="lvl"><small>LEVEL</small>${esc(it.level)}</div>
      <div class="tagbox">${it.added?'<span class="badge">Added on night</span>':''}<div class="tag">${esc(it.tag||'—')}</div><div class="act">${esc(it.act)}</div>
</div></div>
    ${it.dayRemark?`<div class="daynote"><em>Day shift</em>${esc(it.dayRemark)}</div>`:''}
    ${histHtml(it)}
    ${s.by?`<div class="who">Updated by ${esc(authUser&&s.by===authUser.uid?'you':(s.byName||'a teammate'))}${s.at?' · '+new Date(s.at).toLocaleTimeString('en-GB',{hour:'2-digit',minute:'2-digit'}):''}</div>`:''}
    ${trades||'<div class="meta">No trade marked on the DABS. Add the trade that does this job.</div>'}
    ${blk}
    <div class="row-tools">
      ${!needBlk?`<button class="link" type="button" data-act="openblk" data-id="${it.id}">+ Blocker / note</button>`:''}
      ${it.added?`<button class="link danger" type="button" data-act="remove" data-id="${it.id}">Remove</button>`:''}
      ${missing.length?`<select data-act="addtrade" data-id="${it.id}" aria-label="Add trade"><option value="">+ Add trade</option>${missing.map(t=>`<option value="${t.col}">${esc(t.name)}</option>`).join('')}</select>`:''}
    </div>
  </article>`;
}
function histHtml(it){
  const h=historyOf(it.tag);if(!h.length)return '';
  const nd=nightsNotDone(h);const tone=s=>s==='DONE'?'ok':s==='ONGOING'?'go':s==='NOT DONE'?'no':'un';
  return `<details class="hist"><summary><span class="hl">${nd?`Carried over · ${nd} earlier ${nd===1?'night':'nights'} not done`:'Seen on earlier DABS'}</span><span class="hc">${h.length}</span></summary>
    <ul>${h.slice(0,6).map(x=>`<li><span class="hd">${esc(x.label)}</span><span class="hs ${tone(x.s)}">${esc(x.s||'NOT UPDATED')}</span><span class="ha">${esc(x.a)}${x.b.length?' · '+esc(x.b.join(', ')):''}</span></li>`).join('')}</ul></details>`;
}
function renderMain(){
  const m=$('#main');
  renderBanner();
  if(!doc){
    if(!authUser)return;
    if(!ready){m.innerHTML=`<p class="loading">Loading the team's DABS…</p>`;return}
    if(!canWrite){m.innerHTML=`<section class="empty"><h2>No DABS shared yet</h2><p>When the night shift lead loads tonight's DABS it appears here.</p></section>`;return}
    m.innerHTML=`<section class="empty">
    <h2>Share tonight's DABS with the team</h2>
    <p>Load the Excel the day shift sent. Everyone on the night shift then marks jobs from their own phone, even without signal, and the export comes back in the same format with statuses and colours, plus an English handover.</p>
    <label class="drop" for="fileInput"><strong>Choose DABS .xlsx</strong><br><span style="color:var(--muted);font-size:13px">Shared with your approved team.</span></label>
    <ol class="steps"><li>Tap DONE, ONGOING or NOT DONE per trade.</li><li>Pick the blocker and add a short note.</li><li>Export Excel and copy the handover text.</li></ol>
  </section>`;return}
  m.innerHTML=doc.sections.map((sec,si)=>{
    const items=itemsOf(si).filter(visible);
    const addBtn=`<button class="addact" type="button" data-act="addnew" data-sec="${si}">+ Add activity to ${esc(sec.short)}</button>`;
    if(filter.view.startsWith('sec')&&filter.view!=='sec'+si)return '';
    if(!items.length&&filter.view==='open')return '';
    const bh=[...new Set(sec.items.map(i=>i.bh).filter(Boolean))].map(b=>b.replace(/\s*\n\s*(?=\()/g,' ').replace(/\s*\n\s*/g,' · ')).join(' · ');
    return `<section class="sec"><h2>${esc(sec.title)}</h2>${bh?`<p class="meta" style="margin:-4px 0 10px">Black hat: ${esc(bh)}</p>`:''}<div class="list">${items.map(cardHtml).join('')}${addBtn}</div></section>`;
  }).join('')||`<p style="color:var(--muted)">Nothing matches this filter.</p>`;
}
function renderBanner(){
  const b=$('#banner');const lt=latest();
  if(!(doc&&cur&&lt&&lt.id!==cur.id)){b.innerHTML='';return}
  b.innerHTML=`<div class="older"><span>You are viewing an older night (${esc(cur.label)}). Changes are saved to that night.</span><button class="btn alt" type="button" id="goLatest">Open ${esc(lt.label)}</button></div>`;
  $('#goLatest').onclick=async()=>{try{await openNight(latest());window.scrollTo(0,0)}catch(e){toast(e.message)}};
}
function renderAll(){renderTop();if(doc)renderFilters();renderMain()}
function rerenderCard(id,remote){
  const it=findItem(id);const el=$('#card-'+id);if(!el||!it)return;
  const focusNote=document.activeElement&&document.activeElement.id==='note-'+id;
  if(remote&&focusNote)return;
  el.outerHTML=cardHtml(it);renderTop();
  if(focusNote){const t=$('#note-'+id);t.focus();t.setSelectionRange(t.value.length,t.value.length)}
}
function findItem(id){return allItems().find(i=>i.id===id)}

/* ---------- events ---------- */
document.addEventListener('click',e=>{
  const b=e.target.closest('[data-act],[data-view],[data-level]');if(!b)return;
  if(b.dataset.view){filter.view=b.dataset.view;renderFilters();renderMain();return}
  if(b.dataset.level!==undefined&&!b.dataset.act){filter.level=filter.level===b.dataset.level?null:b.dataset.level;renderFilters();renderMain();return}
  if(b.closest('#panel'))return;
  const id=b.dataset.id;const s=st[id];
  if(b.dataset.act==='addnew'){if(canWrite)openAddForm(+b.dataset.sec);return}
  if(!canWrite||!s)return;
  if(b.dataset.act==='st'){const col=+b.dataset.col;s.status[col]=s.status[col]===b.dataset.v?null:b.dataset.v;rerenderCard(id);writeItem(id)}
  else if(b.dataset.act==='blk'){const x=b.dataset.b;s.blockers=s.blockers.includes(x)?s.blockers.filter(y=>y!==x):[...s.blockers,x];rerenderCard(id);writeItem(id)}
  else if(b.dataset.act==='openblk'){s.open=true;rerenderCard(id);writeItem(id);const t=$('#note-'+id);if(t)t.focus()}
  else if(b.dataset.act==='remove'){confirmRemove(id)}
});
document.addEventListener('input',e=>{const t=e.target;if(t.dataset&&t.dataset.act==='note'&&canWrite){st[t.dataset.id].note=t.value;writeItem(t.dataset.id,900)}});
document.addEventListener('change',e=>{const t=e.target;if(t.dataset&&t.dataset.act==='addtrade'&&t.value&&canWrite){st[t.dataset.id].extra.push(+t.value);rerenderCard(t.dataset.id);writeItem(t.dataset.id)}});
$('#btnNew').addEventListener('click',()=>{if(!canWrite){toast('View only: ask the night shift lead for Contributor access');return}$('#fileInput').click()});
$('#fileInput').addEventListener('change',async e=>{
  const f=e.target.files[0];e.target.value='';if(!f)return;
  toast('Sharing '+f.name+'…');
  try{await uploadDabs(f)}
  catch(err){openSheet(`<h3>Could not share this file</h3><p class="err">${esc(err.message||err.code||'Unknown error')}</p><button class="btn" type="button" id="cOk">OK</button>`);$('#cOk').onclick=closeSheet}
});

/* ---------- add activity ---------- */
function openAddForm(si){
  const levels=[...new Set(allItems().map(i=>i.level))].sort((a,b)=>parseFloat(a)-parseFloat(b));
  const acts=[...new Set(allItems().map(i=>i.act.replace(/\s+/g,' ').trim()))];
  const secOpts=doc.sections.map((s,i)=>`<option value="${i}"${i===si?' selected':''}>${esc(s.short)}</option>`).join('');
  openSheet(`<h3>Add activity</h3><p>For work that is not on tonight's DABS. The whole team sees it, and it is exported in its own block under the sheet.</p>
    <form id="addForm" novalidate style="display:flex;flex-direction:column;gap:10px">
      <div class="grid2">
        <label for="aSec">Section<select id="aSec">${secOpts}</select></label>
        <label for="aLvl">Level<select id="aLvl">${levels.map(l=>`<option>${esc(l)}</option>`).join('')}<option value="__other">Other…</option></select></label>
      </div>
      <label for="aLvlOther" id="aLvlOtherWrap" hidden>Level (m)<input id="aLvlOther" inputmode="decimal" placeholder="+8.50"></label>
      <label for="aTag"><span id="aTagLbl"></span><input id="aTag" autocomplete="off" placeholder="e.g. 1CEX2478TY--PS005"></label>
      <label for="aAct">Activity description<input id="aAct" list="actList" autocomplete="off" placeholder="e.g. Support erection"></label>
      <datalist id="actList">${acts.map(a=>`<option value="${esc(a)}"></option>`).join('')}</datalist>
      <fieldset style="border:0;padding:0;margin:0;display:flex;flex-direction:column;gap:6px"><legend class="lbl" style="font-size:12px;color:var(--muted);margin-bottom:4px">Trades</legend><div class="bchips" id="aTrades"></div></fieldset>
      <label for="aWhy">Why added (optional)<input id="aWhy" placeholder="e.g. Fallback, scaffold on GL04 not handed over"></label>
      <p class="err" id="aErr" hidden></p>
      <div class="grid2"><button class="btn alt" type="button" id="aCancel">Cancel</button><button class="btn" type="submit">Add</button></div>
    </form>`);
  const picked=new Set();
  const fillSec=()=>{const sec=doc.sections[+$('#aSec').value];$('#aTagLbl').textContent=sec.tagLabel;picked.clear();
    $('#aTrades').innerHTML=sec.trades.map(t=>`<button type="button" class="bchip pick" data-col="${t.col}" aria-pressed="false">${esc(t.name)}</button>`).join('')};
  fillSec();
  $('#aSec').onchange=fillSec;
  $('#aLvl').onchange=()=>{$('#aLvlOtherWrap').hidden=$('#aLvl').value!=='__other'};
  $('#aTrades').onclick=e=>{const b=e.target.closest('.pick');if(!b)return;const c=+b.dataset.col;
    if(picked.has(c))picked.delete(c);else picked.add(c);b.setAttribute('aria-pressed',String(picked.has(c)))};
  $('#aCancel').onclick=closeSheet;
  $('#addForm').onsubmit=e=>{e.preventDefault();
    const sec=+$('#aSec').value;const tag=$('#aTag').value.trim(),act=$('#aAct').value.trim();
    let lvl=$('#aLvl').value;if(lvl==='__other'){const v=$('#aLvlOther').value.trim().replace(',','.');const n=parseFloat(v);lvl=isNaN(n)?v:fmtLevel(n)}
    const err=$('#aErr');
    if(!act){err.textContent='Enter the activity description.';err.hidden=false;return}
    if(!lvl){err.textContent='Enter the level.';err.hidden=false;return}
    const trades=doc.sections[sec].trades.filter(t=>picked.has(t.col)).map(t=>({col:t.col,name:t.name}));
    const id='a'+Date.now().toString(36)+Math.random().toString(36).slice(2,6);
    added.push({id,sec,added:true,level:lvl,shift:'Night',bh:'',tag,act,dayRemark:'',trades,pre:{},created:Date.now(),createdByName:myName()});
    st[id]={status:{},blockers:[],note:$('#aWhy').value.trim()?'Added on night: '+$('#aWhy').value.trim():'',extra:[]};
    writeItem(id);closeSheet();filter.level=null;renderAll();
    const el=$('#card-'+id);if(el)el.scrollIntoView({block:'center'});toast('Activity added for the team');};
  setTimeout(()=>$('#aTag').focus(),50);
}
function confirmRemove(id){
  const it=findItem(id);if(!it)return;
  openSheet(`<h3>Remove this activity?</h3><p><strong>${esc(it.tag||it.act)}</strong> was added on the night. It is removed for the whole team, with its status and notes.</p>
    <div class="grid2"><button class="btn alt" type="button" id="cNo">Keep</button><button class="btn danger" type="button" id="cYes">Remove</button></div>`);
  $('#cNo').onclick=closeSheet;
  $('#cYes').onclick=async()=>{
    pending.delete(id);added=added.filter(a=>a.id!==id);delete st[id];closeSheet();renderAll();
    clearTimeout(timers[id]);dirty.delete(id);
    db.doc('nights/'+cur.id+'/items/'+id).delete().catch(e=>toast('Not removed ('+(e.code||'error')+')'));scheduleSummary();toast('Activity removed')};
}
/* ---------- sheets ---------- */
function openSheet(html){$('#panel').innerHTML=html;$('#sheet').hidden=false}
function closeSheet(){$('#sheet').hidden=true}
$('#sheet').addEventListener('click',e=>{if(e.target.id==='sheet')closeSheet()});

/* ---------- handover ---------- */
function remarkText(s){
  const parts=[];if(s.blockers.length)parts.push('Blocker: '+s.blockers.join(', '));if(s.note.trim())parts.push(s.note.trim());return parts.join(' - ');
}
function handoverText(){
  const c=counts();const L=[];
  const title=doc.sections.map(s=>s.title).join(' / ');
  L.push('NIGHT SHIFT HANDOVER - '+title);
  L.push('Prepared by: '+prepBy+(prepBy?' (Night Shift)':''));
  L.push('');
  L.push(`Summary: ${c.total} activities | ${c.DONE} DONE | ${c.ONGOING} ONGOING | ${c['NOT DONE']+c.unset} NOT DONE`);
  if(added.length)L.push(`Includes ${added.length} ${added.length===1?'activity':'activities'} added on night (not on DABS).`);
  const owners={};
  doc.sections.forEach((sec,si)=>{
    L.push('');L.push('== '+sec.title.toUpperCase()+' ==');
    for(const grp of ['DONE','ONGOING','NOT DONE']){
      const items=itemsOf(si).filter(it=>(rowStatus(it)||'NOT DONE')===grp);if(!items.length)continue;
      L.push('');L.push(grp+' ('+items.length+')');
      for(const it of items){
        const s=st[it.id];
        const detail=allCols(it).length>1&&grp!=='DONE'?' ['+allCols(it).map(col=>{const t=sec.trades.find(x=>x.col===col);return (t?t.name:colName(col))+': '+(s.status[col]||'NOT DONE')}).join(', ')+']':'';
        const rm=remarkText(s);
        const nd=grp!=='DONE'?nightsNotDone(historyOf(it.tag)):0;
        L.push(`- ${it.tag||'-'} | ${it.act} | Lvl ${it.level}${it.added?' | ADDED ON NIGHT':''}${nd?' | CARRIED OVER ('+nd+' earlier '+(nd===1?'night':'nights')+')':''}${detail}${rm?' | '+rm:''}`);
        if(grp!=='DONE')for(const b of s.blockers){const o=(BLOCKERS.find(x=>x[0]===b)||[])[1]||'Other';(owners[o]=owners[o]||new Set()).add(it.tag||it.act)}
      }
    }
  });
  const ok=Object.keys(owners);
  if(ok.length){L.push('');L.push('== SUPPORT NEEDED FROM DAY SHIFT ==');for(const o of ok)L.push(`- ${o}: ${[...owners[o]].join(', ')}`)}
  if(c.unset)L.push('','Note: '+c.unset+' activities were not updated on site and are reported as NOT DONE.');
  return L.join('\n');
}
$('#btnHandover').addEventListener('click',()=>{
  openSheet(`<h3>Handover</h3>
    <label for="prepBy">Prepared by<input id="prepBy" value="${esc(prepBy)}"></label>
    <textarea class="out" id="hoText" readonly></textarea>
    <div class="grid2"><button class="btn alt" type="button" id="hoShare">Share</button><button class="btn" type="button" id="hoCopy">Copy</button></div>
    <button class="ghost" type="button" id="hoClose">Close</button>`);
  const upd=()=>$('#hoText').value=handoverText();upd();
  $('#prepBy').oninput=e=>{prepBy=e.target.value;meta.prepBy=prepBy;saveLocal();upd()};
  $('#hoClose').onclick=closeSheet;
  $('#hoCopy').onclick=async()=>{const t=$('#hoText');try{await navigator.clipboard.writeText(t.value);toast('Handover copied')}catch(e){t.select();try{document.execCommand('copy');toast('Handover copied')}catch(_){toast('Select the text and copy')}}};
  $('#hoShare').onclick=async()=>{const text=$('#hoText').value;
    if(navigator.share){try{await navigator.share({title:'Night shift handover',text})}catch(e){}}
    else{location.href='mailto:?subject='+encodeURIComponent('Night shift handover - '+doc.sections[0].title)+'&body='+encodeURIComponent(text)}};
});

/* ---------- export ---------- */
$('#btnExport').addEventListener('click',()=>{
  const c=counts();
  if(c.unset){openSheet(`<h3>${c.unset} activities not updated</h3><p>They will be exported as NOT DONE. Go back if you want to mark them first.</p>
    <div class="grid2"><button class="btn alt" type="button" id="cNo">Go back</button><button class="btn" type="button" id="cYes">Export anyway</button></div>`);
    $('#cNo').onclick=closeSheet;$('#cYes').onclick=()=>{closeSheet();doExport()};}
  else doExport();
});
function insertOrdered(root,el){
  const ORDER=['sheetPr','dimension','sheetViews','sheetFormatPr','cols','sheetData','sheetCalcPr','sheetProtection','protectedRanges','scenarios','autoFilter','sortState','dataConsolidate','customSheetViews','mergeCells','phoneticPr','conditionalFormatting','dataValidations','hyperlinks','printOptions','pageMargins','pageSetup','headerFooter','rowBreaks','colBreaks','customProperties','cellWatches','ignoredErrors','smartTags','drawing','legacyDrawing','legacyDrawingHF','drawingHF','picture','oleObjects','controls','webPublishItems','tableParts','extLst'];
  const i=ORDER.indexOf(el.localName);
  const after=Array.from(root.children).find(c=>ORDER.indexOf(c.localName)>i);
  root.insertBefore(el,after||null);
}
function insertStyleOrdered(root,el){
  const ORDER=['numFmts','fonts','fills','borders','cellStyleXfs','cellXfs','cellStyles','dxfs','tableStyles','colors','extLst'];
  const i=ORDER.indexOf(el.localName);
  const after=Array.from(root.children).find(c=>ORDER.indexOf(c.localName)>i);
  root.insertBefore(el,after||null);
}
async function buildXlsx(){
  const zip=await JSZip.loadAsync(bytes);
  const {path}=await sheetPaths(zip);
  const sx=parseXml(await zip.file(path).async('string'));
  const sy=parseXml(await zip.file('xl/styles.xml').async('string'));
  const ss=await readShared(zip);
  const sroot=sy.documentElement,wroot=sx.documentElement;
  const E=(d,n,attrs)=>{const e=d.createElementNS(NS,n);for(const k in attrs||{})e.setAttribute(k,attrs[k]);return e};
  const getOrMake=(name)=>{let e=kid(sroot,name);if(!e){e=E(sy,name,{count:'0'});insertStyleOrdered(sroot,e)}return e};
  const fonts=getOrMake('fonts'),fills=getOrMake('fills'),xfs=getOrMake('cellXfs');
  const recount=e=>e.setAttribute('count',String(e.children.length));
  const COL={DONE:['FFC6EFCE','FF006100'],ONGOING:['FFC6EFCE','FF006100'],'NOT DONE':['FFFFC7CE','FF9C0006']};
  const fillIdx={};
  for(const k of ['FFC6EFCE','FFFFC7CE']){const f=E(sy,'fill'),p=E(sy,'patternFill',{patternType:'solid'});p.appendChild(E(sy,'fgColor',{rgb:k}));p.appendChild(E(sy,'bgColor',{indexed:'64'}));f.appendChild(p);fills.appendChild(f);fillIdx[k]=fills.children.length-1}
  recount(fills);
  const xfCache={};
  function styled(sIdx,status){
    const key=sIdx+'|'+status;if(key in xfCache)return xfCache[key];
    const base=xfs.children[sIdx]||xfs.children[0];const xf=base.cloneNode(true);
    const fBase=fonts.children[+(base.getAttribute('fontId')||0)]||fonts.children[0];
    const fnt=fBase?fBase.cloneNode(true):E(sy,'font');
    if(!kid(fnt,'b'))fnt.insertBefore(E(sy,'b'),fnt.firstChild);
    const oc=kid(fnt,'color');if(oc)fnt.removeChild(oc);
    const col=E(sy,'color',{rgb:COL[status][1]});
    const ref=Array.from(fnt.children).find(c=>['name','family','charset','scheme'].includes(c.localName));
    fnt.insertBefore(col,ref||null);fonts.appendChild(fnt);recount(fonts);
    xf.setAttribute('fontId',String(fonts.children.length-1));xf.setAttribute('applyFont','1');
    xf.setAttribute('fillId',String(fillIdx[COL[status][0]]));xf.setAttribute('applyFill','1');
    let al=kid(xf,'alignment');if(!al){al=E(sy,'alignment');xf.insertBefore(al,xf.firstChild)}
    al.setAttribute('horizontal','center');al.setAttribute('vertical','center');xf.setAttribute('applyAlignment','1');
    xfs.appendChild(xf);recount(xfs);return xfCache[key]=xfs.children.length-1;
  }
  function wrapStyle(sIdx){
    const key=sIdx+'|wrap';if(key in xfCache)return xfCache[key];
    const base=xfs.children[sIdx]||xfs.children[0];const xf=base.cloneNode(true);
    let al=kid(xf,'alignment');if(!al){al=E(sy,'alignment');xf.insertBefore(al,xf.firstChild)}
    al.setAttribute('wrapText','1');al.setAttribute('vertical','center');xf.setAttribute('applyAlignment','1');
    xfs.appendChild(xf);recount(xfs);return xfCache[key]=xfs.children.length-1;
  }
  // dxfs for conditional formatting
  const dxfs=getOrMake('dxfs');
  const mkDxf=(fill,font)=>{const d=E(sy,'dxf'),f=E(sy,'font');f.appendChild(E(sy,'b'));f.appendChild(E(sy,'color',{rgb:font}));d.appendChild(f);
    const fl=E(sy,'fill'),p=E(sy,'patternFill',{patternType:'solid'});p.appendChild(E(sy,'bgColor',{rgb:fill}));fl.appendChild(p);d.appendChild(fl);dxfs.appendChild(d);recount(dxfs);return dxfs.children.length-1};
  const dxGreen=mkDxf('FFC6EFCE','FF006100'),dxRed=mkDxf('FFFFC7CE','FF9C0006');

  const sheetData=kid(wroot,'sheetData');
  const origS=new Map();for(const ce of sheetData.getElementsByTagNameNS(NS,'c'))origS.set(ce.getAttribute('r'),+(ce.getAttribute('s')||0));
  const getS0=(r,c)=>origS.has(colName(c)+r)?origS.get(colName(c)+r):null;
  const rowEl=r=>{let found=null,after=null;for(const re of kids(sheetData,'row')){const n=+re.getAttribute('r');if(n===r){found=re;break}if(n>r){after=re;break}}
    if(!found){found=E(sx,'row',{r:String(r)});sheetData.insertBefore(found,after)}return found};
  const cellEl=(r,c)=>{const re=rowEl(r);re.removeAttribute('spans');let after=null;
    for(const ce of kids(re,'c')){const [cc]=splitRef(ce.getAttribute('r'));if(cc===c)return ce;if(cc>c){after=ce;break}}
    const ce=E(sx,'c',{r:colName(c)+r});re.insertBefore(ce,after);return ce};
  const getS=(r,c)=>{const re=rowEl(r);for(const ce of kids(re,'c'))if(ce.getAttribute('r')===colName(c)+r)return +(ce.getAttribute('s')||0);return null};
  const setText=(ce,text,s)=>{while(ce.firstChild)ce.removeChild(ce.firstChild);ce.removeAttribute('t');
    ce.setAttribute('t','inlineStr');const is=E(sx,'is'),t=E(sx,'t');t.setAttributeNS(XMLNS_XML,'xml:space','preserve');t.textContent=text;is.appendChild(t);ce.appendChild(is);
    if(s!=null)ce.setAttribute('s',String(s))};

  const ranges=[];
  for(const sec of doc.sections){
    if(!sec.items.length||!sec.trades.length)continue;
    const r1=sec.items[0].row,r2=sec.items[sec.items.length-1].row;
    const c1=Math.min(...sec.trades.map(t=>t.col)),c2=Math.max(...sec.trades.map(t=>t.col));
    ranges.push({ref:colName(c1)+r1+':'+colName(c2)+r2,tl:colName(c1)+r1});
    // remark header if it was missing
    {const h=cellValue(cellEl(sec.header,sec.cols.remark),ss);
      if(h==null||h===''){setText(cellEl(sec.header,sec.cols.remark),'Remark/Blocker',getS(sec.header,sec.cols.act))}}
    for(const it of sec.items){
      const s=st[it.id];
      for(const col of allCols(it)){
        const v=s.status[col]||'NOT DONE';
        const base=getS(it.row,col)??getS(it.row,sec.cols.act)??0;
        setText(cellEl(it.row,col),v,styled(base,v));
      }
      const add=remarkText(s);
      if(add){
        const text=it.dayRemark?it.dayRemark+' | '+add:add;
        const base=getS(it.row,sec.cols.remark)??getS(it.row,sec.cols.act)??0;
        setText(cellEl(it.row,sec.cols.remark),text,wrapStyle(base));
      }
    }
  }
  // activities added on the night: own block under the sheet, same columns as the section
  let maxRow=0;for(const re of kids(sheetData,'row'))maxRow=Math.max(maxRow,+re.getAttribute('r'));
  let maxCol=0;
  doc.sections.forEach((sec,si)=>{
    const extra=added.filter(a=>a.sec===si);if(!extra.length)return;
    const first=sec.items[0]?sec.items[0].row:sec.header;
    const hdrCols=sec.hdr.map(([c])=>c);if(!hdrCols.includes(sec.cols.remark))hdrCols.push(sec.cols.remark);
    const c1=Math.min(...hdrCols),c2=Math.max(...hdrCols);maxCol=Math.max(maxCol,c2);
    let r=maxRow+3;
    setText(cellEl(r,sec.titleCol),'NIGHT ADDED - '+sec.title,getS0(sec.titleRow,sec.titleCol));
    r+=2;
    for(let c=c1;c<=c2;c++){const h=sec.hdr.find(x=>x[0]===c);setText(cellEl(r,c),h?h[1]:(c===sec.cols.remark?'Remark/Blocker':''),getS0(sec.header,c))}
    const r1=r+1;
    for(const it of extra){
      r++;const s=st[it.id];
      for(let c=c1;c<=c2;c++){const base=getS0(first,c)??getS0(first,sec.cols.act)??0;
        let v='';
        if(c===sec.cols.level)v=it.level;else if(c===sec.cols.shift)v=it.shift;else if(c===sec.cols.tag)v=it.tag;
        else if(c===sec.cols.act)v=it.act;else if(c===sec.cols.bh)v='';
        else if(c===sec.cols.remark){setText(cellEl(r,c),remarkText(s),wrapStyle(base));continue}
        if(allCols(it).includes(c)){const sv=s.status[c]||'NOT DONE';setText(cellEl(r,c),sv,styled(base,sv));continue}
        setText(cellEl(r,c),v,base);
      }
    }
    if(sec.trades.length){const t1=Math.min(...sec.trades.map(t=>t.col)),t2=Math.max(...sec.trades.map(t=>t.col));
      ranges.push({ref:colName(t1)+r1+':'+colName(t2)+r,tl:colName(t1)+r1})}
    maxRow=r;
  });
  const dim=kid(wroot,'dimension');
  if(dim&&maxRow){const ref=dim.getAttribute('ref')||'A1';const [a,b]=ref.split(':');const end=b||a;const [ec,er]=splitRef(end);
    dim.setAttribute('ref',(b?a:'A1')+':'+colName(Math.max(ec,maxCol))+Math.max(er,maxRow))}
  // conditional formatting + dropdowns so the file keeps working in Excel
  let prio=0;for(const r of wroot.getElementsByTagNameNS(NS,'cfRule'))prio=Math.max(prio,+(r.getAttribute('priority')||0));
  for(const rg of ranges){
    const cf=E(sx,'conditionalFormatting',{sqref:rg.ref});
    const add=(dx,f)=>{const r=E(sx,'cfRule',{type:'expression',dxfId:String(dx),priority:String(++prio)});const fe=E(sx,'formula');fe.textContent=f;r.appendChild(fe);cf.appendChild(r)};
    add(dxGreen,`OR(UPPER(TRIM(${rg.tl}))="DONE",UPPER(TRIM(${rg.tl}))="ONGOING")`);
    add(dxRed,`UPPER(TRIM(${rg.tl}))="NOT DONE"`);
    insertOrdered(wroot,cf);
  }
  if(ranges.length){
    let dvs=kid(wroot,'dataValidations');if(!dvs){dvs=E(sx,'dataValidations',{count:'0'});insertOrdered(wroot,dvs)}
    const dv=E(sx,'dataValidation',{type:'list',allowBlank:'1',showErrorMessage:'1',errorTitle:'Status',error:'Use DONE, ONGOING or NOT DONE',sqref:ranges.map(r=>r.ref).join(' ')});
    const f1=E(sx,'formula1');f1.textContent='"DONE,ONGOING,NOT DONE"';dv.appendChild(f1);dvs.appendChild(dv);dvs.setAttribute('count',String(dvs.children.length));
  }
  zip.file(path,serXml(sx));zip.file('xl/styles.xml',serXml(sy));
  return zip.generateAsync({type:'uint8array',compression:'DEFLATE'});
}
function outName(){const base=doc.fileName.replace(/\.xlsx$/i,'').replace(/\s*-\s*Status$/i,'');return base+' - Status.xlsx'}
async function doExport(){
  let out;try{out=await buildXlsx()}catch(err){openSheet(`<h3>Export failed</h3><p class="err">${esc(err.message)}</p><button class="btn" type="button" id="cOk">OK</button>`);$('#cOk').onclick=closeSheet;return}
  window.__lastExport=out;
  deliverFile(outName(),out,'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet','Excel ready');
}
function deliverFile(name,u8,mime,heading){
  const blob=new Blob([u8],{type:mime});const file=new File([blob],name,{type:mime});
  const canShareFile=!!(navigator.canShare&&navigator.canShare({files:[file]}));
  openSheet(`<h3>${esc(heading)}</h3><p>${esc(name)}</p>
    <div class="grid2"><button class="btn alt" type="button" id="xShare"${canShareFile?'':' hidden'}>Share file</button><button class="btn" type="button" id="xSave">Save to phone</button></div>
    <button class="ghost" type="button" id="xClose">Close</button>`);
  $('#xClose').onclick=closeSheet;
  $('#xSave').onclick=()=>{const a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download=name;document.body.appendChild(a);a.click();
    setTimeout(()=>{URL.revokeObjectURL(a.href);a.remove()},4000);toast('Saved to Downloads')};
  $('#xShare').onclick=async()=>{try{await navigator.share({files:[file],title:name})}catch(e){}};
}

/* ---------- archive + account ---------- */
let archiveOpen=false,installPrompt=null;
window.addEventListener('beforeinstallprompt',e=>{e.preventDefault();installPrompt=e});
const standalone=()=>matchMedia('(display-mode: standalone)').matches||navigator.standalone===true;
function miniBar(sm){if(!sm||!sm.total)return '';const p=n=>(n/sm.total*100)+'%';
  return `<div class="bar sm"><span style="width:${p(sm.DONE)};background:var(--ok)"></span><span style="width:${p(sm.ONGOING)};background:var(--go)"></span><span style="width:${p(sm['NOT DONE'])};background:var(--no)"></span></div>`}
function openArchive(){
  archiveOpen=true;
  const list=byDate();
  const role=isAdmin?'Admin':canWrite?'Contributor':'Viewer';
  openSheet(`<h3>Archive</h3><p>${list.length} ${list.length===1?'night':'nights'} saved for the team.</p>
    <div class="nights">${list.map(n=>{const sm=n.summary;const isCur=cur&&n.id===cur.id;
      return `<div class="night${isCur?' cur':''}" id="nt-${n.id}">
        <div class="nl"><div class="nt"><strong>${esc(n.label||'—')}</strong>${isCur?'<span class="badge">On screen</span>':''}</div>
          <span class="nf">${esc(n.fileName)}</span>${miniBar(sm)}
          <span class="nc">${sm?`DONE ${sm.DONE} · ONGOING ${sm.ONGOING} · NOT DONE ${sm['NOT DONE']}${sm.unset?` · Not updated ${sm.unset}`:''}`:'Not started'}</span>
          ${n.createdByName?`<span class="nc">Shared by ${esc(n.createdByName)} · ${new Date(n.loadedAt).toLocaleString('en-GB',{day:'numeric',month:'short',hour:'2-digit',minute:'2-digit'})}</span>`:''}</div>
        <div class="nb">${isCur?'':`<button class="btn alt" type="button" data-arch="open" data-n="${n.id}">Open</button>`}
          ${isAdmin?`<button class="link danger" type="button" data-arch="del" data-n="${n.id}">Delete</button>`:''}</div></div>`}).join('')}</div>
    <div class="bk"><h4>Account</h4>
      <p>Signed in as <strong>${esc(myName())}</strong> (${esc(authUser&&authUser.email||'')}) · ${role}</p>
      <div class="grid2">${isAdmin?'<button class="btn alt" type="button" id="goTeam">Team & approvals</button>':''}<button class="btn alt" type="button" id="signOut">Sign out</button></div></div>
    ${standalone()?'':`<div class="bk"><h4>Install on your phone</h4>
      ${installPrompt?'<button class="btn" type="button" id="doInstall">Install app</button>':''}
      <p><strong>Android (Chrome):</strong> menu ⋮ → Install app.<br><strong>iPhone (Safari):</strong> Share button → Add to Home Screen.</p></div>`}
    ${canWrite?`<div class="bk"><h4>Bring in the offline file</h4><p>If you used the single-file version before, save a backup there and import it here. Nights already online are skipped.</p>
      <button class="btn alt" type="button" id="bkLoad">Import backup (.json)</button></div>`:''}
    <button class="ghost" type="button" id="arClose">Close</button>`);
  $('#arClose').onclick=()=>{archiveOpen=false;closeSheet()};
  const bl=$('#bkLoad');if(bl)bl.onclick=()=>$('#bkInput').click();
  const gt=$('#goTeam');if(gt)gt.onclick=()=>{archiveOpen=false;openTeam()};
  $('#signOut').onclick=async()=>{archiveOpen=false;closeSheet();await fbSignOut()};
  const di=$('#doInstall');if(di)di.onclick=async()=>{try{installPrompt.prompt();await installPrompt.userChoice}catch(e){}installPrompt=null;openArchive()};
  $('#panel').onclick=async e=>{const b=e.target.closest('[data-arch]');if(!b)return;const n=nights.find(x=>x.id===b.dataset.n);if(!n)return;
    if(b.dataset.arch==='open'){try{archiveOpen=false;closeSheet();await openNight(n);window.scrollTo(0,0);toast('Opened '+n.label)}catch(err){toast(err.message)}}
    else if(b.dataset.arch==='del'){
      const row=$('#nt-'+n.id);row.innerHTML=`<div class="nl"><strong>Delete ${esc(n.label)} for everyone?</strong><span class="nf">Its statuses, notes and added activities are removed for the whole team.</span></div>
        <div class="nb"><button class="btn alt" type="button" data-arch="keep" data-n="${n.id}">Keep</button><button class="btn danger" type="button" data-arch="really" data-n="${n.id}">Delete</button></div>`}
    else if(b.dataset.arch==='keep'){openArchive()}
    else if(b.dataset.arch==='really'){
      if(!navigator.onLine){toast('Deleting a night needs signal');return}
      b.disabled=true;b.textContent='Deleting…';
      try{const its=await db.collection('nights/'+n.id+'/items').get();
        for(const d of its.docs)await db.doc('nights/'+n.id+'/items/'+d.id).delete();
        await db.doc('files/'+n.id).delete();await db.doc('nights/'+n.id).delete();toast('Night deleted')}
      catch(err){toast('Delete failed ('+(err.code||'error')+')');openArchive()}}
  };
}
$('#sheet').addEventListener('click',e=>{if(e.target.id==='sheet')archiveOpen=false});
$('#btnArchive').addEventListener('click',openArchive);
$('#bkInput').addEventListener('change',async e=>{
  const f=e.target.files[0];e.target.value='';if(!f)return;
  let data;try{data=JSON.parse(await f.text())}catch(err){toast('This is not a DABS Night backup file');return}
  if(!data||data.app!=='DABS Night'||!Array.isArray(data.nights)){toast('This is not a DABS Night backup file');return}
  if(!navigator.onLine){toast('Importing needs signal');return}
  let addedN=0,skipped=0;toast('Importing…');
  try{
    for(const n of data.nights){if(!n||!n.b64)continue;
      const u8=b64dec(n.b64);const id='d'+hashStr(n.b64)+u8.length.toString(36);
      if(nights.some(x=>x.id===id)){skipped++;continue}
      const d=await parseDabs(u8,n.fileName);
      await db.doc('files/'+id).set({b64:n.b64,fileName:n.fileName});
      for(const [iid,s] of Object.entries(n.st||{})){
        const a=(n.added||[]).find(x=>x.id===iid);
        const body={status:s.status||{},blockers:s.blockers||[],note:s.note||'',extra:s.extra||[],open:!!s.open,by:authUser.uid,byName:myName(),at:n.updatedAt||Date.now()};
        if(a)body.added={sec:a.sec,level:a.level,tag:a.tag,act:a.act,trades:(a.trades||[]).map(t=>({col:t.col,name:t.name})),created:n.updatedAt||Date.now(),createdByName:myName()};
        await db.doc('nights/'+id+'/items/'+String(iid).replace(/[\/]/g,'_')).set(body);
      }
      await db.doc('nights/'+id).set({fileName:n.fileName,label:n.label||nightLabel(d,n.loadedAt||Date.now()),loadedAt:n.loadedAt||Date.now(),updatedAt:n.updatedAt||Date.now(),createdBy:authUser.uid,createdByName:myName(),summary:n.summary||null,items:n.items||[]});
      addedN++;
    }
    toast(`Imported ${addedN} ${addedN===1?'night':'nights'}`+(skipped?`, ${skipped} already online`:''));
  }catch(err){toast('Import stopped ('+(err.code||err.message||'error')+')')}
});

/* ---------- team approvals (admins) ---------- */
let unsubTeam=null;
const ROLES=[['viewer','Viewer'],['contributor','Contributor'],['admin','Admin']];
function openTeam(){
  openSheet(`<h3>Team & approvals</h3>
    <p>New people appear here after they create an account in the app. <strong>Contributors</strong> mark jobs and add activities, <strong>viewers</strong> only read, <strong>admins</strong> also approve people and delete nights.</p>
    <div id="teamList" class="nights"><p class="loading">Loading…</p></div>
    <button class="ghost" type="button" id="tmClose">Close</button>`);
  const stop=()=>{if(unsubTeam)unsubTeam();unsubTeam=null};
  $('#tmClose').onclick=()=>{stop();closeSheet()};
  stop();
  unsubTeam=db.collection('members').onSnapshot(snap=>{
    const el=$('#teamList');if(!el){stop();return}
    const order={pending:0,contributor:1,admin:1,viewer:1,removed:2};
    const ms=snap.docs.map(d=>Object.assign({uid:d.id},d.data())).sort((a,b)=>(order[a.role]??1)-(order[b.role]??1)||String(a.name||a.email).localeCompare(String(b.name||b.email)));
    el.innerHTML=ms.length?ms.map(m=>`<div class="night" id="tm-${m.uid}"><div class="nl"><div class="nt"><strong>${esc(m.name||m.email)}</strong>${m.role==='pending'?'<span class="badge">Waiting</span>':''}${m.role==='removed'?'<span class="nc">Removed</span>':''}</div><span class="nf">${esc(m.email)}</span></div>
      <div class="nb">${m.role==='pending'?`<button class="btn" type="button" data-tm="approve" data-u="${m.uid}">Approve</button><button class="link danger" type="button" data-tm="remove" data-u="${m.uid}">Reject</button>`
        :m.role==='removed'?`<button class="btn alt" type="button" data-tm="restore" data-u="${m.uid}">Restore</button>`
        :`<select data-tm="role" data-u="${m.uid}" aria-label="Role for ${esc(m.name||m.email)}">${ROLES.map(([v,l])=>`<option value="${v}"${m.role===v?' selected':''}>${l}</option>`).join('')}</select><button class="link danger" type="button" data-tm="remove" data-u="${m.uid}">Remove</button>`}</div></div>`).join('')
      :'<p>No one has asked to join yet. Send your team the app link; when they create an account they appear here.</p>';
  },e=>{const el=$('#teamList');if(el)el.innerHTML=`<p class="err">Could not load the team (${esc(e.code||'error')}).</p>`});
  $('#panel').onclick=async e=>{const b=e.target.closest('button[data-tm]');if(!b)return;const u=b.dataset.u;
    const role={approve:'contributor',remove:'removed',restore:'viewer'}[b.dataset.tm];
    try{await db.doc('members/'+u).update({role});toast({approve:'Approved as Contributor',remove:'Access removed',restore:'Restored as Viewer'}[b.dataset.tm])}
    catch(err){toast('Not changed ('+(err.code||'error')+')')}};
  $('#panel').onchange=async e=>{const s=e.target.closest('select[data-tm="role"]');if(!s)return;
    try{await db.doc('members/'+s.dataset.u).update({role:s.value});toast('Role updated')}catch(err){toast('Not changed ('+(err.code||'error')+')')}};
}

/* ---------- sign in ---------- */
let fsDb=null,auth=null,unsubMember=null;
const AUTH_ERR={'auth/invalid-credential':'Email or password is wrong.','auth/wrong-password':'Email or password is wrong.','auth/user-not-found':'No account with this email. Create one first.',
  'auth/email-already-in-use':'An account with this email already exists. Sign in instead.','auth/weak-password':'Use a password of at least 6 characters.',
  'auth/invalid-email':'Check the email address.','auth/network-request-failed':'No signal. Signing in needs a connection the first time.',
  'auth/popup-closed-by-user':'Google sign-in was closed.','auth/popup-blocked':'The Google window was blocked. Use email and password instead.',
  'auth/too-many-requests':'Too many tries. Wait a few minutes and try again.','auth/operation-not-supported-in-this-environment':'Google sign-in does not work here. Use email and password.'};
const authMsg=e=>AUTH_ERR[e&&e.code]||('Something went wrong ('+((e&&e.code)||'error')+').');
function screen(html){doc=null;cur=null;renderTop();$('#banner').innerHTML='';$('#main').innerHTML=html}
function showAuth(mode){
  const create=mode==='create';
  screen(`<section class="empty auth"><h2>${create?'Create your account':'Sign in'}</h2>
    <p>${create?'After you create it, the night shift lead approves you. Then you can mark jobs from your phone.':'Night shift DABS for the Turbine Building piping team.'}</p>
    <form id="authForm" class="authform" novalidate>
      ${create?'<label for="auName">Your name<input id="auName" autocomplete="name" placeholder="e.g. Lee Shepherd"></label>':''}
      <label for="auEmail">Email<input id="auEmail" type="email" autocomplete="email" inputmode="email"></label>
      <label for="auPass">Password<input id="auPass" type="password" autocomplete="${create?'new-password':'current-password'}" minlength="6"></label>
      <p class="err" id="auErr" hidden></p>
      <button class="btn" type="submit">${create?'Create account':'Sign in'}</button>
      <button class="link" type="button" id="auSwitch">${create?'I already have an account':'New here? Create an account'}</button>
      ${create?'':'<button class="link" type="button" id="auForgot">Forgot password</button>'}
      <div class="or"><span>or</span></div>
      <button class="btn alt" type="button" id="auGoogle">Continue with Google</button>
    </form></section>`);
  const err=m=>{const p=$('#auErr');p.textContent=m;p.hidden=false};
  $('#auSwitch').onclick=()=>showAuth(create?'signin':'create');
  const fg=$('#auForgot');if(fg)fg.onclick=async()=>{const em=$('#auEmail').value.trim();if(!em){err('Enter your email first, then tap Forgot password.');return}
    try{await sendPasswordResetEmail(auth,em);toast('Password reset email sent to '+em)}catch(e){err(authMsg(e))}};
  $('#auGoogle').onclick=async()=>{try{await signInWithPopup(auth,new GoogleAuthProvider())}catch(e){err(authMsg(e))}};
  $('#authForm').onsubmit=async e=>{e.preventDefault();
    const em=$('#auEmail').value.trim(),pw=$('#auPass').value;const nm=create?$('#auName').value.trim():'';
    if(create&&!nm){err('Enter your name so the team knows who updated a job.');return}
    if(!em||!pw){err('Enter your email and password.');return}
    const btn=e.submitter||$('#authForm .btn');btn.disabled=true;
    try{
      if(create){const c=await createUserWithEmailAndPassword(auth,em,pw);await updateProfile(c.user,{displayName:nm});
        try{await sendEmailVerification(c.user)}catch(_){}
        authUser=c.user;startMember(c.user)}
      else await signInWithEmailAndPassword(auth,em,pw);
    }catch(x){err(authMsg(x));btn.disabled=false}};
}
function showWaiting(){
  screen(`<section class="empty"><h2>Waiting for approval</h2>
    <p>Your account <strong>${esc(authUser&&authUser.email||'')}</strong> has been sent to the night shift lead. The DABS opens here as soon as you are approved; you can keep this screen open.</p>
    <button class="btn alt" type="button" id="wOut">Sign out</button></section>`);
  $('#wOut').onclick=fbSignOut;
}
function showVerify(u){
  screen(`<section class="empty"><h2>Verify your email</h2>
    <p>This account manages the team, so the email must be confirmed first. Open the link sent to <strong>${esc(u.email)}</strong>, then come back.</p>
    <div class="grid2"><button class="btn alt" type="button" id="vSend">Send the email again</button><button class="btn" type="button" id="vDone">I've verified</button></div>
    <button class="link" type="button" id="vOut">Sign out</button></section>`);
  $('#vSend').onclick=async()=>{try{await sendEmailVerification(u);toast('Email sent')}catch(e){toast(authMsg(e))}};
  $('#vDone').onclick=async()=>{try{await u.reload();await u.getIdToken(true)}catch(e){}if(auth.currentUser&&auth.currentUser.emailVerified)startMember(auth.currentUser);else toast('Not verified yet. Open the link in the email first.')};
  $('#vOut').onclick=fbSignOut;
}
async function fbSignOut(){if(unsubMember){unsubMember();unsubMember=null}if(unsubTeam){unsubTeam();unsubTeam=null}stopListening();try{await signOut(auth)}catch(e){}}
function startMember(u){
  if(unsubMember){unsubMember();unsubMember=null}
  const owner=[].concat(OWNER_EMAIL||[]).map(x=>String(x).toLowerCase()).includes((u.email||'').toLowerCase());
  if(owner&&!u.emailVerified){showVerify(u);return}
  const r=fdoc(fsDb,'members/'+u.uid);
  unsubMember=fsnap(r,async s=>{
    let role=owner?'admin':(s.exists()?s.data().role:null);
    if(!s.exists()&&!owner){role='pending';
      setDoc(r,{email:u.email||'',name:myName(),role:'pending',requestedAt:Date.now()}).catch(()=>{})}
    else if(s.exists()&&!owner&&myName()&&s.data().name!==myName())updateDoc(r,{name:myName()}).catch(()=>{});
    if(!role||role==='pending'||role==='removed'){canWrite=false;isAdmin=false;stopListening();showWaiting();return}
    canWrite=role==='contributor'||role==='admin';isAdmin=role==='admin';
    document.documentElement.classList.toggle('ro',!canWrite);
    if(!prepBy)prepBy=myName();
    updateSync();
    if(!unsubNights){setSync('load');renderAll();subscribeNights()}else{renderAll();updateSync()}
  },()=>{showWaiting()});
}

/* ---------- boot ---------- */
(()=>{
  loadLocal();prepBy=meta.prepBy||'';renderAll();
  if('serviceWorker' in navigator)navigator.serviceWorker.register('./sw.js').catch(()=>{});
  if(!firebaseConfig||!firebaseConfig.apiKey||/PASTE/.test(firebaseConfig.apiKey)){
    screen(`<section class="empty"><h2>Almost ready</h2><p>Paste your Firebase settings into <strong>firebase-config.js</strong> and upload it again. The setup guide (KURULUM.md) shows where to find them.</p></section>`);return}
  const app=initializeApp(firebaseConfig);
  try{fsDb=initializeFirestore(app,{localCache:persistentLocalCache({tabManager:persistentMultipleTabManager()})})}
  catch(e){fsDb=initializeFirestore(app,{})}
  auth=getAuth(app);
  const wrap=s=>({id:s.id,exists:s.exists(),data:()=>s.data(),metadata:s.metadata});
  const mkDoc=p=>{const r=fdoc(fsDb,p);return {id:r.id,path:p,get:async()=>wrap(await getDoc(r)),set:d=>setDoc(r,d),update:d=>updateDoc(r,d),delete:()=>deleteDoc(r),
    onSnapshot:(n,e,m)=>fsnap(r,{includeMetadataChanges:!!m},s=>n(wrap(s)),e)}};
  const mkQ=(p,cons)=>{const q=()=>cons.length?fquery(fcol(fsDb,p),...cons):fcol(fsDb,p);return {
    orderBy:(f,d)=>mkQ(p,[...cons,forder(f,d||'asc')]),limit:n=>mkQ(p,[...cons,flimit(n)]),
    get:async()=>{const s=await getDocs(q());return {docs:s.docs.map(wrap),size:s.size,empty:s.empty}},
    onSnapshot:(n,e,m)=>fsnap(q(),{includeMetadataChanges:!!m},s=>n({docs:s.docs.map(wrap),size:s.size,empty:s.empty,metadata:s.metadata,
      docChanges:()=>s.docChanges().map(c=>({type:c.type,doc:wrap(c.doc)}))}),e)}};
  db={doc:mkDoc,collection:p=>mkQ(p,[])};
  setSync('load');
  onAuthStateChanged(auth,u=>{
    authUser=u;
    if(!u){if(unsubMember){unsubMember();unsubMember=null}stopListening();canWrite=false;isAdmin=false;showAuth('signin');return}
    startMember(u);
  });
})();
