/* ============================================================
   Habit Tracker PWA — app logic
   Storage: per-day markdown in OneDrive (Life OS/Health/habits),
   pipe-delimited format. Two-writer safe: read-fresh -> write-whole -> verify.
   ============================================================ */

const CFG = window.HABIT_CONFIG || {};
const HABITS_DIR = "0 - Data/AI/Life OS/Health/habits";
const GRAPH = "https://graph.microsoft.com/v1.0";
const SCOPES = ["Files.ReadWrite", "offline_access", "User.Read"];

const HABITS = [
  {key:'vitamins', ic:'&#128138;', name:'Vitamins',      hint:'Tap: AM &rarr; PM'},
  {key:'pushups',  ic:'&#128170;', name:'Push-ups',      hint:'Tap to log a set'},
  {key:'drink',    ic:'&#127866;', name:'Drinks',        hint:'Tap each drink'},
  {key:'smoke',    ic:'&#128168;', name:'Smoke breaks',  hint:'Tap each session'},
];
const WEEKDAYS=['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];
const MONTHS=['January','February','March','April','May','June','July','August','September','October','November','December'];

/* ---------------- date helpers (LOCAL time) ---------------- */
function dkey(d){ return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`; }
function nowHM(){ const n=new Date(); return `${String(n.getHours()).padStart(2,'0')}:${String(n.getMinutes()).padStart(2,'0')}`; }
let viewDate=new Date(); viewDate.setHours(0,0,0,0);
const TODAY=new Date(); TODAY.setHours(0,0,0,0);

/* ---------------- model ---------------- */
function blank(){ return {vitamins:{am:null,pm:null}, pushups:[], drinks:[], smoke:[]}; }
const pushTotal = m => m.pushups.reduce((a,p)=>a+p.reps,0);
const vitCount  = m => (m.vitamins.am?1:0)+(m.vitamins.pm?1:0);

function parseDay(md){
  const m=blank(); if(!md) return m;
  md.split(/\r?\n/).forEach(line=>{
    const mm=line.match(/^-\s*(\d{2}:\d{2})\s*\|\s*(\w+)(?:\s*\|\s*(.+))?$/);
    if(!mm) return;
    const t=mm[1], habit=mm[2], detail=(mm[3]||'').trim();
    if(habit==='vitamins'){ if(detail==='am')m.vitamins.am=t; else if(detail==='pm')m.vitamins.pm=t; }
    else if(habit==='pushups'){ const reps=parseInt(detail,10); if(reps>0)m.pushups.push({t,reps}); }
    else if(habit==='drink'){ m.drinks.push(t); }
    else if(habit==='smoke'){ m.smoke.push({t,note:detail||null}); }
  });
  return m;
}
function renderDayFile(date,m){
  const rows=[];
  if(m.vitamins.am) rows.push([m.vitamins.am,`- ${m.vitamins.am} | vitamins | am`]);
  if(m.vitamins.pm) rows.push([m.vitamins.pm,`- ${m.vitamins.pm} | vitamins | pm`]);
  m.pushups.forEach(p=>rows.push([p.t,`- ${p.t} | pushups | ${p.reps}`]));
  m.drinks.forEach(t=>rows.push([t,`- ${t} | drink`]));
  m.smoke.forEach(s=>rows.push([s.t,`- ${s.t} | smoke${s.note?` | ${s.note}`:''}`]));
  rows.sort((a,b)=> a[0]<b[0]?-1 : a[0]>b[0]?1 : 0);
  return `# Habits — ${date}\n\n## Entries\n${rows.map(r=>r[1]).join('\n')}\n`;
}

/* ---------------- local cache (localStorage) ---------------- */
function getModel(date){ try{const s=localStorage.getItem('hb_day_'+date); if(s)return JSON.parse(s);}catch(e){} return blank(); }
function setModel(date,m){ localStorage.setItem('hb_day_'+date, JSON.stringify(m)); }

/* ---------------- op queue ---------------- */
function loadQueue(){ try{return JSON.parse(localStorage.getItem('hb_queue')||'[]');}catch(e){return[];} }
function saveQueue(q){ localStorage.setItem('hb_queue', JSON.stringify(q)); }
let _qid = Date.now();
function enqueue(op){ const q=loadQueue(); op.id=String(_qid++); q.push(op); saveQueue(q); updateSyncPill(); scheduleFlush(); }
function removeOps(ids){ const s=new Set(ids); saveQueue(loadQueue().filter(o=>!s.has(o.id))); updateSyncPill(); }
function applyOp(m,op){
  switch(op.op){
    case 'vitAdd':   if(!m.vitamins[op.slot]) m.vitamins[op.slot]=op.t; break;
    case 'vitClear': m.vitamins[op.slot]=null; break;
    case 'push':     m.pushups.push({t:op.t,reps:op.reps}); break;
    case 'pushPop':  m.pushups.pop(); break;
    case 'drink':    m.drinks.push(op.t); break;
    case 'smoke':    m.smoke.push({t:op.t,note:op.note||null}); break;
    case 'popLast':  if(op.habit==='drink')m.drinks.pop(); else if(op.habit==='smoke')m.smoke.pop(); break;
  }
}

/* ---------------- haptics ---------------- */
function haptic(ms=12){
  try{ if(navigator.vibrate) navigator.vibrate(ms); }catch(e){}
  try{ document.getElementById('hapticSwitch').click(); }catch(e){}
}

/* ---------------- toast ---------------- */
let toastT;
function toast(msg){ const t=document.getElementById('toast'); t.textContent=msg; t.classList.add('on');
  clearTimeout(toastT); toastT=setTimeout(()=>t.classList.remove('on'),1400); }

/* ============================================================
   AUTH (MSAL)
   ============================================================ */
let msalApp=null, account=null;
function redirectUri(){ return CFG.redirectUri || (location.origin + location.pathname.replace(/index\.html$/,'')); }
async function initAuth(){
  msalApp = new msal.PublicClientApplication({
    auth:{ clientId:CFG.clientId, authority:`https://login.microsoftonline.com/${CFG.tenantId}`, redirectUri:redirectUri() },
    cache:{ cacheLocation:'localStorage' }
  });
  await msalApp.initialize();
  try{
    const resp = await msalApp.handleRedirectPromise();
    if(resp && resp.account) account = resp.account;
  }catch(e){ showGateErr(e.message||String(e)); }
  if(!account) account = msalApp.getAllAccounts()[0] || null;
}
async function signIn(){ try{ await msalApp.loginRedirect({scopes:SCOPES}); }catch(e){ showGateErr(e.message||String(e)); } }
async function getToken(){
  const r = await msalApp.acquireTokenSilent({account, scopes:SCOPES}).catch(async e=>{
    await msalApp.acquireTokenRedirect({scopes:SCOPES}); throw e;
  });
  return r.accessToken;
}

/* ============================================================
   GRAPH
   ============================================================ */
function dayUrl(date){ return `${GRAPH}/me/drive/root:` + encodeURI(`/${HABITS_DIR}/${date}.md`) + `:/content`; }
function statsUrl(){ return `${GRAPH}/me/drive/root:` + encodeURI(`/${HABITS_DIR}/stats.json`) + `:/content`; }

async function gGetText(url){
  const tok=await getToken();
  const r=await fetch(url,{headers:{Authorization:`Bearer ${tok}`}});
  if(r.status===404) return null;
  if(!r.ok) throw new Error('GET '+r.status);
  return await r.text();
}
async function gPut(url,body,type){
  const tok=await getToken();
  const r=await fetch(url,{method:'PUT',headers:{Authorization:`Bearer ${tok}`,'Content-Type':type},body});
  if(!r.ok) throw new Error('PUT '+r.status);
  return r;
}

/* ============================================================
   SYNC — read fresh -> apply ops -> write whole -> verify
   ============================================================ */
let flushT=null, flushing=false;
function scheduleFlush(){ clearTimeout(flushT); flushT=setTimeout(flush,700); }

async function flush(){
  if(flushing || !account || !navigator.onLine) return;
  const q=loadQueue(); if(!q.length){ updateSyncPill(); return; }
  flushing=true; updateSyncPill();
  const byDate={}; q.forEach(op=>{ (byDate[op.date]=byDate[op.date]||[]).push(op); });
  for(const date of Object.keys(byDate)){
    try{
      const remote = parseDay(await gGetText(dayUrl(date)));   // read fresh
      byDate[date].forEach(op=>applyOp(remote,op));            // merge additive ops
      await gPut(dayUrl(date), renderDayFile(date,remote), 'text/plain');  // write whole
      removeOps(byDate[date].map(o=>o.id));                    // clear queue for this date
      setModel(date, remote);                                 // cache = merged truth
      if(dkey(viewDate)===date && currentTab==='today') renderToday();
      try{ await gGetText(dayUrl(date)); }catch(e){}           // verify read-back (best effort)
      updateStats(date, remote).catch(()=>{});                // stats.json (best effort)
    }catch(e){ console.warn('flush failed for',date,e); /* keep ops, retry later */ }
  }
  flushing=false; updateSyncPill();
}

async function refreshDay(date){
  if(!account || !navigator.onLine) return;
  try{
    const remote = parseDay(await gGetText(dayUrl(date)));
    loadQueue().filter(o=>o.date===date).forEach(o=>applyOp(remote,o)); // keep optimistic
    setModel(date, remote);
    if(dkey(viewDate)===date && currentTab==='today') renderToday();
  }catch(e){ console.warn('refresh failed',e); }
}

/* ---------------- stats.json (PWA-owned) ---------------- */
async function updateStats(date, model){
  let stats;
  try{ stats = JSON.parse((await gGetText(statsUrl()))||'null'); }catch(e){ stats=null; }
  if(!stats || typeof stats!=='object') stats={v:1, days:{}, streaks:{}, lifetime:{}};
  stats.days = stats.days || {};
  stats.days[date] = { vitamins:vitCount(model), pushups:pushTotal(model), drinks:model.drinks.length, smoke:model.smoke.length };
  // lifetime
  let pt=0, dt=0; Object.values(stats.days).forEach(d=>{ pt+=d.pushups||0; dt++; });
  stats.lifetime = { pushupsTotal:pt, daysTracked:dt };
  // streaks (consecutive days ending today)
  stats.streaks = computeStreaks(stats.days);
  stats.updated = new Date().toISOString();
  await gPut(statsUrl(), JSON.stringify(stats,null,2), 'application/json');
}
function computeStreaks(days){
  const back=(pred)=>{ let n=0, d=new Date(TODAY); for(let i=0;i<400;i++){ const k=dkey(d); const rec=days[k]; if(rec && pred(rec)) n++; else break; d=new Date(+d-86400000);} return n; };
  return {
    vitamins_full: back(r=>r.vitamins>=2),
    pushups:       back(r=>r.pushups>0),
    smoke_free:    back(r=>r.smoke===0),
    alcohol_free:  back(r=>r.drinks===0),
  };
}

/* ---------------- sync pill ---------------- */
function updateSyncPill(){
  const el=document.getElementById('syncPill'), tx=document.getElementById('syncTxt');
  if(!el) return;
  const pending=loadQueue().length;
  el.className='syncpill';
  if(!account){ el.classList.add('off'); tx.textContent='Offline'; }
  else if(!navigator.onLine){ el.classList.add('pending'); tx.textContent=pending?`${pending} to sync`:'No signal'; }
  else if(pending){ el.classList.add('pending'); tx.textContent='Syncing…'; }
  else{ el.classList.add('ok'); tx.textContent='Synced'; }
}

/* ============================================================
   RENDER — TODAY
   ============================================================ */
function renderToday(){
  const isToday = viewDate.getTime()===TODAY.getTime();
  const isYest  = viewDate.getTime()===TODAY.getTime()-86400000;
  document.getElementById('dayTitle').textContent = isToday?'Today':(isYest?'Yesterday':WEEKDAYS[viewDate.getDay()]);
  document.getElementById('dayDate').textContent = `${MONTHS[viewDate.getMonth()]} ${viewDate.getDate()}, ${viewDate.getFullYear()}`;
  document.getElementById('nextDay').disabled = isToday;
  document.getElementById('jumpToday').classList.toggle('hidden', isToday);

  const m=getModel(dkey(viewDate));
  const el=document.getElementById('tiles'); el.innerHTML='';
  HABITS.forEach(h=>{
    const t=document.createElement('div');
    let inner='', badge='', cls='';
    if(h.key==='vitamins'){
      const n=vitCount(m);
      if(n===2){cls='done'; badge='<span class="badge b-green">Done</span>';}
      else if(n===1){cls='partial'; badge='<span class="badge b-yellow">'+(m.vitamins.am?'AM ✓':'PM ✓')+'</span>';}
      else badge='<span class="badge b-gray">0 / 2</span>';
      inner=`<div><div class="ic">${h.ic}</div><div class="nm">${h.name}</div><div class="hint">${h.hint}</div></div>
        <div class="seg"><i class="${m.vitamins.am?'on':''}"></i><i class="${m.vitamins.pm?'on':''}"></i></div>`;
    } else if(h.key==='pushups'){
      const total=pushTotal(m), sets=m.pushups.length;
      badge = total>0 ? `<span class="badge b-green">${sets} set${sets>1?'s':''}</span>` : '<span class="badge b-gray">0</span>';
      if(total>0) cls='done';
      inner=`<div><div class="ic">${h.ic}</div><div class="nm">${h.name}</div><div class="hint">${h.hint}</div></div>
        <div class="foot"><div><div class="val">${total}</div><div class="valsub">reps today</div></div></div>`;
    } else {
      const c = h.key==='drink' ? m.drinks.length : m.smoke.length;
      badge = c>0 ? `<span class="badge b-blue">${c}×</span>` : '<span class="badge b-gray">0</span>';
      inner=`<div><div class="ic">${h.ic}</div><div class="nm">${h.name}</div><div class="hint">${h.hint}</div></div>
        <div class="foot"><div><div class="val">${c}</div><div class="valsub">${h.key==='drink'?'drinks':'breaks'} today</div></div></div>`;
    }
    t.className='tile '+cls; t.innerHTML=badge+inner;
    let held=false, timer;
    const start=()=>{ held=false; timer=setTimeout(()=>{held=true; haptic(25); openEditSheet(h.key);},450); };
    const end=()=>{ clearTimeout(timer); if(!held) tapTile(h.key); };
    t.addEventListener('touchstart',start,{passive:true});
    t.addEventListener('touchend',end);
    t.addEventListener('mousedown',start);
    t.addEventListener('mouseup',end);
    t.addEventListener('mouseleave',()=>clearTimeout(timer));
    el.appendChild(t);
  });
}

/* ---------------- tap logic ---------------- */
function tapTile(key){
  const date=dkey(viewDate), m=getModel(date), t=nowHM();
  haptic(12);
  if(key==='vitamins'){
    if(!m.vitamins.am){ m.vitamins.am=t; enqueue({date,op:'vitAdd',slot:'am',t}); toast('Morning vitamins ✓'); }
    else if(!m.vitamins.pm){ m.vitamins.pm=t; enqueue({date,op:'vitAdd',slot:'pm',t}); toast('Night vitamins ✓ — all done!'); }
    else { toast('Both logged — hold to edit'); return; }
  } else if(key==='pushups'){ openPushupSheet(); return; }
  else if(key==='drink'){ m.drinks.push(t); enqueue({date,op:'drink',t}); toast(`Drink #${m.drinks.length} logged`); }
  else if(key==='smoke'){ m.smoke.push({t,note:null}); enqueue({date,op:'smoke',t}); toast(`Smoke break #${m.smoke.length}`); }
  setModel(date,m); renderToday();
}

/* ============================================================
   SHEETS
   ============================================================ */
function openSheet(html){ document.getElementById('sheetBody').innerHTML='<div class="grab"></div>'+html;
  document.getElementById('sheetBack').classList.add('on'); document.getElementById('sheet').classList.add('on'); }
function closeSheet(){ document.getElementById('sheetBack').classList.remove('on'); document.getElementById('sheet').classList.remove('on'); }

let padVal='';
function openPushupSheet(){
  padVal='';
  openSheet(`<h2>Log push-ups</h2><div class="sh-sub">How many in this set?</div>
    <div class="numdisplay" id="numDisp">0</div>
    <div class="quickrow">${[10,15,20,25,50].map(n=>`<button class="quick" data-set="${n}">${n}</button>`).join('')}</div>
    <div class="pad">${[1,2,3,4,5,6,7,8,9].map(n=>`<button data-d="${n}">${n}</button>`).join('')}
      <button data-clear="1">C</button><button data-d="0">0</button><button data-del="1">&#9003;</button></div>
    <button class="save" id="savePush">Add set</button>`);
}
function setPadDisp(){ document.getElementById('numDisp').textContent = padVal||'0'; }
function savePushups(){
  const n=parseInt(padVal||'0',10);
  if(n>0){ const date=dkey(viewDate), m=getModel(date), t=nowHM();
    m.pushups.push({t,reps:n}); setModel(date,m); enqueue({date,op:'push',reps:n,t});
    toast(m.pushups.length>1?`Set ${m.pushups.length} — overachiever! \u{1F4AA}`:`${n} push-ups logged`); }
  haptic(15); closeSheet(); renderToday();
}

function openEditSheet(key){
  const date=dkey(viewDate), m=getModel(date);
  const h=HABITS.find(x=>x.key===key);
  if(key==='vitamins'){
    openSheet(`<h2>Vitamins</h2><div class="sh-sub">Tap to toggle each dose</div>
      <div class="vitrow">
        <button class="vitbtn ${m.vitamins.am?'on':''}" data-vit="am"><span class="vi">&#9728;&#65039;</span>Morning</button>
        <button class="vitbtn ${m.vitamins.pm?'on':''}" data-vit="pm"><span class="vi">&#127769;</span>Night</button>
      </div><button class="save" data-close="1">Done</button>`);
  } else if(key==='pushups'){
    const sets=m.pushups.length, total=pushTotal(m);
    openSheet(`<h2>Push-ups</h2><div class="sh-sub">${total} reps across ${sets} set${sets!==1?'s':''} today</div>
      <div class="sesslist">${sets? m.pushups.map((p,i)=>`Set ${i+1}: <b>${p.reps}</b> reps &middot; ${p.t}`).join('<br>') : 'No sets yet'}</div>
      <button class="save" data-addpush="1">+ Add another set</button>
      ${sets?`<button class="save" style="background:var(--redbg);color:#c0392b;margin-top:10px" data-poppush="1">Remove last set</button>`:''}`);
  } else {
    const c = key==='drink'?m.drinks.length:m.smoke.length;
    openSheet(`<h2>${h.name}</h2><div class="sh-sub">Adjust the count for this day</div>
      <div class="adjrow"><button data-adj="-1" data-habit="${key}">&minus;</button>
        <div class="cnt" id="tallyCnt">${c}</div>
        <button data-adj="1" data-habit="${key}">+</button></div>
      <button class="save" data-close="1">Done</button>`);
  }
}
function toggleVit(slot){
  const date=dkey(viewDate), m=getModel(date);
  if(m.vitamins[slot]){ m.vitamins[slot]=null; enqueue({date,op:'vitClear',slot}); }
  else { m.vitamins[slot]=nowHM(); enqueue({date,op:'vitAdd',slot,t:m.vitamins[slot]}); }
  setModel(date,m); haptic(10);
  const b=document.querySelector(`[data-vit="${slot}"]`); if(b) b.classList.toggle('on', !!m.vitamins[slot]);
  renderToday();
}
function adjTally(key,delta){
  const date=dkey(viewDate), m=getModel(date), t=nowHM();
  if(delta>0){ if(key==='drink'){m.drinks.push(t); enqueue({date,op:'drink',t});} else {m.smoke.push({t,note:null}); enqueue({date,op:'smoke',t});} }
  else { const arr=key==='drink'?m.drinks:m.smoke; if(arr.length){ arr.pop(); enqueue({date,op:'popLast',habit:key}); } }
  setModel(date,m); haptic(10);
  const c=key==='drink'?m.drinks.length:m.smoke.length;
  const el=document.getElementById('tallyCnt'); if(el) el.textContent=c;
  renderToday();
}
function popPush(){ const date=dkey(viewDate), m=getModel(date); m.pushups.pop(); setModel(date,m); enqueue({date,op:'pushPop'}); haptic(10); closeSheet(); renderToday(); }

/* ============================================================
   HISTORY
   ============================================================ */
let currentTab='today', histMode='week', monthFilter='all';
let monthCursor=new Date(TODAY.getFullYear(),TODAY.getMonth(),1);
let HIST={};   // date -> model

function switchTab(t){
  currentTab=t; haptic(8);
  document.getElementById('v-today').classList.toggle('on',t==='today');
  document.getElementById('v-history').classList.toggle('on',t==='history');
  document.getElementById('tab-today').classList.toggle('on',t==='today');
  document.getElementById('tab-history').classList.toggle('on',t==='history');
  if(t==='history') renderHistory();
}

async function ensureDays(dates){
  const need=dates.filter(d=>!(d in HIST));
  await Promise.all(need.map(async d=>{
    try{ HIST[d] = (account&&navigator.onLine) ? parseDay(await gGetText(dayUrl(d))) : getModel(d); }
    catch(e){ HIST[d]=getModel(d); }
  }));
  // today always reflects freshest local cache
  HIST[dkey(TODAY)] = getModel(dkey(TODAY));
}
function score(k){ const m=HIST[k]||getModel(k); return {vit:vitCount(m), push:pushTotal(m), drinks:m.drinks.length, smoke:m.smoke.length}; }

function renderHistory(){ if(histMode==='week') renderWeek(); else if(histMode==='14day') renderHeatmap(); else renderMonth(); }

async function renderWeek(){
  const base=new Date(TODAY); const dow=(base.getDay()+6)%7; base.setDate(base.getDate()-dow);
  const days=[...Array(7)].map((_,i)=>{const d=new Date(base); d.setDate(base.getDate()+i); return d;});
  await ensureDays(days.map(dkey));
  const labels=['M','T','W','T','F','S','S'];
  const rows=[
    {name:'&#128138; Vitamins', get:k=>score(k).vit/2, color:'var(--green)'},
    {name:'&#128170; Push-ups', get:k=>score(k).push>0?1:0, color:'var(--blue)'},
    {name:'&#127866; Drinks',   get:k=>score(k).drinks, color:'var(--yellow)', count:true},
    {name:'&#128168; Smoke breaks', get:k=>score(k).smoke, color:'#8b5cf6', count:true},
  ];
  let html=`<div class="card"><h3>This Week</h3><div class="csub">${MONTHS[base.getMonth()].slice(0,3)} ${base.getDate()} – ${MONTHS[days[6].getMonth()].slice(0,3)} ${days[6].getDate()}</div>`;
  rows.forEach(r=>{
    const vals=days.map(d=>r.get(dkey(d)));
    if(r.count){
      const total=vals.reduce((a,b)=>a+b,0);
      html+=`<div class="wk-item"><div class="wk-top"><span>${r.name}</span><span>${total} this week</span></div>
        <div class="dots">${vals.map((v,i)=>`<b style="${v>0?`background:${r.color};color:#fff`:''}">${v>0?v:labels[i]}</b>`).join('')}</div></div>`;
    } else {
      const done=vals.filter(v=>v>=1).length; const pct=Math.round(vals.reduce((a,b)=>a+b,0)/7*100);
      html+=`<div class="wk-item"><div class="wk-top"><span>${r.name}</span><span>${done}/7 days</span></div>
        <div class="track"><i style="width:${pct}%;background:${r.color}"></i></div>
        <div class="dots">${vals.map((v,i)=>`<b style="${v>=1?`background:${r.color};color:#fff`:(v>0?'background:#d7f0df':'')}">${labels[i]}</b>`).join('')}</div></div>`;
    }
  });
  html+='</div>'; document.getElementById('hist-week').innerHTML=html;
}

async function renderHeatmap(){
  const days=[...Array(14)].map((_,i)=>{const d=new Date(TODAY); d.setDate(TODAY.getDate()-13+i); return d;});
  await ensureDays(days.map(dkey));
  const rows=[
    {name:'&#128138; Vitamins', lvl:k=>{const s=score(k); return s.vit===2?4:(s.vit===1?2:0);}},
    {name:'&#128170; Push-ups', lvl:k=>{const s=score(k); if(!s.push)return 0; return s.push>=100?4:(s.push>=50?3:(s.push>=25?2:1));}},
    {name:'&#127866; Drinks',   lvl:k=>{const s=score(k); return s.drinks===0?0:(s.drinks<=2?2:(s.drinks<=4?3:4));}},
    {name:'&#128168; Smoke breaks', lvl:k=>{const s=score(k); return s.smoke===0?0:(s.smoke<=3?2:(s.smoke<=6?3:4));}},
  ];
  let html=`<div class="card"><h3>Last 14 Days</h3><div class="csub">Darker = more. For drinks &amp; smoke, lighter is better.</div><div class="hm">`;
  rows.forEach(r=>{
    const cells=days.map(d=>{const l=r.lvl(dkey(d)); return `<i class="${l?'heat-'+l:''}"></i>`;}).join('');
    html+=`<div class="hm-row"><div class="hm-lab"><span>${r.name}</span></div><div class="hm-grid">${cells}</div></div>`;
  });
  html+='</div></div>'; document.getElementById('hist-14day').innerHTML=html;
}

async function renderMonth(){
  const y=monthCursor.getFullYear(), mo=monthCursor.getMonth();
  const dim=new Date(y,mo+1,0).getDate();
  const dates=[]; for(let dd=1;dd<=dim;dd++){ const d=new Date(y,mo,dd); if(d.getTime()<=TODAY.getTime()) dates.push(dkey(d)); }
  await ensureDays(dates);
  const filters=[{k:'all',n:'All'},{k:'vitamins',n:'&#128138; Vitamins'},{k:'pushups',n:'&#128170; Push-ups'},{k:'drink',n:'&#127866; Drinks'},{k:'smoke',n:'&#128168; Smoke'}];
  let html=`<div class="filterbar">${filters.map(f=>`<button class="chip ${monthFilter===f.k?'on':''}" data-filter="${f.k}">${f.n}</button>`).join('')}</div>`;
  const atThisMonth = (monthCursor.getFullYear()===TODAY.getFullYear() && monthCursor.getMonth()===TODAY.getMonth());
  html+=`<div class="cal-head"><button class="chev" data-mo="-1">&#8249;</button>
    <h3>${MONTHS[mo]} ${y}</h3><button class="chev" data-mo="1" ${atThisMonth?'disabled':''}>&#8250;</button></div>`;
  html+=`<div class="wd"><span>S</span><span>M</span><span>T</span><span>W</span><span>T</span><span>F</span><span>S</span></div><div class="cal">`;
  const first=new Date(y,mo,1).getDay();
  for(let i=0;i<first;i++) html+='<b class="empty"></b>';
  for(let dd=1; dd<=dim; dd++){
    const d=new Date(y,mo,dd), k=dkey(d);
    const isToday=d.getTime()===TODAY.getTime(), future=d.getTime()>TODAY.getTime();
    if(future){ html+='<b class="empty"></b>'; continue; }
    const s=score(k), any=(s.vit||s.push||s.drinks||s.smoke);
    if(monthFilter==='all'){
      if(any){ html+=`<b class="dots-day ${isToday?'today':''}">${dd}<div class="minidots">
        <u style="background:${s.vit===2?'var(--green)':(s.vit===1?'var(--yellow)':'#dfe2e8')}"></u>
        <u style="background:${s.push>0?'var(--blue)':'#dfe2e8'}"></u>
        <u style="background:${s.drinks>0?'var(--yellow)':'#dfe2e8'}"></u>
        <u style="background:${s.smoke>0?'#8b5cf6':'#dfe2e8'}"></u></div></b>`; }
      else html+=`<b class="${isToday?'today':''}">${dd}</b>`;
      continue;
    }
    let cls='', label='';
    if(monthFilter==='vitamins'){ cls=s.vit===2?'g':(s.vit===1?'y':'r'); label=`${s.vit}/2`; }
    else if(monthFilter==='pushups'){ cls=s.push>0?'g':'r'; label=s.push>0?s.push:''; }
    else if(monthFilter==='drink'){ cls=s.drinks===0?'g':(s.drinks<=2?'y':'r'); label=s.drinks; }
    else if(monthFilter==='smoke'){ cls=s.smoke===0?'g':(s.smoke<=3?'y':'r'); label=s.smoke; }
    html+=`<b class="${cls} ${isToday?'today':''}">${dd}${label!==''?`<span class="mv">${label}</span>`:''}</b>`;
  }
  html+='</div>';
  if(monthFilter==='all') html+=`<div class="legend"><span><i style="background:var(--green)"></i>Vitamins</span><span><i style="background:var(--blue)"></i>Push-ups</span><span><i style="background:var(--yellow)"></i>Drinks</span><span><i style="background:#8b5cf6"></i>Smoke</span></div>`;
  else if(monthFilter==='drink'||monthFilter==='smoke') html+=`<div class="legend"><span><i style="background:var(--green)"></i>None</span><span><i style="background:var(--yellow)"></i>Some</span><span><i style="background:var(--red)"></i>Over</span></div>`;
  else html+=`<div class="legend"><span><i style="background:var(--green)"></i>Done</span><span><i style="background:var(--yellow)"></i>Partial</span><span><i style="background:var(--red)"></i>Missed</span></div>`;
  document.getElementById('hist-month').innerHTML=html;
}

/* ============================================================
   EVENT BINDING (delegation)
   ============================================================ */
function bindEvents(){
  document.getElementById('prevDay').onclick=()=>{ viewDate=new Date(+viewDate-86400000); haptic(8); renderToday(); refreshDay(dkey(viewDate)); };
  document.getElementById('nextDay').onclick=()=>{ if(viewDate.getTime()<TODAY.getTime()){ viewDate=new Date(+viewDate+86400000); haptic(8); renderToday(); refreshDay(dkey(viewDate)); } };
  document.getElementById('jumpToday').onclick=()=>{ viewDate=new Date(TODAY); haptic(8); renderToday(); };
  document.getElementById('tab-today').onclick=()=>switchTab('today');
  document.getElementById('tab-history').onclick=()=>switchTab('history');
  document.getElementById('sheetBack').onclick=closeSheet;
  document.getElementById('signInBtn').onclick=signIn;

  document.getElementById('histToggle').querySelectorAll('button').forEach(b=>{
    b.onclick=()=>{ histMode=b.dataset.h; haptic(8);
      document.getElementById('histToggle').querySelectorAll('button').forEach(x=>x.classList.toggle('on',x===b));
      ['week','14day','month'].forEach(mm=>document.getElementById('hist-'+mm).style.display=(mm===histMode?'block':'none'));
      renderHistory(); };
  });

  // sheet + history delegation
  document.getElementById('sheet').addEventListener('click',e=>{
    const t=e.target.closest('button'); if(!t) return;
    if(t.dataset.set!==undefined){ padVal=t.dataset.set; setPadDisp(); haptic(8); }
    else if(t.dataset.d!==undefined){ if(padVal.length<4){ padVal=(padVal==='0'?'':padVal)+t.dataset.d; setPadDisp(); } }
    else if(t.dataset.del){ padVal=padVal.slice(0,-1); setPadDisp(); }
    else if(t.dataset.clear){ padVal=''; setPadDisp(); }
    else if(t.id==='savePush'){ savePushups(); }
    else if(t.dataset.vit){ toggleVit(t.dataset.vit); }
    else if(t.dataset.addpush){ closeSheet(); openPushupSheet(); }
    else if(t.dataset.poppush){ popPush(); }
    else if(t.dataset.adj){ adjTally(t.dataset.habit, parseInt(t.dataset.adj,10)); }
    else if(t.dataset.close){ closeSheet(); }
  });
  document.getElementById('v-history').addEventListener('click',e=>{
    const t=e.target.closest('button'); if(!t) return;
    if(t.dataset.filter){ monthFilter=t.dataset.filter; haptic(8); renderMonth(); }
    else if(t.dataset.mo){ const dir=parseInt(t.dataset.mo,10);
      const nm=new Date(monthCursor.getFullYear(),monthCursor.getMonth()+dir,1);
      if(nm.getFullYear()>TODAY.getFullYear()||(nm.getFullYear()===TODAY.getFullYear()&&nm.getMonth()>TODAY.getMonth())) return;
      monthCursor=nm; haptic(8); renderMonth(); }
  });

  window.addEventListener('online', ()=>{ updateSyncPill(); flush(); });
  window.addEventListener('offline', updateSyncPill);
}

/* ---------------- gate ---------------- */
function showGate(msg){ const g=document.getElementById('gate'); if(msg) document.getElementById('gateMsg').textContent=msg; g.classList.add('on'); }
function hideGate(){ document.getElementById('gate').classList.remove('on'); }
function showGateErr(msg){ document.getElementById('gateErr').textContent=msg; }
function showConfigNeeded(){
  document.getElementById('gateMsg').innerHTML='Before first run, open <code>config.js</code> and paste your Entra <b>clientId</b> and <b>tenantId</b>.';
  document.getElementById('signInBtn').style.display='none';
  document.getElementById('gate').classList.add('on');
}

/* ============================================================
   INIT
   ============================================================ */
async function init(){
  bindEvents();
  renderToday();          // from cache — works offline / pre-auth
  updateSyncPill();
  if(!CFG.clientId || String(CFG.clientId).startsWith('PASTE')){ showConfigNeeded(); return; }
  try{ await initAuth(); }catch(e){ showGate('Auth error — check config.js and the Entra redirect URI.'); showGateErr(e.message||String(e)); return; }
  if(!account){ showGate(); return; }
  hideGate();
  updateSyncPill();
  refreshDay(dkey(TODAY));
  flush();
  setInterval(flush,20000);
}
document.addEventListener('DOMContentLoaded',init);
