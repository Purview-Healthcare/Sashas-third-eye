/* Internal Reporting scopes.
   AR is the report the page has always shown (QA, production, insights, team).
   Payment Posting and Submission each take their own production sheet and get
   their own production views, in the same style as AR's. Nothing uploaded here
   is kept between visits; only the daily targets are remembered in the browser.
   Uses the page's own helpers: esc, num, toast, normDate, normName, fmtD, kpi,
   addRowNumbers, makeSortable, addCopyButtons, addHeadingHelp, PROD_MIN_ATTAIN. */
(function(){
'use strict';

let SCOPE = 'ar';

/* ---------- column names each sheet may use (first match wins) ---------- */
const COLS = {
  pp: {
    date:     ['work date','date worked','worked date','posting date','posted date'],
    empId:    ['emp id','employee id','emp code'],
    biller:   ['biller','posted by','worked by','associate','associate name'],
    tl:       ['team lead','supervisor','supervisior','tl','reports to'],
    insurance:['insurance','payer','insurance name'],
    check:    ['check number','check no','check no.','check #','cheque number','eft number','check/eft number','check / eft number'],
    checkDate:['check date','eft date'],
    client:   ['client name','client','project'],
    practice: ['location','practice','practice name'],
    patient:  ['patient name','patient'],
    dos:      ['dos','service date','date of service'],
    billed:   ['billed amount','billed','charge','charges'],
    allowed:  ['allowed amount','allowed'],
    paid:     ['paid amount','paid'],
    posted:   ['posted amount','posted','amount posted'],
    variance: ['variance'],
    ptResp:   ['patient responsibility','pt responsibility'],
    otherAdj: ['other adjustment','other adjustments','adjustment'],
    comment:  ['comment','comments','remarks']
  },
  sub: {
    client:   ['client','client name','project'],
    practice: ['practice','location','practice name'],
    empId:    ['emp id','employee id','emp code'],
    biller:   ['biller','worked by','associate','associate name'],
    tl:       ['team lead','supervisor','supervisior','tl','reports to'],
    date:     ['date worked','work date','worked date','date of submission','submission date'],
    allocatedTo:['allocated to'],
    allocDate:['allocation date'],
    uid:      ['uid','unique id'],
    claimNo:  ['claim no','claim no.','claim number','claim #','claim id'],
    orderNo:  ['order no','order number'],
    dos:      ['service date','dos','date of service'],
    patient:  ['patient name','patient'],
    insurance:['insurance','payer'],
    type:     ['type'],
    claimSource:['claim source'],
    subType:  ['submission type'],
    billed:   ['billed amount','billed','charge','charges'],
    status:   ['status','status by purview'],
    comment:  ['comment','comments'],
    responsibility:['responsibility'],
    claimType:['claim type'],
    software: ['software']
  }
};
const REQUIRED = { pp:['date','biller'], sub:['date','biller','status'] };

const META = {
  pp: { label:'Payment Posting', unit:'line', units:'lines', Units:'Lines',
        title:'Payment Posting Production',
        sub:'Lines posted by biller, day, client and payer, with checks, amounts and posting variance' },
  sub:{ label:'Submission', unit:'claim', units:'claims', Units:'Claims',
        title:'Submission Production',
        sub:'Claims worked and submitted by biller, day, client, practice and status' }
};

const blankFilters = () => ({from:'', to:'', client:'', practice:'', tl:'', biller:''});
const STATE = {
  pp:  { rows:[], file:'', notes:[], filters:blankFilters() },
  sub: { rows:[], file:'', notes:[], filters:blankFilters() }
};

/* ---------- daily targets, remembered in the browser ---------- */
const TKEY = 's3-scope-targets';
let TARGETS = { pp:{def:null, by:{}}, sub:{def:null, by:{}} };
try{ const t = JSON.parse(localStorage.getItem(TKEY)||'null');
  if(t && t.pp && t.sub) TARGETS = t; }catch(e){}
function saveTargets(){ try{ localStorage.setItem(TKEY, JSON.stringify(TARGETS)); }catch(e){} }
function targetFor(kind, nameKey){
  const t = TARGETS[kind];
  const v = t.by[nameKey];
  return (v>0) ? v : (t.def>0 ? t.def : null);
}

/* ---------- small helpers ---------- */
const hnorm = s => String(s??'').replace(/^﻿/,'').replace(/[:\s]+$/,'').replace(/\s+/g,' ').trim().toLowerCase();
const str = v => (v===null||v===undefined) ? '' : (v instanceof Date ? '' : String(v).replace(/\s+/g,' ').trim());
const keyOf = s => String(s||'').replace(/\s+/g,' ').trim().toLowerCase();
const isBlank = v => v===null || v===undefined || (typeof v==='string' && !v.trim());
const fmtN = n => (Math.round(n*10)/10).toLocaleString('en-US');
const fmtInt = n => Math.round(n).toLocaleString('en-US');
const fmtMoney = n => (Math.round(n*100)/100).toLocaleString('en-US',{minimumFractionDigits:2, maximumFractionDigits:2});
const dollars = n => (n<0?'−$':'$') + fmtMoney(Math.abs(n));
// a KPI tile has room for about eight characters
const kpiMoney = n => { const a = Math.abs(n), sg = n<0?'−$':'$';
  return a>=1e6 ? sg+(a/1e6).toFixed(2)+'M' : a>=1e5 ? sg+(a/1e3).toFixed(1)+'K' : sg+fmtInt(a); };
const pctOf = (n,d) => d>0 ? n/d*100 : null;
const fmtP = p => p===null ? '—' : (Math.round(p*10)/10).toFixed(1)+'%';
function dateKey(v){ const d = normDate(v); return d ? d.key : ''; }
function money(v, tally){
  if(isBlank(v)) return 0;
  if(typeof v==='number') return v;
  const s = String(v).trim();
  if(/^[-–—]$/.test(s) || /^n\/?a$/i.test(s)){ if(tally) tally.dash++; return 0; }
  const neg = /^\(.*\)$/.test(s);
  const n = parseFloat(s.replace(/[$,()\s]/g,''));
  if(isNaN(n)){ if(tally) tally.text++; return 0; }
  return neg ? -n : n;
}
function mondayOf(dk){
  const [y,m,d] = dk.split('-').map(Number);
  const dt = new Date(y, m-1, d);
  const shift = (dt.getDay()+6)%7;
  dt.setDate(dt.getDate()-shift);
  const p = n => String(n).padStart(2,'0');
  return `${dt.getFullYear()}-${p(dt.getMonth()+1)}-${p(dt.getDate())}`;
}
function addDays(dk, n){
  const [y,m,d] = dk.split('-').map(Number);
  const dt = new Date(y, m-1, d+n);
  const p = x => String(x).padStart(2,'0');
  return `${dt.getFullYear()}-${p(dt.getMonth()+1)}-${p(dt.getDate())}`;
}
function groupBy(rows, fn){
  const m = new Map();
  rows.forEach(r=>{ const k = fn(r); if(!m.has(k)) m.set(k, []); m.get(k).push(r); });
  return m;
}
const sum = (rows, f) => rows.reduce((s,r)=>s+(f(r)||0), 0);

/* One spelling per client, practice, payer and biller. "NAEC" and "NAEC ",
   "Eyemed" and "EyeMed" are the same thing; the most used spelling is shown. */
function canon(rows, field, label){
  const by = new Map();
  rows.forEach(r=>{
    const v = r[field]; if(!v) return;
    const k = keyOf(v);
    if(!by.has(k)) by.set(k, new Map());
    const m = by.get(k); m.set(v, (m.get(v)||0)+1);
  });
  const pickName = new Map(), merged = [];
  by.forEach((m,k)=>{
    const best = [...m.entries()].sort((a,b)=>b[1]-a[1])[0][0];
    pickName.set(k, best);
    if(m.size>1) merged.push(`${[...m.keys()].map(x=>'“'+x+'”').join(', ')} counted as “${best}”`);
  });
  rows.forEach(r=>{ if(r[field]) r[field] = pickName.get(keyOf(r[field])); });
  return merged.length ? `${label} spelled more than one way: ${merged.slice(0,10).join('; ')}${merged.length>10?`; and ${merged.length-10} more`:''}.` : '';
}

/* ---------- reading a sheet ---------- */
function readAoA(file, cb){
  if(/\.(xlsx|xlsm|xls)$/i.test(file.name)){
    if(typeof XLSX==='undefined'){ toast('Excel reader did not load on this network. Save the sheet as CSV and upload that instead.', true); return; }
    const fr = new FileReader();
    fr.onload = e=>{
      try{
        const wb = XLSX.read(new Uint8Array(e.target.result), {type:'array', cellDates:true});
        cb(wb.SheetNames.map(n=>({name:n, aoa:XLSX.utils.sheet_to_json(wb.Sheets[n], {header:1, raw:true, defval:null})})));
      }catch(err){ console.error(err); toast('Could not read that Excel file.', true); }
    };
    fr.onerror = ()=>toast('Could not read that file.', true);
    fr.readAsArrayBuffer(file);
  } else {
    if(typeof Papa==='undefined'){ toast('CSV reader did not load on this network.', true); return; }
    Papa.parse(file, {header:false, skipEmptyLines:true,
      complete: res=>cb([{name:file.name, aoa:res.data}]),
      error: ()=>toast('Could not parse that CSV.', true)});
  }
}
function layoutOf(aoa, cols){
  let best = null;
  for(let r=0; r<Math.min(aoa.length, 20); r++){
    const hdr = (aoa[r]||[]).map(hnorm);
    const map = {}; let score = 0;
    Object.entries(cols).forEach(([k, names])=>{
      for(const n of names){ const i = hdr.indexOf(n); if(i>=0){ map[k]=i; score++; return; } }
    });
    if(!best || score>best.score) best = {row:r, map, score};
  }
  return best;
}
function rosterOf(sheets){
  const v = sheets.find(s=>/validation|roster/i.test(s.name));
  const out = { byId:new Map(), byName:new Map(), dupIds:[] };
  if(!v) return out;
  const lay = layoutOf(v.aoa, { empId:['emp id','employee id'], name:['biller','name','associate','employee name'],
    tl:['team lead','supervisor','supervisior','reports to'], status:['status'], doj:['doj','date of joining'] });
  if(!lay || lay.map.name===undefined) return out;
  for(let i=lay.row+1; i<v.aoa.length; i++){
    const r = v.aoa[i]||[]; const g = k => lay.map[k]===undefined ? '' : str(r[lay.map[k]]);
    const name = g('name'); if(!name) continue;
    const e = { name, empId:g('empId'), tl:g('tl'), status:g('status') };
    out.byName.set(normName(name), e);
    if(e.empId){
      const prev = out.byId.get(keyOf(e.empId));
      if(prev && normName(prev.name)!==normName(name)) out.dupIds.push(`${e.empId} (${prev.name} and ${name})`);
      else out.byId.set(keyOf(e.empId), e);
    }
  }
  return out;
}

function parseSheets(kind, sheets){
  const cols = COLS[kind];
  let pick = null;
  sheets.forEach(s=>{
    if(/^(validation|lists?|how to)/i.test(s.name.trim())) return;
    const lay = layoutOf(s.aoa, cols);
    if(lay && REQUIRED[kind].every(k=>lay.map[k]!==undefined) && (!pick || lay.score>pick.lay.score)) pick = {s, lay};
  });
  if(!pick) return { error: `No sheet in this file has the ${REQUIRED[kind].map(k=>({date:kind==='pp'?'Work Date':'Date Worked', biller:'Biller', status:'Status'})[k]).join(', ')} columns. Is this the ${META[kind].label} sheet?` };
  const roster = rosterOf(sheets);
  const {s, lay} = pick, M = lay.map;
  const out = [], tally = {dash:0, text:0};
  let noDate = 0, blankTemplate = 0, noBiller = 0, rosterTL = 0;
  let billerFromId = 0; const billerUnknownId = new Set();
  for(let i=lay.row+1; i<s.aoa.length; i++){
    const raw = s.aoa[i]||[];
    if(raw.every(isBlank)) continue;
    const g = k => M[k]===undefined ? null : raw[M[k]];
    const t = k => str(g(k));
    const date = dateKey(g('date'));
    // the template ships with Emp ID, Biller and Team Lead filled in on otherwise
    // empty rows; a row only counts as work once it has something worked on it
    const workFields = kind==='pp' ? ['paid','posted','patient','check','insurance','billed']
                                   : ['claimNo','status','patient','uid','insurance','billed'];
    const hasWork = workFields.some(k=>!isBlank(g(k)));
    if(!date){ if(hasWork) noDate++; else blankTemplate++; continue; }
    // the Submission template looks Biller and Team Lead up from Emp ID; a file
    // saved before Excel worked them out has them blank, and an unknown Emp ID
    // gives "Not found", so the file's own Validation Sheet answers instead
    const empId = t('empId');
    const ro0 = empId ? roster.byId.get(keyOf(empId)) : null;
    let biller = t('biller'), tl = t('tl');
    if(/^not found$/i.test(tl)) tl = '';
    if(!biller || /^not found$/i.test(biller)){
      if(ro0 && ro0.name){ biller = ro0.name; billerFromId++; }
      else if(empId){ biller = empId; billerUnknownId.add(empId); }
      else biller = '';
    }
    const r = { date, biller: biller || '(biller not stated)', empId,
      tl, client: t('client'), practice: t('practice'), insurance: t('insurance'),
      patient: t('patient'), dos: dateKey(g('dos')), comment: t('comment') };
    if(!biller) noBiller++;
    if(kind==='pp'){
      Object.assign(r, { check: t('check'), checkDate: dateKey(g('checkDate')),
        billed: money(g('billed'), tally), allowed: money(g('allowed'), tally), paid: money(g('paid'), tally),
        posted: money(g('posted'), tally), variance: money(g('variance'), tally),
        ptResp: money(g('ptResp'), tally), otherAdj: money(g('otherAdj'), tally) });
      r.diff = Math.round((r.paid - r.posted)*100)/100;
      r.isVar = Math.abs(r.diff)>=0.005 || Math.abs(r.variance)>=0.005;
    } else {
      Object.assign(r, { claimNo: t('claimNo'), uid: t('uid'), orderNo: t('orderNo'), status: t('status'),
        allocatedTo: t('allocatedTo'), allocDate: dateKey(g('allocDate')), type: t('type'),
        claimSource: t('claimSource'), subType: t('subType'), billed: money(g('billed'), tally),
        responsibility: t('responsibility'), claimType: t('claimType'), software: t('software') });
    }
    r.nameKey = normName(r.biller);
    const ro = (r.empId && roster.byId.get(keyOf(r.empId))) || roster.byName.get(r.nameKey);
    if(!r.tl && ro && ro.tl){ r.tl = ro.tl; rosterTL++; }
    r.rosterStatus = ro ? ro.status : '';
    out.push(r);
  }
  const notes = [];
  ['client','practice','insurance','tl','biller'].forEach(f=>{
    const n = canon(out, f, ({client:'Clients', practice: kind==='pp'?'Locations':'Practices', insurance:'Payers', tl:'Team leads', biller:'Billers'})[f]);
    if(n) notes.push(n);
  });
  out.forEach(r=>{ r.nameKey = normName(r.biller); });
  if(noDate) notes.push(`${noDate} row${noDate!==1?'s':''} with work on ${noDate!==1?'them':'it'} but no ${kind==='pp'?'Work Date':'Date Worked'} ${noDate!==1?'are':'is'} left out, because a day can't be given to ${noDate!==1?'them':'it'}.`);
  if(noBiller) notes.push(`${noBiller} row${noBiller!==1?'s have':' has'} no Biller and ${noBiller!==1?'are':'is'} shown as “(biller not stated)”.`);
  if(billerFromId) notes.push(`${billerFromId} row${billerFromId!==1?'s':''} had no Biller name, so it was taken from the Emp ID on the file’s Validation Sheet.`);
  if(billerUnknownId.size) notes.push(`Emp ID${billerUnknownId.size!==1?'s':''} ${[...billerUnknownId].slice(0,8).join(', ')} ${billerUnknownId.size!==1?'are':'is'} not on the file’s Validation Sheet, so the Emp ID is shown in place of the name.`);
  if(roster.dupIds.length) notes.push(`The Validation Sheet gives the same Emp ID to two people: ${roster.dupIds.join('; ')}. Rows with that Emp ID can be credited to the wrong person; give each person their own Emp ID.`);
  if(rosterTL) notes.push(`${rosterTL} row${rosterTL!==1?'s':''} had no team lead, so the one on the file’s Validation Sheet was used.`);
  if(tally.dash) notes.push(`${tally.dash} amount cell${tally.dash!==1?'s':''} held “-” instead of a number and ${tally.dash!==1?'were':'was'} read as 0.`);
  if(tally.text) notes.push(`${tally.text} amount cell${tally.text!==1?'s':''} held text that isn’t a number and ${tally.text!==1?'were':'was'} read as 0.`);
  return { rows: out, sheet: s.name, notes, blankTemplate };
}

function handleFile(kind, file){
  readAoA(file, sheets=>{
    const res = parseSheets(kind, sheets);
    if(res.error){ toast(res.error, true); return; }
    const st = STATE[kind];
    st.rows = res.rows; st.file = file.name; st.notes = res.notes; st.filters = blankFilters();
    if(!res.rows.length){ toast(`No worked rows in ${file.name} yet — every row is missing its ${kind==='pp'?'Work Date':'Date Worked'}.`, true); }
    else toast(`${META[kind].label}: ${res.rows.length.toLocaleString('en-US')} ${META[kind].units} loaded from “${res.sheet}”.`);
    if(SCOPE!==kind) setScope(kind); else draw();
  });
}

/* ---------- filters ---------- */
function filtered(kind, skip){
  const f = STATE[kind].filters;
  return STATE[kind].rows.filter(r=>
    (skip==='date' || ((!f.from || r.date>=f.from) && (!f.to || r.date<=f.to))) &&
    (skip==='client'   || !f.client   || r.client===f.client) &&
    (skip==='practice' || !f.practice || r.practice===f.practice) &&
    (skip==='tl'       || !f.tl       || r.tl===f.tl) &&
    (skip==='biller'   || !f.biller   || r.nameKey===f.biller));
}
function filtersHTML(kind){
  const st = STATE[kind], f = st.filters, M = META[kind];
  if(!st.rows.length) return '';
  const opts = (rows, fn, cur, allLbl, lblFn) => {
    const vals = [...new Set(rows.map(fn).filter(Boolean))].sort((a,b)=>String(lblFn?lblFn(a):a).localeCompare(String(lblFn?lblFn(b):b)));
    if(cur && !vals.includes(cur)) vals.unshift(cur);
    return `<option value="">${allLbl}</option>` + vals.map(v=>`<option value="${esc(v)}"${v===cur?' selected':''}>${esc(lblFn?lblFn(v):v)}</option>`).join('');
  };
  const names = new Map(st.rows.map(r=>[r.nameKey, r.biller]));
  const dates = st.rows.map(r=>r.date).sort();
  const pLbl = kind==='pp' ? 'Location' : 'Practice';
  return `<div class="period-bar sc-filters" data-kind="${kind}">
    <span class="plabel">${kind==='pp'?'Work Date':'Date Worked'}</span>
    <input type="date" data-f="from" value="${esc(f.from)}" min="${dates[0]}" max="${dates[dates.length-1]}" title="Filters the ${esc(M.label)} sheet on its ${kind==='pp'?'Work Date':'Date Worked'} column" />
    <span class="plabel">to</span>
    <input type="date" data-f="to" value="${esc(f.to)}" min="${dates[0]}" max="${dates[dates.length-1]}" title="Filters the ${esc(M.label)} sheet on its ${kind==='pp'?'Work Date':'Date Worked'} column" />
    <button class="btn small" data-act="full" type="button">Full range</button>
    <span class="plabel">Client</span>
    <select data-f="client">${opts(filtered(kind,'client'), r=>r.client, f.client, 'All clients')}</select>
    <span class="plabel">${pLbl}</span>
    <select data-f="practice">${opts(filtered(kind,'practice'), r=>r.practice, f.practice, `All ${pLbl.toLowerCase()}s`)}</select>
    <span class="plabel">Team Lead</span>
    <select data-f="tl">${opts(filtered(kind,'tl'), r=>r.tl, f.tl, 'All team leads')}</select>
    <span class="plabel">Biller</span>
    <select data-f="biller">${opts(filtered(kind,'biller'), r=>r.nameKey, f.biller, 'All billers', k=>names.get(k)||k)}</select>
  </div>`;
}
function filterText(kind){
  const f = STATE[kind].filters, st = STATE[kind];
  const name = f.biller ? ((st.rows.find(r=>r.nameKey===f.biller)||{}).biller || f.biller) : '';
  const bits = [f.client?'Client '+f.client:'', f.practice?(kind==='pp'?'Location ':'Practice ')+f.practice:'',
    f.tl?'Team Lead '+f.tl:'', name?'Biller '+name:'',
    (f.from||f.to)?'Dates '+(f.from||'start')+' to '+(f.to||'end'):''].filter(Boolean);
  return bits.length ? bits.join(' · ') : 'none — all data';
}

/* ---------- table builder: every heading carries its description ---------- */
function th(h){ return `<th class="${h.num?'num':''}" title="${esc(h.d||'')}">${esc(h.l)}</th>`; }
function td(v, isNum){
  const html = (v && typeof v==='object' && 'html' in v) ? v.html : esc(v==null?'':v);
  return `<td class="${isNum?'num':''}">${html}</td>`;
}
function tableHTML(o){
  const body = o.rows.length ? o.rows.map(r=>`<tr class="${r.cls||''}">${(r.cells||r).map((c,i)=>td(c, o.heads[i].num)).join('')}</tr>`).join('')
    : `<tr><td colspan="${o.heads.length}" style="color:var(--text-faint);">${esc(o.empty||'Nothing in this selection.')}</td></tr>`;
  return `<div class="card">
    <h3>${esc(o.title)}</h3><div class="card-sub">${o.sub||''}</div>
    <div class="table-scroll"><table class="data">
      <thead><tr>${o.heads.map(th).join('')}</tr></thead>
      <tbody>${body}</tbody>
      ${o.foot && o.rows.length ? `<tfoot><tr>${o.foot.map((c,i)=>td(c, o.heads[i].num)).join('')}</tr></tfoot>` : ''}
    </table></div>${o.note?`<div class="sc-note">${o.note}</div>`:''}</div>`;
}
const stamp = (p) => p===null ? '—' : {html:`<span class="stamp ${p>=PROD_MIN_ATTAIN?'good':'bad'}">${fmtP(p)}</span>`};

/* ---------- shared production pieces ---------- */
function billerList(kind, rows, extra){
  return [...groupBy(rows, r=>r.nameKey)].map(([k, rs])=>{
    const days = new Set(rs.map(r=>r.date)).size;
    const tgt = targetFor(kind, k);
    const target = tgt ? tgt*days : null;
    const tlCount = groupBy(rs.filter(r=>r.tl), r=>r.tl);
    const tl = [...tlCount].sort((a,b)=>b[1].length-a[1].length).map(x=>x[0])[0] || '';
    const o = { key:k, name:rs[0].biller, tl, status: rs.find(r=>r.rosterStatus)?.rosterStatus || '',
      days, units: rs.length, perDay: rs.length/days, target, attain: target ? pctOf(rs.length, target) : null, rs };
    return Object.assign(o, extra ? extra(rs) : {});
  }).sort((a,b)=>(a.tl||'~').localeCompare(b.tl||'~') || b.units-a.units);
}
function headerBands(kind, list, rows){
  const M = META[kind];
  const dayCount = new Set(rows.map(r=>r.date)).size;
  const withT = list.filter(b=>b.attain!==null);
  if(!withT.length) return `<div class="alert-band sc-hint"><div class="ab-title">No daily target set yet</div>
    <div class="ab-body">Enter a target in the Daily Target card below to see each biller’s attainment against the ${PROD_MIN_ATTAIN}% minimum.</div></div>`;
  const below = withT.filter(b=>b.attain<PROD_MIN_ATTAIN).sort((a,b)=>a.attain-b.attain);
  return below.length ? `<div class="alert-band bad">
      <div class="ab-title">${below.length} biller${below.length!==1?'s':''} below the ${PROD_MIN_ATTAIN}% minimum</div>
      <div class="ab-body">${below.slice(0,6).map((b,i)=>`${i+1})&nbsp;${esc(b.name)} — ${fmtP(b.attain)}`).join(' &nbsp;·&nbsp; ')}${below.length>6?` · +${below.length-6} more`:''}</div></div>`
    : `<div class="alert-band good"><div class="ab-title">Every biller with a target met the ${PROD_MIN_ATTAIN}% minimum</div>
      <div class="ab-body">${withT.length} biller(s) across ${dayCount} working day(s), ${esc(M.units)} counted one per row.</div></div>`;
}
function targetCardHTML(kind, list){
  const M = META[kind], T = TARGETS[kind];
  return `<div class="card sc-targets" data-kind="${kind}">
    <h3>Daily Target</h3>
    <div class="card-sub">${M.Units} each biller is expected to ${kind==='pp'?'post':'work'} per working day. A biller’s own figure wins over the team default. Remembered in this browser; the sheets are not.</div>
    <div class="sc-tgt-row"><label>Team default <input type="number" min="0" step="1" data-tgt="__def__" value="${T.def>0?T.def:''}" placeholder="—"> ${esc(M.units)} / day</label></div>
    <div class="sc-tgt-grid">${list.map(b=>`<label title="${esc(b.tl?('Team lead '+b.tl):'')}"><span>${esc(b.name)}</span>
      <input type="number" min="0" step="1" data-tgt="${esc(b.key)}" value="${T.by[b.key]>0?T.by[b.key]:''}" placeholder="${T.def>0?T.def:'—'}"></label>`).join('')}</div>
  </div>`;
}
function dailyChartHTML(kind, rows){
  const M = META[kind];
  const by = groupBy(rows, r=>r.date);
  const days = [...by.keys()].sort();
  if(days.length<2) return '';
  const vals = days.map(d=>by.get(d).length);
  const tgts = days.map(d=>{ const rs = by.get(d); const ks = new Set(rs.map(r=>r.nameKey));
    let t = 0, any = false; ks.forEach(k=>{ const v = targetFor(kind, k); if(v){ t+=v; any=true; } }); return any ? t : null; });
  const max = Math.max(...vals, ...tgts.filter(x=>x!==null), 1);
  const W = 900, H = 190, L = 34, B = 28, T = 10;
  const bw = (W-L-6)/days.length;
  const y = v => T + (H-T-B)*(1 - v/max);
  const ticks = [0, Math.round(max/2), Math.round(max)];
  const bars = days.map((d,i)=>{
    const x = L + i*bw + bw*0.15, w = bw*0.7, v = vals[i];
    const tip = `${fmtD(d)}: ${v} ${M.units}${tgts[i]!==null?` · target ${tgts[i]}`:''}`;
    return `<rect x="${x.toFixed(1)}" y="${y(v).toFixed(1)}" width="${w.toFixed(1)}" height="${(H-B-y(v)).toFixed(1)}" rx="2" class="sc-bar"><title>${esc(tip)}</title></rect>`
      + (tgts[i]!==null ? `<line x1="${(x-2).toFixed(1)}" x2="${(x+w+2).toFixed(1)}" y1="${y(tgts[i]).toFixed(1)}" y2="${y(tgts[i]).toFixed(1)}" class="sc-tline"/>` : '')
      + ((days.length<=16 || i%Math.ceil(days.length/16)===0) ? `<text x="${(x+w/2).toFixed(1)}" y="${H-10}" text-anchor="middle" class="sc-ax">${esc(fmtD(d).replace(/ \d{4}$/,''))}</text>` : '');
  }).join('');
  const grid = ticks.map(t=>`<line x1="${L}" x2="${W}" y1="${y(t).toFixed(1)}" y2="${y(t).toFixed(1)}" class="sc-grid"/><text x="${L-5}" y="${(y(t)+4).toFixed(1)}" text-anchor="end" class="sc-ax">${t}</text>`).join('');
  return `<div class="card"><h3>${esc(M.Units)} by Day</h3>
    <div class="card-sub">Every ${esc(M.unit)} in the selection, by the day it was worked.${tgts.some(t=>t!==null)?' The short line on a bar is that day’s target: each biller who worked that day × their target / day.':''}</div>
    <svg viewBox="0 0 ${W} ${H}" class="sc-chart" role="img" aria-label="${esc(M.Units)} by day">${grid}${bars}</svg></div>`;
}
function weekTable(kind, rows, extraHeads, extraCells){
  const M = META[kind];
  const by = groupBy(rows, r=>mondayOf(r.date));
  const list = [...by].sort((a,b)=>a[0]<b[0]?-1:1).map(([wk, rs])=>{
    const days = new Set(rs.map(r=>r.date)).size;
    const bd = new Set(rs.map(r=>r.nameKey+'|'+r.date)).size;
    return [`${fmtD(wk)} – ${fmtD(addDays(wk,6))}`, days, new Set(rs.map(r=>r.nameKey)).size, rs.length, fmtN(rs.length/bd), ...extraCells(rs)];
  });
  const bdAll = new Set(rows.map(r=>r.nameKey+'|'+r.date)).size;
  return tableHTML({ title:'Production by Week', sub:`Monday to Sunday weeks, from the ${kind==='pp'?'Work Date':'Date Worked'} column.`,
    heads:[{l:'Week', d:'Monday to Sunday.'}, {l:'Days', d:'Days in the week with any work.', num:1},
      {l:'Billers', d:'Billers with work that week.', num:1}, {l:M.Units, d:`${M.Units} worked that week, one per row of the sheet.`, num:1},
      {l:`${M.Units} / Biller / Day`, d:`${M.Units} ÷ the biller-days worked (each biller on each day they worked counts once).`, num:1}, ...extraHeads],
    rows:list,
    foot:['TOTAL', new Set(rows.map(r=>r.date)).size, new Set(rows.map(r=>r.nameKey)).size, rows.length, bdAll?fmtN(rows.length/bdAll):'—', ...extraCells(rows)] });
}
function billerDayMatrix(kind, rows){
  const M = META[kind];
  const days = [...new Set(rows.map(r=>r.date))].sort();
  if(!days.length || days.length>62) return '';
  const by = groupBy(rows, r=>r.nameKey);
  const list = [...by].map(([k,rs])=>({name:rs[0].biller, tl:rs[0].tl, c:groupBy(rs, r=>r.date), n:rs.length}))
    .sort((a,b)=>(a.tl||'~').localeCompare(b.tl||'~') || b.n-a.n);
  const heads = [{l:'Biller', d:'The person who did the work.'}, {l:'Team Lead', d:'From the sheet, or its Validation Sheet when the row has none.'},
    ...days.map(d=>({l:fmtD(d).replace(/ \d{4}$/,''), d:`${M.Units} worked on ${fmtD(d)}.`, num:1})), {l:'Total', d:`All ${M.units} in the selection.`, num:1}];
  const rowsOut = list.map(b=>[b.name, b.tl||'—', ...days.map(d=>{ const v = (b.c.get(d)||[]).length;
    const t = targetFor(kind, normName(b.name));
    return v ? (t ? {html:`<span class="${v>=t*PROD_MIN_ATTAIN/100?'sc-ok':'sc-low'}">${v}</span>`} : v) : ''; }), b.n]);
  const foot = ['TOTAL', '', ...days.map(d=>rows.filter(r=>r.date===d).length), rows.length];
  return tableHTML({ title:`Daily ${M.Units} by Biller`, sub:`Each biller’s ${M.units} on each day. With a target set, a figure under ${PROD_MIN_ATTAIN}% of it shows in red.`,
    heads, rows:rowsOut, foot });
}
function notesHTML(kind){
  const n = STATE[kind].notes;
  if(!n.length) return '';
  return `<div class="card"><h3>Sheet Checks</h3><div class="card-sub">What was found while reading ${esc(STATE[kind].file)}. Nothing here changes the sheet itself.</div>
    <ul class="sc-list">${n.map(x=>`<li>${esc(x)}</li>`).join('')}</ul></div>`;
}

/* ---------- Payment Posting ---------- */
function ppHTML(){
  const rows = filtered('pp');
  if(!rows.length) return `<div class="empty-state"><h3>Nothing in this selection</h3><p>Widen the dates or set a filter back to “All”.</p></div>`;
  const checks = rs => new Set(rs.filter(r=>r.check).map(r=>keyOf(r.check))).size;
  const list = billerList('pp', rows, rs=>({ checks:checks(rs), paid:sum(rs,r=>r.paid), posted:sum(rs,r=>r.posted),
    varLines: rs.filter(r=>r.isVar).length }));
  const paid = sum(rows,r=>r.paid), posted = sum(rows,r=>r.posted);
  const varRows = rows.filter(r=>r.isVar);
  const bd = new Set(rows.map(r=>r.nameKey+'|'+r.date)).size;
  const withT = list.filter(b=>b.target);
  const tAll = sum(withT, b=>b.target), uT = sum(withT, b=>b.units);
  const noCheck = rows.filter(r=>!r.check).length;

  const moneyHeads = [{l:'Paid ($)', d:'Paid Amount on the sheet: what the payer paid.', num:1},
    {l:'Posted ($)', d:'Posted Amount on the sheet: what the biller posted.', num:1}];
  const byBiller = tableHTML({ title:'Production by Biller', sub:'One row per line posted. Grouped by team lead, highest volume first.',
    heads:[{l:'Biller', d:'The person who posted the lines.'}, {l:'Team Lead', d:'Supervisor from the sheet, or its Validation Sheet when the row has none.'},
      {l:'Status', d:'Production, OJT, Training or PIP / QIP, from the sheet’s Validation Sheet.'},
      {l:'Days Worked', d:'Days with at least one line posted.', num:1}, {l:'Lines', d:'Lines posted in the selection.', num:1},
      {l:'Lines / Day', d:'Lines ÷ days worked.', num:1}, {l:'Target', d:'Target / day × days worked. Blank until a target is set.', num:1},
      {l:'Attainment', d:`Lines as a share of target. Under ${PROD_MIN_ATTAIN}% is flagged.`, num:1},
      {l:'Checks', d:'Distinct check / EFT numbers the biller posted lines from.', num:1}, ...moneyHeads,
      {l:'Variance Lines', d:'Lines where Paid and Posted differ, or the Variance column isn’t 0.', num:1}],
    rows: list.map(b=>({cls: b.attain!==null && b.attain<PROD_MIN_ATTAIN ? 'is-outlier' : '',
      cells:[b.name, b.tl||'—', b.status||'—', b.days, b.units, fmtN(b.perDay), b.target?fmtInt(b.target):'—', stamp(b.attain),
        b.checks, fmtMoney(b.paid), fmtMoney(b.posted), b.varLines||'']})),
    foot:['TOTAL', '', '', new Set(rows.map(r=>r.date)).size, rows.length, bd?fmtN(rows.length/bd):'—', tAll?fmtInt(tAll):'—',
      stamp(tAll?pctOf(uT,tAll):null), checks(rows), fmtMoney(paid), fmtMoney(posted), varRows.length||''] });

  const dayRows = [...groupBy(rows, r=>r.date)].sort((a,b)=>a[0]<b[0]?-1:1).map(([d,rs])=>
    [fmtD(d), new Set(rs.map(r=>r.nameKey)).size, rs.length, checks(rs), fmtMoney(sum(rs,r=>r.paid)), fmtMoney(sum(rs,r=>r.posted)), rs.filter(r=>r.isVar).length||'']);
  const byDay = tableHTML({ title:'Production by Day', sub:'Every working day in the selection.',
    heads:[{l:'Date', d:'The Work Date.'}, {l:'Billers', d:'Billers who posted that day.', num:1}, {l:'Lines', d:'Lines posted that day.', num:1},
      {l:'Checks', d:'Distinct check / EFT numbers posted from that day.', num:1}, ...moneyHeads,
      {l:'Variance Lines', d:'Lines where Paid and Posted differ, or the Variance column isn’t 0.', num:1}],
    rows: dayRows, foot:['TOTAL', new Set(rows.map(r=>r.nameKey)).size, rows.length, checks(rows), fmtMoney(paid), fmtMoney(posted), varRows.length||''] });

  const byWeek = weekTable('pp', rows, [...moneyHeads], rs=>[fmtMoney(sum(rs,r=>r.paid)), fmtMoney(sum(rs,r=>r.posted))]);

  const cp = [...groupBy(rows, r=>(r.client||'(client not stated)')+'||'+(r.practice||'(location not stated)'))]
    .map(([k,rs])=>{ const [c,p] = k.split('||'); return {c,p,rs}; }).sort((a,b)=>a.c.localeCompare(b.c) || b.rs.length-a.rs.length);
  const byClient = tableHTML({ title:'Production by Client & Location', sub:'Client Name and Location from the sheet. A blank Location shows as “(location not stated)”.',
    heads:[{l:'Client', d:'Client Name on the sheet.'}, {l:'Location', d:'Location on the sheet.'}, {l:'Lines', d:'Lines posted.', num:1},
      {l:'Share', d:'Share of all lines in the selection.', num:1}, {l:'Billers', d:'Billers who posted for it.', num:1},
      {l:'Checks', d:'Distinct check / EFT numbers.', num:1}, ...moneyHeads, {l:'Variance Lines', d:'Lines where Paid and Posted differ, or the Variance column isn’t 0.', num:1}],
    rows: cp.map(x=>[x.c, x.p, x.rs.length, fmtP(pctOf(x.rs.length, rows.length)), new Set(x.rs.map(r=>r.nameKey)).size, checks(x.rs),
      fmtMoney(sum(x.rs,r=>r.paid)), fmtMoney(sum(x.rs,r=>r.posted)), x.rs.filter(r=>r.isVar).length||'']),
    foot:['TOTAL', '', rows.length, '100.0%', new Set(rows.map(r=>r.nameKey)).size, checks(rows), fmtMoney(paid), fmtMoney(posted), varRows.length||''] });

  const ins = [...groupBy(rows, r=>r.insurance||'(payer not stated)')].sort((a,b)=>b[1].length-a[1].length);
  const byPayer = tableHTML({ title:'Production by Payer', sub:'Insurance column on the sheet, spellings that differ only in case or spaces counted together.',
    heads:[{l:'Payer', d:'Insurance on the sheet.'}, {l:'Lines', d:'Lines posted.', num:1}, {l:'Checks', d:'Distinct check / EFT numbers.', num:1},
      {l:'Billed ($)', d:'Billed Amount: what was charged.', num:1}, {l:'Allowed ($)', d:'Allowed Amount: what the payer allows.', num:1}, ...moneyHeads,
      {l:'Patient Resp. ($)', d:'Patient Responsibility on the sheet.', num:1}, {l:'Paid % of Allowed', d:'Paid ÷ Allowed.', num:1}],
    rows: ins.map(([p,rs])=>{ const al = sum(rs,r=>r.allowed), pd = sum(rs,r=>r.paid);
      return [p, rs.length, checks(rs), fmtMoney(sum(rs,r=>r.billed)), fmtMoney(al), fmtMoney(pd), fmtMoney(sum(rs,r=>r.posted)), fmtMoney(sum(rs,r=>r.ptResp)), fmtP(pctOf(pd, al))]; }),
    foot:['TOTAL', rows.length, checks(rows), fmtMoney(sum(rows,r=>r.billed)), fmtMoney(sum(rows,r=>r.allowed)), fmtMoney(paid), fmtMoney(posted), fmtMoney(sum(rows,r=>r.ptResp)), fmtP(pctOf(paid, sum(rows,r=>r.allowed)))] });

  const varList = tableHTML({ title:'Posting Variance Lines', sub:'Every line where Paid and Posted differ, or the sheet’s own Variance column isn’t 0. These are the lines to check.',
    heads:[{l:'Work Date', d:'The day it was posted.'}, {l:'Biller', d:'Who posted it.'}, {l:'Client', d:'Client Name.'}, {l:'Payer', d:'Insurance.'},
      {l:'Check No.', d:'Check / EFT number.'}, {l:'Patient', d:'Patient Name.'}, {l:'DOS', d:'Date of service.'},
      ...moneyHeads, {l:'Paid − Posted ($)', d:'Paid Amount minus Posted Amount.', num:1}, {l:'Variance ($)', d:'The Variance column as entered on the sheet.', num:1}],
    rows: varRows.slice().sort((a,b)=>Math.abs(b.diff)-Math.abs(a.diff)).map(r=>[fmtD(r.date), r.biller, r.client, r.insurance, r.check, r.patient, r.dos?fmtD(r.dos):'',
      fmtMoney(r.paid), fmtMoney(r.posted), fmtMoney(r.diff), fmtMoney(r.variance)]),
    empty:'No variance in this selection: Paid equals Posted on every line.' });

  const ck = [];
  if(noCheck) ck.push(`${noCheck} line${noCheck!==1?'s have':' has'} no check number.`);

  return `${headerBands('pp', list, rows)}
    <div class="kpi-grid">
      ${kpi('Billers', list.length, '', `${new Set(rows.map(r=>r.date)).size} working day(s)`)}
      ${kpi('Lines Posted', fmtInt(rows.length), '', 'one per row of the sheet')}
      ${kpi('Lines / Biller / Day', bd?fmtN(rows.length/bd):'—', '', 'average over biller-days worked')}
      ${kpi('Attainment', tAll?fmtP(pctOf(uT,tAll)):'—', tAll?(pctOf(uT,tAll)>=PROD_MIN_ATTAIN?'good':'bad'):'', tAll?`minimum ${PROD_MIN_ATTAIN}%`:'set a daily target')}
      ${kpi('Checks Posted', fmtInt(checks(rows)), '', noCheck?`${noCheck} line(s) without a check no.`:'distinct check / EFT numbers')}
      ${kpi('Paid Amount', kpiMoney(paid), '', `${dollars(paid)} from the payer`)}
      ${kpi('Posted Amount', kpiMoney(posted), '', `${dollars(posted)} posted by the team`)}
      ${kpi('Variance Lines', varRows.length, varRows.length?'bad':'good', `${dollars(sum(varRows,r=>r.diff))} paid − posted`)}
    </div>
    ${targetCardHTML('pp', list)}
    <div class="section-title">Production by Biller</div>
    ${byBiller}
    ${billerDayMatrix('pp', rows)}
    <div class="section-title">Production by Day &amp; Week</div>
    ${dailyChartHTML('pp', rows)}
    ${byDay}
    ${byWeek}
    <div class="section-title">Production by Client &amp; Location</div>
    ${byClient}
    <div class="section-title">Production by Payer</div>
    ${byPayer}
    <div class="section-title">Posting Variance</div>
    ${varList}
    ${ck.length||STATE.pp.notes.length ? `<div class="section-title">Sheet Checks</div>${notesHTML('pp')}${ck.length?`<div class="card"><h3>In This Selection</h3><ul class="sc-list">${ck.map(x=>`<li>${esc(x)}</li>`).join('')}</ul></div>`:''}` : ''}`;
}

/* ---------- Submission ---------- */
const SUB_GROUPS = [
  ['Submitted', /^submitted$/i],
  ['Partially Submitted', /^partially submitted$/i],
  ['Not Submitted', /^(not submitted|created|untouched)$/i],
  ['Pending / On Hold', /^(pending|on hold|hold)$/i],
  ['Needs Assistance', /(assistance|client responsibility)/i],
  ['Non Workable', /^non[- ]?workable$/i]
];
function subGroup(status){
  if(!status) return 'Status not stated';
  for(const [g, re] of SUB_GROUPS) if(re.test(status.trim())) return g;
  return 'Other';
}
const isSubmitted = r => { const g = subGroup(r.status); return g==='Submitted' || g==='Partially Submitted'; };
function subHTML(){
  const rows = filtered('sub');
  if(!rows.length) return `<div class="empty-state"><h3>Nothing in this selection</h3><p>Widen the dates or set a filter back to “All”.</p></div>`;
  rows.forEach(r=>{ r.group = subGroup(r.status); });
  const cnt = (rs, g) => rs.filter(r=>r.group===g).length;
  const subd = rs => rs.filter(isSubmitted);
  const list = billerList('sub', rows, rs=>({ s:cnt(rs,'Submitted'), ps:cnt(rs,'Partially Submitted'), ns:cnt(rs,'Not Submitted'),
    ph:cnt(rs,'Pending / On Hold'), as:cnt(rs,'Needs Assistance'), billed:sum(subd(rs), r=>r.billed) }));
  const bd = new Set(rows.map(r=>r.nameKey+'|'+r.date)).size;
  const withT = list.filter(b=>b.target);
  const tAll = sum(withT, b=>b.target), uT = sum(withT, b=>b.units);
  const sAll = subd(rows).length;
  const subRate = rs => fmtP(pctOf(subd(rs).length, rs.length));

  const statusHeads = [
    {l:'Submitted', d:'Status “Submitted”.', num:1}, {l:'Partially', d:'Status “Partially Submitted”.', num:1},
    {l:'Not Submitted', d:'Status Not Submitted, Created or Untouched.', num:1},
    {l:'Pending / Hold', d:'Status Pending or On Hold.', num:1},
    {l:'Assistance', d:'Status Client, AR, EV or Posting Assistance, or Client Responsibility.', num:1},
    {l:'Submission %', d:'Submitted plus partially submitted ÷ claims worked.', num:1},
    {l:'Billed Submitted ($)', d:'Billed Amount of the claims submitted or partially submitted.', num:1}];
  const statusCells = rs => [cnt(rs,'Submitted'), cnt(rs,'Partially Submitted'), cnt(rs,'Not Submitted'), cnt(rs,'Pending / On Hold'),
    cnt(rs,'Needs Assistance'), subRate(rs), fmtMoney(sum(subd(rs), r=>r.billed))];

  const byBiller = tableHTML({ title:'Production by Biller', sub:'One row per claim worked. Grouped by team lead, highest volume first.',
    heads:[{l:'Biller', d:'The person who worked the claims.'}, {l:'Team Lead', d:'From the sheet, or its Validation Sheet when the row has none.'},
      {l:'Status', d:'Production, OJT, Training or PIP / QIP, from the sheet’s Validation Sheet.'},
      {l:'Days Worked', d:'Days with at least one claim worked.', num:1}, {l:'Claims', d:'Claims worked in the selection.', num:1},
      {l:'Claims / Day', d:'Claims ÷ days worked.', num:1}, {l:'Target', d:'Target / day × days worked. Blank until a target is set.', num:1},
      {l:'Attainment', d:`Claims as a share of target. Under ${PROD_MIN_ATTAIN}% is flagged.`, num:1}, ...statusHeads],
    rows: list.map(b=>({cls: b.attain!==null && b.attain<PROD_MIN_ATTAIN ? 'is-outlier' : '',
      cells:[b.name, b.tl||'—', b.status||'—', b.days, b.units, fmtN(b.perDay), b.target?fmtInt(b.target):'—', stamp(b.attain), ...statusCells(b.rs)]})),
    foot:['TOTAL', '', '', new Set(rows.map(r=>r.date)).size, rows.length, bd?fmtN(rows.length/bd):'—', tAll?fmtInt(tAll):'—',
      stamp(tAll?pctOf(uT,tAll):null), ...statusCells(rows)] });

  const st = [...groupBy(rows, r=>r.status||'(status not stated)')].sort((a,b)=>b[1].length-a[1].length);
  const byStatus = tableHTML({ title:'Claims by Status', sub:'The Status column as entered, with the group each status is counted under.',
    heads:[{l:'Status', d:'Status on the sheet.'}, {l:'Counted As', d:'The group used in the other tables.'}, {l:'Claims', d:'Claims with this status.', num:1},
      {l:'Share', d:'Share of all claims in the selection.', num:1}, {l:'Billed ($)', d:'Billed Amount of these claims.', num:1}],
    rows: st.map(([s,rs])=>[s, rs[0].group, rs.length, fmtP(pctOf(rs.length, rows.length)), fmtMoney(sum(rs,r=>r.billed))]),
    foot:['TOTAL', '', rows.length, '100.0%', fmtMoney(sum(rows,r=>r.billed))] });

  const cp = [...groupBy(rows, r=>(r.client||'(client not stated)')+'||'+(r.practice||'(practice not stated)'))]
    .map(([k,rs])=>{ const [c,p] = k.split('||'); return {c,p,rs}; }).sort((a,b)=>a.c.localeCompare(b.c) || b.rs.length-a.rs.length);
  const byClient = tableHTML({ title:'Production by Client & Practice', sub:'Client and Practice from the sheet. A blank Practice shows as “(practice not stated)”.',
    heads:[{l:'Client', d:'Client on the sheet.'}, {l:'Practice', d:'Practice on the sheet.'}, {l:'Claims', d:'Claims worked.', num:1},
      {l:'Billers', d:'Billers who worked it.', num:1}, ...statusHeads],
    rows: cp.map(x=>[x.c, x.p, x.rs.length, new Set(x.rs.map(r=>r.nameKey)).size, ...statusCells(x.rs)]),
    foot:['TOTAL', '', rows.length, new Set(rows.map(r=>r.nameKey)).size, ...statusCells(rows)] });

  const small = (field, title, lbl, desc) => {
    const g = [...groupBy(rows, r=>r[field]||'(not stated)')].sort((a,b)=>b[1].length-a[1].length);
    if(g.length===1 && g[0][0]==='(not stated)') return '';
    return tableHTML({ title, sub:`${lbl} column on the sheet.`,
      heads:[{l:lbl, d:desc}, {l:'Claims', d:'Claims worked.', num:1}, {l:'Submitted', d:'Submitted or partially submitted.', num:1},
        {l:'Submission %', d:'Submitted ÷ claims worked.', num:1}, {l:'Billed Submitted ($)', d:'Billed Amount of the claims submitted.', num:1}],
      rows: g.map(([k,rs])=>[k, rs.length, subd(rs).length, subRate(rs), fmtMoney(sum(subd(rs),r=>r.billed))]),
      foot:['TOTAL', rows.length, sAll, subRate(rows), fmtMoney(sum(subd(rows),r=>r.billed))] });
  };
  const byType = small('subType', 'By Submission Type', 'Submission Type', 'Regular, COB, IOF, AR, Lab Claim, MNCL and so on.');
  const bySource = small('claimSource', 'By Claim Source', 'Claim Source', 'Where the claim came from: Regular, AR, Posting Ledger and so on.');
  const byPayer = small('insurance', 'By Payer', 'Insurance', 'Insurance on the sheet.');

  const byDay = tableHTML({ title:'Production by Day', sub:'Every working day in the selection.',
    heads:[{l:'Date', d:'The Date Worked.'}, {l:'Billers', d:'Billers who worked that day.', num:1}, {l:'Claims', d:'Claims worked that day.', num:1},
      {l:'Submitted', d:'Submitted or partially submitted.', num:1}, {l:'Submission %', d:'Submitted ÷ claims worked.', num:1},
      {l:'Billed Submitted ($)', d:'Billed Amount of the claims submitted.', num:1}],
    rows: [...groupBy(rows, r=>r.date)].sort((a,b)=>a[0]<b[0]?-1:1).map(([d,rs])=>[fmtD(d), new Set(rs.map(r=>r.nameKey)).size, rs.length, subd(rs).length, subRate(rs), fmtMoney(sum(subd(rs),r=>r.billed))]),
    foot:['TOTAL', new Set(rows.map(r=>r.nameKey)).size, rows.length, sAll, subRate(rows), fmtMoney(sum(subd(rows),r=>r.billed))] });
  const byWeek = weekTable('sub', rows, [{l:'Submitted', d:'Submitted or partially submitted.', num:1}, {l:'Submission %', d:'Submitted ÷ claims worked.', num:1}],
    rs=>[subd(rs).length, subRate(rs)]);

  // claims allocated to one person and worked by another
  const alloc = rows.filter(r=>r.allocatedTo);
  const allocOther = alloc.filter(r=>normName(r.allocatedTo)!==r.nameKey);
  // the template fills Allocated To from Emp ID, so it only says something once
  // a claim was worked by someone other than the person it was allocated to
  const allocCard = allocOther.length ? tableHTML({ title:'Allocated vs Worked', sub:'Claims with an Allocated To name, and whether the person they were allocated to worked them.',
    heads:[{l:'Allocated To', d:'Allocated To on the sheet.'}, {l:'Claims Allocated', d:'Rows naming this person in Allocated To.', num:1},
      {l:'Worked by Them', d:'Of those, rows where Biller is the same person.', num:1}, {l:'Worked by Someone Else', d:'Rows where Biller is a different person.', num:1}],
    rows: [...groupBy(alloc, r=>normName(r.allocatedTo))].map(([k,rs])=>[rs[0].allocatedTo, rs.length, rs.filter(r=>r.nameKey===k).length, rs.filter(r=>r.nameKey!==k).length||''])
      .sort((a,b)=>b[1]-a[1]),
    foot:['TOTAL', alloc.length, alloc.length-allocOther.length, allocOther.length||''] }) : '';

  const open = rows.filter(r=>!isSubmitted(r));
  const openList = tableHTML({ title:'Claims Not Submitted', sub:'Every claim worked in the selection whose status isn’t Submitted or Partially Submitted, oldest first.',
    heads:[{l:'Date Worked', d:'The day it was worked.'}, {l:'Biller', d:'Who worked it.'}, {l:'Client', d:'Client.'}, {l:'Practice', d:'Practice.'},
      {l:'Claim No.', d:'Claim No, or UID when the claim number is blank.'}, {l:'Payer', d:'Insurance.'}, {l:'Status', d:'Status on the sheet.'},
      {l:'Counted As', d:'The status group.'}, {l:'Billed ($)', d:'Billed Amount.', num:1}, {l:'Comment', d:'Comment on the sheet.'}],
    rows: open.slice().sort((a,b)=>a.date<b.date?-1:a.date>b.date?1:0).map(r=>[fmtD(r.date), r.biller, r.client, r.practice, r.claimNo||r.uid, r.insurance,
      r.status||'(not stated)', r.group, fmtMoney(r.billed), r.comment]),
    empty:'Every claim in this selection was submitted.' });

  // checks on this selection
  const ck = [];
  const miss = f => rows.filter(r=>!r[f]).length;
  [['client','Client'],['practice','Practice'],['claimNo','Claim No'],['status','Status']].forEach(([f,l])=>{
    const n = miss(f); if(n) ck.push(`${n} row${n!==1?'s have':' has'} no ${l}.`); });
  const dup = [...groupBy(rows.filter(r=>r.claimNo), r=>r.nameKey+'|'+r.date+'|'+keyOf(r.client)+'|'+keyOf(r.claimNo))].filter(([,rs])=>rs.length>1);
  if(dup.length) ck.push(`${dup.length} claim${dup.length!==1?'s were':' was'} entered more than once by the same biller on the same day (${dup.slice(0,5).map(([,rs])=>`${rs[0].claimNo} by ${rs[0].biller} on ${fmtD(rs[0].date)}`).join('; ')}${dup.length>5?'; …':''}). Each row is counted.`);
  const other = rows.filter(r=>r.group==='Other');
  if(other.length) ck.push(`${other.length} row${other.length!==1?'s have':' has'} a status the dashboard doesn’t know (${[...new Set(other.map(r=>r.status))].slice(0,6).map(s=>'“'+s+'”').join(', ')}); counted as not submitted.`);

  return `${headerBands('sub', list, rows)}
    <div class="kpi-grid">
      ${kpi('Billers', list.length, '', `${new Set(rows.map(r=>r.date)).size} working day(s)`)}
      ${kpi('Claims Worked', fmtInt(rows.length), '', 'one per row of the sheet')}
      ${kpi('Claims / Biller / Day', bd?fmtN(rows.length/bd):'—', '', 'average over biller-days worked')}
      ${kpi('Attainment', tAll?fmtP(pctOf(uT,tAll)):'—', tAll?(pctOf(uT,tAll)>=PROD_MIN_ATTAIN?'good':'bad'):'', tAll?`minimum ${PROD_MIN_ATTAIN}%`:'set a daily target')}
      ${kpi('Submitted', fmtInt(sAll), '', `${cnt(rows,'Partially Submitted')} of them partially`)}
      ${kpi('Submission Rate', subRate(rows), '', 'submitted ÷ claims worked')}
      ${kpi('Not Submitted', fmtInt(open.length), open.length?'bad':'good', `${cnt(rows,'Pending / On Hold')} pending or on hold`)}
      ${kpi('Billed Submitted', kpiMoney(sum(subd(rows),r=>r.billed)), '', `${dollars(sum(subd(rows),r=>r.billed))} billed on submitted claims`)}
    </div>
    ${targetCardHTML('sub', list)}
    <div class="section-title">Production by Biller</div>
    ${byBiller}
    ${billerDayMatrix('sub', rows)}
    <div class="section-title">Production by Day &amp; Week</div>
    ${dailyChartHTML('sub', rows)}
    ${byDay}
    ${byWeek}
    <div class="section-title">Production by Client &amp; Practice</div>
    ${byClient}
    <div class="section-title">Status, Submission Type &amp; Source</div>
    ${byStatus}
    <div class="grid-2"><div>${byType}</div><div>${bySource}</div></div>
    ${byPayer}
    ${allocCard ? `<div class="section-title">Allocation</div>${allocCard}` : ''}
    <div class="section-title">Claims Not Submitted</div>
    ${openList}
    ${ck.length||STATE.sub.notes.length ? `<div class="section-title">Sheet Checks</div>${notesHTML('sub')}${ck.length?`<div class="card"><h3>In This Selection</h3><ul class="sc-list">${ck.map(x=>`<li>${esc(x)}</li>`).join('')}</ul></div>`:''}` : ''}`;
}

/* ---------- drawing ---------- */
function draw(){
  const area = document.getElementById('scopeArea');
  if(!area) return;
  const kind = SCOPE;
  document.querySelectorAll('.sc-zone').forEach(z=>z.hidden = z.dataset.kind!==kind);
  const st = STATE[kind];
  const badge = document.getElementById(kind+'Badge');
  if(badge){ badge.textContent = st.rows.length ? `${st.rows.length.toLocaleString('en-US')} ${META[kind].units}` : 'none'; badge.classList.toggle('on', !!st.rows.length); }
  const sub = document.getElementById(kind+'Sub');
  if(sub && st.file) sub.textContent = st.file;
  document.getElementById('scopeFilters').innerHTML = filtersHTML(kind);
  const c = document.getElementById('scopeContent');
  if(!st.rows.length){
    c.innerHTML = `<div class="empty-state"><h3>No ${esc(META[kind].label.toLowerCase())} sheet loaded yet</h3>
      <p>${kind==='pp'
        ? 'Upload the Payment Posting sheet above: one row per line posted, with Work Date, Biller, Insurance, Check number, Client Name, Location and the amounts.'
        : 'Upload the Submission production sheet above: one row per claim worked, with Client, Practice, Biller, Date Worked, Claim No and Status.'}</p></div>`;
    return;
  }
  c.innerHTML = kind==='pp' ? ppHTML() : subHTML();
  addRowNumbers(c); makeSortable(c); addCopyButtons(c); addHeadingHelp(c);
  c.classList.add('fade-in'); setTimeout(()=>c.classList.remove('fade-in'), 260);
}
function setScope(next){
  SCOPE = next;
  document.querySelectorAll('.scope-btn').forEach(b=>{ const on = b.dataset.scope===next;
    b.classList.toggle('active', on); b.setAttribute('aria-pressed', on); });
  const main = document.querySelector('.main');
  main.classList.toggle('scope-other', next!=='ar');
  document.getElementById('scopeArea').hidden = next==='ar';
  if(next==='ar'){ if(typeof render==='function') render(); return; }
  document.getElementById('pageTitle').textContent = META[next].title;
  document.getElementById('pageSub').textContent = META[next].sub;
  draw();
}

/* ---------- wiring ---------- */
document.getElementById('scopeBar').addEventListener('click', e=>{
  const b = e.target.closest('.scope-btn'); if(!b || b.dataset.scope===SCOPE) return;
  setScope(b.dataset.scope);
});
['pp','sub'].forEach(kind=>{
  const zone = document.getElementById(kind+'Zone'), input = document.getElementById(kind+'Input');
  input.addEventListener('change', e=>{ const f = e.target.files[0]; if(f) handleFile(kind, f); e.target.value=''; });
  zone.addEventListener('dragover', e=>{ e.preventDefault(); zone.classList.add('dragover'); });
  zone.addEventListener('dragleave', ()=>zone.classList.remove('dragover'));
  zone.addEventListener('drop', e=>{ e.preventDefault(); zone.classList.remove('dragover');
    const f = e.dataTransfer.files[0]; if(f) handleFile(kind, f); });
});
document.getElementById('scopeClear').addEventListener('click', ()=>{
  const st = STATE[SCOPE]; if(!st) return;
  st.rows = []; st.file = ''; st.notes = []; st.filters = blankFilters();
  const sub = document.getElementById(SCOPE+'Sub');
  if(sub) sub.textContent = sub.dataset.def;
  draw(); toast(`${META[SCOPE].label} sheet cleared.`);
});
document.addEventListener('click', e=>{
  if(!e.target.closest('[data-clear="all"]')) return;
  ['pp','sub'].forEach(k=>{ STATE[k].rows=[]; STATE[k].file=''; STATE[k].notes=[]; STATE[k].filters=blankFilters();
    const sub = document.getElementById(k+'Sub'); if(sub) sub.textContent = sub.dataset.def; });
});
document.getElementById('scopeArea').addEventListener('change', e=>{
  const kind = SCOPE; if(kind==='ar') return;
  const f = e.target.closest('[data-f]');
  if(f){ STATE[kind].filters[f.dataset.f] = f.value; draw(); return; }
  const t = e.target.closest('[data-tgt]');
  if(t){
    const v = Math.max(0, +t.value||0);
    if(t.dataset.tgt==='__def__') TARGETS[kind].def = v||null;
    else if(v) TARGETS[kind].by[t.dataset.tgt] = v; else delete TARGETS[kind].by[t.dataset.tgt];
    saveTargets(); draw();
  }
});
document.getElementById('scopeArea').addEventListener('click', e=>{
  if(!e.target.closest('[data-act="full"]')) return;
  STATE[SCOPE].filters.from = ''; STATE[SCOPE].filters.to = ''; draw();
});

/* the print button prints the scope that is showing */
function scopePrintHTML(){
  const kind = SCOPE, M = META[kind];
  const styles = [...document.querySelectorAll('style')].map(x=>x.textContent).join('\n');
  const clone = document.getElementById('scopeContent').cloneNode(true);
  clone.querySelectorAll('.no-print, .col-help').forEach(n=>n.remove());
  clone.querySelectorAll('input[data-tgt]').forEach(i=>{ const s = document.createElement('b'); s.textContent = i.value || i.placeholder; i.replaceWith(s); });
  const rows = filtered(kind), ds = rows.map(r=>r.date).sort();
  const range = ds.length ? (ds[0]===ds[ds.length-1] ? fmtD(ds[0]) : `${fmtD(ds[0])} – ${fmtD(ds[ds.length-1])}`) : '—';
  return `<!DOCTYPE html><html><head><meta charset="utf-8"><title>${esc(M.title)} — ${esc(range)}</title>
<style>${styles}
body{background:#fff;padding:18px 22px;}
.table-scroll{max-height:none !important;overflow:visible !important;}
table.data td,table.data th{white-space:normal;}
table.data td.num,table.data th.num{white-space:nowrap;}
.ph h1{font-size:19px;margin:0 0 3px;color:var(--navy);}
.ph{margin-bottom:14px;padding-bottom:9px;border-bottom:2px solid var(--navy);font-size:11.5px;color:var(--text-dim);}
.report-actions{margin:0 0 14px;display:flex;gap:8px;}
@page{size:A4 landscape;margin:9mm;}
@media print{ .report-actions{display:none !important;} body{padding:0;} }
</style></head><body>
<div class="ph"><h1>${esc(M.title)}</h1>
<div>Sasha&rsquo;s Third Eye &nbsp;·&nbsp; Purview Healthcare &nbsp;·&nbsp; Internal Reporting, ${esc(M.label)} scope</div>
<div><b>Date range:</b> ${esc(range)} &nbsp;·&nbsp; <b>Filters:</b> ${esc(filterText(kind))} &nbsp;·&nbsp; <b>Sheet:</b> ${esc(STATE[kind].file)} &nbsp;·&nbsp; Generated ${esc(new Date().toLocaleString())}</div></div>
<div class="report-actions"><button class="btn primary" onclick="window.print()">Print / Save as PDF</button></div>
${clone.innerHTML}
<scr`+`ipt>window.addEventListener('load',function(){setTimeout(function(){try{window.print();}catch(e){}},400);});<\/scr`+`ipt>
</body></html>`;
}
document.addEventListener('click', e=>{
  if(SCOPE==='ar' || !e.target.closest('#pdfBtn')) return;
  e.stopPropagation(); e.preventDefault();
  if(!STATE[SCOPE].rows.length){ toast(`Upload the ${META[SCOPE].label} sheet first.`, true); return; }
  const html = scopePrintHTML();
  let win = null; try{ win = window.open('', '_blank'); }catch(err){ win = null; }
  if(win && win.document){ win.document.open(); win.document.write(html); win.document.close();
    toast('Print dialog opening — choose "Save as PDF" as the destination.'); return; }
  const url = URL.createObjectURL(new Blob([html], {type:'text/html'}));
  const a = document.createElement('a'); a.href = url; a.download = META[SCOPE].title.toLowerCase().replace(/[^a-z0-9]+/g,'-')+'.html';
  document.body.appendChild(a); a.click(); a.remove(); setTimeout(()=>URL.revokeObjectURL(url), 30000);
  toast('Pop-ups are blocked here — report downloaded instead. Open it and press Ctrl/Cmd + P.', true);
}, true);

/* CSV downloads and their notes name the scope's report and filters */
window.scopeReportName = () => SCOPE==='ar' ? '' : META[SCOPE].title;
window.scopeFilterNote = () => SCOPE==='ar' ? '' : filterText(SCOPE);
window.S3_SCOPES = { setScope, state:STATE, get scope(){ return SCOPE; } };
})();
