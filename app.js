const COLS=['date','migraine','intensity','prostrating','start','start_time','duration','location','light','sound','nausea','vomiting','aura','duties','medication','trigger','could_not_do','ibs_pattern','abdominal_distress','bathroom_limited','ibs_note','bowel_movement','pain_with_bm'];
const KEY='symlog.entries.v1';
const $=s=>document.querySelector(s);
const load=()=>JSON.parse(localStorage.getItem(KEY)||'{}');
const store=o=>localStorage.setItem(KEY,JSON.stringify(o));
const today=()=>{const d=new Date();return new Date(d-d.getTimezoneOffset()*6e4).toISOString().slice(0,10)};
const toast=m=>{const t=$('#toast');t.textContent=m;t.style.display='block';setTimeout(()=>t.style.display='none',1800)};
const esc=s=>String(s??'').replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
let cur={}, curDate=today();

// ---- CSV ----
function toCSV(db){const q=v=>{v=v??'';return /[",\n\r]/.test(v)?'"'+String(v).replace(/"/g,'""')+'"':v};
 return COLS.join(',')+'\n'+Object.keys(db).sort().map(d=>COLS.map(c=>q(db[d][c])).join(',')).join('\n')+'\n';}
function parseCSVRows(t){const rows=[];let r=[],f='',q=false;for(let i=0;i<t.length;i++){const c=t[i];
 if(q){if(c=='"'){if(t[i+1]=='"'){f+='"';i++}else q=false}else f+=c}
 else if(c=='"')q=true;else if(c==','){r.push(f);f=''}else if(c=='\n'||c=='\r'){if(c=='\r'&&t[i+1]=='\n')i++;r.push(f);rows.push(r);r=[];f=''}else f+=c}
 if(f||r.length){r.push(f);rows.push(r)}
 return rows.filter(x=>x.length>1)}

// ---- Excel import ----
const p2=n=>String(n).padStart(2,'0');
function normDate(v){if(v===''||v==null)return '';if(typeof v==='number'){const d=new Date(Math.round((v-25569)*864e5));return d.getUTCFullYear()+'-'+p2(d.getUTCMonth()+1)+'-'+p2(d.getUTCDate())}
 v=String(v).trim();let m=v.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);if(m)return m[1]+'-'+p2(m[2])+'-'+p2(m[3]);
 m=v.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})/);if(m)return (m[3].length==2?'20'+m[3]:m[3])+'-'+p2(m[1])+'-'+p2(m[2]);return v}
function normTime(v){if(v===''||v==null)return '';if(typeof v==='number'){const t=Math.round((v%1)*1440);return p2(Math.floor(t/60))+':'+p2(t%60)}
 v=String(v).trim();const m=v.match(/^(\d{1,2}):(\d{2})(?::\d{2})?\s*([ap]m)?$/i);if(!m)return v;let h=+m[1];if(m[3]){h%=12;if(/p/i.test(m[3]))h+=12}return p2(h)+':'+m[2]}
function rowsToDB(rows){const h=rows.shift().map(x=>String(x??'').trim().replace(/^\ufeff/,'').toLowerCase());const out={};
 rows.forEach(r=>{const o={};COLS.forEach(c=>{const i=h.indexOf(c);let v=i<0?'':r[i];v=v==null?'':v;
  o[c]=c==='date'?normDate(v):c==='start_time'?normTime(v):typeof v==='number'?String(+v.toFixed(4)):String(v)});if(o.date)out[o.date]=o});return out}
function loadXLSX(){return window.XLSX?Promise.resolve():new Promise((ok,no)=>{const s=document.createElement('script');s.src='vendor/xlsx.full.min.js';s.onload=ok;s.onerror=()=>no(new Error('Excel reader not loaded'));document.head.appendChild(s)})}
async function readImport(f){if(/\.(xlsx|xls|xlsm|ods)$/i.test(f.name)||/sheet|excel/i.test(f.type)){await loadXLSX();
 const wb=XLSX.read(await f.arrayBuffer(),{type:'array'});const ws=wb.Sheets[wb.SheetNames[0]];
 return rowsToDB(XLSX.utils.sheet_to_json(ws,{header:1,raw:true,defval:'',blankrows:false}))}
 const t=await f.text();return rowsToDB(parseCSVRows(t))}

// ---- UI helpers ----
function chips(field,opts,multi){return `<div class="row">`+opts.map(([v,l])=>{const on=multi?(cur[field]||'').split(' ').includes(v):cur[field]===v;
 return `<button class="chip${on?' on':''}" data-f="${field}" data-v="${esc(v)}" data-m="${multi?1:''}">${l}</button>`}).join('')+`</div>`}
const yn=f=>chips(f,[['yes','Yes'],['no','No']]);
const L=(t,h)=>`<div class="lbl">${t}</div>${h}`;
function renderLog(){
 const mg=cur.migraine&&cur.migraine!=='no';const scale=[];for(let i=0;i<=10;i+=0.5)scale.push([String(i),String(i)]);
 const medOther=cur.medication&&!['none taken','acetaminophen'].includes(cur.medication);
 const trOther=cur.trigger&&!['unknown','anxiety','stress','sleep','food',''].includes(cur.trigger);
 $('#v-log').innerHTML=`<h1>Evening check-in</h1>
 <div class="card"><input type="date" id="dt" value="${curDate}" max="${today()}"><div class="mu" style="margin-top:6px">${load()[curDate]?'Saved entry: editing':'New entry'}</div></div>
 <button class="big quiet" id="quiet">☾ Quiet day (no migraine, normal IBS)</button>
 <h2>Migraine</h2><div class="card">${chips('migraine',[['no','No'],['yes','Yes'],['still','Still going']])}
 ${mg?`${L('Intensity (0–10)','<div class="row scale">'+chips('intensity',scale).slice(17))}
 ${L('Prostrating (had to lie down)?',yn('prostrating'))}
 ${L('Started',chips('start',[['morning','Morning'],['afternoon','Afternoon'],['evening','Evening'],['night','Night']]))}
 ${L('Start time (optional)',`<input type="time" id="st" value="${esc(cur.start_time)}">`)}
 ${L('Duration',chips('duration',[['under_1h','&lt;1 h'],['h1_4','1–4 h'],['h4_12','4–12 h'],['over_12h','&gt;12 h'],['still','Still going']]))}
 ${L('Location (tap all)',chips('location',[['forehead','Forehead'],['behind_eye','Behind eyes'],['temples','Temples'],['both_sides','Both sides'],['back_of_head','Back of head'],['neck','Neck']],1))}
 ${L('Light sensitivity',yn('light'))}${L('Sound sensitivity',yn('sound'))}${L('Nausea',yn('nausea'))}${L('Vomiting',yn('vomiting'))}
 ${L('Aura',chips('aura',[['yes','Yes'],['no','No'],['unsure','Unsure']]))}
 ${L('Could perform duties?',chips('duties',[['yes','Yes, could perform'],['no','No, could NOT perform']]))}
 ${L('Medication',chips('medication',[['none taken','None'],['acetaminophen','Acetaminophen'],[medOther?cur.medication:'__other','Other…']]))}
 ${medOther||cur.medication==='__other'?`<input id="medtx" placeholder="Medication name" value="${medOther?esc(cur.medication):''}">`:''}
 ${L('Trigger',chips('trigger',[['unknown','Unknown'],['anxiety','Anxiety'],['stress','Stress'],['sleep','Sleep'],['food','Food'],[trOther?cur.trigger:'__other','Other…']]))}
 ${trOther||cur.trigger==='__other'?`<input id="trtx" placeholder="Trigger" value="${trOther?esc(cur.trigger):''}">`:''}
 ${L('What I could not do (optional)',`<textarea id="cnd" rows="2">${esc(cur.could_not_do)}</textarea>`)}`:''}</div>
 <h2>IBS</h2><div class="card">${L('Pattern',chips('ibs_pattern',[['normal','Normal'],['constipation','Constipation'],['diarrhea','Diarrhea'],['mixed','Mixed']]))}
 ${L('Abdominal distress',chips('abdominal_distress',[['none','None'],['episodic','Episodic'],['most','Most of day']]))}
 ${L('Bowel movement today?',yn('bowel_movement'))}
 ${L('Abdominal pain related to a bowel movement?',yn('pain_with_bm'))}
 ${L('Bathroom access limited duties?',yn('bathroom_limited'))}
 ${L('IBS note (optional)',`<textarea id="ibn" rows="2">${esc(cur.ibs_note)}</textarea>`)}</div>
 <button class="big save" id="save">Save ${curDate}</button>`;
}
function readText(){const g=id=>document.getElementById(id);
 if(g('st'))cur.start_time=g('st').value;if(g('cnd'))cur.could_not_do=g('cnd').value;if(g('ibn'))cur.ibs_note=g('ibn').value;
 if(g('medtx'))cur.medication=g('medtx').value||'__other';if(g('trtx'))cur.trigger=g('trtx').value||'__other';}
function openDate(d){curDate=d;cur={...(load()[d]||{date:d})};renderLog();scrollTo(0,0)}
function save(){readText();const o={};COLS.forEach(c=>o[c]=cur[c]??'');o.date=curDate;
 ['medication','trigger'].forEach(k=>{if(o[k]==='__other')o[k]=''});
 if(o.migraine==='no'||!o.migraine)COLS.slice(2,17).forEach(c=>o[c]='');
 const db=load();db[curDate]=o;store(db);toast('Saved '+curDate);renderLog();}
document.addEventListener('click',e=>{const b=e.target.closest('button');if(!b)return;
 if(b.dataset.f){readText();const f=b.dataset.f,v=b.dataset.v;
  if(b.dataset.m){const s=new Set((cur[f]||'').split(' ').filter(Boolean));s.has(v)?s.delete(v):s.add(v);cur[f]=[...s].join(' ')}
  else cur[f]=cur[f]===v?'':v;const y=scrollY;renderLog();scrollTo(0,y);return}
 if(b.id==='quiet'){cur={date:curDate,migraine:'no',ibs_pattern:'normal',abdominal_distress:'none',bathroom_limited:'no',bowel_movement:'',pain_with_bm:'no'};save();return}
 if(b.id==='save')return save();
 if(b.dataset.v&&b.parentNode.tagName==='NAV')return show(b.dataset.v);
 if(b.dataset.d){openDate(b.dataset.d);show('log')}
 if(b.id==='exp')exportCSV();if(b.id==='notif')enableNotif();if(b.id==='ics')downloadICS();if(b.id==='wipe'&&confirm('Delete ALL entries on this phone?')){store({});toast('Cleared');show('set')}
});
document.addEventListener('change',e=>{if(e.target.id==='dt'&&e.target.value)openDate(e.target.value);
 if(e.target.id==='imp'&&e.target.files[0])readImport(e.target.files[0]).then(n=>{const db={...load(),...n};store(db);toast('Imported '+Object.keys(n).length+' days');show('hist')})});
function show(v){document.querySelectorAll('section').forEach(s=>s.hidden=s.id!=='v-'+v);document.querySelectorAll('nav button').forEach(b=>b.classList.toggle('on',b.dataset.v===v));
 if(v==='hist')renderHist();if(v==='set')renderSet();if(v==='log')renderLog();}
function renderHist(){const db=load(),ds=Object.keys(db).sort().reverse(),m={};
 ds.forEach(d=>{const e=db[d],k=d.slice(0,7);m[k]=m[k]||[0,0,0,0,0];const mg=e.migraine&&e.migraine!=='no';
  if(mg)m[k][0]++;if(mg&&e.prostrating==='yes')m[k][1]++;if(mg&&e.duties==='no')m[k][2]++;if(e.ibs_pattern&&e.ibs_pattern!=='normal')m[k][3]++;if(e.pain_with_bm==='yes')m[k][4]++});
 $('#v-hist').innerHTML=`<h1>History</h1><div class="card"><table><tr><th>Month</th><th>Migraine days</th><th>Prostrating</th><th>Couldn't do duties</th><th>IBS days</th><th>BM pain days</th></tr>
 ${Object.keys(m).sort().reverse().map(k=>`<tr><td>${k}</td>${m[k].map(x=>`<td>${x}</td>`).join('')}</tr>`).join('')}</table></div>
 <div class="card hist">${ds.length?ds.map(d=>{const e=db[d],mg=e.migraine&&e.migraine!=='no';
 return `<div><span>${d}<br><span class="mu">${mg?`Migraine ${e.intensity||'?'}/10${e.prostrating==='yes'?' · prostrating':''}${e.duties==='no'?' · could NOT perform duties':''}`:'No migraine'} · IBS ${e.ibs_pattern||'–'}</span></span><button data-d="${d}">Edit</button></div>`}).join(''):'<p class="mu">No entries yet. Import your CSV in Settings.</p>'}</div>`;}
function renderSet(){$('#v-set').innerHTML=`<h1>Settings</h1>
 <div class="card"><b>Export</b><p class="mu">Downloads every entry as a CSV (same columns as your VA log).</p><button class="big" id="exp">Export CSV</button></div>
 <div class="card"><b>Import</b><p class="mu">Load a CSV or Excel file (e.g. your Sept 19 – Oct 20 log). Days in the file replace the same days here.</p><input type="file" id="imp" accept=".csv,.xlsx,.xls,text/csv,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel"></div>
 <div class="card"><b>9 PM reminder</b><p class="mu">Status: ${'Notification' in window?Notification.permission:'not supported here'}. Phone browsers can't reliably fire a timed alert while the app is closed, so also add the calendar reminder.</p>
 <button class="big" id="notif">Enable app notifications</button><div style="height:8px"></div><button class="big" id="ics">Add nightly 9 PM calendar reminder</button></div>
 <div class="card"><b>Data</b><p class="mu">${Object.keys(load()).length} days stored on this phone only. Nothing is sent anywhere. Export regularly as a backup.</p><button id="wipe">Delete all data</button></div>`;}
function exportCSV(){const b=new Blob([toCSV(load())],{type:'text/csv'});const a=document.createElement('a');a.href=URL.createObjectURL(b);a.download=`migraine_ibs_log_${today()}.csv`;a.click();}
function downloadICS(){const s=today().replace(/-/g,'');const url=location.href.split('#')[0];
 const t=`BEGIN:VCALENDAR\r\nVERSION:2.0\r\nPRODID:-//SymptomLog//EN\r\nBEGIN:VEVENT\r\nUID:symlog-nightly@local\r\nDTSTAMP:${s}T000000Z\r\nDTSTART;TZID=America/Chicago:${s}T210000\r\nDURATION:PT5M\r\nRRULE:FREQ=DAILY\r\nSUMMARY:Symptom Log check-in\r\nDESCRIPTION:Open ${url}\r\nURL:${url}\r\nBEGIN:VALARM\r\nACTION:DISPLAY\r\nTRIGGER:PT0M\r\nDESCRIPTION:Symptom Log check-in\r\nEND:VALARM\r\nEND:VEVENT\r\nEND:VCALENDAR\r\n`;
 const a=document.createElement('a');a.href=URL.createObjectURL(new Blob([t],{type:'text/calendar'}));a.download='symptom_log_9pm.ics';a.click();}
async function enableNotif(){if(!('Notification' in window))return toast('On iPhone: add to Home Screen first');
 const p=await Notification.requestPermission();if(p==='granted'){const r=await navigator.serviceWorker.ready;
  try{await r.periodicSync.register('nightly',{minInterval:12*3600e3})}catch(e){}schedule();toast('Notifications on')}renderSet();}
function schedule(){if(!('Notification' in window)||Notification.permission!=='granted')return;const n=new Date(),t=new Date();t.setHours(21,0,0,0);if(t<=n)t.setDate(t.getDate()+1);
 setTimeout(async()=>{if(!load()[today()])(await navigator.serviceWorker.ready).showNotification('Symptom Log',{body:'Time for your evening check-in.',icon:'icons/icon-192.png',tag:'nightly'});schedule()},t-n);}
if('serviceWorker' in navigator)navigator.serviceWorker.register('sw.js');
openDate(today());schedule();
window.__t={toCSV,readImport,load,store};
