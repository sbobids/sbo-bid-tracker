const $=id=>document.getElementById(id);
const defaultChoices=[49,89,66,31,85,75,86,88,76,29,10,19,78,36,57,47,82,8,12,67,7,80,37,72,64,32,53,69,42,1,56,20,24,3,2];
let state={people:[],capacities:{},details:{},pageMap:{},excluded:[],awards:[],selectedName:'',overrides:{}};
$('choices').value='';
const DB_NAME='sbo-bid-tracker-local-v1', DB_STORE='saved';
function openDB(){return new Promise((resolve,reject)=>{let req=indexedDB.open(DB_NAME,1);req.onupgradeneeded=()=>req.result.createObjectStore(DB_STORE);req.onsuccess=()=>resolve(req.result);req.onerror=()=>reject(req.error)})}
async function dbPut(key,value){let db=await openDB();try{await new Promise((resolve,reject)=>{let tx=db.transaction(DB_STORE,'readwrite');tx.objectStore(DB_STORE).put(value,key);tx.oncomplete=resolve;tx.onerror=()=>reject(tx.error)})}finally{db.close()}}
async function dbGet(key){let db=await openDB();try{return await new Promise((resolve,reject)=>{let req=db.transaction(DB_STORE).objectStore(DB_STORE).get(key);req.onsuccess=()=>resolve(req.result);req.onerror=()=>reject(req.error)})}finally{db.close()}}
async function dbClear(){let db=await openDB();try{await new Promise((resolve,reject)=>{let tx=db.transaction(DB_STORE,'readwrite');tx.objectStore(DB_STORE).clear();tx.oncomplete=resolve;tx.onerror=()=>reject(tx.error)})}finally{db.close()}}
function populatePilots(){const sel=$('pilotSelect');sel.replaceChildren(new Option('Choose your name…',''),...state.people.map(p=>new Option(p.name,p.name)));const previous=localStorage.getItem('sbo-selected-pilot')||'';const name=state.people.some(p=>p.name===previous)?previous:'';sel.value=name;selectPilot(name)}
async function restoreSaved(){try{const saved=await dbGet('parsed');if(saved&&Array.isArray(saved.people)&&saved.people.length){state={...state,...saved,awards:[]};populatePilots();simulate();$('lastUpdated').textContent=saved.updatedAt?'Last imported: '+new Date(saved.updatedAt).toLocaleString():'Saved data restored (import date unknown)';$('importStatus').textContent='Saved bidding data restored. To update bids, select only the View All FO Bids file.'}else{$('importStatus').textContent='No saved bidding data on this device yet.'}}catch(e){$('importStatus').textContent='Local storage unavailable: '+e.message}}

document.querySelectorAll('.tab').forEach(b=>b.onclick=()=>{document.querySelectorAll('.tab').forEach(x=>x.classList.toggle('active',x===b));document.querySelectorAll('.page').forEach(x=>x.classList.toggle('hidden',x.id!==b.dataset.tab))});
$('openFiles').onclick=()=>{document.querySelectorAll('.tab').forEach(x=>x.classList.remove('active'));document.querySelectorAll('.page').forEach(x=>x.classList.toggle('hidden',x.id!=='files'));window.scrollTo(0,0)};
$('backHome').onclick=()=>document.querySelector('[data-tab="home"]').click();
$('homeRefresh').onclick=()=>document.querySelector('[data-tab="refresh"]').click();

function norm(s){return String(s||'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase().replace(/[^a-z ]/g,' ').trim().split(/\s+/).filter(Boolean).sort().join(' ')}
function match(a,b){let x=norm(a).split(' '),y=norm(b).split(' ');return x.length>=2&&y.length>=2&&x.filter(t=>y.includes(t)).length>=2}
function parseNums(s){return [...new Set((String(s||'').match(/(?:F\s*)?\d+/gi)||[]).map(x=>Number(x.replace(/[^0-9]/g,''))).filter(x=>x>0&&x<500))]}
function rows(wb,name){let sh=wb.Sheets[name]||wb.Sheets[wb.SheetNames[0]];return XLSX.utils.sheet_to_json(sh,{header:1,defval:''})}
async function excel(file){if(!window.XLSX)throw Error('Excel library unavailable. Connect to the internet and reload.');return XLSX.read(await file.arrayBuffer(),{type:'array'})}
function parseRoster(wb){let r=rows(wb,'777');if(!wb.Sheets['777'])throw Error('777 worksheet not found in monthly Bid List.');let header=r.slice(0,30).flatMap((row,i)=>row.map((v,j)=>({v:String(v).toUpperCase().trim(),i,j})));let exclusions=[];
 for(let [label,reason] of [['FO CQ SCHEDULED','FO CQ'],['VAC FO','FO Vacation']]){let h=header.find(x=>x.v===label);if(!h)throw Error('Missing '+label+' column in 777 worksheet');for(let i=h.i+1;i<r.length;i++){let val=r[i]?.[h.j];if(typeof val!=='string'||!val.trim())break;exclusions.push({name:val.trim(),reason})}}
 return exclusions}
function parseBidRows(r){ // The company export uses repeated Name/Choice/Line rows; preserve displayed order.
 let hIndex=r.findIndex(row=>row.some(v=>/^name$/i.test(String(v).trim()))&&row.some(v=>/^line$/i.test(String(v).trim())));
 if(hIndex<0)throw Error('Could not find Name / Choice / Line headings in exported bids.');
 let h=r[hIndex].map(x=>String(x).trim().toLowerCase()),nameCol=h.indexOf('name'),choiceCol=h.indexOf('choice'),lineCol=h.indexOf('line');
 let people=[],current=null;
 for(let i=hIndex+1;i<r.length;i++){
   let row=r[i]||[],name=String(row[nameCol]||'').trim(),choice=String(row[choiceCol]||'').trim(),line=String(row[lineCol]||'').trim();
   if(/^name$/i.test(name)||/^first officer bids/i.test(name))continue;
   if(name&&/[a-z]/i.test(name)&&!/^page\s+\d+|^total/i.test(name)){
     if(!current||!match(current.name,name)){current={name,bids:[],raw:[]};people.push(current)}
   }
   if(!current)continue;
   let ln=line.match(/\bF?\s*(\d{1,3})\b/i);if(ln&&Number(ln[1])>0){let n=Number(ln[1]);let rank=Number(choice.replace(/\D/g,''));current.raw.push({n,rank:Number.isFinite(rank)?rank:9999,seq:i})}
 }
 for(let p of people){p.raw.sort((a,b)=>a.rank-b.rank||a.seq-b.seq);p.bids=[...new Set(p.raw.map(x=>x.n))];delete p.raw}
 return people.filter(x=>x.name.length>3)}
async function parsePdf(file){
 if(!window.pdfjsLib)throw Error('PDF library unavailable. Connect to the internet and reload.');
 pdfjsLib.GlobalWorkerOptions.workerSrc='https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
 const doc=await pdfjsLib.getDocument({data:await file.arrayBuffer()}).promise;
 const capacities={},details={},pageMap={};
 for(let i=1;i<=doc.numPages;i++){
  const page=await doc.getPage(i),content=await page.getTextContent(),t=content.items.map(x=>x.str).join(' ');
  const id=t.match(/\bF\s*(\d{1,3})\s*[-–]\s*\d+\s*\/\s*(\d+)\b/i);
  if(!id)continue;
  const n=Number(id[1]);capacities[n]=Number(id[2]);pageMap[n]=i;
  const credit=t.match(/Credit\s*time\s*:\s*(\d{1,3}:\d{2})/i),days=t.match(/Duty\s*days\s*:\s*(\d+)/i);
  details[n]={credit:credit?.[1]||'—',duties:days?.[1]||'—'};
 }
 return {capacities,details,pageMap};
}
function simulate(){let left={...state.capacities},awards=[];for(let i=0;i<state.people.length;i++){let p=state.people[i],exc=state.excluded.find(e=>match(e.name,p.name)),bids=state.overrides[p.name]||p.bids,got=null,rank=null;
 if(!exc)for(let j=0;j<bids.length;j++){let n=bids[j];if(left[n]>0){got=n;rank=j+1;left[n]--;break}}
 awards.push({name:p.name,seniority:i+1,line:got,rank,blank:p.bids.length===0,excluded:exc?.reason||'',bidCount:p.bids.length})}
 state.awards=awards;render()}
function render(){let me=state.awards.find(p=>p.name===state.selectedName);$('award').textContent=me?.line?'F'+me.line:'—';$('status').textContent=me?.excluded?'Excluded: '+me.excluded:me?.line?'Provisional award':'No projected line';$('credit').textContent=state.details[me?.line]?.credit||'—';$('duties').textContent=state.details[me?.line]?.duties||'—';$('rank').textContent=me?.rank?'#'+me.rank:'—';let senior=me?state.awards.slice(0,me.seniority-1):[];$('blanks').textContent=me?senior.filter(x=>x.blank&&!x.excluded).length:'—';$('eligible').textContent=state.awards.filter(x=>!x.excluded).length||'—';$('viewProjectedLine').disabled=!me?.line;
 $('whoAmIFlyingWith').disabled=!me?.line;
 if(!$('lineDetail').classList.contains('hidden'))showProjectedLine();
 if(!$('flyingWith').classList.contains('hidden'))showFlyingWith();$('notes').textContent=me?`Seniority position ${me.seniority} in exported bid report. ${senior.filter(x=>x.excluded).length} senior CQ/Vacation exclusions. ${senior.filter(x=>x.blank&&!x.excluded).length} eligible senior blank bids. Projection only; company award rules may differ.`:'Import the three files and select a pilot to calculate a projected line.';renderPilots()}
function renderPilots(){let q=$('search').value.trim().toLowerCase(),list=state.awards.filter(p=>!q||p.name.toLowerCase().includes(q)||('f'+p.line).toLowerCase()===q);$('pilotList').replaceChildren(...list.map(p=>{let d=document.createElement('div');d.className='pilot';let a=document.createElement('div'),b=document.createElement('b'),sm=document.createElement('small');a.textContent=`${p.seniority}. ${p.name}`;sm.textContent=p.excluded?p.excluded:p.blank?'No bid submitted':p.line?`Choice #${p.rank}`:'No available submitted choice';a.append(sm);b.textContent=p.excluded?'Excluded':p.line?'F'+p.line:'—';d.append(a,b);return d}))}
$('search').oninput=renderPilots;
function selectPilot(name){
 state.selectedName=name;
 if(name){try{localStorage.setItem('sbo-selected-pilot',name)}catch(e){console.warn('Name could not be remembered',e)}}
 const pilot=state.people.find(p=>p.name===name);
 $('choices').value=(pilot?(state.overrides[name]||pilot.bids):[]).map(n=>'F'+n).join('\n');
 $('prefMsg').textContent=pilot?'Imported bids shown. Changes are what-if only.':'';
 if(state.people.length)render();
}
$('pilotSelect').onchange=e=>selectPilot(e.target.value);
$('saveChoices').onclick=()=>{
 if(!state.selectedName){$('prefMsg').textContent='Select a pilot on Home first.';return}
 const n=parseNums($('choices').value);
 state.overrides[state.selectedName]=n;
 $('prefMsg').textContent=`Applied ${n.length} what-if choices for ${state.selectedName}. This does not submit bids.`;
 simulate();
};
$('analyze').onclick=async()=>{
 let status=$('importStatus');
 try{
   const pf=$('linesFile').files[0],rf=$('rosterFile').files[0],bf=$('bidsFile').files[0];
   if(!bf)throw Error('Select the updated View All FO Bids Excel report.');
   // Keep the existing monthly schedule and CQ/Vacation exclusions unless replaced.
   const saved=await dbGet('parsed');
   const base=saved&&saved.capacities&&saved.details&&Array.isArray(saved.excluded)?saved:null;
   if(!base&&(!pf||!rf))throw Error('First import on this device: select all three files. Future updates need only the FO bids export.');
   if((pf&&!rf)||(!pf&&rf))throw Error('For a new month, replace BOTH the 777 Lines PDF and monthly Bid List workbook.');
   status.textContent='Reading updated FO bids…';
   const people=parseBidRows(rows(await excel(bf),'Sheet1'));
   if(people.length<30)throw Error(`Only ${people.length} pilots identified in the bid report. Please check the export format.`);
   let capacities,details,excluded,pageMap;
   if(pf&&rf){
     status.textContent='Reading new monthly lines and exclusions…';
     const pdf=await parsePdf(pf);
     capacities=pdf.capacities;details=pdf.details;pageMap=pdf.pageMap;
     excluded=parseRoster(await excel(rf));
     if(Object.keys(capacities).length<40)throw Error(`Only ${Object.keys(capacities).length} lines identified in PDF.`);
   }else{
     capacities={...base.capacities};details={...base.details};pageMap={...(base.pageMap||{})};excluded=base.excluded;
   }
   for(const m of $('capacities').value.matchAll(/F?\\s*(\\d+)\\s*=\\s*(\\d+)/gi))capacities[Number(m[1])]=Number(m[2]);
   const updatedAt=new Date().toISOString();
   const next={people,capacities,details,pageMap,excluded,overrides:{},updatedAt};
   // Save first, then update the live display, so a failed write cannot masquerade as success.
   if(pf)await dbPut('linesPdf',pf);
   await dbPut('parsed',next);
   state={...state,...next};
   populatePilots();simulate();
   $('lastUpdated').textContent='Last imported: '+new Date(updatedAt).toLocaleString();
   status.textContent=`Saved locally: ${people.length} pilots, ${Object.keys(capacities).length} lines, ${excluded.length} CQ/Vacation entries. ${pf?'All monthly files replaced.':'Only FO bids updated; saved lines and exclusions reused.'}`;
   $('bidsFile').value='';$('linesFile').value='';$('rosterFile').value='';
   document.querySelector('[data-tab="home"]').click();
 }catch(e){status.textContent='Import error: '+e.message;console.error(e)}
};

$('clearSaved').onclick=async()=>{if(!confirm('Remove saved bidding data and selected pilot from this device?'))return;await dbClear();localStorage.removeItem('sbo-selected-pilot');location.reload()};
restoreSaved();

// iPhone-friendly refresh flow: open the authenticated portal, then select the
// downloaded file using the operating system's normal file picker.
// Browser security prevents silent reading of Safari downloads or portal cookies.
$('pickLatest').onclick=()=>{$('bidsFile').click()};
$('bidsFile').addEventListener('change',()=>{
 const f=$('bidsFile').files[0];
 $('refreshFileName').textContent=f?`Selected: ${f.name}`:'No report selected';
 $('refreshUpdate').disabled=!f;
});
$('refreshUpdate').onclick=async()=>{
 if(!$('bidsFile').files[0])return;
 $('refreshStatus').textContent='Importing latest FO bids…';
 // Reuse the exact same tested parser, validation and IndexedDB persistence.
 await $('analyze').onclick();
 const msg=$('importStatus').textContent;
 $('refreshStatus').textContent=msg;
 if(msg.startsWith('Saved locally:')){
   $('refreshFileName').textContent='No report selected';
   $('refreshUpdate').disabled=true;
 }else{
   document.querySelector('[data-tab="refresh"]').click();
 }
};

// Render the original Jeppesen line chart directly from the imported PDF.
// No transcription or invented day-by-day data. All PDF data stays on this device.
let lineViewToken=0;
async function showProjectedLine(){
 const me=state.awards.find(p=>p.name===state.selectedName);
 if(!me?.line){$('lineDetail').classList.add('hidden');return}
 const token=++lineViewToken;
 $('lineDetailTitle').textContent='Projected line F'+me.line;
 $('lineDetail').classList.remove('hidden');
 $('lineDetailMessage').textContent='Loading original line chart…';
 $('lineChart').replaceChildren();
 $('lineZoom').classList.add('hidden');
 try{
  const pageNumber=state.pageMap?.[me.line];
  const pdfBlob=await dbGet('linesPdf');
  if(!pageNumber||!pdfBlob){
   $('lineDetailMessage').textContent='To view the original line chart, reimport this month’s 777 Lines PDF and Bid List together with the FO bids report. Your saved projection remains available.';
   return;
  }
  if(!window.pdfjsLib)throw Error('PDF viewer unavailable. Connect to the internet and reload.');
  pdfjsLib.GlobalWorkerOptions.workerSrc='https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
  const pdf=await pdfjsLib.getDocument({data:await pdfBlob.arrayBuffer()}).promise;
  const page=await pdf.getPage(pageNumber);
  // PDF is landscape. The schedule occupies its upper half; crop blank lower page.
  const viewport=page.getViewport({scale:2.2});
  const full=document.createElement('canvas');
  full.width=Math.ceil(viewport.width);full.height=Math.ceil(viewport.height);
  await page.render({canvasContext:full.getContext('2d'),viewport}).promise;
  if(token!==lineViewToken)return;
  const left=Math.floor(full.width*.025),top=Math.floor(full.height*.12);
  const width=Math.floor(full.width*.93),height=Math.floor(full.height*.43);
  const cropped=document.createElement('canvas');
  cropped.width=width;cropped.height=height;
  cropped.getContext('2d').drawImage(full,left,top,width,height,0,0,width,height);
  cropped.setAttribute('role','img');
  cropped.setAttribute('aria-label','Original Kalitta PDF chart for line F'+me.line);
  $('lineChart').append(cropped);
  $('lineZoom').classList.remove('hidden');
  $('lineDetailMessage').textContent='Original chart from your imported PDF. Tap Enlarge for a clearer view, then scroll sideways.';
 }catch(e){
  if(token===lineViewToken)$('lineDetailMessage').textContent='Unable to display the PDF chart: '+e.message;
 }
}

function showFlyingWith(){
 const me=state.awards.find(p=>p.name===state.selectedName);
 if(!me?.line){$('flyingWith').classList.add('hidden');return}
 $('flyingWithTitle').textContent=`Who am I flying with on F${me.line}?`;
 $('flyingWith').classList.remove('hidden');
 const partners=state.awards.filter(p=>p.line===me.line&&p.name!==me.name);
 $('flyingWithList').replaceChildren(...partners.map(p=>{
   const row=document.createElement('div');
   row.className='flying-with-pilot';
   const info=document.createElement('div');
   const name=document.createElement('strong');
   const meta=document.createElement('small');
   name.textContent=p.name;
   meta.textContent=`Seniority #${p.seniority}${p.rank?` · Choice #${p.rank}`:''}`;
   info.append(name,meta);
   const line=document.createElement('span');
   line.className='flying-with-line';
   line.textContent='F'+p.line;
   row.append(info,line);
   return row;
 }));
 $('flyingWithMessage').textContent=partners.length
   ? `${partners.length} other first officer${partners.length===1?'':'s'} ${partners.length===1?'is':'are'} currently projected to hold this line.`
   : 'No other first officer is currently projected to hold this line.';
}
$('whoAmIFlyingWith').onclick=()=>{
 if($('flyingWith').classList.contains('hidden'))showFlyingWith();
 else $('flyingWith').classList.add('hidden');
};
$('closeFlyingWith').onclick=()=>$('flyingWith').classList.add('hidden');

$('viewProjectedLine').onclick=()=>{
 if($('lineDetail').classList.contains('hidden'))showProjectedLine();
 else{$('lineDetail').classList.add('hidden');lineViewToken++;}
};
$('closeLineDetail').onclick=()=>{$('lineDetail').classList.add('hidden');lineViewToken++};
$('lineZoom').onclick=()=>{
 const zoomed=$('lineChart').classList.toggle('zoomed');
 $('lineZoom').textContent=zoomed?'Fit to screen':'Enlarge / scroll';
};

// All dates are UTC. At/after 21:00Z on the 16th, count down to next month.
function nextBidDeadline(now){
  let year=now.getUTCFullYear(), month=now.getUTCMonth();
  let deadline=Date.UTC(year,month,16,21,0,0);
  if(now.getTime()>=deadline) deadline=Date.UTC(year,month+1,16,21,0,0);
  return deadline;
}
function updateBidCountdown(){
  const now=new Date(), deadline=nextBidDeadline(now);
  let remaining=Math.max(0,Math.ceil((deadline-now.getTime())/1000));
  const days=Math.floor(remaining/86400); remaining%=86400;
  const hours=Math.floor(remaining/3600); remaining%=3600;
  const minutes=Math.floor(remaining/60);
  $('bidDays').textContent=String(days).padStart(2,'0');
  $('bidHours').textContent=String(hours).padStart(2,'0');
  $('bidMinutes').textContent=String(minutes).padStart(2,'0');
  $('bidDeadline').textContent='Closes '+new Intl.DateTimeFormat('en-GB',{timeZone:'UTC',day:'numeric',month:'long',year:'numeric'}).format(new Date(deadline))+' at 21:00 UTC';
}
updateBidCountdown();
setInterval(updateBidCountdown,10000);
document.addEventListener('visibilitychange',()=>{if(!document.hidden)updateBidCountdown()});
const $=id=>document.getElementById(id);
const defaultChoices=[49,89,66,31,85,75,86,88,76,29,10,19,78,36,57,47,82,8,12,67,7,80,37,72,64,32,53,69,42,1,56,20,24,3,2];
let state={people:[],capacities:{},details:{},pageMap:{},excluded:[],awards:[],selectedName:'',overrides:{}};
$('choices').value='';
const DB_NAME='sbo-bid-tracker-local-v1', DB_STORE='saved';
function openDB(){return new Promise((resolve,reject)=>{let req=indexedDB.open(DB_NAME,1);req.onupgradeneeded=()=>req.result.createObjectStore(DB_STORE);req.onsuccess=()=>resolve(req.result);req.onerror=()=>reject(req.error)})}
async function dbPut(key,value){let db=await openDB();try{await new Promise((resolve,reject)=>{let tx=db.transaction(DB_STORE,'readwrite');tx.objectStore(DB_STORE).put(value,key);tx.oncomplete=resolve;tx.onerror=()=>reject(tx.error)})}finally{db.close()}}
async function dbGet(key){let db=await openDB();try{return await new Promise((resolve,reject)=>{let req=db.transaction(DB_STORE).objectStore(DB_STORE).get(key);req.onsuccess=()=>resolve(req.result);req.onerror=()=>reject(req.error)})}finally{db.close()}}
async function dbClear(){let db=await openDB();try{await new Promise((resolve,reject)=>{let tx=db.transaction(DB_STORE,'readwrite');tx.objectStore(DB_STORE).clear();tx.oncomplete=resolve;tx.onerror=()=>reject(tx.error)})}finally{db.close()}}
function populatePilots(){const sel=$('pilotSelect');sel.replaceChildren(new Option('Choose your name…',''),...state.people.map(p=>new Option(p.name,p.name)));const previous=localStorage.getItem('sbo-selected-pilot')||'';const name=state.people.some(p=>p.name===previous)?previous:'';sel.value=name;selectPilot(name)}
async function restoreSaved(){try{const saved=await dbGet('parsed');if(saved&&Array.isArray(saved.people)&&saved.people.length){state={...state,...saved,awards:[]};populatePilots();simulate();$('lastUpdated').textContent=saved.updatedAt?'Last imported: '+new Date(saved.updatedAt).toLocaleString():'Saved data restored (import date unknown)';$('importStatus').textContent='Saved bidding data restored. To update bids, select only the View All FO Bids file.'}else{$('importStatus').textContent='No saved bidding data on this device yet.'}}catch(e){$('importStatus').textContent='Local storage unavailable: '+e.message}}

document.querySelectorAll('.tab').forEach(b=>b.onclick=()=>{document.querySelectorAll('.tab').forEach(x=>x.classList.toggle('active',x===b));document.querySelectorAll('.page').forEach(x=>x.classList.toggle('hidden',x.id!==b.dataset.tab))});
$('openFiles').onclick=()=>{document.querySelectorAll('.tab').forEach(x=>x.classList.remove('active'));document.querySelectorAll('.page').forEach(x=>x.classList.toggle('hidden',x.id!=='files'));window.scrollTo(0,0)};
$('backHome').onclick=()=>document.querySelector('[data-tab="home"]').click();
$('homeRefresh').onclick=()=>document.querySelector('[data-tab="refresh"]').click();

function norm(s){return String(s||'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase().replace(/[^a-z ]/g,' ').trim().split(/\s+/).filter(Boolean).sort().join(' ')}
function match(a,b){let x=norm(a).split(' '),y=norm(b).split(' ');return x.length>=2&&y.length>=2&&x.filter(t=>y.includes(t)).length>=2}
function parseNums(s){return [...new Set((String(s||'').match(/(?:F\s*)?\d+/gi)||[]).map(x=>Number(x.replace(/[^0-9]/g,''))).filter(x=>x>0&&x<500))]}
function rows(wb,name){let sh=wb.Sheets[name]||wb.Sheets[wb.SheetNames[0]];return XLSX.utils.sheet_to_json(sh,{header:1,defval:''})}
async function excel(file){if(!window.XLSX)throw Error('Excel library unavailable. Connect to the internet and reload.');return XLSX.read(await file.arrayBuffer(),{type:'array'})}
function parseRoster(wb){let r=rows(wb,'777');if(!wb.Sheets['777'])throw Error('777 worksheet not found in monthly Bid List.');let header=r.slice(0,30).flatMap((row,i)=>row.map((v,j)=>({v:String(v).toUpperCase().trim(),i,j})));let exclusions=[];
 for(let [label,reason] of [['FO CQ SCHEDULED','FO CQ'],['VAC FO','FO Vacation']]){let h=header.find(x=>x.v===label);if(!h)throw Error('Missing '+label+' column in 777 worksheet');for(let i=h.i+1;i<r.length;i++){let val=r[i]?.[h.j];if(typeof val!=='string'||!val.trim())break;exclusions.push({name:val.trim(),reason})}}
 return exclusions}
function parseBidRows(r){ // The company export uses repeated Name/Choice/Line rows; preserve displayed order.
 let hIndex=r.findIndex(row=>row.some(v=>/^name$/i.test(String(v).trim()))&&row.some(v=>/^line$/i.test(String(v).trim())));
 if(hIndex<0)throw Error('Could not find Name / Choice / Line headings in exported bids.');
 let h=r[hIndex].map(x=>String(x).trim().toLowerCase()),nameCol=h.indexOf('name'),choiceCol=h.indexOf('choice'),lineCol=h.indexOf('line');
 let people=[],current=null;
 for(let i=hIndex+1;i<r.length;i++){
   let row=r[i]||[],name=String(row[nameCol]||'').trim(),choice=String(row[choiceCol]||'').trim(),line=String(row[lineCol]||'').trim();
   if(/^name$/i.test(name)||/^first officer bids/i.test(name))continue;
   if(name&&/[a-z]/i.test(name)&&!/^page\s+\d+|^total/i.test(name)){
     if(!current||!match(current.name,name)){current={name,bids:[],raw:[]};people.push(current)}
   }
   if(!current)continue;
   let ln=line.match(/\bF?\s*(\d{1,3})\b/i);if(ln&&Number(ln[1])>0){let n=Number(ln[1]);let rank=Number(choice.replace(/\D/g,''));current.raw.push({n,rank:Number.isFinite(rank)?rank:9999,seq:i})}
 }
 for(let p of people){p.raw.sort((a,b)=>a.rank-b.rank||a.seq-b.seq);p.bids=[...new Set(p.raw.map(x=>x.n))];delete p.raw}
 return people.filter(x=>x.name.length>3)}
async function parsePdf(file){
 if(!window.pdfjsLib)throw Error('PDF library unavailable. Connect to the internet and reload.');
 pdfjsLib.GlobalWorkerOptions.workerSrc='https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
 const doc=await pdfjsLib.getDocument({data:await file.arrayBuffer()}).promise;
 const capacities={},details={},pageMap={};
 for(let i=1;i<=doc.numPages;i++){
  const page=await doc.getPage(i),content=await page.getTextContent(),t=content.items.map(x=>x.str).join(' ');
  const id=t.match(/\bF\s*(\d{1,3})\s*[-–]\s*\d+\s*\/\s*(\d+)\b/i);
  if(!id)continue;
  const n=Number(id[1]);capacities[n]=Number(id[2]);pageMap[n]=i;
  const credit=t.match(/Credit\s*time\s*:\s*(\d{1,3}:\d{2})/i),days=t.match(/Duty\s*days\s*:\s*(\d+)/i);
  details[n]={credit:credit?.[1]||'—',duties:days?.[1]||'—'};
 }
 return {capacities,details,pageMap};
}
function simulate(){let left={...state.capacities},awards=[];for(let i=0;i<state.people.length;i++){let p=state.people[i],exc=state.excluded.find(e=>match(e.name,p.name)),bids=state.overrides[p.name]||p.bids,got=null,rank=null;
 if(!exc)for(let j=0;j<bids.length;j++){let n=bids[j];if(left[n]>0){got=n;rank=j+1;left[n]--;break}}
 awards.push({name:p.name,seniority:i+1,line:got,rank,blank:p.bids.length===0,excluded:exc?.reason||'',bidCount:p.bids.length})}
 state.awards=awards;render()}
function render(){let me=state.awards.find(p=>p.name===state.selectedName);$('award').textContent=me?.line?'F'+me.line:'—';$('status').textContent=me?.excluded?'Excluded: '+me.excluded:me?.line?'Provisional award':'No projected line';$('credit').textContent=state.details[me?.line]?.credit||'—';$('duties').textContent=state.details[me?.line]?.duties||'—';$('rank').textContent=me?.rank?'#'+me.rank:'—';let senior=me?state.awards.slice(0,me.seniority-1):[];$('blanks').textContent=me?senior.filter(x=>x.blank&&!x.excluded).length:'—';$('eligible').textContent=state.awards.filter(x=>!x.excluded).length||'—';$('viewProjectedLine').disabled=!me?.line;
 $('whoAmIFlyingWith').disabled=!me?.line;
 if(!$('lineDetail').classList.contains('hidden'))showProjectedLine();
 if(!$('flyingWith').classList.contains('hidden'))showFlyingWith();$('notes').textContent=me?`Seniority position ${me.seniority} in exported bid report. ${senior.filter(x=>x.excluded).length} senior CQ/Vacation exclusions. ${senior.filter(x=>x.blank&&!x.excluded).length} eligible senior blank bids. Projection only; company award rules may differ.`:'Import the three files and select a pilot to calculate a projected line.';renderPilots();renderBackupLines()}
function renderPilots(){let q=$('search').value.trim().toLowerCase(),list=state.awards.filter(p=>!q||p.name.toLowerCase().includes(q)||('f'+p.line).toLowerCase()===q);$('pilotList').replaceChildren(...list.map(p=>{let d=document.createElement('div');d.className='pilot';let a=document.createElement('div'),b=document.createElement('b'),sm=document.createElement('small');a.textContent=`${p.seniority}. ${p.name}`;sm.textContent=p.excluded?p.excluded:p.blank?'No bid submitted':p.line?`Choice #${p.rank}`:'No available submitted choice';a.append(sm);b.textContent=p.excluded?'Excluded':p.line?'F'+p.line:'—';d.append(a,b);return d}))}
// For each choice, test remaining FO capacity after all more-senior awards.
// The actual next backup is the first available lower-ranked choice if the current award is blocked.
function renderBackupLines(){
 const rows=$('backupRows'),current=$('backupCurrent'),next=$('backupNext');
 const currentTag=$('backupCurrentTag'),nextTag=$('backupNextTag');
 const pilot=state.people.find(p=>p.name===state.selectedName);
 const me=state.awards.find(p=>p.name===state.selectedName);
 rows.replaceChildren();
 if(!pilot||!me){current.textContent='—';next.textContent='—';currentTag.textContent='Awaiting projection';nextTag.textContent='Awaiting projection';const msg=document.createElement('p');msg.className='muted';msg.textContent='Select a pilot with imported bids to see backup lines.';rows.append(msg);return}
 const bids=state.overrides[pilot.name]||pilot.bids;
 current.textContent=me.line?'F'+me.line:'—';currentTag.textContent=me.line?'Projected award':me.excluded?'Excluded':'No projected line';
 // Reconstruct allocations made before this pilot, using the same seniority/capacity rules as simulate().
 const remaining={...state.capacities};
 for(const senior of state.awards.slice(0,me.seniority-1))if(senior.line!=null&&remaining[senior.line]>0)remaining[senior.line]--;
 const currentIndex=bids.indexOf(me.line);
 const backup=me.excluded?null:bids.find((line,i)=>i>currentIndex&&remaining[line]>0&&line!==me.line);
 next.textContent=backup!=null?'F'+backup:'—';
 nextTag.textContent=me.excluded?'Not eligible':backup!=null?'Available':'No backup found';
 nextTag.classList.toggle('red',backup==null);nextTag.classList.toggle('green',backup!=null);
 if(!bids.length){const msg=document.createElement('p');msg.className='muted';msg.textContent='No bids listed for this pilot.';rows.append(msg);return}
 bids.forEach((line,i)=>{
   const isCurrent=line===me.line;
   const canHold=!me.excluded&&remaining[line]>0;
   const item=document.createElement('div');item.className='backup-row';
   const number=document.createElement('span');number.className='backup-number';number.textContent=String(i+1);
   const info=document.createElement('div');info.className='backup-info';
   const title=document.createElement('strong');title.textContent='F'+line;
   const detail=document.createElement('small');
   detail.textContent=isCurrent?'Current projected award':!canHold?'Filled by senior pilots / no FO capacity':line===backup?'Next possible backup':i<currentIndex?'Available if higher choices change':'Possible backup if earlier choices unavailable';
   info.append(title,detail);
   const status=document.createElement('span');status.className='backup-status '+(canHold?'yes':'no');status.textContent=canHold?'✓':'✕';status.setAttribute('aria-label',canHold?'Can hold':'Cannot hold');
   item.append(number,info,status);rows.append(item);
 });
}
$('search').oninput=renderPilots;
function selectPilot(name){
 state.selectedName=name;
 if(name){try{localStorage.setItem('sbo-selected-pilot',name)}catch(e){console.warn('Name could not be remembered',e)}}
 const pilot=state.people.find(p=>p.name===name);
 $('choices').value=(pilot?(state.overrides[name]||pilot.bids):[]).map(n=>'F'+n).join('\n');
 $('prefMsg').textContent=pilot?'Imported bids shown. Changes are what-if only.':'';
 if(state.people.length)render();
}
$('pilotSelect').onchange=e=>selectPilot(e.target.value);
$('saveChoices').onclick=()=>{
 if(!state.selectedName){$('prefMsg').textContent='Select a pilot on Home first.';return}
 const n=parseNums($('choices').value);
 state.overrides[state.selectedName]=n;
 $('prefMsg').textContent=`Applied ${n.length} what-if choices for ${state.selectedName}. This does not submit bids.`;
 simulate();
};
$('analyze').onclick=async()=>{
 let status=$('importStatus');
 try{
   const pf=$('linesFile').files[0],rf=$('rosterFile').files[0],bf=$('bidsFile').files[0];
   if(!bf)throw Error('Select the updated View All FO Bids Excel report.');
   // Keep the existing monthly schedule and CQ/Vacation exclusions unless replaced.
   const saved=await dbGet('parsed');
   const base=saved&&saved.capacities&&saved.details&&Array.isArray(saved.excluded)?saved:null;
   if(!base&&(!pf||!rf))throw Error('First import on this device: select all three files. Future updates need only the FO bids export.');
   if((pf&&!rf)||(!pf&&rf))throw Error('For a new month, replace BOTH the 777 Lines PDF and monthly Bid List workbook.');
   status.textContent='Reading updated FO bids…';
   const people=parseBidRows(rows(await excel(bf),'Sheet1'));
   if(people.length<30)throw Error(`Only ${people.length} pilots identified in the bid report. Please check the export format.`);
   let capacities,details,excluded,pageMap;
   if(pf&&rf){
     status.textContent='Reading new monthly lines and exclusions…';
     const pdf=await parsePdf(pf);
     capacities=pdf.capacities;details=pdf.details;pageMap=pdf.pageMap;
     excluded=parseRoster(await excel(rf));
     if(Object.keys(capacities).length<40)throw Error(`Only ${Object.keys(capacities).length} lines identified in PDF.`);
   }else{
     capacities={...base.capacities};details={...base.details};pageMap={...(base.pageMap||{})};excluded=base.excluded;
   }
   for(const m of $('capacities').value.matchAll(/F?\\s*(\\d+)\\s*=\\s*(\\d+)/gi))capacities[Number(m[1])]=Number(m[2]);
   const updatedAt=new Date().toISOString();
   const next={people,capacities,details,pageMap,excluded,overrides:{},updatedAt};
   // Save first, then update the live display, so a failed write cannot masquerade as success.
   if(pf)await dbPut('linesPdf',pf);
   await dbPut('parsed',next);
   state={...state,...next};
   populatePilots();simulate();
   $('lastUpdated').textContent='Last imported: '+new Date(updatedAt).toLocaleString();
   status.textContent=`Saved locally: ${people.length} pilots, ${Object.keys(capacities).length} lines, ${excluded.length} CQ/Vacation entries. ${pf?'All monthly files replaced.':'Only FO bids updated; saved lines and exclusions reused.'}`;
   $('bidsFile').value='';$('linesFile').value='';$('rosterFile').value='';
   document.querySelector('[data-tab="home"]').click();
 }catch(e){status.textContent='Import error: '+e.message;console.error(e)}
};

$('clearSaved').onclick=async()=>{if(!confirm('Remove saved bidding data and selected pilot from this device?'))return;await dbClear();localStorage.removeItem('sbo-selected-pilot');location.reload()};
restoreSaved();

// iPhone-friendly refresh flow: open the authenticated portal, then select the
// downloaded file using the operating system's normal file picker.
// Browser security prevents silent reading of Safari downloads or portal cookies.
$('pickLatest').onclick=()=>{$('bidsFile').click()};
$('bidsFile').addEventListener('change',()=>{
 const f=$('bidsFile').files[0];
 $('refreshFileName').textContent=f?`Selected: ${f.name}`:'No report selected';
 $('refreshUpdate').disabled=!f;
});
$('refreshUpdate').onclick=async()=>{
 if(!$('bidsFile').files[0])return;
 $('refreshStatus').textContent='Importing latest FO bids…';
 // Reuse the exact same tested parser, validation and IndexedDB persistence.
 await $('analyze').onclick();
 const msg=$('importStatus').textContent;
 $('refreshStatus').textContent=msg;
 if(msg.startsWith('Saved locally:')){
   $('refreshFileName').textContent='No report selected';
   $('refreshUpdate').disabled=true;
 }else{
   document.querySelector('[data-tab="refresh"]').click();
 }
};

// Render the original Jeppesen line chart directly from the imported PDF.
// No transcription or invented day-by-day data. All PDF data stays on this device.
let lineViewToken=0;
async function showProjectedLine(){
 const me=state.awards.find(p=>p.name===state.selectedName);
 if(!me?.line){$('lineDetail').classList.add('hidden');return}
 const token=++lineViewToken;
 $('lineDetailTitle').textContent='Projected line F'+me.line;
 $('lineDetail').classList.remove('hidden');
 $('lineDetailMessage').textContent='Loading original line chart…';
 $('lineChart').replaceChildren();
 $('lineZoom').classList.add('hidden');
 try{
  const pageNumber=state.pageMap?.[me.line];
  const pdfBlob=await dbGet('linesPdf');
  if(!pageNumber||!pdfBlob){
   $('lineDetailMessage').textContent='To view the original line chart, reimport this month’s 777 Lines PDF and Bid List together with the FO bids report. Your saved projection remains available.';
   return;
  }
  if(!window.pdfjsLib)throw Error('PDF viewer unavailable. Connect to the internet and reload.');
  pdfjsLib.GlobalWorkerOptions.workerSrc='https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
  const pdf=await pdfjsLib.getDocument({data:await pdfBlob.arrayBuffer()}).promise;
  const page=await pdf.getPage(pageNumber);
  // PDF is landscape. The schedule occupies its upper half; crop blank lower page.
  const viewport=page.getViewport({scale:2.2});
  const full=document.createElement('canvas');
  full.width=Math.ceil(viewport.width);full.height=Math.ceil(viewport.height);
  await page.render({canvasContext:full.getContext('2d'),viewport}).promise;
  if(token!==lineViewToken)return;
  const left=Math.floor(full.width*.025),top=Math.floor(full.height*.12);
  const width=Math.floor(full.width*.93),height=Math.floor(full.height*.43);
  const cropped=document.createElement('canvas');
  cropped.width=width;cropped.height=height;
  cropped.getContext('2d').drawImage(full,left,top,width,height,0,0,width,height);
  cropped.setAttribute('role','img');
  cropped.setAttribute('aria-label','Original Kalitta PDF chart for line F'+me.line);
  $('lineChart').append(cropped);
  $('lineZoom').classList.remove('hidden');
  $('lineDetailMessage').textContent='Original chart from your imported PDF. Tap Enlarge for a clearer view, then scroll sideways.';
 }catch(e){
  if(token===lineViewToken)$('lineDetailMessage').textContent='Unable to display the PDF chart: '+e.message;
 }
}

function showFlyingWith(){
 const me=state.awards.find(p=>p.name===state.selectedName);
 if(!me?.line){$('flyingWith').classList.add('hidden');return}
 $('flyingWithTitle').textContent=`Who am I flying with on F${me.line}?`;
 $('flyingWith').classList.remove('hidden');
 const partners=state.awards.filter(p=>p.line===me.line&&p.name!==me.name);
 $('flyingWithList').replaceChildren(...partners.map(p=>{
   const row=document.createElement('div');
   row.className='flying-with-pilot';
   const info=document.createElement('div');
   const name=document.createElement('strong');
   const meta=document.createElement('small');
   name.textContent=p.name;
   meta.textContent=`Seniority #${p.seniority}${p.rank?` · Choice #${p.rank}`:''}`;
   info.append(name,meta);
   const line=document.createElement('span');
   line.className='flying-with-line';
   line.textContent='F'+p.line;
   row.append(info,line);
   return row;
 }));
 $('flyingWithMessage').textContent=partners.length
   ? `${partners.length} other first officer${partners.length===1?'':'s'} ${partners.length===1?'is':'are'} currently projected to hold this line.`
   : 'No other first officer is currently projected to hold this line.';
}
$('whoAmIFlyingWith').onclick=()=>{
 if($('flyingWith').classList.contains('hidden'))showFlyingWith();
 else $('flyingWith').classList.add('hidden');
};
$('closeFlyingWith').onclick=()=>$('flyingWith').classList.add('hidden');

$('viewProjectedLine').onclick=()=>{
 if($('lineDetail').classList.contains('hidden'))showProjectedLine();
 else{$('lineDetail').classList.add('hidden');lineViewToken++;}
};
$('closeLineDetail').onclick=()=>{$('lineDetail').classList.add('hidden');lineViewToken++};
$('lineZoom').onclick=()=>{
 const zoomed=$('lineChart').classList.toggle('zoomed');
 $('lineZoom').textContent=zoomed?'Fit to screen':'Enlarge / scroll';
};

// All dates are UTC. At/after 21:00Z on the 16th, count down to next month.
function nextBidDeadline(now){
  let year=now.getUTCFullYear(), month=now.getUTCMonth();
  let deadline=Date.UTC(year,month,16,21,0,0);
  if(now.getTime()>=deadline) deadline=Date.UTC(year,month+1,16,21,0,0);
  return deadline;
}
function updateBidCountdown(){
  const now=new Date(), deadline=nextBidDeadline(now);
  let remaining=Math.max(0,Math.ceil((deadline-now.getTime())/1000));
  const days=Math.floor(remaining/86400); remaining%=86400;
  const hours=Math.floor(remaining/3600); remaining%=3600;
  const minutes=Math.floor(remaining/60);
  $('bidDays').textContent=String(days).padStart(2,'0');
  $('bidHours').textContent=String(hours).padStart(2,'0');
  $('bidMinutes').textContent=String(minutes).padStart(2,'0');
  $('bidDeadline').textContent='Closes '+new Intl.DateTimeFormat('en-GB',{timeZone:'UTC',day:'numeric',month:'long',year:'numeric'}).format(new Date(deadline))+' at 21:00 UTC';
}
updateBidCountdown();
setInterval(updateBidCountdown,10000);
document.addEventListener('visibilitychange',()=>{if(!document.hidden)updateBidCountdown()});
