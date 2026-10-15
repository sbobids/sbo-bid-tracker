const $=id=>document.getElementById(id);
const defaultChoices=[49,89,66,31,85,75,86,88,76,29,10,19,78,36,57,47,82,8,12,67,7,80,37,72,64,32,53,69,42,1,56,20,24,3,2];
let state={people:[],capacities:{},details:{},excluded:[],awards:[],selectedName:'',overrides:{}};
$('choices').value='';
const DB_NAME='sbo-bid-tracker-local-v1', DB_STORE='saved';
function openDB(){return new Promise((resolve,reject)=>{let req=indexedDB.open(DB_NAME,1);req.onupgradeneeded=()=>req.result.createObjectStore(DB_STORE);req.onsuccess=()=>resolve(req.result);req.onerror=()=>reject(req.error)})}
async function dbPut(key,value){let db=await openDB();try{await new Promise((resolve,reject)=>{let tx=db.transaction(DB_STORE,'readwrite');tx.objectStore(DB_STORE).put(value,key);tx.oncomplete=resolve;tx.onerror=()=>reject(tx.error)})}finally{db.close()}}
async function dbGet(key){let db=await openDB();try{return await new Promise((resolve,reject)=>{let req=db.transaction(DB_STORE).objectStore(DB_STORE).get(key);req.onsuccess=()=>resolve(req.result);req.onerror=()=>reject(req.error)})}finally{db.close()}}
async function dbClear(){let db=await openDB();try{await new Promise((resolve,reject)=>{let tx=db.transaction(DB_STORE,'readwrite');tx.objectStore(DB_STORE).clear();tx.oncomplete=resolve;tx.onerror=()=>reject(tx.error)})}finally{db.close()}}
function populatePilots(){const sel=$('pilotSelect');sel.replaceChildren(new Option('Choose your name…',''),...state.people.map(p=>new Option(p.name,p.name)));const previous=localStorage.getItem('sbo-selected-pilot')||'';const name=state.people.some(p=>p.name===previous)?previous:'';sel.value=name;selectPilot(name)}
async function restoreSaved(){try{const saved=await dbGet('parsed');if(saved&&Array.isArray(saved.people)&&saved.people.length){state={...state,...saved,awards:[]};populatePilots();simulate();$('importStatus').textContent='Saved bidding data restored from this device. Import new files to refresh.'}else{$('importStatus').textContent='No saved bidding data on this device yet.'}}catch(e){$('importStatus').textContent='Local storage unavailable: '+e.message}}

document.querySelectorAll('.tab').forEach(b=>b.onclick=()=>{document.querySelectorAll('.tab').forEach(x=>x.classList.toggle('active',x===b));document.querySelectorAll('.page').forEach(x=>x.classList.toggle('hidden',x.id!==b.dataset.tab))});
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
async function parsePdf(file){if(!window.pdfjsLib)throw Error('PDF library unavailable. Connect to the internet and reload.');pdfjsLib.GlobalWorkerOptions.workerSrc='https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';let doc=await pdfjsLib.getDocument({data:await file.arrayBuffer()}).promise;let capacities={},details={};
 for(let i=1;i<=doc.numPages;i++){let page=await doc.getPage(i),content=await page.getTextContent(),items=content.items.map(x=>x.str),t=items.join(' ');let id=t.match(/\bF\s*(\d{1,3})\s*[-–]\s*\d+\s*\/\s*(\d+)\b/i);if(!id)continue;let n=Number(id[1]);capacities[n]=Number(id[2]);let credit=t.match(/Credit\s*time\s*:\s*(\d{1,3}:\d{2})/i),days=t.match(/Duty\s*days\s*:\s*(\d+)/i);details[n]={credit:credit?.[1]||'—',duties:days?.[1]||'—'} }
 return {capacities,details}}
function simulate(){let left={...state.capacities},awards=[];for(let i=0;i<state.people.length;i++){let p=state.people[i],exc=state.excluded.find(e=>match(e.name,p.name)),bids=state.overrides[p.name]||p.bids,got=null,rank=null;
 if(!exc)for(let j=0;j<bids.length;j++){let n=bids[j];if(left[n]>0){got=n;rank=j+1;left[n]--;break}}
 awards.push({name:p.name,seniority:i+1,line:got,rank,blank:p.bids.length===0,excluded:exc?.reason||'',bidCount:p.bids.length})}
 state.awards=awards;render()}
function render(){let me=state.awards.find(p=>p.name===state.selectedName);$('award').textContent=me?.line?'F'+me.line:'—';$('status').textContent=me?.excluded?'Excluded: '+me.excluded:me?.line?'Provisional award':'No projected line';$('credit').textContent=state.details[me?.line]?.credit||'—';$('duties').textContent=state.details[me?.line]?.duties||'—';$('rank').textContent=me?.rank?'#'+me.rank:'—';let senior=me?state.awards.slice(0,me.seniority-1):[];$('blanks').textContent=me?senior.filter(x=>x.blank&&!x.excluded).length:'—';$('eligible').textContent=state.awards.filter(x=>!x.excluded).length||'—';$('notes').textContent=me?`Seniority position ${me.seniority} in exported bid report. ${senior.filter(x=>x.excluded).length} senior CQ/Vacation exclusions. ${senior.filter(x=>x.blank&&!x.excluded).length} eligible senior blank bids. Projection only; company award rules may differ.`:'Import the three files and select a pilot to calculate a projected line.';renderPilots()}
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
$('analyze').onclick=async()=>{let status=$('importStatus');try{let pf=$('linesFile').files[0],rf=$('rosterFile').files[0],bf=$('bidsFile').files[0];if(!pf||!rf||!bf)throw Error('Please select all three files.');status.textContent='Reading line PDF…';let pdf=await parsePdf(pf);status.textContent='Reading FO CQ / Vacation…';let excluded=parseRoster(await excel(rf));status.textContent='Reading submitted FO bids…';let people=parseBidRows(rows(await excel(bf),'Sheet1'));let overrides=$('capacities').value.matchAll(/F?\s*(\d+)\s*=\s*(\d+)/gi);for(let m of overrides)pdf.capacities[Number(m[1])]=Number(m[2]);if(Object.keys(pdf.capacities).length<40)throw Error(`Only ${Object.keys(pdf.capacities).length} lines were identified in the PDF. Please check the PDF format.`);if(people.length<30)throw Error(`Only ${people.length} pilots identified in the bid report. Please check the export format.`);state={...state,...pdf,excluded,people,overrides:{}};populatePilots();simulate();
 try{await dbPut('parsed',{people,capacities:pdf.capacities,details:pdf.details,excluded,overrides:{}});}
 catch(storageError){throw Error('Projection calculated but could not save locally: '+storageError.message)}status.textContent=`Saved locally: ${people.length} pilots, ${Object.keys(pdf.capacities).length} lines, ${excluded.length} CQ/Vacation entries. No files uploaded to a server.`;document.querySelector('[data-tab="home"]').click()}catch(e){status.textContent='Import error: '+e.message;console.error(e)}};

$('clearSaved').onclick=async()=>{if(!confirm('Remove saved bidding data and selected pilot from this device?'))return;await dbClear();localStorage.removeItem('sbo-selected-pilot');location.reload()};
restoreSaved();
