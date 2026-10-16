const $=id=>document.getElementById(id);
const defaultChoices=[49,89,66,31,85,75,86,88,76,29,10,19,78,36,57,47,82,8,12,67,7,80,37,72,64,32,53,69,42,1,56,20,24,3,2];
const emptyGroup=()=>({people:[],excluded:[],awards:[],selectedName:'',overrides:{},updatedAt:null});
let groups={FO:emptyGroup(),CA:emptyGroup()};
let position=localStorage.getItem('sbo-position')==='CA'?'CA':'FO';
let state={...emptyGroup(),capacities:{},details:{},pageMap:{}};
const groupKey=pos=>'parsed-'+pos;
function stashGroup(){groups[position]={people:state.people,excluded:state.excluded,awards:state.awards,selectedName:state.selectedName,overrides:state.overrides,updatedAt:state.updatedAt}}
function setPosition(pos){
 if(pos!==position)stashGroup();
 position=pos;localStorage.setItem('sbo-position',pos);
 const g=groups[pos];Object.assign(state,g);
 $('positionFO').classList.toggle('active',pos==='FO');$('positionCA').classList.toggle('active',pos==='CA');
 $('pilotsPositionFO').classList.toggle('active',pos==='FO');$('pilotsPositionCA').classList.toggle('active',pos==='CA');
 $('positionLabel').textContent='777 '+pos;
 $('portalInstructions').textContent='Sign in with Safari, then go to Reports → View Other Bids. Export View All Captain Bids and View All First Officer Bids separately as Excel files.';
 $('eligibleLabel').textContent='Eligible 777 '+(pos==='CA'?'Captains':'FOs');
 $('refreshTitle').textContent='Refresh Captain & FO bids';

 $('homePilotSearch').placeholder='Search '+(pos==='CA'?'Captain':'First Officer')+' name…';
 updateRefreshFiles();
 $('lastUpdated').textContent=g.updatedAt?'Last imported: '+new Date(g.updatedAt).toLocaleString():'No '+pos+' bids imported yet';
 $('importStatus').textContent=g.people.length?pos+' bids loaded.':'Import the '+pos+' bids report to get started.';
 $('lineDetail').classList.add('hidden');
 $('search').value='';$('choices').value='';$('homePilotSearch').value='';hideHomePilotResults();
 populatePilots();simulate();
}
$('positionFO').onclick=()=>setPosition('FO');
$('positionCA').onclick=()=>setPosition('CA');
$('pilotsPositionFO').onclick=()=>setPosition('FO');
$('pilotsPositionCA').onclick=()=>setPosition('CA');
$('choices').value='';
const DB_NAME='sbo-bid-tracker-local-v1', DB_STORE='saved';
function openDB(){return new Promise((resolve,reject)=>{let req=indexedDB.open(DB_NAME,1);req.onupgradeneeded=()=>req.result.createObjectStore(DB_STORE);req.onsuccess=()=>resolve(req.result);req.onerror=()=>reject(req.error)})}
async function dbPut(key,value){let db=await openDB();try{await new Promise((resolve,reject)=>{let tx=db.transaction(DB_STORE,'readwrite');tx.objectStore(DB_STORE).put(value,key);tx.oncomplete=resolve;tx.onerror=()=>reject(tx.error)})}finally{db.close()}}
async function dbGet(key){let db=await openDB();try{return await new Promise((resolve,reject)=>{let req=db.transaction(DB_STORE).objectStore(DB_STORE).get(key);req.onsuccess=()=>resolve(req.result);req.onerror=()=>reject(req.error)})}finally{db.close()}}
async function dbClear(){let db=await openDB();try{await new Promise((resolve,reject)=>{let tx=db.transaction(DB_STORE,'readwrite');tx.objectStore(DB_STORE).clear();tx.oncomplete=resolve;tx.onerror=()=>reject(tx.error)})}finally{db.close()}}
function populatePilots(){
 const previous=state.selectedName||localStorage.getItem('sbo-selected-pilot-'+position)||(position==='FO'?localStorage.getItem('sbo-selected-pilot'):'')||'';
 const name=state.people.some(p=>p.name===previous)?previous:'';
 selectPilot(name);
}
function hideHomePilotResults(){
 $('homePilotResults').classList.add('hidden');
 $('homePilotSearch').setAttribute('aria-expanded','false');
}
function showHomePilotResults(){
 const input=$('homePilotSearch'),box=$('homePilotResults');
 const query=input.value.trim().toLowerCase();
 box.replaceChildren();
 if(!state.people.length){hideHomePilotResults();return}
 const matches=state.people.filter(p=>!query||p.name.toLowerCase().includes(query)).slice(0,25);
 if(!matches.length){
  const empty=document.createElement('div');empty.className='pilot-search-empty';empty.textContent='No matching pilots';box.append(empty);
 }else{
  for(const p of matches){
   const b=document.createElement('button');b.type='button';b.className='pilot-search-option';b.setAttribute('role','option');b.textContent=p.name;
   b.addEventListener('click',()=>selectPilot(p.name));
   box.append(b);
  }
 }
 box.classList.remove('hidden');input.setAttribute('aria-expanded','true');
}
$('homePilotSearch').addEventListener('focus',showHomePilotResults);
$('homePilotSearch').addEventListener('input',showHomePilotResults);
$('homePilotSearch').addEventListener('keydown',e=>{
 if(e.key==='Escape'){hideHomePilotResults();$('homePilotSearch').blur()}
 if(e.key==='Enter'){
  const first=$('homePilotResults').querySelector('.pilot-search-option');
  if(first){e.preventDefault();first.click();$('homePilotSearch').blur()}
 }
});
document.addEventListener('pointerdown',e=>{
 if(!e.target.closest('.pilot-search-wrap'))hideHomePilotResults();
});

async function restoreSaved(){
 try{
  const legacy=await dbGet('parsed'),monthly=await dbGet('monthly');
  const savedFO=await dbGet(groupKey('FO')),savedCA=await dbGet(groupKey('CA'));
  const shared=monthly||legacy||{};
  const typed=o=>Object.fromEntries(Object.entries(o||{}).map(([k,v])=>[lineKey(k)||k,v]));
  state.capacities=typed(shared.capacities);state.details=typed(shared.details);state.pageMap=typed(shared.pageMap);
  for(const pos of ['FO','CA']){
   const saved=pos==='FO'?(savedFO||legacy):savedCA;
   if(saved?.people?.length)groups[pos]={...emptyGroup(),...saved,awards:[],selectedName:localStorage.getItem('sbo-selected-pilot-'+pos)||saved.selectedName||''};
  }
  setPosition(position);
 }catch(e){$('importStatus').textContent='Local storage unavailable: '+e.message}
}
document.querySelectorAll('.tab').forEach(b=>b.onclick=()=>{document.querySelectorAll('.tab').forEach(x=>x.classList.toggle('active',x===b));document.querySelectorAll('.page').forEach(x=>x.classList.toggle('hidden',x.id!==b.dataset.tab))});
$('openFiles').onclick=()=>{document.querySelectorAll('.tab').forEach(x=>x.classList.remove('active'));document.querySelectorAll('.page').forEach(x=>x.classList.toggle('hidden',x.id!=='files'));window.scrollTo(0,0)};
$('backHome').onclick=()=>document.querySelector('[data-tab="home"]').click();
$('homeRefresh').onclick=()=>document.querySelector('[data-tab="refresh"]').click();

function norm(s){return String(s||'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase().replace(/[^a-z ]/g,' ').trim().split(/\s+/).filter(Boolean).sort().join(' ')}
function match(a,b){let x=norm(a).split(' '),y=norm(b).split(' ');return x.length>=2&&y.length>=2&&x.filter(t=>y.includes(t)).length>=2}
function lineKey(v){const m=String(v??'').trim().match(/^([FQ])?\s*(\d{1,3})$/i);return m&&Number(m[2])>0?(m[1]?.toUpperCase()==='Q'?'Q':'F')+Number(m[2]):null}
function parseNums(s){return [...new Set((String(s||'').match(/\b[FQ]?\s*\d{1,3}\b/gi)||[]).map(lineKey).filter(Boolean))]}
function rows(wb,name){let sh=wb.Sheets[name]||wb.Sheets[wb.SheetNames[0]];return XLSX.utils.sheet_to_json(sh,{header:1,defval:''})}
async function excel(file){if(!window.XLSX)throw Error('Excel library unavailable. Connect to the internet and reload.');return XLSX.read(await file.arrayBuffer(),{type:'array'})}
function parseRoster(wb){let r=rows(wb,'777');if(!wb.Sheets['777'])throw Error('777 worksheet not found in monthly Bid List.');let header=r.slice(0,30).flatMap((row,i)=>row.map((v,j)=>({v:String(v).toUpperCase().trim(),i,j})));let exclusions=[];
 for(let [label,reason] of (position==='CA'?[['CA CQ SCHEDULED','CA CQ'],['VAC CA','CA Vacation']]:[['FO CQ SCHEDULED','FO CQ'],['VAC FO','FO Vacation']])){let h=header.find(x=>x.v===label);if(!h)throw Error('Missing '+label+' column in 777 worksheet');for(let i=h.i+1;i<r.length;i++){let val=r[i]?.[h.j];if(typeof val!=='string'||!val.trim())break;exclusions.push({name:val.trim(),reason})}}
 return exclusions}
function resolveImportedLine(value, capacities){
 // The company's ViewAllBids export stores line IDs as bare numbers (e.g. 98),
 // while the monthly PDF identifies the same line as Q98 or F98.
 const raw=String(value??'').trim();
 const m=raw.match(/^([FQ])?\s*(\d{1,3})(?:\.0+)?$/i);
 if(!m||Number(m[2])<1)return null;
 const n=Number(m[2]), explicit=m[1]?.toUpperCase();
 if(explicit)return explicit+n;
 const f='F'+n,q='Q'+n;
 if(Object.hasOwn(capacities,q)&&!Object.hasOwn(capacities,f))return q;
 if(Object.hasOwn(capacities,f))return f;
 if(Object.hasOwn(capacities,q))return q;
 return null;
}
function parseBidRows(r,capacities){ // Preserve displayed bid order and resolve numeric IDs against PDF.
 let hIndex=r.findIndex(row=>row.some(v=>/^name$/i.test(String(v).trim()))&&row.some(v=>/^line$/i.test(String(v).trim())));
 if(hIndex<0)throw Error('Could not find Name / Choice / Line headings in exported bids.');
 let h=r[hIndex].map(x=>String(x).trim().toLowerCase()),nameCol=h.indexOf('name'),choiceCol=h.indexOf('choice'),lineCol=h.indexOf('line');
 let people=[],current=null;
 for(let i=hIndex+1;i<r.length;i++){
   let row=r[i]||[],name=String(row[nameCol]||'').trim(),choice=String(row[choiceCol]||'').trim(),line=String(row[lineCol]||'').trim();
   if(/^name$/i.test(name)||/^(first officer|captain) bids/i.test(name))continue;
   if(name&&/[a-z]/i.test(name)&&!/^page\s+\d+|^total/i.test(name)){
     if(!current||!match(current.name,name)){current={name,bids:[],raw:[]};people.push(current)}
   }
   if(!current)continue;
   let n=resolveImportedLine(line,capacities);if(n){let rank=Number(choice.replace(/\D/g,''));current.raw.push({n,rank:Number.isFinite(rank)?rank:9999,seq:i})}
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
  const id=t.match(/\b([FQ])\s*(\d{1,3})\s*[-–]\s*\d+\s*\/\s*(\d+)\b/i);
  if(!id)continue;
  const n=id[1].toUpperCase()+Number(id[2]);capacities[n]=Number(id[3]);pageMap[n]=i;
  const credit=t.match(/Credit\s*time\s*:\s*(\d{1,3}:\d{2})/i),days=t.match(/Duty\s*days\s*:\s*(\d+)/i);
  details[n]={credit:credit?.[1]||'—',duties:days?.[1]||'—'};
 }
 return {capacities,details,pageMap};
}
// CQ pilots bid and receive projected lines, but use a separate capacity pool.
// Vacation pilots remain ineligible. Seniority order is preserved within each pool.
function isCQ(reason){return /\bCQ\b/i.test(reason||'')}
function projectGroup(g){
 const pools={regular:Object.fromEntries(Object.entries(state.capacities).filter(([k])=>String(k).startsWith('F')||/^\d+$/.test(k))),cq:Object.fromEntries(Object.entries(state.capacities).filter(([k])=>String(k).startsWith('Q')))};
 return g.people.map((p,i)=>{
  const exc=g.excluded.find(e=>match(e.name,p.name));
  const cq=isCQ(exc?.reason),vac=!!exc&&!cq;
  const bids=(g.overrides[p.name]||p.bids).map(lineKey).filter(Boolean);
  const left=pools[cq?'cq':'regular'];
  let line=null,rank=null;
  if(!vac)for(let j=0;j<bids.length;j++)if(bids[j].startsWith(cq?'Q':'F')&&left[bids[j]]>0){line=bids[j];rank=j+1;left[line]--;break}
  return {name:p.name,seniority:i+1,line,rank,blank:p.bids.length===0,
   excluded:vac?(exc?.reason||''):'',cq,bidCount:p.bids.length};
 });
}
function simulate(){state.awards=projectGroup(state);groups[position].awards=state.awards;recalculateOtherGroup();render()}
function recalculateOtherGroup(){
 const other=position==='CA'?'FO':'CA',g=groups[other];
 g.awards=g.people.length?projectGroup(g):[];
}
function render(){let me=state.awards.find(p=>p.name===state.selectedName);$('award').textContent=me?.line?me.line:'—';$('status').textContent=me?.excluded?'Excluded: '+me.excluded:me?.line?(me.cq?'CQ provisional award':'Provisional award'):'No projected line';$('credit').textContent=state.details[me?.line]?.credit||'—';$('duties').textContent=state.details[me?.line]?.duties||'—';$('rank').textContent=me?.rank?'#'+me.rank:'—';let senior=me?state.awards.slice(0,me.seniority-1):[];$('blanks').textContent=me?senior.filter(x=>x.blank&&!x.excluded).length:'—';$('eligible').textContent=state.awards.filter(x=>!x.excluded&&!x.cq).length||'—';$('viewProjectedLine').disabled=!me?.line; if(!$('lineDetail').classList.contains('hidden')&&$('lineDetail').parentElement===$('home'))showProjectedLine();$('notes').textContent=me?`Seniority position ${me.seniority} in exported bid report. ${senior.filter(x=>x.excluded).length} senior vacation exclusions. ${senior.filter(x=>x.blank&&!x.excluded).length} eligible senior blank bids. Projection only; company award rules may differ.`:'Import the three files and select a pilot to calculate a projected line.';renderPilots();renderCrew();renderMyBid()}
function renderCrew(){
 const summary=$('crewSummary'),members=$('crewMembers');members.replaceChildren();
 const me=state.awards.find(p=>p.name===state.selectedName);
 if(!me?.line){summary.textContent='Select a pilot with a projected line to see everyone projected to hold it.';return}
 const line=me.line;
 const all=['CA','FO'].flatMap(pos=>(groups[pos].awards||[]).filter(p=>p.line===line).map(p=>({...p,position:pos})));
 summary.textContent='All pilots currently projected to hold F'+line+' ('+all.length+' total)';
 for(const pos of ['CA','FO']){
  const heading=document.createElement('h4');heading.textContent=pos==='CA'?'777 Captains':'777 First Officers';members.append(heading);
  const matches=all.filter(p=>p.position===pos);
  if(!groups[pos].people.length){const p=document.createElement('p');p.className='muted';p.textContent='Import the '+(pos==='CA'?'Captain':'First Officer')+' bid report to see this group.';members.append(p);continue}
  if(!matches.length){const p=document.createElement('p');p.className='muted';p.textContent='No projected pilots on this line.';members.append(p);continue}
  for(const p of matches){const div=document.createElement('div');div.className='crew-member';const name=document.createElement('strong');name.textContent=p.name;const detail=document.createElement('small');detail.textContent='Seniority #'+p.seniority+(p.cq?' · CQ':'')+' · Choice #'+(p.rank||'—');div.append(name,detail);members.append(div)}
 }
}
function renderPilots(){let q=$('search').value.trim().toLowerCase(),list=state.awards.filter(p=>!q||p.name.toLowerCase().includes(q)||(String(p.line||'').toLowerCase()).toLowerCase()===q);$('pilotList').replaceChildren(...list.map(p=>{let d=document.createElement('div');d.className='pilot';let a=document.createElement('div'),b=document.createElement('b'),sm=document.createElement('small');a.textContent=`${p.seniority}. ${p.name}`;sm.textContent=p.excluded?p.excluded:(p.cq?'FO CQ · ':'')+(p.blank?'No bid submitted':p.line?`Choice #${p.rank}`:(p.cq?'No Q line awarded':'No available submitted choice'));a.append(sm);b.textContent=p.excluded?'Excluded':p.line?p.line:'—';d.append(a,b);return d}))}
$('search').oninput=renderPilots;
function selectPilot(name){
 state.selectedName=name;groups[position].selectedName=name;
 $('homePilotSearch').value=name;
 hideHomePilotResults();
 if(name){try{localStorage.setItem('sbo-selected-pilot-'+position,name)}catch(e){console.warn('Name could not be remembered',e)}}
 const pilot=state.people.find(p=>p.name===name);
 $('choices').value=(pilot?(state.overrides[name]||pilot.bids):[]).map(n=>lineKey(n)).filter(Boolean).join('\n');
 $('prefMsg').textContent=pilot?'Imported bids shown. Changes are what-if only.':'';
 if(state.people.length)render();
}
$('saveChoices').onclick=()=>{
 if(!state.selectedName){$('prefMsg').textContent='Select a pilot on Home first.';return}
 const n=parseNums($('choices').value);
 state.overrides[state.selectedName]=n;groups[position].overrides=state.overrides;
 $('prefMsg').textContent=`Applied ${n.length} what-if choices for ${state.selectedName}. This does not submit bids.`;
 simulate();
};
async function importBids(){
 const status=$('importStatus');
 const pf=$('linesFile').files[0],rf=$('rosterFile').files[0];
 const selected={CA:$('captainBidsFile').files[0],FO:$('foBidsFile').files[0]};
 // Quick update on Refresh tab is scoped to the active position.
 if($('bidsFile').files[0] && !selected[position])selected[position]=$('bidsFile').files[0];
 if(!selected.CA&&!selected.FO)throw Error('Choose at least one Captain or First Officer bids report.');
 const prior=position;
 const monthly=await dbGet('monthly'),legacy=await dbGet('parsed');
 const base=monthly||legacy;
 if(!base&&(!pf||!rf))throw Error('First import needs the 777 Lines PDF and monthly Bid List workbook too.');
 if((pf&&!rf)||(!pf&&rf))throw Error('Replace the monthly Lines PDF and Bid List together.');
 let capacities,details,pageMap,exclusions=await dbGet('exclusions')||{CA:[],FO:[]};
 if(pf&&rf){
  status.textContent='Reading shared 777 Lines PDF and Bid List…';
  const pdf=await parsePdf(pf);({capacities,details,pageMap}=pdf);
  if(Object.keys(capacities).length<40)throw Error('Only '+Object.keys(capacities).length+' lines found in PDF.');
  const workbook=await excel(rf);
  for(const pos of ['CA','FO']){
   position=pos;
   try{exclusions[pos]=parseRoster(workbook)}finally{position=prior}
  }
 }else{
  capacities=Object.fromEntries(Object.entries(base.capacities||{}).map(([k,v])=>[lineKey(k)||k,v]));details=Object.fromEntries(Object.entries(base.details||{}).map(([k,v])=>[lineKey(k)||k,v]));pageMap=Object.fromEntries(Object.entries(base.pageMap||{}).map(([k,v])=>[lineKey(k)||k,v]));
 }
 for(const m of $('capacities').value.matchAll(/([FQ])\s*(\d+)\s*=\s*(\d+)/gi))capacities[m[1].toUpperCase()+Number(m[2])]=Number(m[3]);
 const pending={};
 for(const pos of ['CA','FO']){
  if(!selected[pos])continue;
  status.textContent='Reading '+(pos==='CA'?'Captain':'First Officer')+' bids…';
  const report=await excel(selected[pos]);
   const reportSheet=report.SheetNames.find(n=>/^ViewAllBids$/i.test(n))||report.SheetNames[0];
   const people=parseBidRows(rows(report,reportSheet),capacities);
  if(people.length<30)throw Error('Only '+people.length+' '+pos+' pilots found. Verify the '+pos+' export.');
  pending[pos]={people,excluded:exclusions[pos]||[],overrides:{},updatedAt:new Date().toISOString(),selectedName:groups[pos].selectedName||''};
 }
 // Parse all files before committing anything, so a bad second report does not partially update data.
 if(pf){await dbPut('linesPdf',pf);await dbPut('monthly',{capacities,details,pageMap});await dbPut('exclusions',exclusions)}
 for(const pos of Object.keys(pending)){await dbPut(groupKey(pos),pending[pos]);groups[pos]={...emptyGroup(),...pending[pos]}}
 // Update the non-imported group with the latest monthly exclusions too.
 if(pf)for(const pos of ['CA','FO'])if(!pending[pos]){groups[pos].excluded=exclusions[pos]||[];if(groups[pos].people.length)await dbPut(groupKey(pos),{...groups[pos],awards:[]})}
 Object.assign(state,{capacities,details,pageMap});
 // Do not stash the old state over newly imported groups.
 position=prior;
 const target=selected[prior]?prior:(selected.CA?'CA':'FO');
 position=target;
 setPosition(target);
 for(const id of ['captainBidsFile','foBidsFile','bidsFile','linesFile','rosterFile'])$(id).value='';
 updateRefreshFiles();
 status.textContent='Saved: '+Object.entries(pending).map(([pos,data])=>data.people.length+' '+pos+' pilots').join(', ')+'. Captain and FO lists are separate.';
 $('refreshStatus').textContent=status.textContent;
 document.querySelector('[data-tab="home"]').click();
}
$('analyze').onclick=async()=>{try{await importBids()}catch(e){$('importStatus').textContent='Import error: '+e.message;console.error(e)}};

$('clearSaved').onclick=async()=>{if(!confirm('Remove saved bidding data and selected pilot from this device?'))return;await dbClear();localStorage.removeItem('sbo-selected-pilot');localStorage.removeItem('sbo-selected-pilot-FO');localStorage.removeItem('sbo-selected-pilot-CA');location.reload()};
restoreSaved();

// Both reports can be refreshed together, without changing the Home position.
function updateRefreshFiles(){
 const ca=$('captainBidsFile').files[0],fo=$('foBidsFile').files[0];
 $('refreshCaptainFileName').textContent=ca?'Selected: '+ca.name:'No Captain report selected';
 $('refreshFOFileName').textContent=fo?'Selected: '+fo.name:'No First Officer report selected';
 $('refreshUpdate').disabled=!(ca||fo);
}
$('pickCaptainLatest').onclick=()=>$('captainBidsFile').click();
$('pickFOLatest').onclick=()=>$('foBidsFile').click();
$('captainBidsFile').addEventListener('change',updateRefreshFiles);
$('foBidsFile').addEventListener('change',updateRefreshFiles);
$('refreshUpdate').onclick=async()=>{
 $('refreshStatus').textContent='Importing selected Captain / First Officer reports…';
 try{await importBids()}
 catch(e){$('refreshStatus').textContent='Import error: '+e.message;console.error(e)}
 updateRefreshFiles();
};
// Render the original Jeppesen line chart directly from the imported PDF.
// No transcription or invented day-by-day data. All PDF data stays on this device.
let lineViewToken=0;
async function showProjectedLine(requestedLine){
 const me=state.awards.find(p=>p.name===state.selectedName);
 const displayLine=typeof requestedLine==='string'?requestedLine:me?.line;
 if(!displayLine){$('lineDetail').classList.add('hidden');return}
 const token=++lineViewToken;
 $('lineDetailTitle').textContent='Projected line '+displayLine;
 $('lineDetail').classList.remove('hidden');
 $('lineDetailMessage').textContent='Loading original line chart…';
 $('lineChart').replaceChildren();
 $('lineZoom').classList.add('hidden');
 try{
  const pageNumber=state.pageMap?.[displayLine];
  const pdfBlob=await dbGet('linesPdf');
  if(!pageNumber||!pdfBlob){
   $('lineDetailMessage').textContent='To view the original line chart, reimport this month’s 777 Lines PDF and Bid List together with the bids report. Your saved projection remains available.';
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
  cropped.setAttribute('aria-label','Original Kalitta PDF chart for line '+displayLine);
  $('lineChart').append(cropped);
  $('lineZoom').classList.remove('hidden');
  $('lineDetailMessage').textContent='Original chart from your imported PDF. Tap Enlarge for a clearer view, then scroll sideways.';
 }catch(e){
  if(token===lineViewToken)$('lineDetailMessage').textContent='Unable to display the PDF chart: '+e.message;
 }
}
// Reuse the exact Home PDF chart viewer for My Bid choices, without duplicating PDF rendering.
const homeLineViewerAnchor=document.createComment('Home line viewer anchor');
$('lineDetail').before(homeLineViewerAnchor);
function moveLineViewerHome(){
 homeLineViewerAnchor.after($('lineDetail'));
}
function openMyBidLineChart(line,item){
 const viewer=$('lineDetail');
 const alreadyOpen=!viewer.classList.contains('hidden')&&viewer.parentElement===item&&viewer.dataset.line===line;
 if(alreadyOpen){viewer.classList.add('hidden');lineViewToken++;return}
 item.append(viewer);
 viewer.dataset.line=line;
 $('lineChart').classList.remove('zoomed');
 $('lineZoom').textContent='Enlarge / scroll';
 showProjectedLine(line);
}
$('viewProjectedLine').onclick=()=>{
 moveLineViewerHome();
 if($('lineDetail').classList.contains('hidden'))showProjectedLine();
 else{$('lineDetail').classList.add('hidden');lineViewToken++;}
};
$('closeLineDetail').onclick=()=>{$('lineDetail').classList.add('hidden');lineViewToken++;if($('lineDetail').parentElement!==$('home'))moveLineViewerHome()};
$('lineZoom').onclick=()=>{
 const zoomed=$('lineChart').classList.toggle('zoomed');
 $('lineZoom').textContent=zoomed?'Fit to screen':'Enlarge / scroll';
};

// Monthly bidding deadline: 16th at 21:00 UTC. Uses UTC to avoid timezone errors.
function updateBidCountdown(){
 const now=new Date();
 let deadline=Date.UTC(now.getUTCFullYear(),now.getUTCMonth(),16,21,0,0);
 if(now.getTime()>=deadline)deadline=Date.UTC(now.getUTCFullYear(),now.getUTCMonth()+1,16,21,0,0);
 const minutes=Math.max(0,Math.ceil((deadline-now.getTime())/60000));
 const days=Math.floor(minutes/1440),hours=Math.floor((minutes%1440)/60),mins=minutes%60;
 $('bidDays').textContent=String(days);
 $('bidHours').textContent=String(hours).padStart(2,'0');
 $('bidMinutes').textContent=String(mins).padStart(2,'0');
 const date=new Date(deadline);
 $('bidDeadline').textContent=date.toLocaleDateString('en-GB',{day:'numeric',month:'short',timeZone:'UTC'})+' · 21:00 UTC';
}
updateBidCountdown();
setInterval(updateBidCountdown,30000);


// My Bid: availability at the selected pilot's turn, after senior pilots take their projected lines.
// These are conditional alternatives, not simultaneous awards or guaranteed future outcomes.
function renderMyBid(){
 const pilot=state.people.find(p=>p.name===state.selectedName), me=state.awards.find(a=>a.name===state.selectedName);
 moveLineViewerHome();$('lineDetail').classList.add('hidden');lineViewToken++;
 const list=$('myBidLines');list.replaceChildren();
 $('myBidPilot').textContent=pilot?pilot.name+' · 777 '+position:'Select a pilot on Home to see their bid outlook.';
 $('myBidAward').textContent=me?.line?me.line:'—';
 $('myBidAwardBadge').textContent=me?.excluded?'Excluded':me?.line?(me.cq?'CQ projected':'Currently projected'):'No projected line';
 $('myBidAwardDetail').textContent=me?.excluded?'Excluded: '+me.excluded:me?.line?(me.cq?'CQ bid projection, separate from regular line capacity':'Based on current imported bids and seniority'):'Import bids and select a pilot to calculate your outlook.';
 $('myBidRank').textContent=me?.rank?'#'+me.rank:'—';
 $('myBidCapacity').textContent=me?.line?String(state.capacities[me.line]??'—'):'—';
 const senior=me?state.awards.slice(0,me.seniority-1):[];
 const blankCount=senior.filter(a=>a.cq===me.cq&&a.blank&&!a.excluded).length;
 $('myBidBlanks').textContent=me?String(blankCount):'—';
 $('myBidNext').textContent='—';$('myBidNextRank').textContent='';
 $('myBidNextNote').textContent='Your next currently open preference will appear here.';
 $('myBidCount').textContent='';
 if(!pilot||!me){const msg=document.createElement('p');msg.className='muted';msg.textContent='Select a pilot on Home to see their ranked choices.';list.append(msg);return}
 const bids=(state.overrides[pilot.name]||pilot.bids).map(lineKey).filter(Boolean);
 if(!bids.length){const msg=document.createElement('p');msg.className='muted';msg.textContent='No ranked choices found for this pilot.';list.append(msg);return}
 const remaining={...state.capacities};
 for(const a of senior)if(a.cq===me.cq&&a.line!=null&&remaining[a.line]!=null)remaining[a.line]=Math.max(0,remaining[a.line]-1);
 const nextIndex=me.line?bids.findIndex((line,i)=>i>=me.rank&&line.startsWith(me.cq?'Q':'F')&&remaining[line]>0):-1;
 if(nextIndex>=0){$('myBidNext').textContent=bids[nextIndex];$('myBidNextRank').textContent='Choice #'+(nextIndex+1);$('myBidNextNote').textContent='Currently open after senior pilots are assigned. This is a conditional alternative, not a guaranteed award.'}
 else $('myBidNextNote').textContent=me.line?'No later choice is currently open after senior pilots are assigned.':'No available choice in the current projection.';
 $('myBidCount').textContent=bids.length+' choices';
 for(let i=0;i<bids.length;i++){
  const line=bids[i],cap=state.capacities[line],remainingSlots=remaining[line]||0;
  const available=line.startsWith(me.cq?'Q':'F')&&remainingSlots>0,projected=me.line===line&&me.rank===i+1;
  const holders=senior.filter(a=>a.cq===me.cq&&a.line===line);
  const item=document.createElement('details');item.className='mybid-line';
  const summary=document.createElement('summary');
  const order=document.createElement('span');order.className='mybid-order';order.textContent=i+1;
  const info=document.createElement('span');info.className='mybid-line-info';
  const title=document.createElement('b');title.textContent=line;
  const sub=document.createElement('small');sub.textContent=cap==null?'Capacity unavailable':cap+' '+position+' position'+(cap===1?'':'s')+' · '+remainingSlots+' remaining';
  info.append(title,sub);
  const badge=document.createElement('span');badge.className='bid-pill '+(available?'bid-yes':'bid-no');badge.textContent=projected?'Projected award':available?'Can hold':'Cannot hold';
  summary.append(order,info,badge);item.append(summary);
 summary.title='View original PDF chart for '+line;
 summary.addEventListener('click',()=>{if(item.open)openMyBidLineChart(line,item);else if($('lineDetail').parentElement===item){$('lineDetail').classList.add('hidden');lineViewToken++;moveLineViewerHome()}});
  const explanation=document.createElement('div');explanation.className='mybid-explanation';
  const description=document.createElement('p');description.textContent=!line.startsWith(me.cq?'Q':'F')?'This line is not eligible for this pilot’s CQ status.':cap==null?'No capacity was found for this line in the imported PDF.':available?'A position remains after currently submitted senior bids are simulated.':holders.length?'All positions are currently occupied by projected senior pilots.':'No position is available under the imported capacity data.';
  explanation.append(description);
  if(holders.length){const heading=document.createElement('small');heading.textContent='Projected senior holders:';explanation.append(heading);const ul=document.createElement('ul');for(const h of holders){const li=document.createElement('li');li.textContent=h.name+' · seniority #'+h.seniority;ul.append(li)}explanation.append(ul)}
  if(projected){const note=document.createElement('small');note.textContent='This is your first available preference in the current projection.';explanation.append(note)}
  item.append(explanation);list.append(item);
 }
 if(blankCount){const note=document.createElement('p');note.className='muted mybid-warning';note.textContent=blankCount+' eligible senior pilot(s) have blank bids. They may change which lines you can hold.';list.append(note)}
}
