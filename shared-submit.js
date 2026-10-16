/* Optional shared report submission; local bid projections are unchanged. */
(() => {
 const $=id=>document.getElementById(id);
 const url='https://zstpsfyrtpyxhzadjwro.supabase.co';
 const key='sb_publishable_Y4F5W8ZPlGVlC8xF9n_N1Q_gu7Wd88Y';
 const siteKey='0x4AAAAAAFYHZU584g5y-MD-';
 const status=$('sharedStatus'),send=$('sharedSend');
 let captchaToken='',widgetId=null,client=null,busy=false;
 const say=s=>{status.textContent=s};
 const ready=()=>{send.disabled=busy||!captchaToken||!$('sharedFile').files[0]};
 const reset=()=>{captchaToken='';if(widgetId!==null&&window.turnstile)window.turnstile.reset(widgetId);ready()};
 $('openSharedSubmit').onclick=()=>{document.querySelectorAll('.page').forEach(x=>x.classList.toggle('hidden',x.id!=='sharedSubmit'));document.querySelectorAll('.tab').forEach(x=>x.classList.remove('active'));window.scrollTo(0,0);init()};
 $('backRefreshShared').onclick=()=>document.querySelector('[data-tab="refresh"]').click();
 $('sharedFile').addEventListener('change',ready);
 async function init(){
  if(widgetId!==null)return;
  if(!window.supabase?.createClient){say('Supabase library did not load. Check internet connection and reload.');return}
  client=window.supabase.createClient(url,key,{auth:{persistSession:true,autoRefreshToken:true,detectSessionInUrl:false}});
  say('Loading verification…');
  for(let n=0;n<50&&!window.turnstile;n++)await new Promise(r=>setTimeout(r,100));
  if(!window.turnstile){say('Cloudflare verification did not load. Reload this page.');return}
  widgetId=window.turnstile.render('#sharedCaptcha',{
   sitekey:siteKey,theme:'auto',
   callback:v=>{captchaToken=v;ready();say('Verified. Select a report to submit.')},
   'expired-callback':()=>{captchaToken='';ready();say('Verification expired. Verify again.')},
   'error-callback':()=>{captchaToken='';ready();say('Verification failed. Try again.')}
  });
 }
 send.onclick=async()=>{
  const file=$('sharedFile').files[0],pos=$('sharedPosition').value;
  if(!file||!captchaToken||busy)return;
  if(!/\.(xls|xlsx)$/i.test(file.name)||file.size<100||file.size>5*1024*1024){say('Choose an .xls or .xlsx report under 5 MB.');return}
  busy=true;ready();say('Signing in securely…');
  try{
   const {data:sessionData}=await client.auth.getSession();
   let access=sessionData?.session?.access_token;
   if(!access){
    const {data,error}=await client.auth.signInAnonymously({options:{captchaToken}});
    if(error)throw Error('Anonymous sign-in failed: '+error.message);
    access=data.session?.access_token;
   }
   if(!access)throw Error('Could not establish a session.');
   say('Uploading report for review…');
   const form=new FormData();form.append('position',pos);form.append('file',file,file.name);
   const response=await fetch(url+'/functions/v1/submit-bid-report',{method:'POST',headers:{Authorization:'Bearer '+access,apikey:key},body:form});
   const body=await response.json().catch(()=>({}));
   if(!response.ok)throw Error(body.error||'Upload failed (HTTP '+response.status+').');
   say('Submitted for review. Submission ID: '+(body.submission_id||'available in Supabase')+'. Not published.');
   $('sharedFile').value='';
  }catch(e){say('Submission error: '+e.message)}
  finally{busy=false;reset()}
 };
})();
