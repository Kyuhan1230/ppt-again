import test,{before,after} from 'node:test';
import assert from 'node:assert/strict';
import {createHtmlServer} from '../src/html-to-pptx/server.mjs';
let app,cookie,csrf;
before(async()=>{app=await createHtmlServer({port:0});const response=await fetch(app.url+'/api/session');cookie=response.headers.get('set-cookie').split(';')[0];csrf=(await response.json()).csrfToken;});
after(async()=>{await app?.close();});
const post=(route,body,cookieValue=cookie,csrfValue=csrf)=>fetch(app.url+route,{method:'POST',headers:{cookie:cookieValue,'content-type':'application/json','x-csrf-token':csrfValue},body:JSON.stringify(body)});
test('candidate preview and choice are bound to their owner and actual selection',async()=>{
  const html='<body style="margin:0"><section style="width:400px;height:300px;background:#eee"><p>Selected page</p></section><aside>Excluded page</aside></body>';
  const response=await post('/api/inspect',{html,fileName:'report.html'}),job=await response.json();
  assert.equal(response.status,200,JSON.stringify(job));assert.equal(job.status,'selection-required');assert.ok(job.previewUrl);
  const image=await fetch(app.url+job.previewUrl,{headers:{cookie}});assert.equal(image.status,200);assert.equal(image.headers.get('content-type'),'image/png');assert.equal(image.headers.get('cache-control'),'no-store');
  const other=await fetch(app.url+'/api/session'),otherCookie=other.headers.get('set-cookie').split(';')[0],otherCsrf=(await other.json()).csrfToken;
  assert.equal((await fetch(app.url+job.previewUrl,{headers:{cookie:otherCookie}})).status,404);
  const choice=job.inspection.choices.find(c=>c.scope==='elements'&&c.convertible);assert.ok(choice);
  const payload={snapshotId:job.inspection.snapshotId,choiceId:choice.id};
  assert.equal((await post(`/api/jobs/${job.id}/select`,payload,otherCookie,otherCsrf)).status,404);
  assert.equal((await post(`/api/jobs/${job.id}/select`,{...payload,bounds:{x:0}})).status,400);
  assert.equal((await post(`/api/jobs/${job.id}/select`,{...payload,pageMode:'paginate'})).status,400);
  assert.equal((await post(`/api/jobs/${job.id}/select`,{...payload,choiceId:'c-forged'})).status,409);
  const selected=await post(`/api/jobs/${job.id}/select`,payload),summary=await selected.json();assert.equal(selected.status,200,JSON.stringify(summary));assert.equal(summary.slideCount,1);assert.equal(summary.exportAvailable,true);assert.equal(summary.pagination,undefined);
  const exported=await post('/api/export',{jobId:job.id}),out=await exported.json();assert.equal(exported.status,200,JSON.stringify(out));assert.ok(out.downloadUrl);
  assert.equal((await fetch(app.url+out.downloadUrl,{headers:{cookie}})).status,200);
  const again=post(`/api/jobs/${job.id}/select`,payload);
  let processing=false;
  for(let i=0;i<100;i++){const r=await fetch(app.url+`/api/jobs/${job.id}`,{headers:{cookie}}),value=await r.json();if(value.status==='analyzing'){processing=true;break;}await new Promise(resolve=>setTimeout(resolve,100));}
  assert.equal(processing,true,'selection should enter a processing state');
  assert.equal((await fetch(app.url+out.downloadUrl,{headers:{cookie}})).status,409,'a previous download must be invalidated during re-selection');
  const deleted=await fetch(app.url+`/api/jobs/${job.id}`,{method:'DELETE',headers:{cookie,'x-csrf-token':csrf}});assert.equal(deleted.status,204);
  await again.catch(()=>{});
  assert.equal((await fetch(app.url+`/api/jobs/${job.id}`,{headers:{cookie}})).status,404);
});

test('an expired candidate job loses both its metadata and preview',async()=>{
  const response=await post('/api/inspect',{html:'<body><section style="width:400px;height:300px">Candidate</section></body>'});
  const job=await response.json();assert.equal(response.status,200,JSON.stringify(job));
  const ownerId=cookie.split('=')[1].split('.')[0],stored=app.jobs.get(ownerId,job.id);assert.ok(stored);
  stored.expiresAt=Date.now()-1;
  assert.equal((await fetch(app.url+job.previewUrl,{headers:{cookie}})).status,404);
  assert.equal((await fetch(app.url+`/api/jobs/${job.id}`,{headers:{cookie}})).status,404);
});

test('a pre-cancelled request token cannot create an orphaned candidate job',async()=>{
  const token='a'.repeat(32),url=app.url+`/api/requests/${token}`;
  const cancelled=await fetch(url,{method:'DELETE',headers:{cookie,'x-csrf-token':csrf}});assert.equal(cancelled.status,204);
  const response=await fetch(app.url+'/api/inspect',{method:'POST',headers:{cookie,'content-type':'application/json','x-csrf-token':csrf,'x-job-request-id':token},body:JSON.stringify({html:'<section style="width:400px;height:300px">PAGE</section>'})});
  assert.equal(response.status,409);assert.equal(app.jobs.size,0);
});

test('request tokens are owner-scoped, single-use and remove completed jobs',async()=>{
  const token='b'.repeat(32),headers={cookie,'content-type':'application/json','x-csrf-token':csrf,'x-job-request-id':token};
  const response=await fetch(app.url+'/api/inspect',{method:'POST',headers,body:JSON.stringify({html:'<section style="width:400px;height:300px">PAGE</section>'})});
  const job=await response.json();assert.equal(response.status,200,JSON.stringify(job));
  const duplicate=await fetch(app.url+'/api/inspect',{method:'POST',headers,body:JSON.stringify({html:'<section>OTHER</section>'})});assert.equal(duplicate.status,409);
  const other=await fetch(app.url+'/api/session'),otherCookie=other.headers.get('set-cookie').split(';')[0],otherCsrf=(await other.json()).csrfToken;
  const alien=await fetch(app.url+`/api/requests/${token}`,{method:'DELETE',headers:{cookie:otherCookie,'x-csrf-token':otherCsrf}});assert.equal(alien.status,204);
  assert.equal((await fetch(app.url+`/api/jobs/${job.id}`,{headers:{cookie}})).status,200,'other owner cannot remove this job');
  const badCsrf=await fetch(app.url+`/api/requests/${token}`,{method:'DELETE',headers:{cookie,'x-csrf-token':'wrong'}});assert.equal(badCsrf.status,403);
  const cancel=()=>fetch(app.url+`/api/requests/${token}`,{method:'DELETE',headers:{cookie,'x-csrf-token':csrf}});
  assert.equal((await cancel()).status,204);assert.equal((await cancel()).status,204);
  assert.equal((await fetch(app.url+`/api/jobs/${job.id}`,{headers:{cookie}})).status,404);
  const reused=await fetch(app.url+'/api/inspect',{method:'POST',headers,body:JSON.stringify({html:'<section>REUSE</section>'})});assert.equal(reused.status,409);
});

test('cancelling a token while its worker runs prevents a late orphan',async()=>{
  const token='c'.repeat(32),headers={cookie,'content-type':'application/json','x-csrf-token':csrf,'x-job-request-id':token};
  const pending=fetch(app.url+'/api/inspect',{method:'POST',headers,body:JSON.stringify({html:'<section style="width:400px;height:300px">PAGE</section>'})});
  let started=false;for(let i=0;i<100;i++){if(app.jobs.size>0){started=true;break;}await new Promise(resolve=>setTimeout(resolve,50));}
  assert.equal(started,true,'worker job should start before cancellation');
  const cancel=await fetch(app.url+`/api/requests/${token}`,{method:'DELETE',headers:{cookie,'x-csrf-token':csrf}});assert.equal(cancel.status,204);
  await pending.catch(()=>{});assert.equal(app.jobs.size,0);
  const replay=await fetch(app.url+'/api/inspect',{method:'POST',headers,body:JSON.stringify({html:'<section>REPLAY</section>'})});assert.equal(replay.status,409);
});


test('many retired cancellation tokens do not consume active request capacity',async()=>{
  for(let i=0;i<130;i++){
    const token=(i+1).toString(16).padStart(32,'0');
    const response=await fetch(app.url+`/api/requests/${token}`,{method:'DELETE',headers:{cookie,'x-csrf-token':csrf}});
    assert.equal(response.status,204,`cancel token ${i}`);
  }
  const fresh='d'.repeat(32);
  const response=await fetch(app.url+'/api/inspect',{method:'POST',headers:{cookie,'content-type':'application/json','x-csrf-token':csrf,'x-job-request-id':fresh},body:'{}'});
  assert.equal(response.status,400,'retired tokens should not fill the active-request budget');
});

test('invalid inspection inputs expose safe specific diagnostics with a 4xx status',async()=>{
  const empty=await post('/api/inspect',{html:''}),emptyBody=await empty.json();
  assert.equal(empty.status,400,JSON.stringify(emptyBody));
  const missing=await post('/api/inspect',{html:'<section><img src="private-asset-name.png"></section>'}),missingBody=await missing.json();
  assert.equal(missing.status,400,JSON.stringify(missingBody));
  assert.ok(missingBody.diagnostics.some(d=>d.code==='ASSET_MISSING'),JSON.stringify(missingBody));
  assert.doesNotMatch(JSON.stringify(missingBody),/private-asset-name/);
});


test('invalid completed requests do not accumulate against active token capacity',async()=>{
  for(let i=0;i<130;i++){
    const token=(i+1000).toString(16).padStart(32,'0');
    const response=await fetch(app.url+'/api/inspect',{method:'POST',headers:{cookie,'content-type':'application/json','x-csrf-token':csrf,'x-job-request-id':token},body:'{}'});
    assert.equal(response.status,400,`invalid request ${i} must not consume the active budget`);
  }
});
