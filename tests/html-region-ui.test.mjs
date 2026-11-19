import test,{before,after} from 'node:test';
import assert from 'node:assert/strict';
import {chromium} from 'playwright-core';
import {createHtmlServer} from '../src/html-to-pptx/server.mjs';

let app,browser;
before(async()=>{app=await createHtmlServer({port:0});browser=await chromium.launch({channel:'msedge',headless:true});});
after(async()=>{await browser?.close();await app?.close();});
const summary={id:'finished',fileName:'cards.html',sourceType:'html',targetType:'pptx',status:'analyzed',slideCount:1,slides:[],exportAvailable:true,partialExportAvailable:false,diagnostics:[]};
const inspection={snapshotId:'snapshot-1',regions:[
  {id:'r1',bounds:{x:0,y:0,width:400,height:300},visibility:'visible',parentId:null,overlaps:[],evidence:['visible-content']},
  {id:'r2',bounds:{x:450,y:0,width:400,height:300},visibility:'offscreen',parentId:null,overlaps:[],evidence:['visible-content']},
  {id:'r3',bounds:null,visibility:'hidden',parentId:'r1',overlaps:[],evidence:['hidden']}
],choices:[
  {id:'single',scope:'elements',regionIds:['r1'],evidence:['visible-content'],outsideContentCount:2,outsideContentCountAccuracy:'estimate',convertible:true,reasons:[]},
  {id:'pair',scope:'elements',regionIds:['r1','r2'],evidence:['repeated-sibling'],outsideContentCount:0,convertible:true,reasons:[]},
  {id:'whole',scope:'document',regionIds:[],bounds:{x:0,y:0,width:900,height:600},evidence:['document'],outsideContentCount:0,convertible:false,reasons:['different-size']}
],uncertainties:['hidden-content-not-observed']};
async function upload(page,name='cards.html'){
 await page.locator('#html-file').setInputFiles({name,mimeType:'text/html',buffer:Buffer.from('<div>cards</div>')});
}
async function mockFlow(page,onSelect=()=>summary){
 await page.route('**/api/analyze',route=>route.fulfill({json:{...summary,id:'initial',status:'failed',exportAvailable:false,diagnostics:[{code:'SELECTION_REQUIRED',severity:'error',message:'장표 범위를 선택해야 합니다.'}]}}));
 await page.route('**/api/inspect',route=>route.fulfill(Object.hasOwn(route.request().postDataJSON(),'sourceType')?{status:400,json:{error:'unexpected field'}}:{json:{id:'inspect-1',status:'selection-required',inspection,previewUrl:'/api/jobs/inspect-1/inspection.png'}}));
 await page.route('**/api/jobs/inspect-1/inspection.png',route=>route.fulfill({status:200,contentType:'image/svg+xml',body:'<svg xmlns="http://www.w3.org/2000/svg" width="900" height="600"/>'}));
 await page.route('**/api/jobs/inspect-1/select',async route=>{const result=await onSelect(route.request().postDataJSON());await route.fulfill({json:result}).catch(()=>{});});
}
test('candidate screen shows visual regions, hidden and offscreen explanations, and converts chosen scope',async()=>{
 const context=await browser.newContext({viewport:{width:1280,height:900}}),page=await context.newPage();let request;
 try{await mockFlow(page,body=>{request=body;return summary;});await page.goto(app.url);await upload(page);
  const region=page.locator('#region-area');await region.waitFor({state:'visible'});
  assert.match(await region.innerText(),/장표로 사용할 영역/);
  assert.match(await page.locator('#region-summary').innerText(),/화면 밖.*1/);
  assert.match(await page.locator('#region-uncertainties').innerText(),/숨/);
  assert.equal(await page.locator('#region-map-image').isVisible(),true);
  assert.equal(await page.locator('#region-use-only-button').isVisible(),false);
  assert.equal(await page.locator('#region-select-button').isVisible(),true);
  assert.ok(await page.locator('#region-map-overlays button').count()>=2);
  assert.equal(await page.locator('[data-choice-id="whole"]').isDisabled(),true);
  await page.locator('[data-choice-id="single"]').click();
  assert.equal(await page.locator('#region-select-button').isEnabled(),true);
  await page.locator('#region-select-button').click();await page.waitForFunction(()=>document.querySelector('#status-badge').dataset.state==='analyzed');
  assert.deepEqual(request,{snapshotId:'snapshot-1',choiceId:'single'});
  assert.equal(await page.locator('#download-button').isEnabled(),true);
  assert.equal(await region.isHidden(),true);
 }finally{await context.close();}
});
test('stale selection leaves candidate view available and prevents old output download',async()=>{
 const context=await browser.newContext(),page=await context.newPage();
 try{await mockFlow(page);await page.route('**/api/jobs/inspect-1/select',route=>route.fulfill({status:409,json:{error:'화면이 바뀌었습니다. 다시 분석해 주세요.'}}));
  await page.goto(app.url);await upload(page);await page.locator('#region-area').waitFor({state:'visible'});
  await page.locator('[data-choice-id="pair"]').click();await page.locator('#region-select-button').click();
  await page.waitForFunction(()=>document.querySelector('#status-message').textContent.includes('화면이 바뀌었습니다'));
  assert.equal(await page.locator('#region-area').isVisible(),true);assert.equal(await page.locator('#download-button').isDisabled(),true);
  assert.equal(await page.locator('#region-select-button').isEnabled(),true);
 }finally{await context.close();}
});
test('a clipped document can be reinspected at the suggested screen size',async()=>{
 const context=await browser.newContext(),page=await context.newPage(),viewports=[];let selection;
 try{
  await mockFlow(page,body=>{selection=body;return summary;});
  await page.route('**/api/inspect',route=>{
    const viewport=route.request().postDataJSON().viewport;
    viewports.push(viewport);
    const choices=inspection.choices.map(c=>({...c,convertible:Boolean(viewport&&c.id==='single'),reasons:viewport&&c.id==='single'?[]:['inspection-incomplete']}));
    const scanned={...inspection,snapshotId:viewport?'fitted':'clipped',choices,...(viewport?{}:{suggestedViewport:{width:1600,height:900}})};
    return route.fulfill({json:{id:'inspect-1',status:'selection-required',inspection:scanned,previewUrl:viewport?'/api/jobs/inspect-1/inspection.png':null}});
  });
  await page.goto(app.url);await upload(page);
  const fit=page.locator('#region-fit-button');await fit.waitFor({state:'visible'});
  assert.match(await fit.innerText(),/1600 × 900/);
  assert.equal(await page.locator('[data-choice-id="single"]').isDisabled(),true);
  await fit.click();
  const useOnly=page.getByRole('button',{name:'영역 1로 분석'});
  await useOnly.waitFor();
  assert.deepEqual(viewports,[undefined,{width:1600,height:900}]);
  assert.equal(await fit.isHidden(),true);
  assert.match(await page.locator('#region-preview-note').innerText(),/1600 × 900/);
  assert.equal(await page.locator('#region-map-overlays .region-map-outline').count(),1);
  assert.equal(await page.locator('#region-map-overlays button').count(),0);
  assert.equal(await page.locator('#region-select-button').isVisible(),false);
  assert.equal(await page.locator('#region-list').isVisible(),false);
  await page.locator('#region-review-details summary').click();
  assert.equal(await page.locator('[data-choice-id="pair"]').isVisible(),true);
  assert.equal(await page.locator('[data-choice-id="pair"]').isDisabled(),true);
  assert.equal(selection,undefined);
  await useOnly.click();
  await page.waitForFunction(()=>document.querySelector('#status-badge').dataset.state==='analyzed');
  assert.deepEqual(selection,{snapshotId:'fitted',choiceId:'single'});
  assert.equal(await page.locator('#download-button').isEnabled(),true);
 }finally{await context.close();}
});
for(const matches of [true,false])test((matches?'the same':'a different')+' fixed canvas is checked after the automatic viewport retry',async()=>{
 const context=await browser.newContext(),page=await context.newPage(),viewports=[];let selected;
 try{
  await mockFlow(page,body=>{selected=body;return summary;});
  await page.route('**/api/inspect',route=>{
    const viewport=route.request().postDataJSON().viewport;
    viewports.push(viewport);
    const choices=inspection.choices.map(c=>({...c,convertible:Boolean(viewport&&c.id==='single')}));
    const scanned={...inspection,choices,snapshotId:viewport?'fitted':'clipped',...(viewport?{
      verifiedViewportFrame:{key:matches?'frame-a':'frame-b',choiceId:'single'}
    }:{suggestedViewport:{width:1600,height:900},suggestedRegionKey:'frame-a'})};
    return route.fulfill({json:{id:viewport?'inspect-1':'clipped-1',status:'selection-required',inspection:scanned,previewUrl:null}});
  });
  await page.goto(app.url);await upload(page);
  if(matches){
    await page.waitForFunction(()=>document.querySelector('#status-badge').dataset.state==='analyzed',null,{timeout:5000});
    assert.deepEqual(selected,{snapshotId:'fitted',choiceId:'single'});
  }else{
    await page.getByRole('button',{name:'영역 1로 분석'}).waitFor({timeout:5000});
    assert.equal(selected,undefined);
    assert.equal(await page.locator('#download-button').isDisabled(),true);
  }
  assert.deepEqual(viewports,[undefined,{width:1600,height:900}]);
 }finally{await context.close();}
});
test('mobile candidate view fits the preview and offers the same choices by keyboard',async()=>{
 const context=await browser.newContext({viewport:{width:390,height:844},isMobile:true,hasTouch:true}),page=await context.newPage();
 try{await mockFlow(page);await page.goto(app.url);await upload(page);await page.locator('#region-area').waitFor({state:'visible'});
  const stage=await page.locator('#region-map-stage').boundingBox(),image=await page.locator('#region-map-image').boundingBox();
  assert.ok(image.width<=stage.width&&image.height<=stage.height,'full preview should fit the stage');
  const choice=page.locator('[data-choice-id="pair"]');await choice.focus();await page.keyboard.press('Enter');
  assert.equal(await choice.getAttribute('aria-pressed'),'true');assert.equal(await page.locator('#region-select-button').isEnabled(),true);
  if(process.env.UI_CHECK_SCREENSHOT)await page.screenshot({path:process.env.UI_CHECK_SCREENSHOT,fullPage:true});
 }finally{await context.close();}
});
test('changing file while a selection is pending never enables its old download',async()=>{
 const context=await browser.newContext(),page=await context.newPage();let release;
 try{await mockFlow(page,()=>new Promise(resolve=>{release=()=>resolve(summary);}));
  await page.route('**/api/analyze',route=>route.fulfill({json:route.request().postDataJSON().fileName==='new.html'?{...summary,id:'new',status:'failed',exportAvailable:false,diagnostics:[{code:'INPUT_INVALID',severity:'error',message:'입력 확인'}]}:{...summary,id:'initial',status:'failed',exportAvailable:false,diagnostics:[{code:'SELECTION_REQUIRED',severity:'error',message:'선택 필요'}]}}));
  await page.goto(app.url);await upload(page);await page.locator('#region-area').waitFor({state:'visible'});await page.locator('[data-choice-id="single"]').click();await page.locator('#region-select-button').click();
  await page.waitForFunction(()=>document.querySelector('#status-badge').dataset.state==='selecting');
  await upload(page,'new.html');await page.waitForFunction(()=>document.querySelector('#status-badge').dataset.state==='failed');
  release?.();await page.waitForTimeout(100);
  assert.equal(await page.locator('#download-button').isDisabled(),true);assert.equal(await page.locator('#region-area').isHidden(),true);assert.match(await page.locator('#status-message').textContent(),/항목|영역|확인/);
 }finally{release?.();await context.close();}
});
test('real server allows a candidate to be selected and downloaded as editable PPT',async()=>{
 const context=await browser.newContext({viewport:{width:1280,height:900},acceptDownloads:true}),page=await context.newPage();
 try{
  await page.goto(app.url);
  const html='<body style="margin:0"><section style="width:400px;height:300px;background:#eee"><p>Selected page</p></section><aside>Outside page</aside></body>';
  await page.locator('#html-file').setInputFiles({name:'live.html',mimeType:'text/html',buffer:Buffer.from(html)});
  await page.locator('#region-area').waitFor({state:'visible',timeout:120000});
  await page.locator('#region-map-image').waitFor({state:'visible'});
  const first=page.locator('#region-list button:not(:disabled)').first();await first.waitFor();
  if(process.env.UI_REAL_DESKTOP_SCREENSHOT)await page.screenshot({path:process.env.UI_REAL_DESKTOP_SCREENSHOT,fullPage:true});
  await page.setViewportSize({width:390,height:844});
  if(process.env.UI_REAL_MOBILE_SCREENSHOT)await page.screenshot({path:process.env.UI_REAL_MOBILE_SCREENSHOT,fullPage:true});
  await first.click();await page.locator('#region-select-button').click();
  await page.waitForFunction(()=>['analyzed','failed','selection-required'].includes(document.querySelector('#status-badge').dataset.state),null,{timeout:120000});
  assert.equal(await page.locator('#status-badge').getAttribute('data-state'),'analyzed',await page.locator('#status-message').textContent());
  const pending=page.waitForEvent('download',{timeout:120000});await page.locator('#download-button').click();
  const file=await pending,bytes=await (await import('node:fs/promises')).readFile(await file.path());
  const {inspectPptxIndependently}=await import('./inspect-pptx.mjs');const ppt=await inspectPptxIndependently(new Uint8Array(bytes));
  const texts=ppt.slides.flatMap(slide=>slide.objects).filter(object=>object.kind==='text').map(object=>object.text);
  assert.ok(texts.includes('Selected page'));assert.ok(!texts.includes('Outside page'));
  await page.waitForFunction(()=>document.querySelector('#status-badge').dataset.state==='generated');await page.locator('#delete-button').click();await page.waitForFunction(()=>document.querySelector('#status-badge').dataset.state==='deleted');
 }finally{await context.close();}
});
test('failed candidate analysis keeps selection available and displays diagnostic details',async()=>{
 const context=await browser.newContext(),page=await context.newPage();
 try{await mockFlow(page);await page.route('**/api/jobs/inspect-1/select',route=>route.fulfill({status:422,json:{...summary,id:'inspect-1',status:'failed',exportAvailable:false,diagnostics:[{code:'UNSUPPORTED_OBJECT',severity:'error',message:'지원되지 않는 효과가 있습니다.'}]}}));
  await page.goto(app.url);await upload(page);await page.locator('#region-area').waitFor({state:'visible'});
  await page.locator('[data-choice-id="single"]').click();await page.locator('#region-select-button').click();
  await page.waitForFunction(()=>!document.querySelector('#diagnostics').hidden,null,{timeout:5000});
  assert.equal(await page.locator('#region-area').isVisible(),true);assert.equal(await page.locator('#download-button').isDisabled(),true);
  await page.locator('#diagnostics > summary').click();await page.locator('#diagnostic-list details > summary').first().click();assert.match(await page.locator('#diagnostic-list').innerText(),/지원되지 않는 효과/);
  assert.equal(await page.locator('[data-choice-id="pair"]').isEnabled(),true);
 }finally{await context.close();}
});
test('candidate explanations identify unobserved content and why a scope cannot convert',async()=>{
 const context=await browser.newContext(),page=await context.newPage();
 try{await mockFlow(page);await page.route('**/api/inspect',route=>route.fulfill({json:{id:'inspect-1',status:'selection-required',previewUrl:null,inspection:{...inspection,uncertainties:['hidden-content-not-observed','clipped-or-scroll-content-not-observed','script-content-not-observed','generated-content-not-fully-observed','css-compositing-not-fully-observed','small-elements-not-proposed'],choices:[{...inspection.choices[2],reasons:['document-size-not-fixed','overlap-or-nesting','inspection-incomplete','ancestor-appearance-not-supported','external-overlay-not-supported','root-appearance-not-supported']}]}}}));
  await page.goto(app.url);await upload(page);await page.locator('#region-area').waitFor({state:'visible'});
  const limits=await page.locator('#region-uncertainties').innerText();assert.match(limits,/숨겨진/);assert.match(limits,/스크롤/);assert.match(limits,/스크립트/);assert.match(limits,/CSS로 생성/);assert.match(limits,/CSS 효과/);assert.match(limits,/작은 글자/);
  const unsupported=await page.locator('[data-choice-id="whole"]').innerText();assert.match(unsupported,/문서의 크기가 고정되지/);assert.match(unsupported,/겹치거나 포함 관계/);assert.match(unsupported,/화면이나 자산을 완전히 확인하지 못해/);assert.match(unsupported,/상위 요소에 있는 배경이나 효과/);assert.match(unsupported,/겹친 바깥 요소나 효과/);assert.match(unsupported,/선택 영역의 배경이나 테두리 효과/);
  assert.match(await page.locator('#region-preview-note').innerText(),/목록에서 범위/);
 }finally{await context.close();}
});
test('after successful selection the user can safely choose another area',async()=>{
 const context=await browser.newContext(),page=await context.newPage();let inspections=0,deletions=0;
 try{await mockFlow(page,()=>({...summary,id:'inspect-1'}));
  await page.route('**/api/inspect',route=>{inspections++;return route.fulfill({json:{id:'inspect-1',status:'selection-required',inspection,previewUrl:'/api/jobs/inspect-1/inspection.png'}});});
  await page.route('**/api/jobs/inspect-1',route=>{if(route.request().method()==='DELETE')deletions++;return route.fulfill({status:204});});
  await page.goto(app.url);await upload(page);await page.locator('#region-area').waitFor({state:'visible'});
  await page.locator('[data-choice-id="single"]').click();await page.locator('#region-select-button').click();
  await page.waitForFunction(()=>document.querySelector('#status-badge').dataset.state==='analyzed');
  assert.equal(await page.locator('#download-button').isEnabled(),true);
  const change=page.getByRole('button',{name:'다른 영역 선택'});await change.waitFor({state:'visible',timeout:5000});await change.click();
  assert.equal(await page.locator('#download-button').isDisabled(),true);
  await page.waitForFunction(()=>document.querySelector('#status-badge').dataset.state==='selection-required');
  assert.equal(inspections,2);assert.ok(deletions>=1);assert.equal(await page.locator('#region-area').isVisible(),true);
 }finally{await context.close();}
});
test('failed deletion blocks reinspection and retains the old job for retry',async()=>{
 const context=await browser.newContext(),page=await context.newPage();let deletes=0,inspections=0,analyses=0;
 try{await mockFlow(page,()=>({...summary,id:'inspect-1'}));
  await page.route('**/api/analyze',route=>{analyses++;return route.fulfill({json:{...summary,id:'initial',status:'failed',exportAvailable:false,diagnostics:[{code:'SELECTION_REQUIRED',severity:'error',message:'선택 필요'}]}});});
  await page.route('**/api/inspect',route=>{inspections++;return route.fulfill({json:{id:'inspect-1',status:'selection-required',inspection,previewUrl:'/api/jobs/inspect-1/inspection.png'}});});
  await page.route('**/api/jobs/inspect-1',route=>{deletes++;return route.fulfill(deletes===1?{status:500,json:{error:'삭제 실패'}}:{status:204});});
  await page.goto(app.url);await upload(page);await page.locator('#region-area').waitFor({state:'visible'});
  await page.locator('[data-choice-id="single"]').click();await page.locator('#region-select-button').click();await page.waitForFunction(()=>document.querySelector('#status-badge').dataset.state==='analyzed');
  await page.getByRole('button',{name:'다른 영역 선택'}).click();await page.waitForFunction(()=>document.querySelector('#status-badge').dataset.state==='failed');
  assert.equal(analyses,1);assert.equal(inspections,1);assert.equal(await page.locator('#download-button').isDisabled(),true);assert.equal(await page.locator('#delete-button').isEnabled(),true);
  await page.locator('#analyze-button').click();await page.waitForFunction(()=>document.querySelector('#status-badge').dataset.state==='selection-required');
  assert.equal(deletes,2);assert.equal(analyses,2);assert.equal(inspections,2);
 }finally{await context.close();}
});
test('changing files waits for earlier job deletion and allows retry after deletion fails',async()=>{
 const context=await browser.newContext(),page=await context.newPage();let deletes=0,analyses=0;
 try{await mockFlow(page,()=>({...summary,id:'inspect-1'}));
  await page.route('**/api/analyze',route=>{analyses++;const file=route.request().postDataJSON().fileName;return route.fulfill({json:file==='new.html'?{...summary,id:'new',fileName:file}:{...summary,id:'initial',status:'failed',exportAvailable:false,diagnostics:[{code:'SELECTION_REQUIRED',severity:'error',message:'선택 필요'}]}});});
  await page.route('**/api/jobs/inspect-1',route=>{deletes++;return route.fulfill(deletes===1?{status:500,json:{error:'삭제 실패'}}:{status:204});});
  await page.goto(app.url);await upload(page);await page.locator('#region-area').waitFor({state:'visible'});await page.locator('[data-choice-id="single"]').click();await page.locator('#region-select-button').click();
  await page.waitForFunction(()=>document.querySelector('#status-badge').dataset.state==='analyzed');await upload(page,'new.html');
  await page.waitForFunction(()=>document.querySelector('#status-badge').dataset.state==='failed',null,{timeout:5000});
  assert.equal(analyses,1);assert.equal(await page.locator('#download-button').isDisabled(),true);assert.equal(await page.locator('#delete-button').isEnabled(),true);
  await page.locator('#analyze-button').click();await page.waitForFunction(()=>document.querySelector('#status-badge').dataset.state==='analyzed');
  assert.equal(deletes,2);assert.equal(analyses,2);
 }finally{await context.close();}
});
test('cancel after server reply but before reading job ID removes the retained server job',async()=>{
 const context=await browser.newContext(),page=await context.newPage();
 try{
  await page.addInitScript(()=>{const read=Response.prototype.json;Response.prototype.json=async function(){if(this.url.endsWith('/api/analyze')){window.__answerWaiting=true;await new Promise(resolve=>{window.__releaseAnswer=resolve;});}return read.call(this);};});
  await page.goto(app.url);
  const html='<section data-slide style="width:400px;height:300px"><p>Sample</p></section>';
  await page.locator('#html-file').setInputFiles({name:'first.html',mimeType:'text/html',buffer:Buffer.from(html)});
  await page.waitForFunction(()=>window.__answerWaiting===true,null,{timeout:120000});assert.equal(app.jobs.size,1);
  await page.locator('#html-file').setInputFiles({name:'invalid.txt',mimeType:'text/plain',buffer:Buffer.from('x')});
  await page.evaluate(()=>window.__releaseAnswer());
  for(let i=0;i<100&&app.jobs.size;i++)await new Promise(resolve=>setTimeout(resolve,100));
  assert.equal(app.jobs.size,0,'cancelled operation must be removed even before its job ID is read');
 }finally{await context.close();}
});
test('cancel button cleans the server job before its response ID is read',async()=>{
 const context=await browser.newContext(),page=await context.newPage();
 try{
  await page.addInitScript(()=>{const read=Response.prototype.json;Response.prototype.json=async function(){if(this.url.endsWith('/api/analyze')){window.__answerWaiting=true;await new Promise(resolve=>{window.__releaseAnswer=resolve;});}return read.call(this);};});
  await page.goto(app.url);const html='<section data-slide style="width:400px;height:300px"><p>Sample</p></section>';
  await page.locator('#html-file').setInputFiles({name:'first.html',mimeType:'text/html',buffer:Buffer.from(html)});
  await page.waitForFunction(()=>window.__answerWaiting===true,null,{timeout:120000});assert.equal(app.jobs.size,1);
  await page.locator('#cancel-button').click();for(let i=0;i<100&&app.jobs.size;i++)await new Promise(resolve=>setTimeout(resolve,100));
  await page.evaluate(()=>window.__releaseAnswer());await page.waitForFunction(()=>document.querySelector('#status-badge').dataset.state==='deleted');
  assert.equal(app.jobs.size,0);assert.equal(await page.locator('#download-button').isDisabled(),true);
 }finally{await context.close();}
});
