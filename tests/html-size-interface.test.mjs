import test,{before,after} from 'node:test';
import assert from 'node:assert/strict';
import {chromium} from 'playwright-core';
import {createHtmlServer} from '../src/html-to-pptx/server.mjs';
import {runConversion} from '../src/html-to-pptx/job-runner.mjs';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import {execFile} from 'node:child_process';
import {promisify} from 'node:util';

let app,browser;
before(async()=>{app=await createHtmlServer({port:0});browser=await chromium.launch({channel:'msedge',headless:true});});
after(async()=>{await browser?.close();await app?.close();});

test('slide size control sends an explicit policy, reports the applied size and clears an old result on change',async()=>{
 const context=await browser.newContext(),page=await context.newPage();let submitted;page.setDefaultTimeout(3000);
 try{
  await page.route('**/api/analyze',route=>{submitted=route.request().postDataJSON();return route.fulfill({json:{id:'size-job',sourceType:'html',targetType:'pptx',status:'analyzed',slideCount:1,slides:[],exportAvailable:true,partialExportAvailable:false,diagnostics:[],sizePolicy:{mode:'contain',targetWidthEmu:12192000,targetHeightEmu:6858000,marginColor:'#FFFFFF',source:'custom',slides:[{slideId:'slide-1',sourceWidth:800,sourceHeight:600,scale:1.5,offsetX:200,offsetY:0}]}}});});
  await page.route('**/api/jobs/size-job',route=>route.fulfill({status:204}));
  await page.goto(app.url);
  await page.locator('#html-options').evaluate(el=>el.open=true);
  await page.locator('#slide-size-policy').selectOption('wide');
  await page.locator('#html-file').setInputFiles({name:'slide.html',mimeType:'text/html',buffer:Buffer.from('<section data-slide style="width:800px;height:600px">Page</section>')});
  await page.waitForFunction(()=>document.querySelector('#status-badge').dataset.state==='analyzed');
  assert.deepEqual(submitted.sizePolicy,{mode:'contain',targetWidthEmu:12192000,targetHeightEmu:6858000});
  assert.match(await page.locator('#slide-size-report').innerText(),/비율.*유지/);
  assert.match(await page.locator('#slide-size-report').innerText(),/여백/);
  assert.equal(await page.locator('#download-button').isEnabled(),true);
  await page.locator('#slide-size-policy').selectOption('standard');
  assert.equal(await page.locator('#download-button').isDisabled(),true);
  assert.equal(await page.locator('#slide-size-report').isHidden(),true);
  assert.equal(await page.locator('#slide-size-details').isHidden(),true);
 }finally{await context.close();}
});

test('policy input rejects malformed dimensions and extra fields before worker execution',async()=>{
 const html='<section data-slide style="width:400px;height:300px">Page</section>';
 for(const sizePolicy of [null,{}, {mode:'stretch'}, {mode:'contain',targetWidthEmu:1}, {mode:'contain',targetWidthEmu:0,targetHeightEmu:1}, {mode:'contain',targetWidthEmu:51206401,targetHeightEmu:1}, {mode:'contain',marginColor:'red'}, {mode:'contain',unexpected:true}]){
  await assert.rejects(runConversion({operation:'inspect-html',input:{html,sizePolicy}}),error=>error.code==='INPUT_INVALID');
 }
});
test('custom size accepts the one and 56 inch boundaries and disables hidden invalid fields',async()=>{
 const context=await browser.newContext(),page=await context.newPage();let submitted;page.setDefaultTimeout(3000);
 try{
  await page.route('**/api/analyze',route=>{submitted=route.request().postDataJSON();return route.fulfill({json:{id:'size-job',sourceType:'html',targetType:'pptx',status:'analyzed',slideCount:1,slides:[],exportAvailable:true,partialExportAvailable:false,diagnostics:[]}});});
  await page.route('**/api/jobs/size-job',route=>route.fulfill({status:204}));
  await page.goto(app.url);await page.locator('#html-options').evaluate(el=>el.open=true);await page.locator('#slide-size-policy').selectOption('custom');
  await page.locator('#slide-width').fill('0.99');await page.locator('#slide-height').fill('56');assert.equal(await page.locator('#slide-width').evaluate(el=>el.checkValidity()),false);
  await page.locator('#slide-size-policy').selectOption('auto');assert.equal(await page.locator('#slide-width').isDisabled(),true);
  await page.locator('#slide-size-policy').selectOption('custom');await page.locator('#slide-width').fill('1');await page.locator('#html-file').setInputFiles({name:'slide.html',mimeType:'text/html',buffer:Buffer.from('<section data-slide>Page</section>')});
  await page.waitForFunction(()=>document.querySelector('#status-badge').dataset.state==='analyzed');assert.deepEqual(submitted.sizePolicy,{mode:'contain',targetWidthEmu:914400,targetHeightEmu:51206400});
 }finally{await context.close();}
});

test('UI and CLI explain how to fix a content-box candidate and direct selector',async t=>{
 const html='<!doctype html><html><body style="margin:0"><section style="width:400px;height:300px;padding:10px;background:white">Page</section></body></html>';
 const dir=await fs.mkdtemp(path.join(os.tmpdir(),'ppt-box-help-')),input=path.join(dir,'input.html'),output=path.join(dir,'output.pptx');
 t.after(async()=>{await fs.unlink(input).catch(()=>{});await fs.unlink(output).catch(()=>{});await fs.rmdir(dir).catch(()=>{});});await fs.writeFile(input,html);
 const context=await browser.newContext(),page=await context.newPage();page.setDefaultTimeout(120000);
 try{
  await page.goto(app.url);await page.locator('#html-file').setInputFiles(input);
  await page.waitForFunction(()=>['selection-required','failed'].includes(document.querySelector('#status-badge').dataset.state));
  assert.equal(await page.locator('#status-badge').getAttribute('data-state'),'selection-required');
  assert.match(await page.locator('#region-list').innerText(),/box-sizing:border-box/);assert.match(await page.locator('#region-list').innerText(),/패딩과 테두리를 제거/);
  await page.locator('#html-options').evaluate(el=>el.open=true);await page.locator('#selector').fill('section');await page.locator('#selector').dispatchEvent('change');await page.locator('#analyze-button').click();
  await page.waitForFunction(()=>document.querySelector('#status-badge').dataset.state==='failed');
  await page.locator('#diagnostics > summary').click();await page.locator('#diagnostic-list details > summary').first().click();assert.match(await page.locator('#diagnostic-list').innerText(),/box-sizing:border-box/);
  const run=promisify(execFile),options={timeout:120000,maxBuffer:1024*1024};
  const inspected=await run(process.execPath,['src/html-to-pptx/cli.mjs','inspect',input],options);assert.match(inspected.stdout,/box-sizing:border-box/);assert.match(inspected.stdout,/패딩과 테두리를 제거/);
  await assert.rejects(run(process.execPath,['src/html-to-pptx/cli.mjs','convert',input,'--output',output,'--selector','section'],options),error=>/box-sizing:border-box/.test(error.stderr)&&/패딩과 테두리를 제거/.test(error.stderr));
 }finally{await context.close();}
});

test('server applies viewport consistently to direct analysis, inspection and selection',async()=>{
 const session=await fetch(app.url+'/api/session'),cookie=session.headers.get('set-cookie').split(';')[0],csrf=(await session.json()).csrfToken;
 const post=(route,body)=>fetch(app.url+route,{method:'POST',headers:{cookie,'content-type':'application/json','x-csrf-token':csrf},body:JSON.stringify(body)});
 const html='<!doctype html><html><head><style>section{position:relative;width:400px;height:300px;background:white}@media(min-width:900px){section{width:600px}}</style></head><body style="margin:0"><section>Page</section></body></html>';
 const viewport={width:800,height:600},base={html,slideSelector:'section',viewport};
 let response=await post('/api/analyze',base),direct=await response.json();assert.equal(response.status,200,JSON.stringify(direct));assert.equal(direct.sizePolicy.slides[0].sourceWidth,400);
 response=await post('/api/inspect',base);const inspected=await response.json();assert.equal(response.status,200,JSON.stringify(inspected));
 const choice=inspected.inspection.choices.find(c=>c.scope==='elements'&&c.convertible);assert.ok(choice,JSON.stringify(inspected.inspection.choices));
 response=await post(`/api/jobs/${inspected.id}/select`,{snapshotId:inspected.inspection.snapshotId,choiceId:choice.id});const selected=await response.json();assert.equal(response.status,200,JSON.stringify(selected));assert.equal(selected.sizePolicy.slides[0].sourceWidth,400);
 response=await post('/api/analyze',{...base,viewport:{width:1000,height:600}});const wider=await response.json();assert.equal(response.status,200,JSON.stringify(wider));assert.equal(wider.sizePolicy.slides[0].sourceWidth,600);
 response=await post('/api/inspect',{...base,viewport:{width:1000,height:600}});const widerInspection=await response.json();assert.equal(response.status,200,JSON.stringify(widerInspection));assert.notEqual(widerInspection.inspection.snapshotId,inspected.inspection.snapshotId);
});
