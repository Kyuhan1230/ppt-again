import test,{before,after} from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import {chromium} from 'playwright-core';
import {createHtmlServer} from '../src/html-to-pptx/server.mjs';
import {readPptx} from '../src/html-to-pptx/read-pptx.mjs';
import {inspectPptxIndependently} from './inspect-pptx.mjs';
const chart={type:'bar',categories:['Review','Plan'],series:[{name:'Total',values:[8,0],color:'#164e63'}],title:'Current values',legend:true,axes:{min:0,max:20,majorUnit:4}};
const source=`<style>section{position:relative;width:800px;height:600px;background:white;font:16px Arial}p{position:absolute;left:40px;top:20px;margin:0}table{position:absolute;left:40px;top:100px;width:250px;height:150px;border-collapse:collapse;table-layout:fixed}td{border:1px solid #334455;padding:6px}figure{position:absolute;left:340px;top:90px;width:420px;height:310px;margin:0;font:16px Arial}</style><section data-slide><p id="fixed">Current review</p><table id="table"><tr><td>Amount</td><td>120</td></tr><tr><td>Difference</td><td>0</td></tr></table><figure data-chart-id="chart"></figure><script type="application/json" data-chart-for="chart">${JSON.stringify(chart)}</script></section>`;
let app,browser;
before(async()=>{app=await createHtmlServer({port:0});browser=await chromium.launch({channel:'msedge',headless:true});});
after(async()=>{await browser?.close();await app?.close();});
async function select(page,name,buffer){await page.locator('#html-file').setInputFiles({name,mimeType:name.endsWith('.pptx')?'application/vnd.openxmlformats-officedocument.presentationml.presentation':'text/html',buffer:Buffer.from(buffer)});}
async function analyze(page){await page.waitForFunction(()=>['analyzed','failed'].includes(document.getElementById('status-badge').dataset.state));assert.equal(await page.locator('#status-badge').getAttribute('data-state'),'analyzed',await page.locator('#diagnostic-list').textContent());}
async function download(page){const pending=page.waitForEvent('download');await page.locator('#download-button').click();const file=await pending,bytes=await fs.readFile(await file.path());await page.waitForFunction(()=>document.getElementById('status-badge').dataset.state==='generated');return {name:file.suggestedFilename(),bytes};}
test('large embedded decks show a waiting notice during analysis and clear it for the next file',async()=>{
 const context=await browser.newContext(),page=await context.newPage();let release;
 try{
  await page.route('**/api/analyze',async route=>{if(route.request().postDataJSON().fileName==='large.html')await new Promise(resolve=>{release=resolve;});await route.fulfill({json:{id:'synthetic',fileName:route.request().postDataJSON().fileName,sourceType:'html',targetType:'pptx',status:'analyzed',slideCount:route.request().postDataJSON().fileName==='large.html'?21:route.request().postDataJSON().fileName==='twenty.html'?20:1,slides:[],exportAvailable:true,diagnostics:[]}});});
  await page.goto(app.url);const data=JSON.stringify({pages:Array.from({length:21},()=>({html:'<div></div>',source:'a'})),sources:{a:{styles:''}}});await select(page,'large.html',`<script type="application/json">${data}</script>`);
  await page.waitForFunction(()=>document.getElementById('duration-note')&&!document.getElementById('duration-note').hidden);assert.match(await page.locator('#duration-note').innerText(),/21장.*오래 걸릴/s);while(!release)await new Promise(resolve=>setTimeout(resolve,10));release();await analyze(page);
  await select(page,'twenty.html',source);await analyze(page);assert.equal(await page.locator('#duration-note').isHidden(),true);
  await select(page,'small.html',source);await analyze(page);assert.equal(await page.locator('#duration-note').isHidden(),true);
 }finally{release?.();await context.close();}
});
test('file selection starts analysis and switches from upload to preview automatically',async()=>{
 const context=await browser.newContext(),page=await context.newPage();try{
  await page.goto(app.url);assert.equal(await page.locator('.result-panel').isVisible(),false);
  await select(page,'preview.html',source);await page.waitForFunction(()=>['analyzed','failed'].includes(document.getElementById('status-badge').dataset.state));
  assert.equal(await page.locator('#status-badge').getAttribute('data-state'),'analyzed');assert.equal(await page.locator('.result-panel').isVisible(),true);assert.equal(await page.locator('#download-button').isEnabled(),true);assert.equal(await page.locator('#preview-area').isVisible(),true);
  await page.locator('#delete-button').click();await page.waitForFunction(()=>document.getElementById('status-badge').dataset.state==='deleted');
 }finally{await context.close();}
});
test('dropping a single file starts analysis and multiple files cannot retain an old download',async()=>{
 const context=await browser.newContext(),page=await context.newPage();try{
  await page.goto(app.url);await page.locator('.file-area').dispatchEvent('drop',{dataTransfer:await page.evaluateHandle(html=>{const data=new DataTransfer();data.items.add(new File([html],'drop.html',{type:'text/html'}));return data;},source)});
  await analyze(page);assert.equal(await page.locator('#download-button').isEnabled(),true);
  const deleted=page.waitForResponse(r=>r.request().method()==='DELETE'&&r.url().includes('/api/jobs/'));
  await page.locator('.file-area').dispatchEvent('drop',{dataTransfer:await page.evaluateHandle(()=>{const data=new DataTransfer();data.items.add(new File(['a'],'a.html'));data.items.add(new File(['b'],'b.html'));return data;})});await deleted;
  assert.equal(await page.locator('#download-button').isDisabled(),true);assert.match(await page.locator('#status-message').textContent(),/한 개/);
 }finally{await context.close();}
});
test('repeated warnings are grouped into slide ranges without hiding errors',async()=>{
 const context=await browser.newContext(),page=await context.newPage();try{
  await page.route('**/api/analyze',route=>route.fulfill({json:{id:'synthetic',fileName:'sample.html',sourceType:'html',targetType:'pptx',status:'failed',slideCount:25,slides:[],exportAvailable:false,diagnostics:[...Array.from({length:25},(_,i)=>({code:'FONT_FORMAT_UNSUPPORTED',severity:'warning',message:'글꼴 포함 제한',slideNumber:i+1})),{code:'UNSUPPORTED_OBJECT',severity:'error',message:'표현 오류',slideNumber:3}]}}));
  await page.goto(app.url);await select(page,'sample.html',source);await page.waitForFunction(()=>document.getElementById('status-badge').dataset.state==='failed');
  const entries=page.locator('#diagnostic-list li');assert.equal(await entries.count(),2);assert.match(await entries.nth(0).innerText(),/1–25번 장표/);assert.match(await entries.nth(1).innerText(),/확인 필요.*표현 오류/s);assert.equal(await page.locator('#download-button').isDisabled(),true);
 }finally{await context.close();}
});
test('partial export button describes omitted slides and downloads the supported slides',async()=>{
 const context=await browser.newContext({acceptDownloads:true}),page=await context.newPage();try{
  const html='<section data-slide style="width:800px;height:600px;font:20px Arial"><p>Good</p></section><section data-slide style="width:800px;height:600px;font:20px Arial"><p style="filter:blur(2px)">Blocked</p></section>';
  await page.goto(app.url);await select(page,'partial.html',html);await page.waitForFunction(()=>document.getElementById('status-badge').dataset.state==='failed');
  assert.equal(await page.locator('#download-button').isDisabled(),true);assert.match(await page.locator('#partial-download-button').innerText(),/1장/);assert.match(await page.locator('#partial-download-note').innerText(),/2번 장표/);
  const pending=page.waitForEvent('download');await page.locator('#partial-download-button').click();const file=await pending;assert.match(file.suggestedFilename(),/-partial\.pptx$/);await page.waitForFunction(()=>document.getElementById('status-badge').dataset.state==='generated');assert.match(await page.locator('#status-message').innerText(),/1장/);assert.match(await page.locator('#status-message').innerText(),/포함되지 않았습니다/);
  await page.locator('#delete-button').click();await page.waitForFunction(()=>document.getElementById('status-badge').dataset.state==='deleted');
 }finally{await context.close();}
});
test('one file screen downloads editable tables/charts in both directions and retains current HTML edits',async()=>{
 const context=await browser.newContext({acceptDownloads:true}),page=await context.newPage();
 try{
  await page.goto(app.url);await select(page,'review.html',source);assert.match(await page.locator('#conversion-direction').textContent(),/HTML → PPTX/);await analyze(page);
  const ppt=await download(page);assert.equal(ppt.name,'review.pptx');const audit=await inspectPptxIndependently(ppt.bytes);assert.equal(audit.nativeTableCount,1);assert.equal(audit.nativeChartCount,1);assert.equal(audit.pictureCount,0);
  const original=await readPptx(ppt.bytes);assert.equal(original.status,'parsed');
  await select(page,'revised.pptx',ppt.bytes);assert.match(await page.locator('#conversion-direction').textContent(),/PPTX → HTML/);assert.equal(await page.locator('#html-options').isHidden(),true);await analyze(page);assert.equal(await page.locator('#download-button').textContent(),'HTML 다운로드');
  const html=await download(page);assert.equal(html.name,'revised.html');const text=html.bytes.toString('utf8');assert.match(text,/script-src 'none'/);assert.match(text,/data-chart-for/);assert.doesNotMatch(text,/<script[^>]*src=/);
  const edited=text.replace('>120<','>135<').replace(/(<script\b[^>]*data-chart-for[^>]*>)([\s\S]*?)(<\/script>)/g,(_,start,json,end)=>{const c=JSON.parse(json);c.series[0].values[1]=7;return start+JSON.stringify(c)+end;});assert.notEqual(edited,text);
  await select(page,'current.html',edited);assert.equal(await page.locator('#html-options').isVisible(),true);await analyze(page);const output=await download(page),current=await readPptx(output.bytes);assert.equal(current.status,'parsed',JSON.stringify(current.diagnostics));
  const nodes=current.model.slides[0].nodes;assert.equal(nodes.find(n=>n.kind==='table').content.cells[1].text,'135');assert.deepEqual(nodes.find(n=>n.kind==='chart').content.series[0].values,[8,7]);
  assert.deepEqual(nodes.find(n=>n.id==='fixed').content,original.model.slides[0].nodes.find(n=>n.id==='fixed').content);assert.deepEqual(nodes.find(n=>n.id==='fixed').boxCssPx,original.model.slides[0].nodes.find(n=>n.id==='fixed').boxCssPx);
  await page.locator('#delete-button').click();await page.waitForFunction(()=>document.getElementById('status-badge').dataset.state==='deleted');assert.equal(app.jobs.size,0);assert.equal(await page.locator('#download-button').isDisabled(),true);
 }finally{await context.close();}
});
test('invalid extensions and malformed PPT bytes cannot retain an old download and the screen retries',async()=>{
 const context=await browser.newContext(),page=await context.newPage();
 try{
  await page.goto(app.url);await select(page,'wrong.txt',source);assert.equal(await page.locator('#analyze-button').isDisabled(),true);assert.equal(await page.locator('#status-badge').getAttribute('data-state'),'failed');
  await select(page,'broken.pptx','not a zip');await page.waitForFunction(()=>document.getElementById('status-badge').dataset.state==='failed');assert.equal(await page.locator('#download-button').isDisabled(),true);assert.equal(await page.locator('#preview-area').isHidden(),true);
  await select(page,'retry.html',source);await analyze(page);assert.equal(await page.locator('#download-button').isDisabled(),false);
  await page.locator('#delete-button').click();await page.waitForFunction(()=>document.getElementById('status-badge').dataset.state==='deleted');assert.equal(app.jobs.size,0);
 }finally{await context.close();}
});
test('a cancelled slow file read cannot use the next analysis request signal',async()=>{
 const context=await browser.newContext(),page=await context.newPage(),sent=[];
 try{
  await page.addInitScript(()=>{const read=File.prototype.arrayBuffer;File.prototype.arrayBuffer=async function(){if(this.name==='slow.html')await new Promise(resolve=>{window.releaseOldRead=resolve;});const value=await read.call(this);if(this.name==='slow.html')window.oldReadDone=true;return value;};});
  page.on('request',request=>{if(request.url().endsWith('/api/analyze'))sent.push(request.postDataJSON().fileName);});
  await page.goto(app.url);await select(page,'slow.html',source);await page.waitForFunction(()=>typeof window.releaseOldRead==='function');
  await select(page,'current.html',source);await analyze(page);await page.evaluate(()=>window.releaseOldRead());await page.waitForFunction(()=>window.oldReadDone===true);
  // Allow the released file-read continuation to reach its potential network request.
  await page.evaluate(()=>new Promise(resolve=>setTimeout(resolve,100)));
  assert.deepEqual(sent,['current.html']);assert.equal(app.jobs.size,1);assert.match(await page.locator('#status-message').textContent(),/current\.html/);
  await page.locator('#delete-button').click();await page.waitForFunction(()=>document.getElementById('status-badge').dataset.state==='deleted');assert.equal(app.jobs.size,0);
 }finally{await context.close();}
});
