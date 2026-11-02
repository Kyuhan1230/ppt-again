import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import PptxGenJS from 'pptxgenjs';
import {runConversion} from '../src/html-to-pptx/job-runner.mjs';
import {createTestSession} from './http-session.mjs';
import {createHtmlServer,summarizeJob} from '../src/html-to-pptx/server.mjs';
const fixture=async()=>{const p=new PptxGenJS();p.layout='LAYOUT_WIDE';const s=p.addSlide();s.background={color:'FFFFFF'};s.addText('Import options: 현재 문구',{x:1,y:1,w:6,h:1,fontFace:'Arial',fontSize:20,margin:0,objectName:'heading'});return new Uint8Array(await p.write({outputType:'uint8array'}));};

test('worker retains legacy PPTX bytes and accepts validated options wrappers',async()=>{
 const bytes=await fixture();
 for(const input of [bytes,{bytes,options:{fontPolicy:'original',allowPartial:false}}]){const r=await runConversion({operation:'read-pptx',input});assert.equal(r.status,'parsed');assert.equal(r.model.slides[0].nodes[0].content.text,'Import options: 현재 문구');}
});
test('PPTX options cannot select arbitrary policy or alter byte-budget validation',async()=>{
 const bytes=await fixture();
 for(const options of [{fontPolicy:'arbitrary'},{allowPartial:'yes'},{allowPartial:true,skipSecurity:true},null])await assert.rejects(runConversion({operation:'read-pptx',input:{bytes,options}}),e=>e.code==='INPUT_INVALID');
 await assert.rejects(runConversion({operation:'analyze-pptx',input:{bytes:new Uint8Array([1,2]),options:{}}}),e=>e.code==='INPUT_INVALID');
});
test('PPTX HTML export remains available without a browser preview and keeps current text',async()=>{
 const bytes=await fixture();const analyzed=await runConversion({operation:'analyze-pptx',input:{bytes,options:{fontPolicy:'original',allowPartial:false}}});
 assert.ok(analyzed.measured.model);assert.equal(analyzed.measured.diagnostics.some(d=>d.severity==='error'),false,JSON.stringify(analyzed.measured.diagnostics));
 const result=await runConversion({operation:'export-html',input:analyzed.measured.model});assert.equal(result.status,'generated');assert.match(result.html,/Import options: 현재 문구/);
 if(analyzed.measured.previews.length===0)assert.ok(analyzed.measured.diagnostics.some(d=>d.reasonKey==='PREVIEW_UNAVAILABLE'&&d.severity==='warning'));
 const summary=summarizeJob({id:'fixture',fileName:'fixture.pptx',sourceType:'pptx',targetType:'html',input:analyzed.input,measured:analyzed.measured});assert.equal(summary.exportAvailable,true);assert.equal(summary.slideCount,1);
});
test('new UI settings are explicit and serialized; diagnostic details use text nodes',async()=>{
 const html=await fs.readFile(new URL('../web/html-to-pptx/index.html',import.meta.url),'utf8'),js=await fs.readFile(new URL('../web/html-to-pptx/app.js',import.meta.url),'utf8');
 for(const policy of ['original','malgun-gothic','pretendard'])assert.ok(html.includes(`value="${policy}"`));
 assert.match(js,/pptxOptions:\{fontPolicy:/);assert.match(js,/allowPartial:\$\('pptx-allow-partial'\)\.checked/);assert.match(js,/설정이 바뀌었습니다/);const diagnosticView=await fs.readFile(new URL('../web/html-to-pptx/diagnostic-view.js',import.meta.url),'utf8');assert.match(diagnosticView,/d\.objectName/);assert.ok(!diagnosticView.includes('innerHTML'));assert.ok(!js.includes('innerHTML'));
});

test('HTTP import applies the selected font and downloads HTML even without preview',async()=>{
 const app=await createHtmlServer({port:0});try{
  const request=await createTestSession(app.url),post=(route,body)=>request(route,{method:'POST',headers:{'content-type':'application/json'},body:JSON.stringify(body)});
  const bytes=await fixture(),body={sourceType:'pptx',fileName:'selected.pptx',pptxBase64:Buffer.from(bytes).toString('base64')};
  for(const fontPolicy of ['malgun-gothic','pretendard']){
   const response=await post('/api/analyze',{...body,pptxOptions:{fontPolicy,allowPartial:true}}),job=await response.json();assert.equal(response.status,200,JSON.stringify(job));assert.equal(job.exportAvailable,true);
   const exported=await post('/api/export',{jobId:job.id,target:'html'}),out=await exported.json();assert.equal(exported.status,200,JSON.stringify(out));
   const downloaded=await request(out.downloadUrl);assert.equal(downloaded.status,200);const html=await downloaded.text();assert.match(html,/Import options: 현재 문구/);assert.ok(html.includes(fontPolicy==='pretendard'?'Pretendard':'Malgun Gothic'));
  }
  for(const pptxOptions of [null,{fontPolicy:'unknown'},{allowPartial:'yes'},{skipSecurity:true}])assert.equal((await post('/api/analyze',{...body,pptxOptions})).status,400);
 }finally{await app.close();}
});

test('file selection rejects oversize before reading and recovers on a valid file',async()=>{
 const {FILE_LIMIT_BYTES,selectedFileState}=await import('../web/html-to-pptx/input-policy.js');
 const limit=FILE_LIMIT_BYTES.pptx;assert.equal(limit,64*1024*1024);
 assert.equal(selectedFileState({name:'exact.pptx',size:limit}).valid,true);
 const bad=selectedFileState({name:'large.pptx',size:limit+1});assert.equal(bad.valid,false);assert.match(bad.message,/67,108,865/);assert.match(bad.message,/64MiB/);
 assert.equal(selectedFileState({name:'valid.pptx',size:1024}).valid,true);assert.equal(selectedFileState({name:'html.html',size:FILE_LIMIT_BYTES.html+1}).valid,false);
 const js=await fs.readFile(new URL('../web/html-to-pptx/app.js',import.meta.url),'utf8');assert.match(js,/selectedFileState.*valid\)\$\('analyze-form'\)\.requestSubmit/);assert.match(js,/dispatchEvent\(new Event\('change'/);
});
