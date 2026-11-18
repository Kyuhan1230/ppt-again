import test from 'node:test';
import assert from 'node:assert/strict';
import {spawnSync} from 'node:child_process';
import {createHtmlServer,summarizeJob} from '../src/html-to-pptx/server.mjs';
import {runConversion} from '../src/html-to-pptx/job-runner.mjs';
import {publicDiagnostic} from '../src/html-to-pptx/public-diagnostics.mjs';
import {validSizePolicy} from '../src/html-to-pptx/size-policy-input.mjs';

test('rejects invalid size policies before launching the worker',async()=>{
 for(const sizePolicy of [null,{}, {mode:'stretch'}, {mode:'contain',targetWidthEmu:1}, {mode:'contain',targetWidthEmu:0,targetHeightEmu:1}, {mode:'contain',targetWidthEmu:51206401,targetHeightEmu:1}, {mode:'contain',marginColor:'red'}, {mode:'contain',unexpected:true}]){
  await assert.rejects(runConversion({operation:'inspect-html',input:{html:'<p>test</p>',sizePolicy}}),error=>error.code==='INPUT_INVALID');
 }
});
test('ordinary HTML analysis rejects an invalid viewport before launching the worker',async()=>{
 for(const viewport of [null,{}, {width:0,height:600},{width:800,height:9000},{width:800,height:600,extra:true}]){
  await assert.rejects(runConversion({operation:'analyze-html',input:{html:'<p>test</p>',viewport}}),error=>error.code==='INPUT_INVALID');
 }
});
test('PPT dimensions must be between exactly one and 56 inches per axis',()=>{
 for(const value of [914399,51206401])assert.equal(validSizePolicy({mode:'contain',targetWidthEmu:value,targetHeightEmu:914400}),false);
 for(const value of [914400,51206400])assert.equal(validSizePolicy({mode:'contain',targetWidthEmu:value,targetHeightEmu:914400}),true);
});
test('CLI validates output size and color before reading input',()=>{
 for(const args of [['--slide-size','57x3'],['--slide-size','4x0'],['--slide-size','0.99x1'],['--slide-size','1x0.99'],['--slide-size','fit'],['--margin-color','red']]){
  const result=spawnSync(process.execPath,['src/html-to-pptx/cli.mjs','convert','missing.html','--output','missing.pptx',...args],{encoding:'utf8'});
  assert.equal(result.status,2);assert.match(result.stderr,/PPT 크기/);
 }
 for(const size of ['1x56','56x1']){const result=spawnSync(process.execPath,['src/html-to-pptx/cli.mjs','convert','missing.html','--output','missing.pptx','--slide-size',size],{encoding:'utf8'});assert.equal(result.status,2);assert.match(result.stderr,/입력 파일을 찾을 수 없습니다/);}
});
test('public policy report preserves calculation without exposing source slide IDs',()=>{
 const job={id:'test',fileName:'a.html',sourceType:'html',targetType:'pptx',input:{profile:'html',source:{sha256:'digest'}},measured:{diagnostics:[],model:{slides:[{id:'private-id',nodes:[]}]},previews:[],environment:{previewKind:'source'},sizePolicy:{mode:'contain',source:'custom',targetWidthEmu:12192000,targetHeightEmu:6858000,marginColor:'#ffffff',slides:[{slideId:'private-id',sourceWidth:800,sourceHeight:600,scale:1.5,offsetX:200,offsetY:0}]}}};
 const result=summarizeJob(job);assert.equal(result.sizePolicy.slides[0].slideId,'slide-1');assert.equal(result.sizePolicy.slides[0].offsetX,200);assert.equal(result.previewKind,'source');assert.ok(!JSON.stringify(result).includes('private-id'));
});
test('server rejects invalid size settings at both HTML entry points',async t=>{
 const app=await createHtmlServer({port:0});t.after(()=>app.close());
 const session=await fetch(app.url+'/api/session'),cookie=session.headers.get('set-cookie').split(';')[0],csrf=(await session.json()).csrfToken;
 for(const route of ['/api/analyze','/api/inspect']){
  const response=await fetch(app.url+route,{method:'POST',headers:{cookie,'content-type':'application/json','x-csrf-token':csrf},body:JSON.stringify({html:'<p>x</p>',sizePolicy:{mode:'contain',targetWidthEmu:914399,targetHeightEmu:914400}})});
  assert.equal(response.status,400);
  const invalidViewport=await fetch(app.url+route,{method:'POST',headers:{cookie,'content-type':'application/json','x-csrf-token':csrf},body:JSON.stringify({html:'<p>x</p>',viewport:null})});
  assert.equal(invalidViewport.status,400);
 }
});
test('content-box selection error retains the actionable fix in public diagnostics',()=>{
 const message='장표 루트의 content-box에 패딩이나 테두리가 있습니다. box-sizing:border-box를 사용하거나 패딩과 테두리를 제거해 주세요.';
 const result=publicDiagnostic({code:'UNSUPPORTED_OBJECT',severity:'error',message});
 assert.match(result.message,/box-sizing:border-box/);assert.match(result.message,/패딩과 테두리를 제거/);
});
