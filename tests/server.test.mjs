import test,{before,after} from 'node:test';
import assert from 'node:assert/strict';
import http from 'node:http';
import {createHtmlServer} from '../src/html-to-pptx/server.mjs';
import {inspectPptxIndependently} from './inspect-pptx.mjs';
let app;
before(async()=>{app=await createHtmlServer({port:0});});
after(async()=>{await app?.close();});
const post=(route,body,headers={})=>fetch(app.url+route,{method:'POST',headers:{'content-type':'application/json',...headers},body:JSON.stringify(body)});

for(const example of [
  {file:'wide.html',slides:2,texts:2,ratio:16/9,title:'Current review',html:'<section data-slide style="width:960px;height:540px"><p style="font:24px Arial">Current review</p></section><section data-slide style="width:960px;height:540px"><p style="font:24px Arial">Next steps</p></section>'},
  {file:'standard.html',slides:1,texts:1,ratio:4/3,title:'Plan',html:'<svg xmlns="http://www.w3.org/2000/svg" width="800" height="600" viewBox="0 0 800 600"><text x="40" y="80" style="font:24px Arial">Plan</text></svg>'},
  {file:'portrait.html',slides:1,texts:1,ratio:3/4,title:'Notes',html:'<section data-slide style="width:600px;height:800px"><p style="font:24px Arial">Notes</p></section>'}
])test(`${example.file}: analyze, download and inspect editable output`,async()=>{
  const response=await post('/api/analyze',{html:example.html,fileName:example.file});
  const job=await response.json();assert.equal(response.status,200,JSON.stringify(job));assert.equal(job.slideCount,example.slides);
  const preview=await fetch(app.url+job.slides[0].previewUrl);assert.equal(preview.status,200);assert.equal(preview.headers.get('content-type'),'image/png');
  const output=await post('/api/export',{jobId:job.id}),result=await output.json();assert.equal(output.status,200,JSON.stringify(result));
  assert.equal(result.objectMapUrl,null);assert.equal(result.visualQuality,'unmeasured');
  const download=await fetch(app.url+result.downloadUrl);assert.equal(download.status,200);
  const bytes=new Uint8Array(await download.arrayBuffer()),ppt=await inspectPptxIndependently(bytes);
  assert.equal(ppt.slides.length,example.slides);assert.equal(ppt.nativeTextCount,example.texts);assert.equal(ppt.fullSlideRasterReplacementCount,0);
  assert.ok(Math.abs(ppt.widthEmu/ppt.heightEmu-example.ratio)<0.00001);
  assert.ok(ppt.slides.flatMap(s=>s.objects).some(o=>o.kind==='text'&&o.text===example.title));
  assert.equal((await fetch(app.url+`/api/jobs/${job.id}/object-map`)).status,404);
});

test('private files, reports and path traversal are unavailable',async()=>{
  for(const path of ['/results/latest','/results/source.html','/task.md','/spec.md','/DEVELOPMENT.md','/.git/config','/.env','/package.json','/../README.md','/results/%2e%2e%2fREADME.md'])assert.equal((await fetch(app.url+path)).status,404,path);
  const page=await fetch(app.url);assert.equal(page.status,200);assert.match(page.headers.get('content-security-policy'),/frame-ancestors 'none'/);
  assert.doesNotMatch(await page.text(),/object-map|\/results\/|sourceRef/);
});

test('rejects wrong host, foreign origin and malformed requests',async()=>{
  assert.equal((await post('/api/analyze',{}, {origin:'https://foreign.invalid'})).status,403);
  const wrongHost=await new Promise((resolve,reject)=>{
    const request=http.request(app.url+'/api/analyze',{method:'POST',headers:{host:'foreign.invalid'}},response=>{response.resume();response.on('end',()=>resolve(response.statusCode));});
    request.on('error',reject);request.end('{}');
  });
  assert.equal(wrongHost,403);
  assert.equal((await fetch(app.url+'/api/analyze',{method:'POST',body:'invalid'})).status,400);
  assert.equal((await post('/api/export',{jobId:'unknown'})).status,404);
  assert.equal((await fetch(app.url+'/api/jobs/0000/preview/1.png')).status,404);
  await assert.rejects(createHtmlServer({host:'0.0.0.0'}),/루프백/);
});

test('input failures do not echo source content or object locations',async()=>{
  const marker='sensitive-example-value';
  const response=await post('/api/analyze',{html:`<section data-slide style="width:800px;height:600px"><img src="${marker}.png"></section>`});
  const data=await response.json();assert.equal(response.status,422);assert.equal(data.exportAvailable,false);
  assert.doesNotMatch(JSON.stringify(data),new RegExp(marker));assert.ok(data.diagnostics.some(d=>d.code==='ASSET_MISSING'));assert.ok(data.diagnostics.every(d=>!d.sourceRef&&!d.nodeId));
});
