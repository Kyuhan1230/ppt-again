import test from 'node:test';
import assert from 'node:assert/strict';
import {createHtmlServer} from '../src/html-to-pptx/server.mjs';
import {createTestSession} from './http-session.mjs';
import {unzipSync,strFromU8} from 'fflate';

test('download only supported slides while retaining failed slide diagnostics',async()=>{
 const server=await createHtmlServer({port:0});try{const request=await createTestSession(server.url);const post=(route,body)=>request(route,{method:'POST',headers:{'content-type':'application/json'},body:JSON.stringify(body)});
 const html='<section data-slide style="width:800px;height:600px;font:20px Arial"><p>First slide</p></section><section data-slide style="width:800px;height:600px;font:20px Arial"><p style="filter:blur(2px)">Unsupported slide</p></section><section data-slide style="width:800px;height:600px;font:20px Arial"><p>Last slide</p></section>';
 const response=await post('/api/analyze',{html,fileName:'sample.html'}),job=await response.json();assert.equal(response.status,422);assert.equal(job.exportAvailable,false);assert.equal(job.partialExportAvailable,true);assert.equal(job.partialSlideCount,2);assert.deepEqual(job.excludedSlideNumbers,[2]);
 assert.equal((await post('/api/export',{jobId:job.id})).status,422);
 const exported=await post('/api/export',{jobId:job.id,partial:true}),result=await exported.json();assert.equal(exported.status,200,JSON.stringify(result));assert.equal(result.exportStatus,'generated');assert.equal(result.slideCount,3);assert.equal(result.downloadSlideCount,2);assert.match(result.downloadFileName,/partial/);assert.ok(result.diagnostics.some(d=>d.severity==='error'&&d.slideNumber===2));
 const file=await request(result.downloadUrl);const zip=unzipSync(new Uint8Array(await file.arrayBuffer()));const xml=Object.keys(zip).filter(x=>/^ppt\/slides\/slide\d+\.xml$/.test(x)).map(x=>strFromU8(zip[x])).join('');assert.match(xml,/First slide/);assert.match(xml,/Last slide/);assert.doesNotMatch(xml,/Unsupported slide/);
 }finally{await server.close();}
});

test('global analysis errors cannot be bypassed with a partial download',async()=>{
 const server=await createHtmlServer({port:0});try{const request=await createTestSession(server.url);const post=(route,body)=>request(route,{method:'POST',headers:{'content-type':'application/json'},body:JSON.stringify(body)});
 const response=await post('/api/analyze',{html:'<p>No slide root</p>',fileName:'sample.html'}),job=await response.json();assert.equal(response.status,422);assert.equal(job.partialExportAvailable,false);assert.equal((await post('/api/export',{jobId:job.id,partial:true})).status,422);
 }finally{await server.close();}
});
