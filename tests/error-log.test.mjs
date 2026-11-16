import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import net from 'node:net';
import os from 'node:os';
import path from 'node:path';
import {randomUUID} from 'node:crypto';
import {once} from 'node:events';
import {createErrorLogger} from '../src/html-to-pptx/error-log.mjs';
import {createHtmlServer} from '../src/html-to-pptx/server.mjs';

const now=()=>new Date('2026-11-09T01:02:03.000Z');
const activeFile='errors-2026-11-09.log';
const archiveFile='errors-2026-11-09.1.log';

async function tempDirectory(t){
  const directory=await fs.mkdtemp(path.join(os.tmpdir(),'ppt-again-error-log-'));
  t.after(()=>fs.rm(directory,{recursive:true,force:true}));
  return directory;
}

async function readRecords(directory,file=activeFile){
  return (await fs.readFile(path.join(directory,file),'utf8')).trim().split('\n').map(line=>JSON.parse(line));
}

test('local error logs keep allowlisted metadata and omit private content',async t=>{
  const directory=await tempDirectory(t),logger=createErrorLogger({directory,now});
  const jobId=randomUUID(),requestId=randomUUID();
  const cause=Object.assign(new Error('private-raw-error'),{code:5,killed:false,stderr:'private-command-output'});
  const error=Object.assign(new Error('private-original-message',{cause}),{
    code:'JOB_OBSERVER_FAILED',observerPhase:'closing',locations:['src/html-to-pptx/job-runner.mjs:42:6','src/html-to-pptx/private-file.mjs:1:1','C:/private-folder/path.mjs:1:1']
  });
  assert.equal(await logger({
    operation:'analyze-pptx',jobId,requestId,sourceType:'pptx',targetType:'html',inputBytes:400,statusCode:500,
    fileName:'private-name.pptx',html:'private-body',token:'private-token',error,
    resourceUsage:{elapsedMs:1500,rssSamples:1,peakRssBytes:2048,terminationMs:0,ownedPids:[11111]},
    diagnostics:[{code:'PPTX_PREVIEW_UNAVAILABLE',severity:'warning',slideId:'private-slide',reasonKey:'PREVIEW_UNAVAILABLE',message:'private-diagnostic',objectName:'private-object',fontFamily:'private-font',sourceRef:{path:'private-source'}}],
    slides:[{id:'other'},{id:'private-slide'}]
  }),true);
  const log=await fs.readFile(path.join(directory,activeFile),'utf8');
  assert.doesNotMatch(log,/private-|11111|ownedPids/);
  const [record]=await readRecords(directory);
  assert.equal(record.time,now().toISOString());
  assert.equal(record.jobId,jobId);
  assert.equal(record.requestId,requestId);
  assert.deepEqual(record.error,{
    code:'JOB_OBSERVER_FAILED',name:'Error',locations:['src/html-to-pptx/job-runner.mjs:42:6'],observerPhase:'closing',
    cause:{code:'UNKNOWN_ERROR',name:'Error',locations:[],exitCode:5,killed:false}
  });
  assert.deepEqual(record.resourceUsage,{elapsedMs:1500,rssSamples:1,peakRssBytes:2048,terminationMs:0});
  assert.equal(record.diagnosticCount,1);
  assert.equal(record.diagnostics[0].slideNumber,2);
  assert.equal(record.diagnostics[0].reasonKey,'PREVIEW_UNAVAILABLE');
  assert.match(record.diagnostics[0].message,/미리보기/);
});

test('concurrent writes stay ordered and only one bounded archive is retained',async t=>{
  const directory=await tempDirectory(t),maxBytes=400;
  const logger=createErrorLogger({directory,now,maxBytes});
  const results=await Promise.all(Array.from({length:8},(_,inputBytes)=>logger({operation:'cli',inputBytes})));
  assert.ok(results.every(Boolean));
  assert.deepEqual((await fs.readdir(directory)).sort(),[archiveFile,activeFile]);
  const archive=await readRecords(directory,archiveFile),active=await readRecords(directory);
  const kept=[...archive,...active].map(record=>record.inputBytes);
  assert.equal(kept.at(-1),7);
  assert.deepEqual(kept,Array.from({length:kept.length},(_,index)=>8-kept.length+index));
  for(const file of [activeFile,archiveFile])assert.ok((await fs.stat(path.join(directory,file))).size<=maxBytes);
});

test('writes prune expired log dates and preserve unrelated files',async t=>{
  const directory=await tempDirectory(t);
  for(const file of ['errors-2026-11-02.log','errors-2026-11-02.1.log','errors-2026-11-03.log','unrelated.log']){
    await fs.writeFile(path.join(directory,file),'existing');
  }
  assert.equal(await createErrorLogger({directory,now})({operation:'cli'}),true);
  assert.deepEqual((await fs.readdir(directory)).sort(),['errors-2026-11-03.log',activeFile,'unrelated.log']);
});

test('an unavailable log path does not break the request or poison later writes',async t=>{
  const directory=path.join(await tempDirectory(t),'blocked');
  await fs.writeFile(directory,'not a directory');
  const logger=createErrorLogger({directory,now});
  assert.equal(await logger({operation:'cli'}),false);
  const app=await createHtmlServer({port:0,errorLogger:logger});
  t.after(()=>app.close());
  const session=await fetch(app.url+'/api/session');
  const cookie=session.headers.get('set-cookie').split(';')[0],csrf=(await session.json()).csrfToken;
  const response=await fetch(app.url+'/api/analyze',{method:'POST',headers:{cookie,'x-csrf-token':csrf},body:'bad JSON'});
  assert.equal(response.status,400);
  await response.json();
  await fs.unlink(directory);
  assert.equal(await logger({operation:'cli'}),true);
  assert.equal((await readRecords(directory)).length,1);
});

test('server logs request exceptions and identifies rejected jobs without logging the input',async t=>{
  const directory=await tempDirectory(t);
  const app=await createHtmlServer({port:0,errorLogger:createErrorLogger({directory,now})});
  t.after(()=>app.close());
  const session=await fetch(app.url+'/api/session');
  const cookie=session.headers.get('set-cookie').split(';')[0],csrf=(await session.json()).csrfToken;
  const post=body=>fetch(app.url+'/api/analyze',{method:'POST',headers:{cookie,'x-csrf-token':csrf},body});

  const malformed=await post('private-bad-request');
  assert.equal(malformed.status,400);
  await malformed.json();
  const invalid=await post(JSON.stringify({sourceType:'pptx',fileName:'private-upload.pptx',pptxBase64:Buffer.from('private-payload').toString('base64')}));
  assert.equal(invalid.status,400);
  await invalid.json();

  const [request,job]=await readRecords(directory);
  assert.equal(request.operation,'http');
  assert.equal(request.statusCode,400);
  assert.match(request.error.locations[0],/^src\/html-to-pptx\/server\.mjs:/);
  assert.equal(job.operation,'analyze-pptx');
  assert.equal(job.error.code,'INPUT_INVALID');
  assert.ok(job.jobId);
  assert.notEqual(job.requestId,request.requestId);
  assert.equal(job.resourceUsage.rssSamples,0);
  assert.doesNotMatch(await fs.readFile(path.join(directory,activeFile),'utf8'),/private-|csrfToken/);
});

test('server startup preserves the bind error and logs its code',async t=>{
  const directory=await tempDirectory(t);
  const listener=net.createServer().listen(0,'127.0.0.1');
  t.after(()=>new Promise(resolve=>listener.close(resolve)));
  await once(listener,'listening');
  const errorLogger=createErrorLogger({directory,now});
  await assert.rejects(createHtmlServer({port:listener.address().port,errorLogger}),{code:'EADDRINUSE'});
  const [record]=await readRecords(directory);
  assert.equal(record.operation,'server-start');
  assert.equal(record.error.code,'EADDRINUSE');
});
