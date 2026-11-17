import test from 'node:test';
import assert from 'node:assert/strict';
import {Worker} from 'node:worker_threads';
import {once} from 'node:events';

test('a stalled DOM observation reports a timeout and requests context cleanup',async()=>{
  const moduleUrl=new URL('../src/html-to-pptx/html-pagination-observer.mjs',import.meta.url).href;
  const code=`const {parentPort,workerData}=require('node:worker_threads');
    (async()=>{const {observePaginationDocument}=await import(workerData.moduleUrl);let closed=false;
      const viewport={width:100,height:100},session={page:{viewportSize:()=>viewport,evaluate:()=>new Promise(()=>{})},context:{close:async()=>{closed=true;}},sourceHadScripts:false,blocked:[]};
      const result=await observePaginationDocument({session,selector:'#doc',viewport,timeoutMs:80});
      parentPort.postMessage({status:result.status,diagnostics:result.diagnostics,cleanupRequired:result.cleanupRequired,closed});
    })().catch(error=>parentPort.postMessage({workerError:String(error)}));`;
  const worker=new Worker(code,{eval:true,workerData:{moduleUrl}});let timer;
  try{
    const [result]=await Promise.race([once(worker,'message'),new Promise((_,reject)=>{timer=setTimeout(()=>reject(new Error('observer did not enforce its deadline')),3000);})]);
    assert.equal(result.status,'unsupported',JSON.stringify(result));assert.equal(result.cleanupRequired,true);assert.equal(result.closed,true);
    assert.ok(result.diagnostics.some(item=>item.code==='PAGINATION_UNSUPPORTED'&&item.reasonKey==='PAGINATION_TIMEOUT'),JSON.stringify(result));
  }finally{clearTimeout(timer);await worker.terminate();}
});
