import test from 'node:test';
import assert from 'node:assert/strict';
import {Worker} from 'node:worker_threads';
import {once} from 'node:events';

const moduleUrl=new URL('../src/html-to-pptx/html-pagination-tiles.mjs',import.meta.url).href;
const workerCode=`
  const {parentPort,workerData}=require('node:worker_threads');
  (async()=>{
    const {captureDocumentTiles}=await import(workerData.moduleUrl);
    let closed=false,screenshots=0;
    const viewport={width:100,height:100};
    const page={viewportSize:()=>viewport,close:async()=>{closed=true;},screenshot:()=>{screenshots++;return new Promise(()=>{});}};
    const readState=()=>workerData.block==='state'?new Promise(()=>{}):Promise.resolve({viewport:{...viewport,dpr:1,scrollX:0,scrollY:0},documentSize:{width:100,height:100}});
    try{await captureDocumentTiles({page,rect:{x:0,y:0,width:100,height:100},viewport,readState,timeoutMs:80});parentPort.postMessage({unexpectedSuccess:true});}
    catch(error){parentPort.postMessage({code:error.code,reasonKey:error.reasonKey,cleanupRequired:error.cleanupRequired,closed,screenshots});}
  })().catch(error=>parentPort.postMessage({workerError:String(error)}));`;

for(const block of ['state','screenshot'])test(`tile deadline terminates a stalled ${block} operation and requests page closure`,async()=>{
  const worker=new Worker(workerCode,{eval:true,workerData:{moduleUrl,block}});let timer;
  try{
    const [result]=await Promise.race([once(worker,'message'),new Promise((_,reject)=>{timer=setTimeout(()=>reject(new Error('tile capture did not enforce its deadline')),3000);})]);
    assert.equal(result.code,'PAGINATION_UNSUPPORTED',JSON.stringify(result));assert.equal(result.reasonKey,'PAGINATION_TIMEOUT');assert.equal(result.cleanupRequired,true);assert.equal(result.closed,true,JSON.stringify(result));
    assert.equal(result.screenshots,block==='state'?0:1);
  }finally{clearTimeout(timer);await worker.terminate();}
});
