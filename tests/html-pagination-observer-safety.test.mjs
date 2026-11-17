import test from 'node:test';
import assert from 'node:assert/strict';
import {observePaginationDocument} from '../src/html-to-pptx/html-pagination-observer.mjs';

test('observer rejects invalid sessions, active source scripts and blocked assets',async()=>{
  const invalid=await observePaginationDocument({});
  assert.equal(invalid.status,'unsupported');assert.ok(invalid.diagnostics.some(d=>d.severity==='error'));
  let evaluated=false;
  const page={viewportSize:()=>({width:800,height:600}),evaluate:async()=>{evaluated=true;throw new Error('should not evaluate');}};
  for(const session of [{page,sourceHadScripts:true,blocked:[]},{page,sourceHadScripts:false,blocked:['blocked asset']}]){
    const result=await observePaginationDocument({session,selector:'body',viewport:{width:800,height:600}});
    assert.equal(result.status,'unsupported');assert.ok(result.diagnostics.some(d=>d.severity==='error'));assert.equal(evaluated,false);
  }
});

test('observer refuses a viewport different from the selected inspection viewport',async()=>{
  let evaluated=false;
  const page={viewportSize:()=>({width:800,height:600}),evaluate:async()=>{evaluated=true;throw new Error('should not evaluate');}};
  const result=await observePaginationDocument({session:{page,sourceHadScripts:false,blocked:[]},selector:'body',viewport:{width:1280,height:960}});
  assert.equal(result.status,'unsupported');assert.ok(result.diagnostics.some(d=>d.severity==='error'));assert.equal(evaluated,false);
});

test('observer deadline closes the temporary context when DOM observation stalls',async()=>{
  let closed=false;
  const session={page:{viewportSize:()=>({width:800,height:600}),evaluate:()=>new Promise(()=>{})},context:{close:async()=>{closed=true;}},sourceHadScripts:false,blocked:[]};
  const result=await observePaginationDocument({session,selector:'body',viewport:{width:800,height:600},timeoutMs:30});
  assert.equal(result.status,'unsupported');assert.equal(closed,true);
});
