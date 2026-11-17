import test from 'node:test';
import assert from 'node:assert/strict';
import {spawnSync} from 'node:child_process';
import {createHtmlServer,summarizeJob} from '../src/html-to-pptx/server.mjs';
import {runConversion} from '../src/html-to-pptx/job-runner.mjs';
import {publicDiagnostic} from '../src/html-to-pptx/public-diagnostics.mjs';

const pagination={
  mode:'paginate',sourceWidth:800,sourceHeight:900,pageHeight:450,
  oversizePolicy:'stop',textEditable:true,textObjectGranularity:'line',
  pages:[{fromY:0,toY:450},{fromY:450,toY:900}]
};

function job(report){
  return {
    id:'test-job',fileName:'synthetic.html',sourceType:'html',targetType:'pptx',pageMode:'paginate',
    input:{profile:'static-slides-v1',source:{sha256:'synthetic-digest'}},
    measured:{
      diagnostics:[],previews:[],pagination:report,
      model:{slides:[{id:'private-1',nodes:[]},{id:'private-2',nodes:[]}]}
    }
  };
}

test('worker rejects pagination combined with a selector or prior candidate choice',async()=>{
  for(const fields of [{slideSelector:'p'},{choiceId:'candidate'},{snapshotId:'saved'},{pageMode:'single'},{scope:'region'}]){
    await assert.rejects(runConversion({operation:'analyze-html',input:{html:'<p>sample</p>',pageMode:'paginate',...fields}}),{code:'INPUT_INVALID'});
  }
  await assert.rejects(runConversion({operation:'inspect-html',input:{html:'<p>sample</p>',pageMode:'paginate'}}),{code:'INPUT_INVALID'});
  for(const operation of ['verify-html-selection','analyze-html-selection']){
    await assert.rejects(runConversion({operation,input:{html:'<p>sample</p>',pageMode:'paginate',snapshotId:'saved',choiceId:'candidate'}}),{code:'INPUT_INVALID'});
  }
});

test('CLI requires document scope explicitly when requesting pagination',()=>{
  const base=['src/html-to-pptx/cli.mjs','convert','no-such-pagination-input.html','--output','no-such-pagination-output.pptx'];
  for(const args of [
    ['--page-mode','paginate'],
    ['--scope','region','--page-mode','paginate'],
    ['--scope','document','--page-mode','paginate','--selector','p'],
    ['--scope','document','--page-mode','paginate','--inspection','saved.json','--region','candidate']
  ]){
    const result=spawnSync(process.execPath,[...base,...args],{encoding:'utf8'});
    assert.equal(result.status,2);
    assert.match(result.stderr,/Usage:/);
  }
  const result=spawnSync(process.execPath,[...base,'--scope','document','--page-mode','paginate'],{encoding:'utf8'});
  assert.equal(result.status,2);
  assert.match(result.stderr,/입력 파일을 찾을 수 없습니다/);
});

test('HTTP only accepts pagination for HTML document analysis with no selector or candidate choice',async t=>{
  const app=await createHtmlServer({port:0});t.after(()=>app.close());
  const session=await fetch(app.url+'/api/session');
  const cookie=session.headers.get('set-cookie').split(';')[0];
  const csrf=(await session.json()).csrfToken;
  const source={html:'<p>sample</p>',pageMode:'paginate'};
  for(const [route,body] of [
    ['/api/inspect',source],
    ['/api/analyze',{...source,slideSelector:'p'}],
    ['/api/analyze',{...source,snapshotId:'saved'}],
    ['/api/analyze',{...source,scope:'region'}],
    ['/api/analyze',{...source,pageMode:'other'}],
    ['/api/analyze',{sourceType:'pptx',pptxBase64:'UEs=',pageMode:'paginate'}]
  ]){
    const response=await fetch(app.url+route,{method:'POST',headers:{cookie,'content-type':'application/json','x-csrf-token':csrf},body:JSON.stringify(body)});
    assert.equal(response.status,400);
    await response.json();
  }
});

test('pagination reports public page ranges and omits private source identifiers',()=>{
  const report={...pagination,sizeSource:'default-16:9',editableTableCount:2,imageCount:1,
    privateMetadata:'secret',pages:pagination.pages.map(page=>({...page,sourceId:'secret'}))};
  const result=summarizeJob(job(report));
  assert.deepEqual(result.pagination,{...pagination,sizeSource:'default-16:9',editableTableCount:2,imageCount:1,
    pages:pagination.pages.map((page,i)=>({number:i+1,...page}))});
  assert.ok(!JSON.stringify(result).includes('secret'));
  assert.equal(result.exportAvailable,true);
});

test('pagination without an intact page range cannot be exported as a partial document',()=>{
  for(const report of [undefined,{...pagination,pages:[{fromY:0,toY:449},{fromY:450,toY:900}]}]){
    const result=summarizeJob(job(report));
    assert.equal(result.status,'failed');
    assert.equal(result.exportAvailable,false);
    assert.equal(result.partialExportAvailable,false);
    assert.equal(result.pagination,undefined);
  }
});

test('pagination reason codes preserve guidance without forwarding raw object details',()=>{
  const diagnostic=publicDiagnostic({code:'PAGINATION_UNSUPPORTED',severity:'error',reasonKey:'PAGINATION_OVERSIZE',message:'private content'});
  assert.equal(diagnostic.code,'PAGINATION_UNSUPPORTED');
  assert.match(diagnostic.message,/한 장/);
  assert.ok(!JSON.stringify(diagnostic).includes('private content'));
});
