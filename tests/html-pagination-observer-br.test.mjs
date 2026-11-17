import test from 'node:test';
import assert from 'node:assert/strict';
import {readInput} from '../src/html-to-pptx/input.mjs';
import {createBrowserRuntime} from '../src/html-to-pptx/browser-runtime.mjs';
import {observePaginationDocument} from '../src/html-to-pptx/html-pagination-observer.mjs';
import {adaptMappedPagination} from '../src/html-to-pptx/html-pagination-adapter.mjs';
import {validateSourceCoverage} from '../src/html-to-pptx/html-pagination.mjs';

test('a source BR is observed once and belongs to its actual line even across pages',async()=>{
  const viewport={width:800,height:600};
  const html='<!doctype html><html><body style="margin:0"><main id="doc" style="width:600px;height:100px"><p style="margin:10px;font:20px/30px Arial">Alpha<br>Beta</p></main></body></html>';
  const runtime=await createBrowserRuntime({capture:false,checkLimits:true});
  try{const session=await runtime.render(await readInput({html}),{viewport});try{
    const result=await observePaginationDocument({session,selector:'#doc',viewport});
    assert.equal(result.status,'observed',JSON.stringify(result.diagnostics));
    assert.deepEqual(result.sourceLedger.map(s=>s.kind),['text','br','text']);
    assert.equal(result.displayMapping.tokens.filter(t=>t.kind==='hard-break').length,1);
    const output=adaptMappedPagination({...result,mode:'paginate',pageHeight:40});
    assert.deepEqual(output.diagnostics,[],JSON.stringify(output.diagnostics));
    const parts=output.pages.flatMap(page=>page.parts);
    assert.deepEqual(parts.map(part=>part.content.text),['Alpha','Beta']);
    const expected=result.sourceLedger.flatMap(s=>s.ranges.map(range=>({sourceId:s.sourceId,...range})));
    assert.equal(validateSourceCoverage(expected,parts.flatMap(part=>part.sourceRanges)).valid,true);
    assert.equal(parts.flatMap(part=>part.tokens).filter(t=>t.kind==='hard-break'&&t.handledBy==='page-boundary').length,1);
  }finally{await session.context.close();}}finally{await runtime.close();}
});
