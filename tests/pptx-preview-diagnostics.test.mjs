import test from 'node:test';
import assert from 'node:assert/strict';
import PptxGenJS from 'pptxgenjs';
import {runConversion} from '../src/html-to-pptx/job-runner.mjs';

test('PPTX preview reports existing placeholders once and does not claim a rendered slide failed',async()=>{
  const ppt=new PptxGenJS();ppt.layout='LAYOUT_WIDE';const slide=ppt.addSlide();slide.background={color:'FFFFFF'};
  slide.addShape(ppt.ShapeType.triangle,{x:1,y:1,w:2,h:1,fill:{color:'225588'},line:{color:'000000'},objectName:'Unsupported triangle'});
  slide.addTable([['first\nsecond\nthird\nfourth\nfifth']],{x:4,y:1,w:2,h:.3,fontFace:'Arial',fontSize:20,objectName:'Tall text'});
  const bytes=new Uint8Array(await ppt.write({outputType:'uint8array'}));
  const {measured}=await runConversion({operation:'analyze-pptx',input:{bytes,options:{allowPartial:true}}});
  assert.equal(measured.previews.length,1,JSON.stringify(measured.diagnostics));
  assert.ok(measured.diagnostics.some(d=>d.code==='PPTX_OBJECT_PLACEHOLDER'));
  assert.equal(measured.diagnostics.some(d=>d.code==='SOURCE_PLACEHOLDER_PRESENT'),false);
  assert.equal(measured.diagnostics.some(d=>d.code==='PPTX_PREVIEW_UNAVAILABLE'),false);
});
