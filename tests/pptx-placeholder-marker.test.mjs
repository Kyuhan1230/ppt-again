import test from 'node:test';
import assert from 'node:assert/strict';
import PptxGenJS from 'pptxgenjs';
import {readPptx} from '../src/html-to-pptx/read-pptx.mjs';
import {exportHtml} from '../src/html-to-pptx/export-html.mjs';
import {readInput} from '../src/html-to-pptx/input.mjs';
import {createBrowserRuntime} from '../src/html-to-pptx/browser-runtime.mjs';
import {measureInput} from '../src/html-to-pptx/measure.mjs';
import {exportPptx} from '../src/html-to-pptx/export.mjs';
import {publicDiagnostic} from '../src/html-to-pptx/public-diagnostics.mjs';

test('an unsupported object is identified by a compact accessible marker at its original position',async()=>{
  const ppt=new PptxGenJS();ppt.layout='LAYOUT_WIDE';const slide=ppt.addSlide();
  slide.addShape(ppt.ShapeType.triangle,{x:1,y:1,w:4,h:2,fill:{color:'225588'},line:{color:'000000'},objectName:'Unsupported triangle'});
  const parsed=await readPptx(new Uint8Array(await ppt.write({outputType:'uint8array'})),{allowPartial:true});
  assert.equal(parsed.status,'parsed',JSON.stringify(parsed.diagnostics));
  assert.ok(parsed.diagnostics.some(d=>d.code==='PPTX_OBJECT_PLACEHOLDER'));
  const output=exportHtml(parsed.model);assert.equal(output.status,'generated',JSON.stringify(output.diagnostics));
  const runtime=await createBrowserRuntime({capture:false});
  const input=await readInput({html:output.html});const session=await runtime.render(input);
  try {
    const marker=session.page.locator('[data-placeholder-marker]');
    assert.equal(await marker.count(),1);
    assert.equal(await marker.innerText(),'!');
    assert.equal(await marker.getAttribute('role'),'img');
    assert.equal(await marker.getAttribute('tabindex'),'0');
    const label=await marker.getAttribute('aria-label');
    assert.match(label,/Unsupported triangle/);
    assert.match(label,/지원/);
    assert.equal(await marker.getAttribute('title'),label);
    const bounds=await marker.evaluate(el=>({left:parseFloat(el.style.left),top:parseFloat(el.style.top),width:parseFloat(el.style.width),height:parseFloat(el.style.height)}));
    assert.deepEqual(bounds,{left:96,top:96,width:20,height:20});
    const summary=session.page.locator('details[data-placeholder-summary]');
    assert.match(await summary.locator('summary').innerText(),/변환하지 못한 항목 1/);
    await summary.locator('summary').click();
    assert.match(await summary.locator('li').innerText(),/Unsupported triangle/);
    const altered=output.html.replace(/aria-label="[^"]*"/, 'aria-label="UNTRUSTED MESSAGE"');
    const measured=await measureInput(await readInput({html:altered}),runtime);
    const warning=measured.diagnostics.find(d=>d.code==='SOURCE_PLACEHOLDER_PRESENT');
    assert.equal(warning?.severity,'warning',JSON.stringify(measured.diagnostics));
    assert.doesNotMatch(JSON.stringify(publicDiagnostic(warning)),/UNTRUSTED MESSAGE/);
    const pptx=await exportPptx(measured);
    assert.equal(pptx.status,'generated',JSON.stringify(pptx.diagnostics));
    assert.ok(pptx.diagnostics.some(d=>d.code==='SOURCE_PLACEHOLDER_PRESENT'&&d.severity==='warning'));
  } finally {await session.context.close();await runtime.close();}
});
