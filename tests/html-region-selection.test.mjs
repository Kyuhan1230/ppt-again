import test from 'node:test';
import assert from 'node:assert/strict';
import {runConversion} from '../src/html-to-pptx/job-runner.mjs';

const source=body=>`<!doctype html><html><body style="margin:0">${body}</body></html>`;
test('selected fixed-size candidate is measured in a new worker without accepting coordinates',async()=>{
  const html=source('<section style="width:400px;height:300px;background:#eee"><p>Selected page</p></section><aside>Outside page</aside>');
  const input={html,assets:[],allowTableOverlay:true};
  const {inspection}=await runConversion({operation:'inspect-html',input});
  const choice=inspection.choices.find(c=>c.scope==='elements'&&c.convertible);
  assert.ok(choice,'a fixed-size candidate must be available for selection');
  const selected=await runConversion({operation:'analyze-html-selection',input:{...input,snapshotId:inspection.snapshotId,choiceId:choice.id}});
  assert.equal(selected.measured.model.slides.length,1);
  assert.ok(selected.measured.model.slides[0].nodes.some(n=>n.kind==='text'&&JSON.stringify(n).includes('Selected page')));
  assert.ok(!JSON.stringify(selected.measured.model).includes('Outside page'));
  await assert.rejects(()=>runConversion({operation:'analyze-html-selection',input:{...input,snapshotId:inspection.snapshotId,choiceId:choice.id,bounds:{x:0,y:0,width:10,height:10}}}),{code:'INPUT_INVALID'});
});

test('responsive content is measured at the viewport used for selection',async()=>{
  const html=source('<style>@media(max-width:900px){.changing{display:none}}</style><section style="width:400px;height:600px;background:#eee"><p class="changing">Changes with viewport</p></section>');
  const input={html,assets:[],allowTableOverlay:true}, {inspection}=await runConversion({operation:'inspect-html',input});
  const choice=inspection.choices.find(c=>c.scope==='elements'&&c.convertible);assert.ok(choice);
  const result=await runConversion({operation:'analyze-html-selection',input:{...input,snapshotId:inspection.snapshotId,choiceId:choice.id}});
  assert.ok(JSON.stringify(result.measured.model).includes('Changes with viewport'));
});

test('compatible repeated candidates preserve slide count and order in the PPT',async()=>{
  const html=source('<section style="width:400px;height:300px;background:#eee"><p>First page</p></section><section style="width:400px;height:300px;background:#eee"><p>Second page</p></section>');
  const input={html,assets:[],allowTableOverlay:true},{inspection}=await runConversion({operation:'inspect-html',input});
  const choice=inspection.choices.find(c=>c.evidence.includes('repeated-sibling')&&c.regionIds.length===2&&c.convertible);assert.ok(choice,JSON.stringify(inspection.choices));
  const selected=await runConversion({operation:'analyze-html-selection',input:{...input,snapshotId:inspection.snapshotId,choiceId:choice.id}});
  assert.equal(selected.measured.model?.slides.length,2,JSON.stringify(selected.measured.diagnostics));
  const {inspectPptxIndependently}=await import('./inspect-pptx.mjs');
  const exported=await runConversion({operation:'export-pptx',input:selected.measured});assert.equal(exported.status,'generated',JSON.stringify(exported.diagnostics));
  const ppt=await inspectPptxIndependently(exported.pptxBytes);assert.equal(ppt.slides.length,2);
  assert.ok(JSON.stringify(ppt.slides[0]).includes('First page'));assert.ok(JSON.stringify(ppt.slides[1]).includes('Second page'));
});

test('an explicitly selected fixed-size complete document can be measured as one slide',async()=>{
  const html='<!doctype html><html><body style="margin:0;width:1280px;height:960px"><p style="margin:0">Document content</p></body></html>';
  const input={html,assets:[],allowTableOverlay:true},{inspection}=await runConversion({operation:'inspect-html',input});
  const choice=inspection.choices.find(c=>c.scope==='document');assert.ok(choice?.convertible,JSON.stringify(choice));
  const selected=await runConversion({operation:'analyze-html-selection',input:{...input,snapshotId:inspection.snapshotId,choiceId:choice.id}});
  assert.equal(selected.measured.model?.slides.length,1,JSON.stringify(selected.measured.diagnostics));
});

test('candidate selection does not silently lose parent paint or external overlays',async()=>{
  const {createBrowserRuntime}=await import('../src/html-to-pptx/browser-runtime.mjs');
  const {readInput}=await import('../src/html-to-pptx/input.mjs');
  const {inspectInput}=await import('../src/html-to-pptx/html-regions.mjs');
  const runtime=await createBrowserRuntime({capture:false,checkLimits:true});
  const getChildChoice=async html=>{
    const inspection=await inspectInput(await readInput({html:source(html)}),runtime);
    const candidate=inspection.regions.find(r=>r.parentId&&r.bounds?.width===400&&r.bounds?.height===300);
    assert.ok(candidate,JSON.stringify(inspection.regions));
    return inspection.choices.find(c=>c.regionIds?.length===1&&c.regionIds[0]===candidate.id);
  };
  try{
    const parent=await getChildChoice('<main style="width:400px;height:300px;background:red"><section style="width:400px;height:300px;background:transparent">TEXT</section></main>');
    assert.equal(parent?.convertible,false,'the parent background would be lost');
    const overlay=await getChildChoice('<main style="position:relative;width:400px;height:300px"><section style="width:400px;height:300px">UNDER</section><aside style="position:absolute;left:0;top:0;width:400px;height:300px;background:red;z-index:2"></aside></main>');
    assert.equal(overlay?.convertible,false,'an overlapping sibling would be lost');
    const rounded=await getChildChoice('<main style="width:400px;height:300px;background:red"><section style="width:400px;height:300px;background:white;border-radius:80px">TEXT</section></main>');
    assert.equal(rounded?.convertible,false,'rounded corners expose the parent background');
    const clipped=await getChildChoice('<main style="width:400px;height:300px;background:red"><section style="width:400px;height:300px;background:white;background-clip:content-box;padding:20px;box-sizing:border-box">TEXT</section></main>');
    assert.equal(clipped?.convertible,false,'background clipping exposes the parent background');
    const opaque=await getChildChoice('<main style="width:400px;height:300px;background:red"><section style="width:400px;height:300px;background:white">TEXT</section></main>');
    assert.equal(opaque?.convertible,true,'an opaque slide covers the parent background');
  }finally{await runtime.close();}
});

test('explicit selectors reject visible external composition but retain outside text warning',async()=>{
  const {createBrowserRuntime}=await import('../src/html-to-pptx/browser-runtime.mjs');
  const {readInput}=await import('../src/html-to-pptx/input.mjs');
  const {measureInput}=await import('../src/html-to-pptx/measure.mjs');
  const runtime=await createBrowserRuntime({capture:false,checkLimits:true});
  const measure=async html=>measureInput(await readInput({html:source(html),slideSelector:'section'}),runtime);
  try{
    const parent=await measure('<main style="width:400px;height:300px;background:red"><section style="width:400px;height:300px">TEXT</section></main>');
    assert.ok(parent.diagnostics.some(d=>d.code==='UNSUPPORTED_OBJECT'&&d.severity==='error'),JSON.stringify(parent.diagnostics));
    const overlay=await measure('<main style="position:relative;width:400px;height:300px"><section style="width:400px;height:300px">UNDER</section><aside style="position:absolute;left:0;top:0;width:400px;height:300px;background:red;z-index:2"></aside></main>');
    assert.ok(overlay.diagnostics.some(d=>d.code==='UNSUPPORTED_OBJECT'&&d.severity==='error'),JSON.stringify(overlay.diagnostics));
    const safe=await measure('<section style="width:400px;height:300px;background:white">IN</section><aside>OUTSIDE</aside>');
    assert.ok(safe.model);assert.ok(!safe.diagnostics.some(d=>d.severity==='error'),JSON.stringify(safe.diagnostics));
    assert.ok(safe.diagnostics.some(d=>d.code==='SLIDE_CONTENT_OUTSIDE'&&d.severity==='warning'));
  }finally{await runtime.close();}
});

test('root and whole-document choices never promise backgrounds the PPT cannot preserve',async()=>{
  const {createBrowserRuntime}=await import('../src/html-to-pptx/browser-runtime.mjs');
  const {readInput}=await import('../src/html-to-pptx/input.mjs');
  const {inspectInput}=await import('../src/html-to-pptx/html-regions.mjs');
  const {measureInput}=await import('../src/html-to-pptx/measure.mjs');
  const runtime=await createBrowserRuntime({capture:false,checkLimits:true});
  try{
    const redParent='<!doctype html><html style="background:red"><body style="margin:0;width:1280px;height:960px;background:transparent"><p style="margin:0">PAGE</p></body></html>';
    const inspected=await inspectInput(await readInput({html:redParent}),runtime);
    assert.equal(inspected.choices.find(c=>c.scope==='document').convertible,false,'html background remains visible through body');
    const rounded=source('<section style="width:400px;height:300px;background:red;border-radius:80px">TEXT</section>');
    const candidates=await inspectInput(await readInput({html:rounded}),runtime);
    assert.equal(candidates.choices.find(c=>c.scope==='elements')?.convertible,false,'rounded root cannot become square slide background');
    const bordered=source('<section style="width:400px;height:300px;background:white;border:4px solid red">TEXT</section>');
    assert.equal((await inspectInput(await readInput({html:bordered}),runtime)).choices.find(c=>c.scope==='elements')?.convertible,false,'slide-root border is not a slide background');
    const borderedMeasure=await measureInput(await readInput({html:bordered,slideSelector:'section'}),runtime);
    assert.ok(borderedMeasure.diagnostics.some(d=>d.code==='UNSUPPORTED_OBJECT'&&d.severity==='error'),JSON.stringify(borderedMeasure.diagnostics));
    const measured=await measureInput(await readInput({html:rounded,slideSelector:'section'}),runtime);
    assert.ok(measured.diagnostics.some(d=>d.code==='UNSUPPORTED_OBJECT'&&d.severity==='error'),JSON.stringify(measured.diagnostics));
    const normal='<!doctype html><html><body style="margin:0;width:1280px;height:960px"><p style="margin:0">PAGE</p></body></html>';
    assert.equal((await inspectInput(await readInput({html:normal}),runtime)).choices.find(c=>c.scope==='document').convertible,true);
    const opaque=redParent.replace('background:transparent','background:white');
    assert.equal((await inspectInput(await readInput({html:opaque}),runtime)).choices.find(c=>c.scope==='document').convertible,true,'opaque body hides the html background');
  }finally{await runtime.close();}
});

test('content-box padding and transparent borders cannot change the selected slide size',async()=>{
  const {createBrowserRuntime}=await import('../src/html-to-pptx/browser-runtime.mjs');
  const {readInput}=await import('../src/html-to-pptx/input.mjs');
  const {inspectInput}=await import('../src/html-to-pptx/html-regions.mjs');
  const {measureInput}=await import('../src/html-to-pptx/measure.mjs');
  const runtime=await createBrowserRuntime({capture:false,checkLimits:true});
  try{
    for(const extra of ['padding:20px','border:8px solid transparent']){
      const html=source(`<section style="width:400px;height:300px;${extra}">TEXT</section>`);
      const inspected=await inspectInput(await readInput({html}),runtime);
      const choice=inspected.choices.find(c=>c.scope==='elements');
      assert.equal(choice?.convertible,false,extra);
      assert.ok(choice.reasons.includes('root-size-mismatch'),JSON.stringify(choice));
      const measured=await measureInput(await readInput({html,slideSelector:'section'}),runtime);
      assert.ok(measured.diagnostics.some(d=>d.code==='UNSUPPORTED_OBJECT'&&d.severity==='error'&&d.message.includes('content-box')&&d.message.includes('box-sizing:border-box')),JSON.stringify(measured.diagnostics));
    }
    const fixed=source('<section style="width:400px;height:300px;padding:20px;box-sizing:border-box">TEXT</section>');
    assert.equal((await inspectInput(await readInput({html:fixed}),runtime)).choices.find(c=>c.scope==='elements')?.convertible,true);
  }finally{await runtime.close();}
});

test('explicit selectors check root and parent appearance in the measured viewport',async()=>{
  const {createBrowserRuntime}=await import('../src/html-to-pptx/browser-runtime.mjs');
  const {readInput}=await import('../src/html-to-pptx/input.mjs');
  const {measureInput}=await import('../src/html-to-pptx/measure.mjs');
  const runtime=await createBrowserRuntime({capture:false,checkLimits:true});
  const measure=async html=>measureInput(await readInput({html:source(html),slideSelector:'section'}),runtime);
  try{
    const initiallyRounded=await measure('<style>@media(min-width:1000px){section{border-radius:80px}}</style><section style="width:400px;height:600px;background:red">TEXT</section>');
    assert.ok(initiallyRounded.diagnostics.some(d=>d.code==='UNSUPPORTED_OBJECT'&&d.severity==='error'),JSON.stringify(initiallyRounded.diagnostics));
    const finallyRounded=await measure('<style>@media(max-width:900px){section{border-radius:80px}}</style><section style="width:400px;height:600px;background:red">TEXT</section>');
    assert.ok(!finallyRounded.diagnostics.some(d=>d.severity==='error'),JSON.stringify(finallyRounded.diagnostics));
    const initialParent=await measure('<style>@media(min-width:1000px){main{background:red}}</style><main style="width:400px;height:600px"><section style="width:400px;height:600px">TEXT</section></main>');
    assert.ok(initialParent.diagnostics.some(d=>d.code==='UNSUPPORTED_OBJECT'&&d.severity==='error'),JSON.stringify(initialParent.diagnostics));
  }finally{await runtime.close();}
});
