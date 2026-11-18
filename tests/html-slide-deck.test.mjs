import test, {after, before} from 'node:test';
import assert from 'node:assert/strict';
import {unzipSync, strFromU8} from 'fflate';
import {createBrowserRuntime} from '../src/html-to-pptx/browser-runtime.mjs';
import {exportPptx} from '../src/html-to-pptx/export.mjs';
import {installedNativeAliases} from '../src/html-to-pptx/font-profiles.mjs';
import {inspectInput} from '../src/html-to-pptx/html-regions.mjs';
import {readInput} from '../src/html-to-pptx/input.mjs';
import {measureInput} from '../src/html-to-pptx/measure.mjs';

let runtime;
before(async()=>{runtime=await createBrowserRuntime();});
after(async()=>{await runtime?.close();});

const deck=(extra='',second='')=>`<!doctype html><style>
  body{font:20px Arial}
  main{position:relative;width:400px;height:300px}
  .slide{display:none;position:absolute;inset:0;width:400px;height:300px;background:white}
  .active{display:block}
  .label{position:absolute;left:16px;top:24px;margin:0}
  @media print{.slide{display:block;position:relative}.label{left:99px}.private{display:block!important}}
</style>
<main>
  <section data-slide class="slide active"><p class="label">One</p></section>
  <section data-slide class="slide" ${second}>
    <p class="label">Two</p><aside class="private" style="display:none">Hidden detail</aside>
    <script>document.querySelector('.label').textContent='Unexpected';</script>
  </section>
  ${extra}
</main>`;

test('declared deck measures one screen frame at a time in document order',async()=>{
  const input=await readInput({html:deck('<section class="slide"><p>Undeclared</p></section>')});
  const inspected=await inspectInput(input,runtime);
  const candidates=inspected.regions.filter(r=>r.evidence.includes('declared-slide'));
  assert.equal(candidates.length,2);
  assert.equal(candidates[1].visibility,'hidden');
  assert.ok(candidates[1].evidence.includes('declared-deck-frame'));

  const measured=await measureInput(input,runtime);
  assert.deepEqual(measured.diagnostics,[]);
  assert.deepEqual(measured.measurements.map(m=>[m.slideId,m.referenceSize.width,m.referenceSize.height,m.layoutStable]),[
    ['slide-1',400,300,true],['slide-2',400,300,true],
  ]);
  assert.equal(measured.previews.length,2);
  assert.deepEqual(measured.model.slides.map(s=>s.nodes.filter(n=>n.kind==='text').map(n=>[n.content.text,n.boxCssPx.x])),[[['One',16]],[['Two',16]]]);
  const output=await exportPptx(measured);
  assert.equal(output.status,'generated',JSON.stringify(output.diagnostics));
  const zip=unzipSync(output.pptxBytes);
  assert.match(strFromU8(zip['ppt/slides/slide2.xml']),/>Two</);
  assert.doesNotMatch(strFromU8(zip['ppt/slides/slide2.xml']),/Hidden detail|Undeclared|Unexpected/);
});

test('nonmatching and semantically hidden roots stay unmeasured',async()=>{
  for(const second of ['hidden','style="width:500px"','style="visibility:hidden"','style="left:20px"']){
    const input=await readInput({html:deck('',second)});
    const measured=await measureInput(input,runtime);
    assert.equal(measured.model.slides.length,1,second);
    assert.ok(measured.diagnostics.some(d=>d.code==='SLIDE_ROOT_REQUIRED'&&d.reasonKey==='HTML_DECLARED_SLIDE_HIDDEN'&&d.slideId==='slide-2'&&d.slideNumber===2),second);
  }
});

test('an unavailable font in a hidden SVG frame uses a real embedded replacement',async()=>{
  const html=deck().replace('<p class="label">Two</p>',
    '<svg width="200" height="100" viewBox="0 0 200 100"><text x="10" y="35" font-family="NoSuchPptAgainFont" font-size="20">Editable</text></svg>');
  const measured=await measureInput(await readInput({html}),runtime);
  assert.deepEqual(measured.diagnostics.map(d=>[d.code,d.slideId]),[['FONT_SUBSTITUTED','slide-2']]);
  assert.equal(measured.model.slides[1].nodes.find(n=>n.kind==='text').style.fontFamily,'PPTAgainFallback');
  const output=await exportPptx(measured,{fontEmbedding:'auto'});
  assert.equal(output.status,'generated',JSON.stringify(output.diagnostics));
  const zip=unzipSync(output.pptxBytes);
  assert.ok(Object.keys(zip).some(path=>path.startsWith('ppt/fonts/')));
  assert.match(strFromU8(zip['ppt/slides/slide2.xml']),/typeface="Pretendard"/);
});

test('a verified installed font stays native in a hidden SVG frame',async t=>{
  if(!installedNativeAliases().some(alias=>alias.family==='Pretendard')){
    t.skip('Verified Pretendard face is not installed');
    return;
  }
  const html=deck().replace('<p class="label">Two</p>',
    '<svg width="200" height="100" viewBox="0 0 200 100"><text x="10" y="35" font-family="Pretendard" font-size="20">Editable</text></svg>');
  const measured=await measureInput(await readInput({html}),runtime);
  assert.deepEqual(measured.diagnostics,[]);
  assert.equal(measured.model.slides[1].nodes.find(n=>n.kind==='text').style.fontFamily,'Pretendard');
  const output=await exportPptx(measured);
  assert.equal(output.status,'generated',JSON.stringify(output.diagnostics));
  assert.match(strFromU8(unzipSync(output.pptxBytes)['ppt/slides/slide2.xml']),/typeface="Pretendard"/);
});
