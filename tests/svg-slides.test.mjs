import test,{before,after} from 'node:test';
import assert from 'node:assert/strict';
import {unzipSync,strFromU8} from 'fflate';
import {readInput} from '../src/html-to-pptx/input.mjs';
import {measureInput} from '../src/html-to-pptx/measure.mjs';
import {createBrowserRuntime} from '../src/html-to-pptx/browser-runtime.mjs';
import {exportPptx} from '../src/html-to-pptx/export.mjs';
import {inspectPptxIndependently} from './inspect-pptx.mjs';
let runtime;before(async()=>{runtime=await createBrowserRuntime();});after(async()=>runtime?.close());
const svg=(w=800,h=600)=>`<svg viewBox="0 0 ${w} ${h}" xmlns="http://www.w3.org/2000/svg" style="width:100%;height:auto"><defs><linearGradient id="shade"><stop stop-color="#ddeeff"/><stop offset="1" stop-color="#ffffff"/></linearGradient></defs><rect width="${w}" height="${h}" fill="url(#shade)"/><rect id="box" x="40" y="140" width="200" height="90" rx="12" fill="#123456" stroke="#112233" stroke-width="4"/><text x="40" y="70" font-family="Arial" font-size="28">Keep <tspan fill="#cc2222" font-weight="700">editable</tspan></text><line x1="40" y1="300" x2="250" y2="250" stroke="#111111" stroke-width="3"/></svg>`;
test('responsive SVG pages become native text, shapes and lines at their viewBox size',async()=>{
  const input=await readInput({html:`<style>main{width:90vw}</style><main>${svg()}</main>`});
  assert.equal(input.profile,'svg-slides-v1');
  const m=await measureInput(input,runtime),r=await exportPptx(m);assert.equal(r.status,'generated',JSON.stringify(r.diagnostics));
  const a=await inspectPptxIndependently(r.pptxBytes);
  assert.equal(a.nativeTextCount,1);assert.equal(a.nativeShapeCount,1);assert.equal(a.nativeLineCount,1);assert.equal(a.pictureCount,1);assert.equal(a.fullSlideRasterReplacementCount,0);
  const text=a.slides[0].objects.find(o=>o.text);assert.equal(text.text,'Keep editable');assert.equal(text.runs[1].bold,true);assert.equal(text.runs[1].color,'CC2222');
  const shape=a.slides[0].objects.find(o=>o.name==='box');assert.equal(shape.boxEmu.x,609600);assert.equal(shape.boxEmu.w,3048000);
  assert.deepEqual([a.widthEmu,a.heightEmu],[12192000,9144000]);assert.deepEqual([m.previews[0].width,m.previews[0].height],[1280,960]);assert.equal(m.environment.sourceScriptsExecuted,false);
});
test('SVG deck detection ignores toolbar icons and preserves portrait pages',async()=>{
  const input=await readInput({html:`<nav><button><svg viewBox="0 0 24 24"><path d="M0 0L20 20"/></svg>Next</button></nav><main>${svg(600,800)}${svg(600,800)}</main><script>document.body.remove()</script>`});
  const m=await measureInput(input,runtime),r=await exportPptx(m);assert.equal(r.status,'generated',JSON.stringify(r.diagnostics));
  const a=await inspectPptxIndependently(r.pptxBytes);assert.equal(a.slides.length,2);assert.deepEqual([a.widthEmu,a.heightEmu],[9144000,12192000]);assert.equal(a.nativeTextCount,2);
});
test('automatic SVG detection does not discard surrounding document content',async()=>{
  const input=await readInput({html:`<p>Keep this paragraph too</p>${svg()}`}),m=await measureInput(input,runtime),r=await exportPptx(m);
  assert.equal(r.status,'failed');assert.ok(r.diagnostics.some(d=>d.code==='UNSUPPORTED_OBJECT'),JSON.stringify(r.diagnostics));
});
test('unsupported SVG text placement fails explicitly instead of flattening the slide',async()=>{
  const input=await readInput({html:svg().replace('<text x="40"','<text rotate="15" x="40"')}),m=await measureInput(input,runtime),r=await exportPptx(m);
  assert.equal(r.status,'failed');assert.ok(r.diagnostics.some(d=>d.code==='UNSUPPORTED_OBJECT'));
});
test('positioned SVG lines retain styled spans as editable text',async()=>{
  const html='<section data-slide style="width:800px;height:600px"><svg width="400" height="200" viewBox="0 0 400 200"><text x="100" y="40" font-family="Arial" font-size="20"><tspan x="20" y="40"><tspan font-weight="700">Bold </tspan><tspan fill="#cc2222">first</tspan></tspan><tspan x="20" y="70"><tspan>Second </tspan><tspan font-weight="700">line</tspan></tspan></text></svg></section>';
  const m=await measureInput(await readInput({html}),runtime),r=await exportPptx(m);
  assert.equal(r.status,'generated',JSON.stringify(r.diagnostics));
  const texts=m.model.slides[0].nodes.filter(n=>n.kind==='text');
  assert.deepEqual(texts.map(n=>n.content.text),['Bold first','Second line']);
  assert.equal(texts[0].content.runs[0].style.fontWeight,700);
  assert.equal(texts[0].content.runs[1].style.color,'rgb(204, 34, 34)');
  const audit=await inspectPptxIndependently(r.pptxBytes);assert.equal(audit.nativeTextCount,2);
});
test('embedded SVG image remains a picture beside editable text',async()=>{
  const png='data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jRZkAAAAASUVORK5CYII=';
  const html=`<section data-slide style="width:800px;height:600px"><svg width="400" height="200" viewBox="0 0 400 200"><image x="20" y="20" width="100" height="100" href="${png}"/><text x="150" y="80" font-family="Arial" font-size="20">editable</text></svg></section>`;
  const m=await measureInput(await readInput({html}),runtime),r=await exportPptx(m);assert.equal(r.status,'generated',JSON.stringify(r.diagnostics));
  const audit=await inspectPptxIndependently(r.pptxBytes);assert.equal(audit.pictureCount,1);assert.equal(audit.nativeTextCount,1);
});
test('a text-free SVG containing an embedded image becomes a PPT picture',async()=>{
 const png='data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jRZkAAAAASUVORK5CYII=';
 const html=`<section data-slide style="width:800px;height:600px"><svg width="200" height="100" viewBox="0 0 200 100"><image x="0" y="0" width="100" height="100" href="${png}"/></svg></section>`;
 const m=await measureInput(await readInput({html}),runtime),r=await exportPptx(m);assert.equal(r.status,'generated',JSON.stringify(r.diagnostics));const audit=await inspectPptxIndependently(r.pptxBytes);assert.equal(audit.pictureCount,1);
});
test('a rotated and clipped SVG image is preserved as an isolated picture',async()=>{
 const png='data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jRZkAAAAASUVORK5CYII=';
 const html=`<section data-slide style="width:800px;height:600px"><svg width="200" height="100" viewBox="0 0 200 100" style="overflow:hidden"><image x="50" y="0" width="100" height="100" transform="rotate(45 100 50)" href="${png}"/></svg></section>`;
 const m=await measureInput(await readInput({html}),runtime),r=await exportPptx(m);assert.equal(r.status,'generated',JSON.stringify(r.diagnostics));const audit=await inspectPptxIndependently(r.pptxBytes);assert.equal(audit.pictureCount,1);
});
test('a blended rotated image over an opaque solid background is preserved',async()=>{
 const png='data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jRZkAAAAASUVORK5CYII=';
 const html=`<section data-slide style="width:800px;height:600px;background:rgb(19,18,37)"><div style="mix-blend-mode:lighten"><svg width="200" height="100" viewBox="0 0 200 100" style="overflow:hidden"><image x="50" y="0" width="100" height="100" transform="rotate(45 100 50)" href="${png}"/></svg></div></section>`;
 const m=await measureInput(await readInput({html}),runtime),r=await exportPptx(m);assert.equal(r.status,'generated',JSON.stringify(r.diagnostics));
});
test('a blended image with overlapping content is not flattened silently',async()=>{
 const png='data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jRZkAAAAASUVORK5CYII=';
 const html=`<section data-slide style="width:800px;height:600px;background:rgb(19,18,37);position:relative"><p style="position:absolute;top:20px;left:30px">Do not cover</p><div style="mix-blend-mode:lighten"><svg width="200" height="100" viewBox="0 0 200 100" style="overflow:hidden"><image x="50" y="0" width="100" height="100" transform="rotate(45 100 50)" href="${png}"/></svg></div></section>`;
 const m=await measureInput(await readInput({html}),runtime);assert.ok(m.diagnostics.some(d=>d.severity==='error'));
});
test('scaled SVG glyph text stays editable with a visible approximation warning',async()=>{
  const html='<section data-slide style="width:800px;height:600px"><svg width="400" height="200" viewBox="0 0 400 200"><text x="20" y="50" font-family="Arial" font-size="20" textLength="100" lengthAdjust="spacingAndGlyphs">Wide text</text></svg></section>';
  const m=await measureInput(await readInput({html}),runtime),r=await exportPptx(m);assert.equal(r.status,'generated',JSON.stringify(r.diagnostics));
  const node=m.model.slides[0].nodes.find(n=>n.kind==='text');assert.equal(node.content.text,'Wide text');assert.notEqual(node.style.letterSpacing,'normal');assert.ok(m.diagnostics.some(d=>d.code==='SVG_TEXT_APPROXIMATED'&&d.severity==='warning'));
  const audit=await inspectPptxIndependently(r.pptxBytes);assert.equal(audit.nativeTextCount,1);
});
test('nonuniformly scaled SVG text remains editable with an approximation warning',async()=>{
  const html='<section data-slide style="width:800px;height:600px"><svg width="400" height="200" viewBox="0 0 400 200"><text transform="scale(1.2 0.9)" x="20" y="50" font-family="Arial" font-size="20">Wide text</text></svg></section>';
  const m=await measureInput(await readInput({html}),runtime),r=await exportPptx(m);assert.equal(r.status,'generated',JSON.stringify(r.diagnostics));
  const node=m.model.slides[0].nodes.find(n=>n.kind==='text');assert.equal(node.content.text,'Wide text');assert.ok(Math.abs(node.style.fontSize-18)<.01);assert.ok(m.diagnostics.some(d=>d.code==='SVG_TEXT_APPROXIMATED'&&d.severity==='warning'));
});
test('ancestor clipping and marker arrows stop export instead of silently losing visible content',async()=>{
  for(const html of ['<svg viewBox="0 0 800 600"><defs><clipPath id="c"><rect width="20" height="20"/></clipPath></defs><g clip-path="url(#c)"><rect width="300" height="100"/></g></svg>','<svg viewBox="0 0 800 600"><defs><marker id="a" markerWidth="10" markerHeight="10"><path d="M0 0L10 5L0 10Z"/></marker></defs><line x1="20" y1="50" x2="300" y2="50" stroke="black" marker-end="url(#a)"/></svg>']){
    const m=await measureInput(await readInput({html}),runtime),r=await exportPptx(m);assert.equal(r.status,'failed');assert.ok(r.diagnostics.some(d=>d.code==='UNSUPPORTED_OBJECT'));
  }
});
test('independent SVG geometry preserves shifted viewBox, groups and CSS gradient stops',async()=>{
  const input=await readInput({html:'<svg viewBox="100 100 800 600"><style>.red{stop-color:rgb(220,20,40);stop-opacity:.6}</style><defs><linearGradient id="g"><stop class="red"/><stop offset="1" stop-color="blue"/></linearGradient></defs><rect x="100" y="100" width="800" height="600" fill="url(#g)"/><g transform="translate(10 20)"><circle id="disc" cx="200" cy="210" r="30" fill="#223344"/><text x="200" y="310" font-family="Arial" font-size="24" text-anchor="middle">Centered text</text></g></svg>'});
  const m=await measureInput(input,runtime),r=await exportPptx(m);assert.equal(r.status,'generated',JSON.stringify(r.diagnostics));
  const a=await inspectPptxIndependently(r.pptxBytes),disc=a.slides[0].objects.find(n=>n.name==='disc');
  assert.equal(a.groupCount,1);assert.equal(disc.shape,'ellipse');assert.equal(disc.boxEmu.x,1219200);assert.equal(disc.boxEmu.y,1524000);
  const artwork=Object.entries(unzipSync(r.pptxBytes)).filter(([name])=>name.endsWith('.svg')).map(([,b])=>strFromU8(b)).join('');
  assert.match(artwork,/stop-color: rgb\(220, 20, 40\)/);assert.match(artwork,/stop-opacity: 0.6/);assert.ok(m.measurements.every(m=>m.layoutStable));
});
test('mixed HTML and SVG preserve paint order and editability at two slide ratios',async()=>{
  for(const [width,height] of [[800,600],[600,800]]){
    const html=`<section data-slide style="position:relative;width:${width}px;height:${height}px"><p id="before" style="position:absolute;left:10px;top:10px;font:20px Arial">HTML 전</p><svg id="art" style="position:absolute;left:120px;top:100px" width="200" height="100" viewBox="20 10 100 50"><g transform="translate(10 0)"><rect id="base" x="25" y="15" width="50" height="20" fill="#225577"/><text id="label" x="30" y="30" font-family="Arial" font-size="8">SVG 글자</text></g><rect id="cover" x="70" y="20" width="10" height="10" fill="#ff0000"/></svg><p id="after" style="position:absolute;left:10px;top:300px;font:20px Arial">HTML 후</p></section>`;
    const m=await measureInput(await readInput({html}),runtime),r=await exportPptx(m);
    assert.equal(r.status,'generated',JSON.stringify(r.diagnostics));
    assert.deepEqual(m.model.slides[0].nodes.filter(n=>n.kind!=='group').map(n=>n.id),['before','art::base','art::label','art::cover','after']);
    const audit=await inspectPptxIndependently(r.pptxBytes);
    assert.equal(audit.nativeTextCount,3);assert.equal(audit.nativeShapeCount,2);assert.equal(audit.groupCount,1);assert.equal(audit.pictureCount,0);
  }
});
test('clipped mixed SVG text is rejected rather than silently extended',async()=>{
  const html='<section data-slide style="width:800px;height:600px"><svg width="100" height="50" viewBox="0 0 100 50" style="overflow:hidden"><text x="95" y="30" font-family="Arial" font-size="20">clipped text</text></svg></section>';
  const m=await measureInput(await readInput({html}),runtime),r=await exportPptx(m);
  assert.equal(r.status,'failed');assert.ok(r.diagnostics.some(d=>d.code==='UNSUPPORTED_OBJECT'),JSON.stringify(r.diagnostics));
});

test('an HTML slide containing only SVG text does not duplicate the text as HTML',async()=>{
  const html='<section data-slide style="width:800px;height:600px"><svg width="200" height="100" viewBox="0 0 200 100"><text x="10" y="30" font-family="Arial" font-size="20">only once</text></svg></section>';
  const m=await measureInput(await readInput({html}),runtime),r=await exportPptx(m);
  assert.equal(r.status,'generated',JSON.stringify(r.diagnostics));
  assert.deepEqual(m.model.slides[0].nodes.filter(n=>n.kind==='text').map(n=>n.content.text),['only once']);
});
test('mixed SVG root decoration is explicitly rejected instead of lost',async()=>{
  const html='<section data-slide style="width:800px;height:600px"><svg width="200" height="100" viewBox="0 0 200 100" style="background-color:red"><text x="10" y="30" font-family="Arial" font-size="20">text</text></svg></section>';
  const m=await measureInput(await readInput({html}),runtime),r=await exportPptx(m);
  assert.equal(r.status,'failed');assert.ok(r.diagnostics.some(d=>d.code==='UNSUPPORTED_OBJECT'),JSON.stringify(r.diagnostics));
});

test('SVG definitions containing text cannot silently become artwork',async()=>{
  const html='<section data-slide style="width:800px;height:600px"><svg width="200" height="100" viewBox="0 0 200 100"><defs><pattern id="p" width="30" height="30" patternUnits="userSpaceOnUse"><text x="2" y="20" font-family="Arial" font-size="10">pattern text</text></pattern></defs><rect width="100" height="50" fill="url(#p)"/><text x="110" y="40" font-family="Arial" font-size="12">visible text</text></svg></section>';
  const m=await measureInput(await readInput({html}),runtime),r=await exportPptx(m);
  assert.equal(r.status,'failed');assert.ok(r.diagnostics.some(d=>d.code==='UNSUPPORTED_OBJECT'),JSON.stringify(r.diagnostics));
});

test('a clipped SVG stroke is rejected even when the shape bounds fit',async()=>{
  const html='<section data-slide style="width:800px;height:600px"><svg style="position:absolute;left:120px;top:80px;overflow:hidden" width="100" height="100" viewBox="0 0 100 100"><rect x="0" y="0" width="100" height="100" fill="white" stroke="red" stroke-width="8"/><text x="20" y="30" font-family="Arial" font-size="10">text</text></svg></section>';
  const m=await measureInput(await readInput({html}),runtime),r=await exportPptx(m);
  assert.equal(r.status,'failed');assert.ok(r.diagnostics.some(d=>d.code==='UNSUPPORTED_OBJECT'),JSON.stringify(r.diagnostics));
});

test('central baseline and one positioned tspan keep measured editable text boxes',async()=>{
  const html='<section data-slide style="position:relative;width:800px;height:600px"><svg style="position:absolute;left:120px;top:90px" width="200" height="100" viewBox="0 0 200 100"><text x="20" y="25" font-family="Arial" font-size="20" dominant-baseline="central">central</text><text y="55" font-family="Arial" font-size="20"><tspan x="20" dy="8">positioned</tspan></text></svg></section>';
  const input=await readInput({html}),session=await runtime.render(input);
  let bounds;
  try{bounds=await session.page.evaluate(()=>{const slide=document.querySelector('[data-slide]').getBoundingClientRect();return [...document.querySelectorAll('svg text')].map(t=>{const r=t.getBoundingClientRect();return {x:r.x-slide.x,y:r.y-slide.y,w:r.width,h:r.height};});});}finally{await session.context.close();}
  const m=await measureInput(input,runtime),r=await exportPptx(m);
  assert.equal(r.status,'generated',JSON.stringify(r.diagnostics));
  const nodes=m.model.slides[0].nodes.filter(n=>n.kind==='text');assert.deepEqual(nodes.map(n=>n.content.text),['central','positioned']);
  nodes.forEach((node,i)=>{for(const key of ['x','y','w','h'])assert.ok(Math.abs(node.boxCssPx[key]-bounds[i][key])<.05,`${key}: ${node.boxCssPx[key]} vs ${bounds[i][key]}`);});
  const audit=await inspectPptxIndependently(r.pptxBytes);assert.equal(audit.nativeTextCount,2);assert.equal(audit.pictureCount,0);
  for(const [i,node] of nodes.entries()){const object=audit.slides[0].objects.find(o=>o.text===node.content.text);for(const key of ['x','y','w','h'])assert.ok(Math.abs(object.boxEmu[key]/audit.widthEmu*800-bounds[i][key])<.05);}
});

test('mixed direct text and positioned tspans remain unsupported',async()=>{
  const html='<section data-slide style="width:800px;height:600px"><svg width="300" height="120" viewBox="0 0 300 120"><text y="20" font-family="Arial">prefix<tspan x="10" dy="20">first</tspan><tspan x="10" dy="20">second</tspan></text></svg></section>';
  const m=await measureInput(await readInput({html}),runtime),r=await exportPptx(m);
  assert.equal(r.status,'failed');assert.ok(r.diagnostics.some(d=>d.code==='UNSUPPORTED_OBJECT'),JSON.stringify(r.diagnostics));
});

test('direct positioned tspan lines retain current content, styles and measured coordinates',async()=>{
  const html='<section data-slide style="position:relative;width:800px;height:600px"><svg style="position:absolute;left:120px;top:90px" width="200" height="140" viewBox="0 0 200 140"><text font-family="Arial" font-size="16"><tspan x="20" y="25">first</tspan><tspan x="20" y="50" fill="red" font-weight="700">second</tspan></text><text y="70" font-family="Arial" font-size="16" dominant-baseline="central"><tspan x="20" dy="10">third</tspan><tspan x="20" dy="25">fourth</tspan></text></svg></section>';
  const input=await readInput({html}),session=await runtime.render(input);
  let bounds;try{bounds=await session.page.evaluate(()=>{const slide=document.querySelector('[data-slide]').getBoundingClientRect();return [...document.querySelectorAll('svg tspan')].map(t=>{const r=t.getBoundingClientRect();return {x:r.x-slide.x,y:r.y-slide.y,w:r.width,h:r.height};});});}finally{await session.context.close();}
  const m=await measureInput(input,runtime),r=await exportPptx(m);assert.equal(r.status,'generated',JSON.stringify(r.diagnostics));
  const nodes=m.model.slides[0].nodes.filter(n=>n.kind==='text');assert.deepEqual(nodes.map(n=>n.content.text),['first','second','third','fourth']);assert.equal(nodes[1].style.fontWeight,700);
  nodes.forEach((node,i)=>{for(const key of ['x','y','w','h'])assert.ok(Math.abs(node.boxCssPx[key]-bounds[i][key])<.05);});
  const audit=await inspectPptxIndependently(r.pptxBytes);assert.equal(audit.nativeTextCount,4);assert.equal(audit.pictureCount,0);
  for(const [i,node] of nodes.entries()){const object=audit.slides[0].objects.find(o=>o.text===node.content.text);for(const key of ['x','y','w','h'])assert.ok(Math.abs(object.boxEmu[key]/audit.widthEmu*800-bounds[i][key])<.05);}
});

test('positioned tspan outlines are not silently omitted',async()=>{
  const html='<section data-slide style="width:800px;height:600px"><svg width="200" height="100" viewBox="0 0 200 100"><text y="30" font-family="Arial"><tspan x="10" dy="10" stroke="red" stroke-width="1">outline</tspan></text></svg></section>';
  const m=await measureInput(await readInput({html}),runtime),r=await exportPptx(m);assert.equal(r.status,'failed');assert.ok(r.diagnostics.some(d=>d.code==='UNSUPPORTED_OBJECT'),JSON.stringify(r.diagnostics));
});
test('a simple solid SVG text outline remains editable with its outline',async()=>{
 const html='<section data-slide style="width:400px;height:300px"><svg width="200" height="100" viewBox="0 0 200 100"><text x="10" y="35" font-family="Arial" font-size="20" fill="white" stroke="#123456" stroke-width="1">Editable</text></svg></section>';
 const m=await measureInput(await readInput({html}),runtime),r=await exportPptx(m);assert.equal(r.status,'generated',JSON.stringify(r.diagnostics));const node=m.model.slides[0].nodes.find(n=>n.kind==='text');assert.equal(node.content.text,'Editable');assert.ok(Math.abs(node.content.runs[0].style.outline.width-1)<0.001);
});

test('SVG positioned tspans with tab separated character coordinates remain unsupported',async()=>{
  for(const attr of ['x="20&#9;100" dy="8"','x="20" dy="0&#9;20"']){
    const html=`<section data-slide style="width:800px;height:600px"><svg width="200" height="100" viewBox="0 0 200 100"><text y="50" font-family="Arial"><tspan ${attr}>AB</tspan></text></svg></section>`;
    const m=await measureInput(await readInput({html}),runtime),r=await exportPptx(m);assert.equal(r.status,'failed');assert.ok(r.diagnostics.some(d=>d.code==='UNSUPPORTED_OBJECT'),JSON.stringify(r.diagnostics));
  }
});

test('a text-free nested SVG icon preserves its clipped viewport as one ordered picture',async()=>{
  const html='<style>.cut{fill:#e84040}</style><section data-slide style="position:relative;width:400px;height:300px"><svg id="art" width="200" height="140" viewBox="0 0 200 140"><rect id="below" x="100" y="80" width="20" height="20"/><svg id="icon" x="20" y="10" width="60" height="40" viewBox="0 0 30 20" style="overflow:hidden"><rect class="cut" x="-5" y="0" width="20" height="20"/><circle cx="27" cy="10" r="7" fill="#164e63"/></svg><rect id="above" x="150" y="80" width="20" height="20"/><text id="label" x="20" y="90" font-family="Arial">editable</text></svg></section>';
  const input=await readInput({html}),session=await runtime.render(input);let original;
  try{const slide=await session.page.locator('[data-slide]').boundingBox();original=await session.page.screenshot({clip:{x:slide.x+20,y:slide.y+10,width:60,height:40},omitBackground:true});}finally{await session.context.close();}
  const m=await measureInput(input,runtime),r=await exportPptx(m);assert.equal(r.status,'generated',JSON.stringify(r.diagnostics));
  const nodes=m.model.slides[0].nodes.filter(n=>n.kind!=='group');assert.deepEqual(nodes.map(n=>n.id),['art::below','art::icon','art::above','art::label']);
  const picture=nodes[1],asset=m.model.assets.find(a=>a.id===picture.content.assetId);assert.equal(picture.kind,'image');assert.equal(picture.content.artworkOnly,true);assert.deepEqual(picture.boxCssPx,{x:20,y:10,w:60,h:40});
  const second=await runtime.render(input);try{
    const page=await second.context.newPage();await page.setViewportSize({width:60,height:40});
    await page.setContent(`<style>html,body{margin:0;padding:0}</style><img width="60" height="40" src="data:image/svg+xml;base64,${Buffer.from(asset.bytes).toString('base64')}">`);
    const candidate=await page.locator('img').screenshot({omitBackground:true});
    const diff=await page.evaluate(async pair=>{const images=[];for(const value of pair){const im=new Image();im.src='data:image/png;base64,'+value;await im.decode();images.push(im);}const cvs=document.createElement('canvas');cvs.width=60;cvs.height=40;const ctx=cvs.getContext('2d');const pixels=images.map(im=>{ctx.clearRect(0,0,60,40);ctx.drawImage(im,0,0);return ctx.getImageData(0,0,60,40).data;});let changed=0;for(let i=0;i<pixels[0].length;i+=4)if([0,1,2,3].some(k=>pixels[0][i+k]!==pixels[1][i+k]))changed++;return changed;},[original.toString('base64'),candidate.toString('base64')]);
    assert.equal(diff,0,'browser pixels changed');
  }finally{await second.context.close();}
});
test('nested SVG text or references remain unsupported',async()=>{
  for(const inner of ['<text x="1" y="10">hidden</text><rect width="30" height="20"/>','<use href="#shape"/>']){
    const html=`<section data-slide style="width:400px;height:300px"><svg width="200" height="140" viewBox="0 0 200 140"><svg x="20" y="10" width="60" height="40" viewBox="0 0 30 20">${inner}</svg><text x="20" y="90" font-family="Arial">outside</text></svg></section>`;
    const m=await measureInput(await readInput({html}),runtime),r=await exportPptx(m);assert.equal(r.status,'failed');assert.ok(r.diagnostics.some(d=>d.code==='UNSUPPORTED_OBJECT'));
  }
});
test('nested SVG separates its text from non-overlapping artwork',async()=>{
 const html='<section data-slide style="width:400px;height:300px"><svg width="200" height="140" viewBox="0 0 200 140"><svg x="20" y="10" width="120" height="80" viewBox="0 0 120 80" style="overflow:hidden"><rect width="120" height="80" fill="#123456"/><text x="10" y="25" font-family="Arial" font-size="16" fill="white">inside</text><path d="M10 60L100 60" stroke="red"/></svg><text x="20" y="120" font-family="Arial">outside</text></svg></section>';
 const m=await measureInput(await readInput({html}),runtime),r=await exportPptx(m);assert.equal(r.status,'generated',JSON.stringify(r.diagnostics));const nodes=m.model.slides[0].nodes.filter(n=>n.kind!=='group');assert.deepEqual(nodes.map(n=>n.kind),['image','text','text']);assert.equal(nodes[1].content.text,'inside');assert.equal(nodes[2].content.text,'outside');
});
test('nested SVG artwork can use its own local gradient and retains editable text outside',async()=>{
  const html='<section data-slide style="width:400px;height:300px"><svg width="200" height="140" viewBox="0 0 200 140"><svg x="20" y="10" width="60" height="40" viewBox="0 0 30 20" style="overflow:hidden"><defs><linearGradient id="local-g"><stop stop-color="red"/><stop offset="1" stop-color="blue"/></linearGradient></defs><rect width="30" height="20" fill="url(#local-g)"/></svg><text x="20" y="90" font-family="Arial">outside</text></svg></section>';
  const m=await measureInput(await readInput({html}),runtime),r=await exportPptx(m);assert.equal(r.status,'generated',JSON.stringify(r.diagnostics));
  const nodes=m.model.slides[0].nodes.filter(n=>n.kind!=='group');assert.deepEqual(nodes.map(n=>n.kind),['image','text']);assert.equal(nodes[1].content.text,'outside');
  const asset=m.model.assets.find(a=>a.id===nodes[0].content.assetId);assert.match(Buffer.from(asset.bytes).toString('utf8'),/linearGradient/);
});
test('nested SVG artwork does not borrow a gradient outside its own viewport',async()=>{
  const html='<section data-slide style="width:400px;height:300px"><svg width="200" height="140" viewBox="0 0 200 140"><defs><linearGradient id="outside"><stop stop-color="red"/></linearGradient></defs><svg x="20" y="10" width="60" height="40" viewBox="0 0 30 20" style="overflow:hidden"><rect width="30" height="20" fill="url(#outside)"/></svg><text x="20" y="90" font-family="Arial">outside</text></svg></section>';
  const m=await measureInput(await readInput({html}),runtime);assert.ok(m.diagnostics.some(d=>d.severity==='error'));
});

test('hidden SVG text descendants cannot leak into an editable PPT',async()=>{
  for(const style of ['display:none','visibility:hidden','opacity:0']){
    const html=`<section data-slide style="width:400px;height:300px"><svg width="200" height="140" viewBox="0 0 200 140"><text x="10" y="30" font-family="Arial">VISIBLE<tspan style="${style}">HIDDEN</tspan></text></svg></section>`;
    const m=await measureInput(await readInput({html}),runtime),r=await exportPptx(m);
    assert.equal(r.status,'failed');assert.ok(r.diagnostics.some(d=>d.code==='UNSUPPORTED_OBJECT'));
  }
});

test('visible overflow on a nested SVG cannot silently clip artwork',async()=>{
  const html='<section data-slide style="width:400px;height:300px"><svg width="200" height="140" viewBox="0 0 200 140"><svg x="20" y="10" width="60" height="40" viewBox="0 0 30 20" style="overflow:visible"><rect x="25" y="2" width="15" height="12"/></svg><text x="20" y="90" font-family="Arial">text</text></svg></section>';
  const m=await measureInput(await readInput({html}),runtime),r=await exportPptx(m);
  assert.equal(r.status,'failed');assert.ok(r.diagnostics.some(d=>d.code==='UNSUPPORTED_OBJECT'));
});

test('fractional nested SVG viewport remains supported',async()=>{
  const html='<section data-slide style="width:400px;height:300px"><svg width="200" height="140" viewBox="0 0 200 140"><svg x="20.4" y="10.82" width="47.52" height="47.52" viewBox="0 0 48 48" style="overflow:hidden"><rect width="48" height="48"/></svg><text x="20" y="90" font-family="Arial">text</text></svg></section>';
  const m=await measureInput(await readInput({html}),runtime),r=await exportPptx(m);assert.equal(r.status,'generated',JSON.stringify(r.diagnostics));
  const icon=m.model.slides[0].nodes.find(n=>n.kind==='image');assert.ok(Math.abs(icon.boxCssPx.x-20.4)<.001);assert.ok(Math.abs(icon.boxCssPx.w-47.52)<.001);
});

test('a transform on a nested SVG viewport cannot silently move its picture',async()=>{
  for(const transform of ['transform="translate(10 0)"','style="overflow:hidden;transform:translateX(10px)"']){
    const html=`<section data-slide style="width:400px;height:300px"><svg width="200" height="140" viewBox="0 0 200 140"><svg x="20" y="10" width="60" height="40" viewBox="0 0 30 20" ${transform}><rect width="20" height="20"/></svg><text x="20" y="90" font-family="Arial">text</text></svg></section>`;
    const m=await measureInput(await readInput({html}),runtime),r=await exportPptx(m);assert.equal(r.status,'failed');assert.ok(r.diagnostics.some(d=>d.code==='UNSUPPORTED_OBJECT'));
  }
});

test('mixed SVG marker artwork preserves viewport clipping without rasterizing text',async()=>{
  const html='<section data-slide style="width:400px;height:300px"><svg id="art" width="200" height="140" viewBox="0 0 200 140" style="overflow:hidden"><defs><marker id="arrow" markerWidth="12" markerHeight="10" refX="10" refY="5" orient="auto"><path d="M0 0 L12 5 L0 10Z" fill="#164e63"/></marker></defs><polyline id="line" points="20,70 195,70" fill="none" stroke="#164e63" stroke-width="2" marker-end="url(#arrow)"/><text x="20" y="110" font-family="Arial">editable</text></svg></section>';
  const input=await readInput({html}),session=await runtime.render(input);let original;
  try{await session.page.locator('svg text').evaluate(el=>el.style.display='none');original=await session.page.locator('svg').screenshot({omitBackground:true});}finally{await session.context.close();}
  const m=await measureInput(input,runtime),r=await exportPptx(m);assert.equal(r.status,'generated',JSON.stringify(r.diagnostics));
  const nodes=m.model.slides[0].nodes.filter(n=>n.kind!=='group');assert.deepEqual(nodes.map(n=>n.kind),['image','text']);assert.equal(nodes[1].content.text,'editable');
  const asset=m.model.assets.find(a=>a.id===nodes[0].content.assetId);assert.ok(asset);assert.doesNotMatch(Buffer.from(asset.bytes).toString('utf8'),/<text\b/i);
  const second=await runtime.render(input);try{
    const page=await second.context.newPage();await page.setViewportSize({width:200,height:140});
    await page.setContent(`<style>html,body{margin:0;padding:0}</style><img width="200" height="140" src="data:image/svg+xml;base64,${Buffer.from(asset.bytes).toString('base64')}">`);
    const candidate=await page.locator('img').screenshot({omitBackground:true});
    const diff=await page.evaluate(async pair=>{const images=[];for(const value of pair){const im=new Image();im.src='data:image/png;base64,'+value;await im.decode();images.push(im);}const cvs=document.createElement('canvas');cvs.width=200;cvs.height=140;const ctx=cvs.getContext('2d');const pixels=images.map(im=>{ctx.clearRect(0,0,200,140);ctx.drawImage(im,0,0);return ctx.getImageData(0,0,200,140).data;});let changed=0;for(let i=0;i<pixels[0].length;i+=4)if([0,1,2,3].some(k=>pixels[0][i+k]!==pixels[1][i+k]))changed++;return changed;},[original.toString('base64'),candidate.toString('base64')]);
    assert.equal(diff,0,'marker appearance or clipping changed');
  }finally{await second.context.close();}
});

test('marker definitions with text or unsupported placement stay explicit errors',async()=>{
  for(const [definition,attribute,overflow] of [
    ['<text x="1" y="5">hidden marker text</text>','marker-end="url(#arrow)"','hidden'],
    ['<path d="M0 0 L8 4 L0 8Z"/>','marker-start="url(#arrow)"','hidden'],
    ['<path d="M0 0 L8 4 L0 8Z"/>','marker-end="url(#arrow)"','visible']
  ]){
    const html=`<section data-slide style="width:400px;height:300px"><svg width="200" height="140" viewBox="0 0 200 140" style="overflow:${overflow}"><defs><marker id="arrow" markerWidth="8" markerHeight="8">${definition}</marker></defs><polyline points="20,70 150,70" fill="none" stroke="black" ${attribute}/><text x="20" y="110" font-family="Arial">editable</text></svg></section>`;
    const m=await measureInput(await readInput({html}),runtime),r=await exportPptx(m);assert.equal(r.status,'failed');assert.ok(r.diagnostics.some(d=>d.code==='UNSUPPORTED_OBJECT'));
  }
});

test('line and path markers beyond a visible SVG viewport preserve the slide clip',async()=>{
  for(const element of ['<line x1="30" y1="60" x2="130" y2="60"','<path d="M30 60 L130 60"']){
    const html=`<section data-slide style="width:400px;height:300px"><div style="width:400px;height:300px;overflow:hidden"><svg width="100" height="100" viewBox="0 0 100 100" style="overflow:visible"><defs><marker id="arrow" markerWidth="10" markerHeight="10" refX="5" refY="5" orient="auto"><path d="M0 0 L10 5 L0 10Z" fill="#164e63"/></marker></defs><text x="5" y="30" font-family="Arial">before</text>${element} fill="none" stroke="#164e63" stroke-width="2" marker-end="url(#arrow)"/><text x="5" y="90" font-family="Arial">after</text></svg></div></section>`;
    const input=await readInput({html}),first=await runtime.render(input);let original;
    try{await first.page.locator('svg text').evaluateAll(els=>els.forEach(el=>el.style.display='none'));original=await first.page.locator('[data-slide]').screenshot({omitBackground:true});}finally{await first.context.close();}
    const m=await measureInput(input,runtime),r=await exportPptx(m);assert.equal(r.status,'generated',JSON.stringify(r.diagnostics));
    const nodes=m.model.slides[0].nodes.filter(n=>n.kind!=='group');assert.deepEqual(nodes.map(n=>n.kind),['text','image','text']);
    assert.deepEqual(nodes[1].boxCssPx,{x:0,y:0,w:400,h:300});
    const asset=m.model.assets.find(a=>a.id===nodes[1].content.assetId);assert.ok(asset);assert.doesNotMatch(Buffer.from(asset.bytes).toString('utf8'),/<text\b/i);
    const next=await runtime.render(input);try{const page=await next.context.newPage();await page.setViewportSize({width:400,height:300});await page.setContent(`<style>html,body{margin:0;padding:0}</style><img width="400" height="300" src="data:image/svg+xml;base64,${Buffer.from(asset.bytes).toString('base64')}">`);
      const candidate=await page.locator('img').screenshot({omitBackground:true});assert.deepEqual(candidate,original);
    }finally{await next.context.close();}
  }
});

test('visible SVG marker with a smaller ancestor clip stays unsupported',async()=>{
  const html='<section data-slide style="width:400px;height:300px"><div style="width:100px;height:100px;overflow:hidden"><svg width="100" height="100" style="overflow:visible"><defs><marker id="arrow"><path d="M0 0 L8 4 L0 8Z"/></marker></defs><line x1="10" y1="50" x2="150" y2="50" stroke="black" marker-end="url(#arrow)"/><text x="2" y="10" font-family="Arial">editable</text></svg></div></section>';
  const r=await exportPptx(await measureInput(await readInput({html}),runtime));assert.equal(r.status,'failed');assert.ok(r.diagnostics.some(d=>d.code==='UNSUPPORTED_OBJECT'));
});

test('clipped marker artwork refuses smaller ancestor clips and rotated SVG viewports',async()=>{
  for(const [wrapper,svgStyle] of [['width:100px;height:100px;overflow:hidden','overflow:hidden'],['width:400px;height:300px','overflow:hidden;transform:rotate(20deg)']]){
    const html=`<section data-slide style="width:400px;height:300px"><div style="${wrapper}"><svg width="200" height="140" style="${svgStyle}"><defs><marker id="arrow"><path d="M0 0 L8 4 L0 8Z"/></marker></defs><polyline points="10,50 150,50" fill="none" stroke="black" marker-end="url(#arrow)"/><text x="2" y="30" font-family="Arial">editable</text></svg></div></section>`;
    const r=await exportPptx(await measureInput(await readInput({html}),runtime));assert.equal(r.status,'failed');assert.ok(r.diagnostics.some(d=>d.code==='UNSUPPORTED_OBJECT'));
  }
});
