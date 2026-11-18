import test, {before, after} from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import {spawnSync} from 'node:child_process';
import {readInput} from '../src/html-to-pptx/input.mjs';
import {measureInput} from '../src/html-to-pptx/measure.mjs';
import {createBrowserRuntime} from '../src/html-to-pptx/browser-runtime.mjs';
import {exportPptx} from '../src/html-to-pptx/export.mjs';
import {unzipSync,strFromU8} from 'fflate';

let runtime;
before(async () => { runtime = await createBrowserRuntime({capture:false}); });
after(async () => { await runtime?.close(); });

const png = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+j1ioAAAAASUVORK5CYII=';
const card = '<div id="slide" style="width:400px;height:300px;background:white"><p style="margin:0;font:20px Arial">CARD</p></div>';
const documentWith = (before = '', after = '') => `<!doctype html><html><body style="margin:0">${before}${card}${after}</body></html>`;
const inspect = async (html, slideSelector = '#slide') => measureInput(await readInput({html,slideSelector}),runtime);
const outside = result => result.diagnostics.filter(d => d.code === 'SLIDE_CONTENT_OUTSIDE');

test('explicit selection warns about visible text before and after its candidate',async()=>{
  const measured=await inspect(documentWith('<h1>REPORT</h1>','<p>FOOTER</p>'));
  assert.ok(outside(measured).some(d=>d.severity==='warning'),JSON.stringify(measured.diagnostics));
  assert.equal((await exportPptx(measured)).status,'generated');
});

test('explicit selection warns about a visible image outside its candidate',async()=>{
  const measured=await inspect(documentWith('',`<img src="${png}" width="40" height="40" alt="outside">`));
  assert.ok(outside(measured).some(d=>d.severity==='warning'),JSON.stringify(measured.diagnostics));
});

test('explicit selection inspects direct text on an ancestor of the candidate',async()=>{
  const measured=await inspect(`<body style="margin:0"><main>Intro${card}Outro</main></body>`);
  assert.ok(outside(measured).some(d=>d.severity==='warning'),JSON.stringify(measured.diagnostics));
});

test('explicit selection checks document content below the current viewport',async()=>{
  const measured=await inspect(documentWith('','<p style="position:absolute;top:3000px">BELOW</p>'));
  assert.ok(outside(measured).some(d=>d.severity==='warning'),JSON.stringify(measured.diagnostics));
});

test('explicit selector preserves the chosen card and warns about outside content',async()=>{
  const measured=await inspect(documentWith('<h1>REPORT</h1>','<p>FOOTER</p>'),'#slide');
  assert.ok(outside(measured).some(d=>d.severity==='warning'),JSON.stringify(measured.diagnostics));
  assert.equal(outside(measured).some(d=>d.severity==='error'),false,JSON.stringify(measured.diagnostics));
  assert.equal(measured.model?.slides.length,1);
  const text=measured.model.slides[0].nodes.filter(n=>n.kind==='text').map(n=>n.content.text).join(' ');
  assert.match(text,/CARD/);
  assert.doesNotMatch(text,/REPORT|FOOTER/);
  const result=await exportPptx(measured);
  assert.equal(result.status,'generated',JSON.stringify(result.diagnostics));
  const xml=strFromU8(unzipSync(result.pptxBytes)['ppt/slides/slide1.xml']);
  assert.match(xml,/CARD/);
  assert.doesNotMatch(xml,/REPORT|FOOTER/);
  assert.ok(result.diagnostics.some(d=>d.code==='SLIDE_CONTENT_OUTSIDE'&&d.severity==='warning'));
});

test('display none outside the candidate does not produce an outside warning',async()=>{
  const measured=await inspect(documentWith('<aside style="display:none">HIDDEN</aside>'));
  assert.deepEqual(outside(measured),[],JSON.stringify(measured.diagnostics));
  assert.equal((await exportPptx(measured)).status,'generated');
});

test('transparent ancestor outside the candidate does not produce an outside warning',async()=>{
  const measured=await inspect(documentWith('<aside style="opacity:0"><span>HIDDEN</span></aside>'));
  assert.deepEqual(outside(measured),[],JSON.stringify(measured.diagnostics));
  assert.equal((await exportPptx(measured)).status,'generated');
});

test('fully clipped outside text does not produce an outside warning',async()=>{
  const clipped='<aside style="position:relative;width:50px;height:30px;overflow:hidden"><span style="position:absolute;left:100px;top:0">HIDDEN</span></aside>';
  const measured=await inspect(documentWith(clipped));
  assert.deepEqual(outside(measured),[],JSON.stringify(measured.diagnostics));
  assert.equal((await exportPptx(measured)).status,'generated');
});

test('declared data-slide retains its existing interpretation',async()=>{
  const html='<h1>PAGE CHROME</h1><section data-slide style="width:400px;height:300px"><p style="font:20px Arial">SLIDE</p></section>';
  const measured=await inspect(html,null);
  assert.deepEqual(outside(measured),[],JSON.stringify(measured.diagnostics));
  assert.equal((await exportPptx(measured)).status,'generated',JSON.stringify(measured.diagnostics));
});

test('existing single SVG slide retains its existing interpretation',async()=>{
  const html='<svg viewBox="0 0 800 600" width="800" height="600" xmlns="http://www.w3.org/2000/svg"><rect x="20" y="20" width="200" height="100" fill="#336699"/></svg>';
  const input=await readInput({html});
  assert.equal(input.profile,'svg-slides-v1');
  const measured=await measureInput(input,runtime);
  assert.deepEqual(outside(measured),[],JSON.stringify(measured.diagnostics));
  assert.equal((await exportPptx(measured)).status,'generated',JSON.stringify(measured.diagnostics));
});

test('CLI requires a selection and does not create a PPT for unmarked HTML',async t=>{
  const dir=await fs.mkdtemp(path.join(os.tmpdir(),'ppt-region-safety-'));
  const input=path.join(dir,'input.html'),output=path.join(dir,'output.pptx');
  t.after(async()=>{
    await fs.unlink(input).catch(()=>{});
    await fs.unlink(output).catch(()=>{});
    await fs.rmdir(dir);
  });
  await fs.writeFile(input,documentWith('<h1>PRIVATE_HEADING</h1>','<p>PRIVATE_FOOTER</p>'));
  const result=spawnSync(process.execPath,[path.resolve('src/html-to-pptx/cli.mjs'),'convert',input,'--output',output,'--json'],{encoding:'utf8',timeout:120000});
  assert.equal(result.status,1,result.stdout||result.stderr);
  const report=JSON.parse(result.stdout);
  assert.equal(report.status,'failed');
  assert.equal(report.code,'SELECTION_REQUIRED');
  assert.ok(report.diagnostics.some(d=>d.code==='SELECTION_REQUIRED'&&d.severity==='error'),JSON.stringify(report));
  assert.doesNotMatch(result.stdout,/PRIVATE_HEADING|PRIVATE_FOOTER/);
  await assert.rejects(fs.access(output),{code:'ENOENT'});
});
test('explicit selection warns about a small SVG outside its candidate',async()=>{
  const icon='<svg width="40" height="40" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg"><rect x="0" y="0" width="24" height="24" fill="#cc0000"/></svg>';
  const input=await readInput({html:documentWith('',icon),slideSelector:'#slide'});
  assert.equal(input.profile,'static-slides-v1');
  const measured=await measureInput(input,runtime);
  assert.ok(outside(measured).some(d=>d.severity==='warning'),JSON.stringify(measured.diagnostics));
});

test('explicit selection warns about a visible colored non-text element outside its candidate',async()=>{
  const decoration='<aside style="width:40px;height:30px;background-color:#cc0000"></aside>';
  const measured=await inspect(documentWith('',decoration));
  assert.ok(outside(measured).some(d=>d.severity==='warning'),JSON.stringify(measured.diagnostics));
});
test('offscreen loaded image is detected while an empty transparent wrapper is ignored',async()=>{
  const empty=await inspect(documentWith('','<aside style="width:40px;height:30px;background:transparent"></aside>'));
  assert.deepEqual(outside(empty),[]);
  const image=await inspect(documentWith('',`<img src="${png}" style="position:absolute;top:3000px;width:40px;height:40px">`));
  assert.ok(outside(image).some(d=>d.severity==='warning'));
});

test('a partially transparent colored area is detected',async()=>{
  const measured=await inspect(documentWith('','<aside style="width:40px;height:30px;background:rgba(255,0,0,.2)"></aside>'));
  assert.ok(outside(measured).some(d=>d.severity==='warning'));
});

test('hidden content and executable scripts report an observation limit',async()=>{
  for(const addition of ['<aside style="display:none">HIDDEN</aside>','<script>document.body.append("DYNAMIC")</script>']){
    const measured=await inspect(documentWith(addition));
    assert.ok(measured.diagnostics.some(d=>d.code==='SLIDE_CONTENT_UNOBSERVED'&&d.severity==='warning'));
    assert.deepEqual(outside(measured),[]);
  }
  const json=await inspect(documentWith('<script type="application/json">{"value":1}</script>'));
  assert.equal(json.diagnostics.some(d=>d.code==='SLIDE_CONTENT_UNOBSERVED'),false);
});

test('direct text fully clipped by its own element is not outside content',async()=>{
  const html=documentWith('<aside style="width:20px;height:20px;overflow:hidden;text-indent:100px;white-space:nowrap">HIDDEN</aside>');
  const measured=await inspect(html);
  assert.deepEqual(outside(measured),[],JSON.stringify(measured.diagnostics));
});

test('transparent text is ignored but visible text overriding inherited visibility is detected',async()=>{
  const transparent=await inspect(documentWith('<span style="color:transparent">HIDDEN</span>'));
  assert.deepEqual(outside(transparent),[],JSON.stringify(transparent.diagnostics));
  const visible=await inspect(documentWith('<aside style="visibility:hidden"><span style="visibility:visible">VISIBLE</span></aside>'));
  assert.ok(outside(visible).some(d=>d.severity==='warning'));
});

test('document scrolling does not hide loaded text or images below the viewport',async()=>{
  for(const item of ['<p style="position:absolute;top:3000px">BELOW</p>',`<img src="${png}" style="position:absolute;top:3000px;width:40px;height:40px">`]){
    const html=`<style>html,body{height:100%;overflow:auto}</style><body>${card}${item}</body>`;
    const measured=await inspect(html);
    assert.ok(outside(measured).some(d=>d.severity==='warning'),JSON.stringify(measured.diagnostics));
  }
});

test('script MIME parameters still report the observation limit',async()=>{
  const measured=await inspect(documentWith('<script type="text/javascript; charset=utf-8">void 0</script>'));
  assert.ok(measured.diagnostics.some(d=>d.code==='SLIDE_CONTENT_UNOBSERVED'&&d.severity==='warning'));
});

test('overflow on a non-replaced inline element does not hide visible text',async()=>{
  const measured=await inspect(documentWith('<span style="overflow:hidden">VISIBLE</span>'));
  assert.ok(outside(measured).some(d=>d.severity==='warning'),JSON.stringify(measured.diagnostics));
});

test('an internal scroll container clips content outside its visible area',async()=>{
  const hidden='<aside style="position:relative;width:50px;height:30px;overflow:auto"><span style="position:absolute;top:100px">HIDDEN</span></aside>';
  const measured=await inspect(documentWith(hidden));
  assert.deepEqual(outside(measured),[],JSON.stringify(measured.diagnostics));
  assert.ok(measured.diagnostics.some(d=>d.code==='SLIDE_CONTENT_UNOBSERVED'&&d.severity==='warning'));
});

test('generated content outside the slide is detected for automatic and explicit selection',async()=>{
  const html=documentWith('<style>aside::before{content:"PRIVATE_GENERATED"}</style><aside></aside>');
  const auto=await inspect(html,null);
  assert.ok(auto.diagnostics.some(d=>d.code==='SELECTION_REQUIRED'&&d.severity==='error'),JSON.stringify(auto.diagnostics));
  const explicit=await inspect(html,'#slide');
  assert.ok(outside(explicit).some(d=>d.severity==='warning'),JSON.stringify(explicit.diagnostics));
  assert.doesNotMatch(JSON.stringify(outside(explicit)),/PRIVATE_GENERATED|#slide/);
});

test('hidden, transparent, clipped, and absent generated content are not visible outside',async()=>{
  for(const style of ['content:none','content:"HIDDEN";display:none','content:"HIDDEN";opacity:0','content:"HIDDEN";color:transparent','content:"HIDDEN";position:absolute;left:100px']){
    const html=documentWith(`<style>aside::before{${style}}</style><aside style="position:relative;width:20px;height:20px;overflow:hidden"></aside>`);
    const measured=await inspect(html);
    assert.deepEqual(outside(measured),[],JSON.stringify(measured.diagnostics));
  }
});

test('generated-content scan budget reports an unexamined explicit scope',async()=>{
  const boxes='<aside></aside>'.repeat(129);
  const html=documentWith(`<style>aside{position:relative;width:1px;height:1px;overflow:hidden}aside::before{content:"x";position:absolute;left:100px}</style>${boxes}`);
  const measured=await inspect(html);
  assert.ok(measured.diagnostics.some(d=>d.code==='SLIDE_CONTENT_UNOBSERVED'&&d.severity==='warning'));
});

test('empty transparent generated boxes are not visible content',async()=>{
  const html=documentWith('<style>aside::before{content:"";display:block;width:40px;height:30px;background:transparent}</style><aside></aside>');
  const measured=await inspect(html);
  assert.deepEqual(outside(measured),[],JSON.stringify(measured.diagnostics));
});

test('painted ancestor outside the selected card is detected without treating default body background as content',async()=>{
  const painted=await inspect(`<body style="margin:0"><main style="padding:20px;background:red">${card}</main></body>`);
  assert.ok(outside(painted).some(d=>d.severity==='warning'),JSON.stringify(painted.diagnostics));
  const body=await inspect(`<body style="margin:0;background:white">${card}</body>`);
  assert.deepEqual(outside(body),[],JSON.stringify(body.diagnostics));
});

test('visible nondefault body background, border, and gradient are not silently excluded',async()=>{
  for(const style of ['background:red','border:4px solid red','background:linear-gradient(red,blue)']){
    const measured=await inspect(`<body style="margin:0;${style}">${card}</body>`);
    assert.ok(measured.diagnostics.some(d=>d.code==='SLIDE_CONTENT_OUTSIDE'||d.code==='SLIDE_CONTENT_UNOBSERVED'),JSON.stringify(measured.diagnostics));
  }
});

test('nondefault document root background is not silently excluded',async()=>{
  for(const background of ['transparent','white']){
    const measured=await inspect(`<html style="background:${background}"><body style="margin:0">${card}</body></html>`);
    assert.deepEqual(outside(measured),[],JSON.stringify(measured.diagnostics));
  }
  const measured=await inspect(`<html style="background:red"><body style="margin:0">${card}</body></html>`);
  assert.ok(measured.diagnostics.some(d=>d.code==='SLIDE_CONTENT_OUTSIDE'||d.code==='SLIDE_CONTENT_UNOBSERVED'),JSON.stringify(measured.diagnostics));
});

test('outside generated content is checked in the selected viewport',async()=>{
  const slide='<div id="slide" style="width:400px;height:600px;background:white">CARD</div>';
  for(const [media,visibleAtWide] of [['@media(max-width:1000px){aside::before{content:none}}',true],['@media(min-width:1001px){aside::before{content:none}}',false]]){
    const html=`<style>aside::before{content:"OUTSIDE"}${media}</style><body style="margin:0">${slide}<aside></aside></body>`;
    const automatic=await inspect(html,null);
    assert.ok(automatic.diagnostics.some(d=>d.code==='SELECTION_REQUIRED'&&d.severity==='error'),JSON.stringify(automatic.diagnostics));
    for(const [width,expected] of [[1280,visibleAtWide],[800,!visibleAtWide]]){
      const explicit=await measureInput(await readInput({html,slideSelector:'#slide'}),runtime,{viewport:{width,height:960}});
      assert.equal(outside(explicit).some(d=>d.severity==='warning'),expected,JSON.stringify(explicit.diagnostics));
    }
  }
});

test('clipped generated content does not mutate selector conditions or ancestor layout',async()=>{
  const html=`<style>aside:empty{width:20px;height:20px;overflow:hidden} aside{position:relative} aside::before{content:'X';position:absolute;left:100px}</style><main>${card}<aside></aside></main>`;
  let observation;
  const trackedRuntime={...runtime,async render(input){
    const session=await runtime.render(input),evaluate=session.page.evaluate.bind(session.page);
    session.page.evaluate=async(fn,arg)=>{
      if(!String(fn).includes('pseudoScan:'))return evaluate(fn,arg);
      const snapshot=()=>{
        const selected=[document.querySelector('#slide'),document.querySelector('main'),document.querySelector('aside')];
        return selected.map(el=>{const r=el.getBoundingClientRect(),s=getComputedStyle(el);return {x:r.x,y:r.y,width:r.width,height:r.height,overflow:s.overflow,position:s.position,children:el.childNodes.length};});
      };
      const before=await evaluate(snapshot);
      await evaluate(()=>{globalThis.__regionMutationCount=0;globalThis.__regionObserver=new MutationObserver(records=>globalThis.__regionMutationCount+=records.length);globalThis.__regionObserver.observe(document,{subtree:true,childList:true,attributes:true});});
      const result=await evaluate(fn,arg);
      const changes=await evaluate(()=>{const count=globalThis.__regionMutationCount+globalThis.__regionObserver.takeRecords().length;globalThis.__regionObserver.disconnect();return count;});
      observation={before,after:await evaluate(snapshot),changes};
      return result;
    };
    return session;
  }};
  const measured=await measureInput(await readInput({html,slideSelector:'#slide'}),trackedRuntime);
  assert.deepEqual(outside(measured),[],JSON.stringify(measured.diagnostics));
  assert.ok(measured.diagnostics.some(d=>d.code==='SLIDE_CONTENT_UNOBSERVED'));
  assert.deepEqual(observation.before,observation.after);
  assert.equal(observation.changes,0);
});

test('outside shadow and outline cannot be silently ignored',async()=>{
  for(const decoration of ['box-shadow:0 0 0 5px red','outline:5px solid red']){
    const measured=await inspect(documentWith('',`<aside style="width:40px;height:30px;${decoration}"></aside>`));
    assert.ok(measured.diagnostics.some(d=>d.code==='SLIDE_CONTENT_OUTSIDE'||d.code==='SLIDE_CONTENT_UNOBSERVED'),JSON.stringify(measured.diagnostics));
  }
});

test('fully clipped outside images report the observation limit',async()=>{
  const html=documentWith('',`<aside style="position:relative;width:20px;height:20px;overflow:hidden"><img src="${png}" style="position:absolute;left:100px;width:40px;height:40px"></aside>`);
  const measured=await inspect(html);
  assert.deepEqual(outside(measured),[],JSON.stringify(measured.diagnostics));
  assert.ok(measured.diagnostics.some(d=>d.code==='SLIDE_CONTENT_UNOBSERVED'),JSON.stringify(measured.diagnostics));
});

test('CLI explicit selection generates a PPT and warns without exposing source content',async t=>{
  const dir=await fs.mkdtemp(path.join(os.tmpdir(),'ppt-region-selector-'));
  const input=path.join(dir,'input.html'),output=path.join(dir,'output.pptx');
  t.after(async()=>{await fs.unlink(input).catch(()=>{});await fs.unlink(output).catch(()=>{});await fs.rmdir(dir);});
  await fs.writeFile(input,documentWith('<h1>PRIVATE_HEADING</h1>','<p>PRIVATE_FOOTER</p>'));
  const result=spawnSync(process.execPath,[path.resolve('src/html-to-pptx/cli.mjs'),'convert',input,'--output',output,'--selector','#slide','--json'],{encoding:'utf8',timeout:120000});
  assert.equal(result.status,0,result.stdout||result.stderr);
  const report=JSON.parse(result.stdout);
  assert.equal(report.status,'generated');
  assert.ok(report.diagnostics.some(d=>d.code==='SLIDE_CONTENT_OUTSIDE'&&d.severity==='warning'));
  assert.doesNotMatch(result.stdout,/PRIVATE_HEADING|PRIVATE_FOOTER|#slide/);
  await fs.access(output);
});

test('unmarked HTML never becomes a slide without an explicit choice even for a single fixed-size card',async()=>{
  for(const html of [documentWith(),documentWith('<h1>REPORT</h1>'),documentWith('',`<img src="${png}" width="40" height="40">`)]){
    const measured=await inspect(html,null);
    assert.ok(measured.diagnostics.some(d=>d.code==='SELECTION_REQUIRED'&&d.severity==='error'),JSON.stringify(measured.diagnostics));
    assert.equal(measured.model,null);
    assert.equal((await exportPptx(measured)).status,'failed');
  }
});
