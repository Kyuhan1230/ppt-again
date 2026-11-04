import test from 'node:test';
import assert from 'node:assert/strict';
import {DOMParser} from '@xmldom/xmldom';
import {unzipSync,strFromU8} from 'fflate';
import {readInput} from '../src/html-to-pptx/input.mjs';
import {measureInput} from '../src/html-to-pptx/measure.mjs';
import {createBrowserRuntime} from '../src/html-to-pptx/browser-runtime.mjs';
import {exportPptx} from '../src/html-to-pptx/export.mjs';

test('nested groups and painted text stay together when re-exported to PPTX',async()=>{
 const html=`<!doctype html><html data-slidehandoff="2"><body><section data-slide data-slide-id="slide-1" data-ppt-width-emu="9144000" data-ppt-height-emu="5143500" style="position:relative;width:960px;height:540px;background:white">
 <div data-source-id="outer" data-kind="group" style="position:absolute;left:50px;top:50px;width:500px;height:250px"></div>
 <div data-source-id="inner" data-kind="group" data-parent-id="outer" style="position:absolute;left:60px;top:60px;width:300px;height:180px"></div>
 <div data-source-id="label" data-kind="text" data-parent-id="inner" style="position:absolute;left:80px;top:80px;width:200px;height:70px;background:#ddd;font:24px Arial">Test label</div>
 <div data-source-id="sibling" data-kind="shape" data-parent-id="outer" style="position:absolute;left:380px;top:80px;width:100px;height:70px;background:#5599aa"></div>
 </section></body></html>`;
 const runtime=await createBrowserRuntime({capture:false});try{
 const measured=await measureInput(await readInput({html}),runtime);
 assert.equal(measured.diagnostics.filter(d=>d.severity==='error').length,0,JSON.stringify(measured.diagnostics));
 const out=await exportPptx(measured);assert.equal(out.status,'generated',JSON.stringify(out.diagnostics));
 const doc=new DOMParser().parseFromString(strFromU8(unzipSync(out.pptxBytes)['ppt/slides/slide1.xml']),'application/xml');
 const groups=[...doc.getElementsByTagNameNS('*','grpSp')];assert.equal(groups.length,2);
 const outer=groups.find(g=>g.getElementsByTagNameNS('*','cNvPr')[0].getAttribute('name')==='outer');
 const inner=groups.find(g=>g.getElementsByTagNameNS('*','cNvPr')[0].getAttribute('name')==='inner');
 assert.equal(inner.parentNode,outer);assert.ok(inner.textContent.includes('Test label'));
 const innerNames=[...inner.getElementsByTagNameNS('*','cNvPr')].map(n=>n.getAttribute('name'));assert.ok(innerNames.includes('label::background'));assert.ok(innerNames.includes('label'));
 }finally{await runtime.close();}
});


test('a bold first run does not make subsequent normal runs bold',async()=>{
 const html=`<section data-slide style="position:relative;width:960px;height:540px;background:white"><p id="mixed" style="position:absolute;left:20px;top:20px;width:500px;height:100px;font:700 24px Arial"><b>Heading</b><br><span style="font-weight:400">Normal text</span></p></section>`;
 const runtime=await createBrowserRuntime({capture:false});try{
  const measured=await measureInput(await readInput({html}),runtime);const out=await exportPptx(measured);assert.equal(out.status,'generated',JSON.stringify(out.diagnostics));
  const doc=new DOMParser().parseFromString(strFromU8(unzipSync(out.pptxBytes)['ppt/slides/slide1.xml']),'application/xml');
  const runs=[...doc.getElementsByTagNameNS('*','r')],run=runs.find(r=>r.textContent.includes('Normal text'));assert.ok(run);
  assert.notEqual(run.getElementsByTagNameNS('*','rPr')[0].getAttribute('b'),'1');
 }finally{await runtime.close();}
});


test('text export keeps distinct padding and explicit no-wrap across multiple lines',async()=>{
 const html=`<section data-slide style="position:relative;width:960px;height:540px;background:white"><p id="nowrap" style="position:absolute;left:20px;top:20px;width:500px;height:100px;box-sizing:border-box;padding:4px 8px 12px 16px;white-space:pre;font:24px Arial">First line<br>Second line</p></section>`;
 const runtime=await createBrowserRuntime({capture:false});try{
  const measured=await measureInput(await readInput({html}),runtime),out=await exportPptx(measured);assert.equal(out.status,'generated',JSON.stringify(out.diagnostics));
  const doc=new DOMParser().parseFromString(strFromU8(unzipSync(out.pptxBytes)['ppt/slides/slide1.xml']),'application/xml');const body=[...doc.getElementsByTagNameNS('*','sp')].find(s=>s.textContent.includes('First line')).getElementsByTagNameNS('*','bodyPr')[0];
  assert.equal(body.getAttribute('wrap'),'none');for(const [k,v] of Object.entries({lIns:203200,rIns:101600,bIns:152400,tIns:50800}))assert.equal(Number(body.getAttribute(k)),v,k);
 }finally{await runtime.close();}
});
