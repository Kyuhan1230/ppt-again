import test from 'node:test';
import assert from 'node:assert/strict';
import PptxGenJS from 'pptxgenjs';
import {unzipSync,zipSync,strFromU8,strToU8} from 'fflate';
import {chromium} from 'playwright-core';
import {readPptx} from '../src/html-to-pptx/read-pptx.mjs';
import {exportHtml} from '../src/html-to-pptx/export-html.mjs';
import {readInput} from '../src/html-to-pptx/input.mjs';
import {measureInput} from '../src/html-to-pptx/measure.mjs';
import {createBrowserRuntime} from '../src/html-to-pptx/browser-runtime.mjs';
import {exportPptx} from '../src/html-to-pptx/export.mjs';

test('superscript and subscript use the native reduced glyph size and retain their baseline on reimport',async()=>{
 const ppt=new PptxGenJS(),slide=ppt.addSlide();slide.background={color:'FFFFFF'};
 slide.addText([{text:'Ordinary '},{text:'up',options:{superscript:true}},{text:' down',options:{subscript:true}}],{x:1,y:1,w:5,h:1,fontFace:'Arial',fontSize:18,objectName:'baseline'});
 const parsed=await readPptx(new Uint8Array(await ppt.write({outputType:'uint8array'})),{allowPartial:true});assert.equal(parsed.status,'parsed',JSON.stringify(parsed.diagnostics));
 const node=parsed.model.slides[0].nodes.find(n=>n.id==='baseline');
 assert.deepEqual(node.content.runs.filter(r=>r.style.baseline).map(r=>r.style.baseline),[30000,-40000]);
 const html=exportHtml(parsed.model);assert.equal(html.status,'generated',JSON.stringify(html.diagnostics));
 const browser=await chromium.launch({channel:'msedge',headless:true});try{
  const page=await browser.newPage();await page.setContent(html.html);
  const spans=await page.locator('[data-source-id="baseline"] [data-pptx-baseline]').evaluateAll(els=>els.map(e=>({text:e.textContent,size:parseFloat(getComputedStyle(e).fontSize),shift:parseFloat(getComputedStyle(e).verticalAlign)})));
  assert.deepEqual(spans.map(s=>s.text),['up',' down']);
  assert.ok(spans.every(s=>Math.abs(s.size-16)<.01));assert.ok(Math.abs(spans[0].shift-7.2)<.01);assert.ok(Math.abs(spans[1].shift+9.6)<.01);
 }finally{await browser.close();}
 const runtime=await createBrowserRuntime({capture:false});try{
  const measured=await measureInput(await readInput({html:html.html}),runtime),result=await exportPptx(measured);
  assert.equal(result.status,'generated',JSON.stringify(result.diagnostics));
  const xml=strFromU8(unzipSync(result.pptxBytes)['ppt/slides/slide1.xml']);
  assert.match(xml,/sz="1800"[^>]*baseline="30000"|baseline="30000"[^>]*sz="1800"/);
  assert.match(xml,/sz="1800"[^>]*baseline="-40000"|baseline="-40000"[^>]*sz="1800"/);
 }finally{await runtime.close();}
});
test('a paragraph consisting only of superscript uses the reduced line height',async()=>{
 const ppt=new PptxGenJS(),slide=ppt.addSlide();slide.background={color:'FFFFFF'};
 slide.addText('one\ntwo',{x:1,y:1,w:4,h:2,fontFace:'Arial',fontSize:18,objectName:'notes'});
 const parts=unzipSync(new Uint8Array(await ppt.write({outputType:'uint8array'}))),key='ppt/slides/slide1.xml';
 let xml=strFromU8(parts[key]);xml=xml.replaceAll('<a:rPr ','<a:rPr baseline="30000" ').replace(/<a:pPr([^>]*)>/g,'<a:pPr$1><a:lnSpc><a:spcPct val="100000"/></a:lnSpc>');parts[key]=strToU8(xml);
 const parsed=await readPptx(zipSync(parts),{allowPartial:true});assert.equal(parsed.status,'parsed',JSON.stringify(parsed.diagnostics));
 const html=exportHtml(parsed.model),browser=await chromium.launch({channel:'msedge',headless:true});try{
  const page=await browser.newPage();await page.setContent(html.html);
  const heights=await page.locator('[data-source-id="notes"] [data-pptx-paragraph]').evaluateAll(p=>p.map(e=>e.getBoundingClientRect().height));
  assert.equal(heights.length,2);assert.ok(heights.every(h=>Math.abs(h-19.2)<.1),JSON.stringify(heights));
 }finally{await browser.close();}
});
