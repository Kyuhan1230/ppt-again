import test from 'node:test';import assert from 'node:assert/strict';import PptxGenJS from 'pptxgenjs';
import {unzipSync,zipSync,strFromU8,strToU8} from 'fflate';import {chromium} from 'playwright-core';
import {readPptx} from '../src/html-to-pptx/read-pptx.mjs';import {exportHtml} from '../src/html-to-pptx/export-html.mjs';
import {readInput} from '../src/html-to-pptx/input.mjs';import {measureInput} from '../src/html-to-pptx/measure.mjs';import {createBrowserRuntime} from '../src/html-to-pptx/browser-runtime.mjs';import {exportPptx} from '../src/html-to-pptx/export.mjs';
test('character spacing stays with its text run in HTML and on PPT re-export',async()=>{
 const ppt=new PptxGenJS(),slide=ppt.addSlide();slide.background={color:'FFFFFF'};
 slide.addText([{text:'AB'},{text:'CD'}],{x:1,y:1,w:4,h:1,fontFace:'Arial',fontSize:20,objectName:'spacing'});
 const parts=unzipSync(new Uint8Array(await ppt.write({outputType:'uint8array'}))),key='ppt/slides/slide1.xml';let xml=strFromU8(parts[key]);
 xml=xml.replace('<a:rPr ','<a:rPr spc="150" ');parts[key]=strToU8(xml);
 const parsed=await readPptx(zipSync(parts));assert.equal(parsed.status,'parsed',JSON.stringify(parsed.diagnostics));
 const runs=parsed.model.slides[0].nodes.find(n=>n.id==='spacing').content.runs;assert.equal(runs.find(r=>r.text==='AB').style.letterSpacing,2);assert.equal(runs.find(r=>r.text==='CD').style.letterSpacing,0);
 const html=exportHtml(parsed.model).html,browser=await chromium.launch({channel:'msedge',headless:true});try{
  const page=await browser.newPage();await page.setContent(html);
  const spacings=await page.locator('[data-source-id="spacing"] span span').evaluateAll(els=>els.map(el=>[el.textContent,getComputedStyle(el).letterSpacing]));
  assert.deepEqual(spacings[0],['AB','2px']);assert.equal(spacings[1][0],'CD');assert.ok(['normal','0px'].includes(spacings[1][1]));
 }finally{await browser.close();}
 const runtime=await createBrowserRuntime({capture:false});try{
  const measured=await measureInput(await readInput({html}),runtime),result=await exportPptx(measured);assert.equal(result.status,'generated',JSON.stringify(result.diagnostics));
  assert.match(strFromU8(unzipSync(result.pptxBytes)['ppt/slides/slide1.xml']),/spc="150"/);
 }finally{await runtime.close();}
});
