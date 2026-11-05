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

test('PowerPoint kerning thresholds affect text and tables and disabled kerning survives export',async()=>{
 const ppt=new PptxGenJS();ppt.layout='LAYOUT_WIDE';
 for(const size of [10,20,10]){
  const slide=ppt.addSlide();slide.background={color:'FFFFFF'};
  slide.addText('AVATAR',{x:1,y:1,w:3,h:1,fontFace:'Arial',fontSize:size,margin:0,objectName:'shape'});
  slide.addTable([['AVATAR']],{x:1,y:2.5,w:3,h:1,fontFace:'Arial',fontSize:size,objectName:'table'});
 }
 const parts=unzipSync(new Uint8Array(await ppt.write({outputType:'uint8array'})));
 for(const [index,kern] of [1200,0,100].entries()){
  const name=`ppt/slides/slide${index+1}.xml`,xml=strFromU8(parts[name]),patched=xml.replaceAll('<a:rPr ',`<a:rPr kern="${kern}" `);
  assert.notEqual(xml,patched);parts[name]=strToU8(patched);
 }
 const parsed=await readPptx(zipSync(parts),{allowPartial:true});assert.equal(parsed.status,'parsed',JSON.stringify(parsed.diagnostics));
 for(const [index,slide] of parsed.model.slides.entries()){
  const shape=slide.nodes.find(node=>node.id==='shape'),cell=slide.nodes.find(node=>node.id==='table').content.cells[0];
  assert.equal(Boolean(shape.content.runs[0].style.disableKerning),index!==2);
  assert.equal(Boolean(cell.runs[0].style.disableKerning),index!==2);
 }
 const html=exportHtml(parsed.model).html,browser=await chromium.launch({channel:'msedge',headless:true});
 try{
  const page=await browser.newPage();await page.setContent(html);
  for(let i=0;i<3;i++)for(const selector of ['[data-source-id="shape"] span','[data-source-id="table"] td span']){
   const style=await page.locator('section[data-slide]').nth(i).locator(selector).last().evaluate(el=>getComputedStyle(el).fontKerning);
   assert.equal(style,i===2?'auto':'none');
  }
 }finally{await browser.close();}
 const runtime=await createBrowserRuntime({capture:false});try{
  const measured=await measureInput(await readInput({html}),runtime),output=await exportPptx(measured);
  assert.equal(output.status,'generated',JSON.stringify(output.diagnostics));
  const result=unzipSync(output.pptxBytes);
  for(const index of [1,2]){
   const xml=strFromU8(result[`ppt/slides/slide${index}.xml`]);
   assert.ok((xml.match(/<a:rPr\b[^>]*\bkern="0"/g)??[]).length>=2,xml);
  }
  const edited=html.replaceAll('font-kerning:none;','font-kerning:normal;');assert.notEqual(edited,html);
  const changed=await measureInput(await readInput({html:edited}),runtime),changedOutput=await exportPptx(changed);
  assert.equal(changedOutput.status,'generated',JSON.stringify(changedOutput.diagnostics));
  assert.ok(!changed.model.slides[0].nodes.find(node=>node.id==='shape').content.runs[0].style.disableKerning);
  assert.doesNotMatch(strFromU8(unzipSync(changedOutput.pptxBytes)['ppt/slides/slide1.xml']),/<a:rPr\b[^>]*\bkern="0"/);
 }finally{await runtime.close();}
});
