import test from 'node:test';
import assert from 'node:assert/strict';
import PptxGenJS from 'pptxgenjs';
import {unzipSync,zipSync,strFromU8,strToU8} from 'fflate';
import {readPptx} from '../src/html-to-pptx/read-pptx.mjs';
import {exportHtml} from '../src/html-to-pptx/export-html.mjs';
import {chromium} from 'playwright-core';
import {readInput} from '../src/html-to-pptx/input.mjs';
import {measureInput} from '../src/html-to-pptx/measure.mjs';
import {createBrowserRuntime} from '../src/html-to-pptx/browser-runtime.mjs';

async function fixture(firstLast=false){
 const ppt=new PptxGenJS(),slide=ppt.addSlide();slide.background={color:'FFFFFF'};
 slide.addText('First\nSecond\nThird',{x:1,y:1,w:5,h:4,fontFace:'Arial',fontSize:20,margin:0,objectName:'paragraphs'});
 const parts=unzipSync(new Uint8Array(await ppt.write({outputType:'uint8array'}))),key='ppt/slides/slide1.xml';
 let xml=strFromU8(parts[key]);let i=0;
 xml=xml.replace(/<a:pPr([^>]*)>/g,(_,attrs)=>`<a:pPr${i===1?attrs.replace(/\s+(?:marL|marR|indent)="[^"]*"/g,''):attrs}${i===1?' marL="720000" marR="360000" indent="-360000"':''}>${[
  '<a:lnSpc><a:spcPct val="200000"/></a:lnSpc><a:spcBef><a:spcPts val="1200"/></a:spcBef>',
  '<a:lnSpc><a:spcPct val="100000"/></a:lnSpc><a:spcBef><a:spcPts val="600"/></a:spcBef><a:spcAft><a:spcPts val="300"/></a:spcAft>',
  '<a:lnSpc><a:spcPct val="150000"/></a:lnSpc><a:spcAft><a:spcPts val="900"/></a:spcAft>'
 ][i++]??''}`);
 assert.equal(i,3);if(firstLast)xml=xml.replace('<a:bodyPr ','<a:bodyPr spcFirstLastPara="1" ');
 parts[key]=strToU8(xml);return zipSync(parts);
}
test('HTML preserves each paragraph line and before/after spacing, excluding edge spacing by default',async()=>{
 const parsed=await readPptx(await fixture(),{allowPartial:true});assert.equal(parsed.status,'parsed',JSON.stringify(parsed.diagnostics));
 const node=parsed.model.slides[0].nodes.find(n=>n.id==='paragraphs');assert.equal(node.content.text,'First\nSecond\nThird');
 const html=exportHtml(parsed.model).html;const browser=await chromium.launch({channel:'msedge',headless:true});try{
  const page=await browser.newPage();await page.setContent(html);
  const result=await page.locator('[data-source-id="paragraphs"]').evaluate(el=>({text:el.textContent,paras:[...el.querySelectorAll('[data-pptx-paragraph]')].map(p=>{const s=getComputedStyle(p);return {before:parseFloat(s.paddingTop),after:parseFloat(s.paddingBottom),left:parseFloat(s.paddingLeft),right:parseFloat(s.paddingRight),indent:parseFloat(s.textIndent),lineHeight:parseFloat(s.lineHeight),fontSize:parseFloat(s.fontSize)};})}));
  assert.equal(result.text,'First\nSecond\nThird');assert.equal(result.paras.length,3);
  assert.equal(result.paras[0].before,0);assert.equal(result.paras[1].before,8);assert.equal(result.paras[1].after,4);assert.equal(result.paras[2].after,0);
  assert.ok(Math.abs(result.paras[1].left-720000/9525)<.01);assert.ok(Math.abs(result.paras[1].right-360000/9525)<.01);assert.ok(Math.abs(result.paras[1].indent+360000/9525)<.01);
  for(const [i,ratio] of [2.4,1.2,1.8].entries())assert.ok(Math.abs(result.paras[i].lineHeight-result.paras[i].fontSize*ratio)<.1);
 }finally{await browser.close();}
});
test('top-aligned expanded line spacing positions glyphs without changing paragraph boxes',async()=>{
 const parsed=await readPptx(await fixture(),{allowPartial:true}),node=parsed.model.slides[0].nodes.find(n=>n.id==='paragraphs');
 node.style.valign='top';
 const html=exportHtml(parsed.model).html,browser=await chromium.launch({channel:'msedge',headless:true});
 try{
  const page=await browser.newPage();await page.setContent(html);
  const info=await page.locator('[data-source-id="paragraphs"]').evaluate(el=>[...el.querySelectorAll('[data-pptx-paragraph]')].map(p=>{
   const span=[...p.children].find(c=>c.textContent.trim()),box=p.getBoundingClientRect();
   const range=document.createRange();range.selectNodeContents(span);const before=range.getBoundingClientRect();
   for(const child of p.children){child.style.position='static';child.style.top='auto';}
   const after=range.getBoundingClientRect(),newBox=p.getBoundingClientRect();
   return {shift:before.top-after.top,originalHeight:box.height,newHeight:newBox.height,originalLeft:before.left,newLeft:after.left};
  }));
  for(const [index,expected] of [8,0,4].entries()){
   assert.ok(Math.abs(info[index].shift-expected)<.1,JSON.stringify(info));
   assert.ok(Math.abs(info[index].originalHeight-info[index].newHeight)<.1,JSON.stringify(info));
   assert.ok(Math.abs(info[index].originalLeft-info[index].newLeft)<.1,JSON.stringify(info));
  }
  for(const align of ['mid','bottom']){
   node.style.valign=align;await page.setContent(exportHtml(parsed.model).html);
   const positioned=await page.locator('[data-pptx-paragraph] > span').evaluateAll(spans=>spans.some(span=>span.style.position==='relative'&&parseFloat(span.style.top)>0));
   assert.equal(positioned,false,align);
  }
 }finally{await browser.close();}
 const runtime=await createBrowserRuntime({capture:false});try{
  const edited=html.replace('>First</span>','>Changed</span>');assert.notEqual(edited,html);
  const measured=await measureInput(await readInput({html:edited}),runtime);
  assert.equal(measured.diagnostics.filter(d=>d.severity==='error').length,0,JSON.stringify(measured.diagnostics));
  assert.equal(measured.model.slides[0].nodes.find(n=>n.id==='paragraphs').content.text,'Changed\nSecond\nThird');
 }finally{await runtime.close();}
});
test('first and last paragraph spacing is used when explicitly enabled',async()=>{
 const parsed=await readPptx(await fixture(true),{allowPartial:true}),node=parsed.model.slides[0].nodes.find(n=>n.id==='paragraphs');
 assert.equal(node.content.paragraphs[0].beforePx,16);assert.equal(node.content.paragraphs[2].afterPx,12);
});
test('zero percentage paragraph spacing does not disable other paragraph layout',async()=>{
 const parts=unzipSync(await fixture()),key='ppt/slides/slide1.xml',before=strFromU8(parts[key]);
 const after=before.replace('<a:spcAft><a:spcPts val="900"/></a:spcAft>','<a:spcAft><a:spcPct val="0"/></a:spcAft>');assert.notEqual(after,before);parts[key]=strToU8(after);
 const result=await readPptx(zipSync(parts),{allowPartial:true});assert.equal(result.status,'parsed',JSON.stringify(result.diagnostics));
 const paragraphs=result.model.slides[0].nodes.find(n=>n.id==='paragraphs').content.paragraphs;
 assert.equal(paragraphs.length,3);assert.equal(paragraphs[2].afterPx,0);assert.equal(paragraphs[1].beforePx,8);
});
test('empty paragraphs retain their line and spacing without adding text on HTML reimport',async()=>{
 const parts=unzipSync(await fixture()),key='ppt/slides/slide1.xml';
 parts[key]=strToU8(strFromU8(parts[key]).replace('>Second</a:t>','></a:t>'));
 const parsed=await readPptx(zipSync(parts),{allowPartial:true});assert.equal(parsed.status,'parsed',JSON.stringify(parsed.diagnostics));
 const html=exportHtml(parsed.model).html,browser=await chromium.launch({channel:'msedge',headless:true});try{
  const page=await browser.newPage();await page.setContent(html);
  const actual=await page.locator('[data-source-id="paragraphs"]').evaluate(el=>({text:el.textContent,height:el.querySelectorAll('[data-pptx-paragraph]')[1]?.getBoundingClientRect().height}));
  assert.equal(actual.text,'First\n\nThird');assert.ok(Math.abs(actual.height-44)<.1,`empty paragraph height was ${actual.height}`);
 }finally{await browser.close();}
 const runtime=await createBrowserRuntime({capture:false});try{
  const measured=await measureInput(await readInput({html}),runtime);
  assert.equal(measured.diagnostics.filter(d=>d.severity==='error').length,0,JSON.stringify(measured.diagnostics));
  assert.equal(measured.model.slides[0].nodes.find(n=>n.id==='paragraphs').content.text,'First\n\nThird');
  assert.ok(measured.diagnostics.some(d=>d.code==='PARAGRAPH_SPACING_APPROXIMATED'));
 }finally{await runtime.close();}
});
