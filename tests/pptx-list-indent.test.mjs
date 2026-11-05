import test from 'node:test';import assert from 'node:assert/strict';import PptxGenJS from 'pptxgenjs';
import {unzipSync,zipSync,strFromU8,strToU8} from 'fflate';import {chromium} from 'playwright-core';
import {readPptx} from '../src/html-to-pptx/read-pptx.mjs';import {exportHtml} from '../src/html-to-pptx/export-html.mjs';

test('numbered bullets end at the PowerPoint hanging-indent text position',async()=>{
 const ppt=new PptxGenJS(),slide=ppt.addSlide();slide.background={color:'FFFFFF'};
 slide.addText('Alpha\nBeta',{x:1,y:1,w:4,h:2,fontFace:'Arial',fontSize:18,margin:0,objectName:'list'});
 const parts=unzipSync(new Uint8Array(await ppt.write({outputType:'uint8array'}))),key='ppt/slides/slide1.xml';let xml=strFromU8(parts[key]);
 xml=xml.replace(/<a:pPr([^>]*)><a:buNone\/>/g,(_,attrs)=>`<a:pPr${attrs.replace(/\s+(?:marL|indent)="[^"]*"/g,'')} marL="342900" indent="-342900"><a:spcBef><a:spcPts val="1200"/></a:spcBef><a:buAutoNum type="arabicPeriod"/>`);parts[key]=strToU8(xml);
 const result=await readPptx(zipSync(parts),{allowPartial:true});assert.equal(result.status,'parsed',JSON.stringify(result.diagnostics));
 const html=exportHtml(result.model),browser=await chromium.launch({channel:'msedge',headless:true});try{
  const page=await browser.newPage();await page.setContent(html.html);
  const data=await page.locator('[data-source-id="list"] [data-pptx-paragraph]').first().evaluate(p=>{const marker=p.querySelector('[data-pptx-list-marker]'),text=marker?.nextElementSibling;return {value:p.textContent.trim(),markerWidth:marker?.getBoundingClientRect().width,start:text&&text.getBoundingClientRect().left-p.getBoundingClientRect().left};});
  assert.equal(data.value,'1. Alpha');assert.ok(Math.abs(data.markerWidth-36)<.1,JSON.stringify(data));assert.ok(Math.abs(data.start-36)<.1,JSON.stringify(data));
 }finally{await browser.close();}
});
