import test from 'node:test';import assert from 'node:assert/strict';import PptxGenJS from 'pptxgenjs';import {chromium} from 'playwright-core';
import {readPptx} from '../src/html-to-pptx/read-pptx.mjs';import {exportHtml} from '../src/html-to-pptx/export-html.mjs';
test('default PowerPoint tab stops line up consecutive paragraphs one inch apart',async()=>{
 const ppt=new PptxGenJS(),slide=ppt.addSlide();slide.background={color:'FFFFFF'};
 slide.addText('ABC\tFirst\n\tSecond',{x:1,y:1,w:5,h:2,fontFace:'Arial',fontSize:16,margin:0,objectName:'tabs'});
 const parsed=await readPptx(new Uint8Array(await ppt.write({outputType:'uint8array'})),{allowPartial:true});assert.equal(parsed.status,'parsed',JSON.stringify(parsed.diagnostics));
 const html=exportHtml(parsed.model).html,browser=await chromium.launch({channel:'msedge',headless:true});try{
  const page=await browser.newPage();await page.setContent(html);
  const positions=await page.locator('[data-source-id="tabs"]').evaluate(el=>{
   const start=el.getBoundingClientRect().left,walker=document.createTreeWalker(el,NodeFilter.SHOW_TEXT),result={};
   for(let n;n=walker.nextNode();)for(const word of ['First','Second']){const i=n.data.indexOf(word);if(i<0)continue;const range=document.createRange();range.setStart(n,i);range.setEnd(n,i+word.length);result[word]=range.getBoundingClientRect().left-start;}
   return result;
  });
  assert.ok(Math.abs(positions.First-96)<.2,JSON.stringify(positions));assert.ok(Math.abs(positions.Second-96)<.2,JSON.stringify(positions));
 }finally{await browser.close();}
});
