import test from 'node:test';
import assert from 'node:assert/strict';
import PptxGenJS from 'pptxgenjs';
import {unzipSync,strFromU8} from 'fflate';
import {readPptx} from '../src/html-to-pptx/read-pptx.mjs';
import {exportHtml} from '../src/html-to-pptx/export-html.mjs';
import {chromium} from 'playwright-core';

test('uniform percentage line spacing is rendered in the editable HTML',async()=>{
 const ppt=new PptxGenJS(),slide=ppt.addSlide();slide.background={color:'FFFFFF'};
 slide.addText('First\nSecond',{x:1,y:1,w:4,h:2,fontFace:'Arial',fontSize:20,lineSpacingMultiple:2,margin:0,objectName:'spaced'});
 const bytes=new Uint8Array(await ppt.write({outputType:'uint8array'}));
 assert.match(strFromU8(unzipSync(bytes)['ppt/slides/slide1.xml']),/<a:lnSpc><a:spcPct val="200000"\/><\/a:lnSpc>/);
 const parsed=await readPptx(bytes,{allowPartial:true});assert.equal(parsed.status,'parsed',JSON.stringify(parsed.diagnostics));
 const node=parsed.model.slides[0].nodes.find(n=>n.id==='spaced');assert.ok(node);assert.equal(node.style.lineHeight,'2.4');
 const html=exportHtml(parsed.model).html;assert.match(html,/line-height:2\.4;/);
 const browser=await chromium.launch({channel:'msedge',headless:true});try{
  const page=await browser.newPage();await page.setContent(html);const lineHeight=await page.locator('[data-source-id="spaced"]').evaluate(el=>parseFloat(getComputedStyle(el).lineHeight));
  assert.ok(Math.abs(lineHeight-2.4*node.style.fontSize)<0.05,`${lineHeight} versus ${node.style.fontSize}`);
 }finally{await browser.close();}
});

test('implicit single line spacing uses the PowerPoint line advance',async()=>{
 const ppt=new PptxGenJS(),slide=ppt.addSlide();slide.background={color:'FFFFFF'};
 slide.addText('Single line',{x:1,y:1,w:4,h:1,fontFace:'Arial',fontSize:20,margin:0,objectName:'single'});
 const parsed=await readPptx(new Uint8Array(await ppt.write({outputType:'uint8array'})),{allowPartial:true});
 const node=parsed.model.slides[0].nodes.find(n=>n.id==='single');assert.equal(node.style.lineHeight,'1.2');
 const html=exportHtml(parsed.model).html;assert.match(html,/line-height:1\.2;/);
});
