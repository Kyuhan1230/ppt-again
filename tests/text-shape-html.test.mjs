import test from 'node:test';
import assert from 'node:assert/strict';
import PptxGenJS from 'pptxgenjs';
import {readPptx} from '../src/html-to-pptx/read-pptx.mjs';
import {exportHtml} from '../src/html-to-pptx/export-html.mjs';
import {chromium} from 'playwright-core';
import {unzipSync,zipSync,strFromU8,strToU8} from 'fflate';

test('HTML preserves a text node’s fill, border, and rounded corners', async()=>{
  const ppt=new PptxGenJS(),slide=ppt.addSlide();slide.background={color:'FFFFFF'};
  slide.addText('Shape text',{x:1,y:1,w:3,h:1,fontFace:'Arial',fontSize:18});
  const parsed=await readPptx(new Uint8Array(await ppt.write({outputType:'uint8array'})));
  assert.equal(parsed.status,'parsed',JSON.stringify(parsed.diagnostics));
  const node=parsed.model.slides[0].nodes.find(n=>n.kind==='text');
  Object.assign(node.style,{fill:'#123456',stroke:'#ABCDEF',strokeWidth:2,radius:8});
  const result=exportHtml(parsed.model);
  assert.equal(result.status,'generated',JSON.stringify(result.diagnostics));
  assert.match(result.html,/background:#123456/);
  assert.match(result.html,/outline:2px solid #ABCDEF/);
  assert.match(result.html,/outline-offset:-1px/);
  assert.match(result.html,/border-radius:8px/);
  assert.match(result.html,/Shape text/);
});

test('underline on the first run does not leak to later text',async()=>{
  const ppt=new PptxGenJS(),slide=ppt.addSlide();slide.background={color:'FFFFFF'};
  slide.addText([{text:'Heading',options:{underline:true}},{text:' plain',options:{underline:false}}],{x:1,y:1,w:3,h:1,fontFace:'Arial',fontSize:18,objectName:'mixed-decoration'});
  const parsed=await readPptx(new Uint8Array(await ppt.write({outputType:'uint8array'})),{allowPartial:true});
  assert.equal(parsed.status,'parsed',JSON.stringify(parsed.diagnostics));
  const result=exportHtml(parsed.model);assert.equal(result.status,'generated',JSON.stringify(result.diagnostics));
  const browser=await chromium.launch({channel:'msedge',headless:true});try{
    const page=await browser.newPage();await page.setContent(result.html);
    const decorations=await page.locator('[data-source-id="mixed-decoration"]').evaluate(el=>({outer:getComputedStyle(el).textDecorationLine,children:[...(el.querySelector('[data-pptx-text-content]')??el).children].map(c=>getComputedStyle(c).textDecorationLine)}));
    assert.equal(decorations.outer,'none');assert.deepEqual(decorations.children,['underline','none']);
  }finally{await browser.close();}
});

test('vertically centered overflowing text remains centered in the original box',async()=>{
  const ppt=new PptxGenJS(),slide=ppt.addSlide();slide.background={color:'FFFFFF'};
  slide.addText('First\nSecond\nThird',{x:1,y:1,w:4,h:1,fontFace:'Arial',fontSize:15,margin:0,valign:'mid',objectName:'overflowing'});
  const parsed=await readPptx(new Uint8Array(await ppt.write({outputType:'uint8array'})),{allowPartial:true});
  const node=parsed.model.slides[0].nodes.find(n=>n.id==='overflowing');node.boxCssPx.h=40;node.style.lineHeight='24px';
  const result=exportHtml(parsed.model),browser=await chromium.launch({channel:'msedge',headless:true});try{
    const page=await browser.newPage();await page.setContent(result.html);
    const actual=await page.locator('[data-source-id="overflowing"]').evaluate(el=>{const a=el.getBoundingClientRect(),inner=el.querySelector('[data-pptx-text-content]'),b=inner?.getBoundingClientRect();return {height:b?.height,offset:b?b.top+b.height/2-a.top-a.height/2:null};});
    assert.ok(Math.abs(actual.height-72)<.1,JSON.stringify(actual));assert.ok(Math.abs(actual.offset)<.1,JSON.stringify(actual));
  }finally{await browser.close();}
});

test('auto-fit text may use a small part of the right inset before wrapping',async()=>{
  const ppt=new PptxGenJS(),slide=ppt.addSlide();slide.background={color:'FFFFFF'};
  slide.addText('Near the edge',{x:1,y:1,w:3,h:1,fontFace:'Arial',fontSize:18,objectName:'autofit'});
  const parts=unzipSync(new Uint8Array(await ppt.write({outputType:'uint8array'}))),file='ppt/slides/slide1.xml';
  parts[file]=strToU8(strFromU8(parts[file]).replace('</a:bodyPr>','<a:spAutoFit/></a:bodyPr>'));
  const parsed=await readPptx(zipSync(parts),{allowPartial:true});assert.equal(parsed.status,'parsed',JSON.stringify(parsed.diagnostics));
  const html=exportHtml(parsed.model),browser=await chromium.launch({channel:'msedge',headless:true});try{
    const page=await browser.newPage();await page.setContent(html.html);
    const actual=await page.locator('[data-source-id="autofit"]').evaluate(el=>{const s=getComputedStyle(el),inner=el.querySelector('[data-pptx-text-content]');return {right:parseFloat(s.paddingRight),extra:inner.getBoundingClientRect().width-(el.clientWidth-parseFloat(s.paddingLeft)-parseFloat(s.paddingRight))};});
    assert.equal(actual.right,9.6);assert.ok(Math.abs(actual.extra-4.8)<.1,JSON.stringify(actual));
  }finally{await browser.close();}
});
