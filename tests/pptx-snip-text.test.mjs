import test from 'node:test';
import assert from 'node:assert/strict';
import PptxGenJS from 'pptxgenjs';
import {unzipSync,zipSync,strFromU8,strToU8} from 'fflate';
import {chromium} from 'playwright-core';
import {readPptx} from '../src/html-to-pptx/read-pptx.mjs';
import {exportHtml} from '../src/html-to-pptx/export-html.mjs';

async function fixture(fontReference=false){
 const ppt=new PptxGenJS(),slide=ppt.addSlide();slide.background={color:'FFFFFF'};
 slide.addText('Furnace',{x:1,y:1,w:2,h:1,fontFace:'Arial',fontSize:18,color:'FFFFFF',fill:{color:'A6A6A6'},line:{transparency:100},margin:0,objectName:'snipped'});
 const parts=unzipSync(new Uint8Array(await ppt.write({outputType:'uint8array'}))),file='ppt/slides/slide1.xml';
 let xml=strFromU8(parts[file]);
 xml=xml.replace('<p:spPr><a:xfrm>','<p:spPr><a:xfrm flipH="1">').replace('<a:prstGeom prst="rect"><a:avLst></a:avLst>','<a:prstGeom prst="snip1Rect"><a:avLst><a:gd name="adj" fmla="val 50000"/></a:avLst>');
 if(fontReference)xml=xml.replace('<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:latin','<a:latin').replace('</p:spPr><p:txBody>','</p:spPr><p:style><a:fontRef idx="minor"><a:srgbClr val="FFFFFF"/></a:fontRef></p:style><p:txBody>');
 assert.match(xml,/flipH="1"/);assert.match(xml,/prst="snip1Rect"/);parts[file]=strToU8(xml);return zipSync(parts);
}
test('a horizontally flipped snipped text shape keeps its silhouette and editable text',async()=>{
 const result=await readPptx(await fixture(),{allowPartial:true});assert.equal(result.status,'parsed',JSON.stringify(result.diagnostics));
 const slide=result.model.slides[0],text=slide.nodes.find(n=>n.id==='snipped'),background=slide.nodes.find(n=>n.id==='snipped::background');
 assert.equal(text.kind,'text');assert.equal(text.content.text,'Furnace');assert.equal(background.kind,'image');
 assert.ok(Math.abs(text.style.padding.left-24)<.01);assert.ok(Math.abs(text.style.padding.top-24)<.01);
 assert.ok(background.zOrder<text.zOrder);assert.ok(!slide.nodes.some(n=>n.content.placeholderDescription));
 const html=exportHtml(result.model);assert.equal(html.status,'generated',JSON.stringify(html.diagnostics));
 const browser=await chromium.launch({channel:'msedge',headless:true});try{
  const page=await browser.newPage();await page.setContent(html.html);
  const samples=await page.locator('[data-source-id="snipped::background"]').evaluate(async img=>{await img.decode();const canvas=document.createElement('canvas');canvas.width=img.naturalWidth;canvas.height=img.naturalHeight;const c=canvas.getContext('2d');c.drawImage(img,0,0);return [[1,1],[canvas.width-2,1],[canvas.width/2,canvas.height/2]].map(([x,y])=>[...c.getImageData(x,y,1,1).data]);});
  assert.equal(samples[0][3],0);assert.equal(samples[1][3],255);assert.deepEqual(samples[2].slice(0,3),[166,166,166]);
  assert.equal(await page.locator('[data-source-id="snipped"]').innerText(),'Furnace');
 }finally{await browser.close();}
});
test('a shape font reference supplies its text color when the run does not specify one',async()=>{
 const result=await readPptx(await fixture(true),{allowPartial:true});assert.equal(result.status,'parsed',JSON.stringify(result.diagnostics));
 const node=result.model.slides[0].nodes.find(n=>n.id==='snipped');
 assert.equal(node.content.runs[0].style.color,'#ffffff');
});
