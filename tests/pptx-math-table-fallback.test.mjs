import test from 'node:test';
import assert from 'node:assert/strict';
import PptxGenJS from 'pptxgenjs';
import {unzipSync,zipSync,strFromU8,strToU8} from 'fflate';
import {readPptx} from '../src/html-to-pptx/read-pptx.mjs';
import {exportHtml} from '../src/html-to-pptx/export-html.mjs';
import {publicDiagnostic} from '../src/html-to-pptx/public-diagnostics.mjs';
import {chromium} from 'playwright-core';

const png=Buffer.from('iVBORw0KGgoAAAANSUhEUgAAABQAAAAUCAIAAAAC64paAAAAJUlEQVR4nGP8z0A+YKJAL8Oo5pERYIwMDP9Hmp+ZRjWTBigKMAB8+QIlC4gOEgAAAABJRU5ErkJggg==','base64');
async function fixture(){
 const ppt=new PptxGenJS(),slide=ppt.addSlide();slide.background={color:'FFFFFF'};
 slide.addImage({data:'data:image/png;base64,'+png.toString('base64'),x:0,y:0,w:1,h:1});
 slide.addTable([[''],['']],{x:1,y:1,w:2,h:2,fontFace:'Arial',fontSize:16,objectName:'math-table'});
 const parts=unzipSync(new Uint8Array(await ppt.write({outputType:'uint8array'}))),key='ppt/slides/slide1.xml';
 let xml=strFromU8(parts[key]);const picture=xml.match(/<p:pic>[\s\S]*?<\/p:pic>/)?.[0],frame=xml.match(/<p:graphicFrame>[\s\S]*?<\/p:graphicFrame>/)?.[0];assert.ok(picture&&frame);
 let row=0;const fallback=frame.replace(/<a:tcPr\b[^>]*>/g,opening=>`${opening}<a:blipFill><a:blip r:embed="rId1"/><a:stretch><a:fillRect ${row++?'b':'t'}="-100000"/></a:stretch></a:blipFill>`);assert.equal(row,2);
 const alt=`<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><mc:Choice Requires="a14"><p:sp><p:txBody><a:p><a14:m><m:oMath><m:r><m:t>x</m:t></m:r></m:oMath></a14:m></a:p></p:txBody></p:sp></mc:Choice><mc:Fallback>${fallback}</mc:Fallback></mc:AlternateContent>`;
 parts[key]=strToU8(xml.replace(picture,'').replace(frame,alt));return parts;
}
test('partial import clips each math image table row to the requested region',async()=>{
 const bytes=zipSync(await fixture());assert.equal((await readPptx(bytes)).status,'failed');
 const result=await readPptx(bytes,{allowPartial:true});assert.equal(result.status,'parsed',JSON.stringify(result.diagnostics));
 const images=result.model.slides[0].nodes.filter(n=>n.kind==='image');assert.equal(images.length,2);
 assert.equal(images[0].content.assetId,images[1].content.assetId);
 assert.deepEqual(images[0].content.crop,{left:0,top:50000,right:0,bottom:0});
 assert.deepEqual(images[1].content.crop,{left:0,top:0,right:0,bottom:50000});
 assert.ok(Math.abs(images[0].boxCssPx.x-96)<.01&&Math.abs(images[0].boxCssPx.y-96)<.01);
 const warning=result.diagnostics.find(d=>d.code==='PPTX_MATH_IMAGE_FALLBACK');assert.equal(warning?.severity,'warning');assert.match(publicDiagnostic(warning).message,/수식.*편집/);
 assert.equal((exportHtml(result.model).html.match(/data:image\/png;base64,/g)??[]).length,2);
});
test('unrendered visible fallback table borders are reported',async()=>{
 const parts=await fixture(),key='ppt/slides/slide1.xml',xml=strFromU8(parts[key]);
 parts[key]=strToU8(xml.replace(/<a:tcPr\b[^>]*>/,opening=>`${opening}<a:lnT w="12700"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:lnT>`));
 const result=await readPptx(zipSync(parts),{allowPartial:true});assert.equal(result.status,'parsed',JSON.stringify(result.diagnostics));
 const warning=result.diagnostics.find(d=>d.code==='PPTX_MATH_TABLE_BORDER_OMITTED');assert.equal(warning?.severity,'warning');assert.match(publicDiagnostic(warning).message,/테두리/);
});
test('rendered row crops show their respective PNG regions',async()=>{
 const result=await readPptx(zipSync(await fixture()),{allowPartial:true});
 const html=exportHtml(result.model).html,browser=await chromium.launch({channel:'msedge',headless:true});
 try{
  const page=await browser.newPage({viewport:{width:1200,height:800}});await page.setContent(html);
  const rows=page.locator('img[data-kind="image"]');assert.equal(await rows.count(),2);
  async function color(index){
   const screenshot=await rows.nth(index).screenshot();
   return page.evaluate(async base64=>{
    const image=new Image();image.src='data:image/png;base64,'+base64;await image.decode();
    const canvas=document.createElement('canvas');canvas.width=image.width;canvas.height=image.height;
    const context=canvas.getContext('2d');context.drawImage(image,0,0);
    return [...context.getImageData(Math.floor(image.width/2),Math.floor(image.height/2),1,1).data];
   },screenshot.toString('base64'));
  }
  assert.deepEqual(await color(0),[0,0,255,255]);
  assert.deepEqual(await color(1),[255,0,0,255]);
 }finally{await browser.close();}
});
test('malformed table image relationships, bytes and fill rectangles are not accepted',async()=>{
 for(const mode of ['missing','bad-bytes','external','invalid-rect','empty-rect','oversized-rect']){
  const parts=await fixture(),slide='ppt/slides/slide1.xml',rels='ppt/slides/_rels/slide1.xml.rels',media=Object.keys(parts).find(k=>/^ppt\/media\/.*\.png$/.test(k));
  if(mode==='missing')parts[slide]=strToU8(strFromU8(parts[slide]).replaceAll('r:embed="rId1"','r:embed="missing"'));
  if(mode==='bad-bytes')parts[media]=strToU8('<svg onload="alert(1)"></svg>');
  if(mode==='external')parts[rels]=strToU8(strFromU8(parts[rels]).replace(/Target="\.\.\/media\/[^\"]+"/,'Target="https://example.com/x.png" TargetMode="External"'));
  if(mode==='invalid-rect')parts[slide]=strToU8(strFromU8(parts[slide]).replace('t="-100000"','t="bad"'));
  if(mode==='empty-rect')parts[slide]=strToU8(strFromU8(parts[slide]).replace('t="-100000"','t=""'));
  if(mode==='oversized-rect')parts[slide]=strToU8(strFromU8(parts[slide]).replace('t="-100000"','t="-1000001"'));
  const result=await readPptx(zipSync(parts),{allowPartial:true});assert.equal(result.status,'failed',mode);assert.ok(result.diagnostics.some(d=>d.blocking),mode);
 }
});
test('non-math and unsupported table fallbacks retain an explicit placeholder',async()=>{
 for(const mode of ['non-math','two-columns']){
  const parts=await fixture(),key='ppt/slides/slide1.xml';let xml=strFromU8(parts[key]);
  if(mode==='non-math')xml=xml.replace('<a14:m><m:oMath><m:r><m:t>x</m:t></m:r></m:oMath></a14:m>','');
  if(mode==='two-columns')xml=xml.replace(/(<a:gridCol\b[^>]*\/>)/,'$1$1');
  parts[key]=strToU8(xml);const result=await readPptx(zipSync(parts),{allowPartial:true});
  assert.equal(result.status,'parsed',mode);assert.ok(result.diagnostics.some(d=>d.code==='PPTX_OBJECT_PLACEHOLDER'),mode);
  assert.equal(result.model.slides[0].nodes.filter(n=>n.kind==='image').length,0,mode);
 }
});
