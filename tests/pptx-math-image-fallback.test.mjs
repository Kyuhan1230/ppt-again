import test from 'node:test';
import assert from 'node:assert/strict';
import PptxGenJS from 'pptxgenjs';
import {unzipSync,zipSync,strFromU8,strToU8} from 'fflate';
import {readPptx} from '../src/html-to-pptx/read-pptx.mjs';
import {exportHtml} from '../src/html-to-pptx/export-html.mjs';
import {publicDiagnostic} from '../src/html-to-pptx/public-diagnostics.mjs';
import {chromium} from 'playwright-core';
import {readInput} from '../src/html-to-pptx/input.mjs';
import {createBrowserRuntime} from '../src/html-to-pptx/browser-runtime.mjs';
import {measureInput} from '../src/html-to-pptx/measure.mjs';
import {exportPptx} from '../src/html-to-pptx/export.mjs';

const png=Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGNgYGD4DwABBAEAX+XDSwAAAABJRU5ErkJggg==','base64');
async function fixture(){
 const ppt=new PptxGenJS(),slide=ppt.addSlide();slide.background={color:'FFFFFF'};
 slide.addImage({data:'data:image/png;base64,'+png.toString('base64'),x:1,y:2,w:3,h:1});
 const parts=unzipSync(new Uint8Array(await ppt.write({outputType:'uint8array'}))),key='ppt/slides/slide1.xml';
 const xml=strFromU8(parts[key]),original=xml.match(/<p:pic>[\s\S]*?<\/p:pic>/)?.[0];assert.ok(original);
 const xfrm=original.match(/<a:xfrm>[\s\S]*?<\/a:xfrm>/)?.[0];assert.ok(xfrm);
 const alt=`<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><mc:Choice Requires="a14"><p:sp><p:txBody><a:p><a14:m><m:oMath/></a14:m></a:p></p:txBody></p:sp></mc:Choice><mc:Fallback><p:sp><p:nvSpPr><p:cNvPr id="2" name="Equation image"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr>${xfrm}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:blipFill><a:blip r:embed="rId1"/><a:stretch><a:fillRect l="-168"/></a:stretch></a:blipFill></p:spPr><p:txBody><a:p><a:r><a:t> </a:t></a:r></a:p></p:txBody></p:sp></mc:Fallback></mc:AlternateContent>`;
 parts[key]=strToU8(xml.replace(original,alt));return parts;
}
test('partial import preserves a math image fallback and reports that the formula is not editable',async()=>{
 const parts=await fixture(),bytes=zipSync(parts);assert.equal((await readPptx(bytes)).status,'failed');
 const result=await readPptx(bytes,{allowPartial:true});assert.equal(result.status,'parsed',JSON.stringify(result.diagnostics));
 const image=result.model.slides[0].nodes.find(n=>n.kind==='image');assert.ok(image);assert.equal(image.boxCssPx.x,96);assert.equal(image.boxCssPx.y,192);assert.equal(image.boxCssPx.w,288);assert.equal(image.boxCssPx.h,96);
 assert.deepEqual(image.content.crop,{left:168,top:0,right:0,bottom:0});
 const d=result.diagnostics.find(d=>d.code==='PPTX_MATH_IMAGE_FALLBACK');assert.equal(d?.severity,'warning');assert.match(publicDiagnostic(d).message,/수식.*편집/);assert.match(exportHtml(result.model).html,/data:image\/png;base64,/);
});
test('math image fill uses positive insets and clips negative insets on all sides',async()=>{
 const parts=await fixture(),key='ppt/slides/slide1.xml';
 parts[key]=strToU8(strFromU8(parts[key]).replace('l="-168"','l="20000" t="-50000" r="-50000" b="25000"'));
 const result=await readPptx(zipSync(parts),{allowPartial:true});assert.equal(result.status,'parsed',JSON.stringify(result.diagnostics));
 const image=result.model.slides[0].nodes.find(n=>n.kind==='image');
 assert.ok(Math.abs(image.boxCssPx.x-153.6)<.001);assert.equal(image.boxCssPx.y,192);
 assert.ok(Math.abs(image.boxCssPx.w-230.4)<.001);assert.equal(image.boxCssPx.h,72);
 assert.deepEqual(image.content.crop,{left:0,top:40000,right:38462,bottom:0});
 const html=exportHtml(result.model).html;assert.match(html,/object-view-box:inset\(40% 38\.462% 0% 0%\)/);
 const runtime=await createBrowserRuntime({capture:false});try{
  const measured=await measureInput(await readInput({html}),runtime),restored=measured.model.slides[0].nodes.find(n=>n.kind==='image');
  assert.deepEqual(restored.content.crop,image.content.crop);
  const output=await exportPptx(measured);assert.equal(output.status,'generated',JSON.stringify(output.diagnostics));
  const rect=strFromU8(unzipSync(output.pptxBytes)['ppt/slides/slide1.xml']).match(/<a:srcRect\b[^>]*>/)?.[0];
  assert.match(rect,/\bt="40000"/);assert.match(rect,/\br="38462"/);
 }finally{await runtime.close();}
});
test('math fill cropping shows the intended part of the fallback image',async()=>{
 const parts=await fixture(),key='ppt/slides/slide1.xml',media=Object.keys(parts).find(k=>/^ppt\/media\/.*\.png$/.test(k));
 parts[key]=strToU8(strFromU8(parts[key]).replace('l="-168"','t="-100000"'));
 parts[media]=Buffer.from('iVBORw0KGgoAAAANSUhEUgAAABQAAAAUCAIAAAAC64paAAAAJUlEQVR4nGP8z0A+YKJAL8Oo5pERYIwMDP9Hmp+ZRjWTBigKMAB8+QIlC4gOEgAAAABJRU5ErkJggg==','base64');
 const result=await readPptx(zipSync(parts),{allowPartial:true});assert.equal(result.status,'parsed',JSON.stringify(result.diagnostics));
 const browser=await chromium.launch({channel:'msedge',headless:true});try{
  const page=await browser.newPage({viewport:{width:1200,height:800}});await page.setContent(exportHtml(result.model).html);
  const screenshot=await page.locator('img[data-kind="image"]').screenshot();
  const pixel=await page.evaluate(async encoded=>{const image=new Image();image.src='data:image/png;base64,'+encoded;await image.decode();const canvas=document.createElement('canvas');canvas.width=image.width;canvas.height=image.height;const ctx=canvas.getContext('2d');ctx.drawImage(image,0,0);return [...ctx.getImageData(Math.floor(image.width/2),Math.floor(image.height/2),1,1).data];},screenshot.toString('base64'));
  assert.deepEqual(pixel,[0,0,255,255]);
 }finally{await browser.close();}
});
test('missing, unsafe or malformed fallback image data cannot bypass validation',async()=>{
 for(const mode of ['missing','bad-bytes','truncated-png','bad-png-crc','bad-png-stream','external','invalid-rect','empty-rect','oversized-rect','negative-size','invalid-rotation','invalid-flip','fractional-coordinate','scientific-coordinate','unknown-rect-attribute','unknown-rect-child']){const parts=await fixture(),slide='ppt/slides/slide1.xml',rels='ppt/slides/_rels/slide1.xml.rels',media=Object.keys(parts).find(k=>/^ppt\/media\/.*\.png$/.test(k));
  if(mode==='missing')parts[slide]=strToU8(strFromU8(parts[slide]).replace('r:embed="rId1"','r:embed="missing"'));
  if(mode==='bad-bytes')parts[media]=strToU8('<svg onload="alert(1)"></svg>');
  if(mode==='truncated-png')parts[media]=png.subarray(0,8);
  if(mode==='bad-png-crc'){parts[media]=Buffer.from(png);parts[media][45]^=1;}
  if(mode==='bad-png-stream')parts[media]=Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAACElEQVRub3QtemxpYq351UYAAAAASUVORK5CYII=','base64');
  if(mode==='external')parts[rels]=strToU8(strFromU8(parts[rels]).replace(/Target="\.\.\/media\/[^\"]+"/,'Target="https://example.com/x.png" TargetMode="External"'));
  if(mode==='invalid-rect')parts[slide]=strToU8(strFromU8(parts[slide]).replace('l="-168"','l="bad"'));
  if(mode==='empty-rect')parts[slide]=strToU8(strFromU8(parts[slide]).replace('l="-168"','l=""'));
  if(mode==='oversized-rect')parts[slide]=strToU8(strFromU8(parts[slide]).replace('l="-168"','l="-1000001"'));
  if(mode==='negative-size')parts[slide]=strToU8(strFromU8(parts[slide]).replace('l="-168"','l="80000" r="80000"'));
  if(mode==='invalid-rotation')parts[slide]=strToU8(strFromU8(parts[slide]).replace(/(<mc:Fallback>[\s\S]*?)<a:xfrm>/,'$1<a:xfrm rot="bad">'));
  if(mode==='invalid-flip')parts[slide]=strToU8(strFromU8(parts[slide]).replace(/(<mc:Fallback>[\s\S]*?)<a:xfrm>/,'$1<a:xfrm flipH="bad">'));
  if(mode==='fractional-coordinate')parts[slide]=strToU8(strFromU8(parts[slide]).replace('cx="2743200"','cx="12.5"'));
  if(mode==='scientific-coordinate')parts[slide]=strToU8(strFromU8(parts[slide]).replace('x="914400"','x="1e6"'));
  if(mode==='unknown-rect-attribute')parts[slide]=strToU8(strFromU8(parts[slide]).replace('l="-168"','l="-168" bad="1"'));
  if(mode==='unknown-rect-child')parts[slide]=strToU8(strFromU8(parts[slide]).replace('<a:fillRect l="-168"/>','<a:fillRect l="-168"><a:unexpected/></a:fillRect>'));
  const result=await readPptx(zipSync(parts),{allowPartial:true});assert.equal(result.status,'failed',mode);assert.ok(result.diagnostics.some(d=>d.blocking),mode);
 }
});
test('non-math alternates do not use the math image fallback',async()=>{
 const parts=await fixture(),key='ppt/slides/slide1.xml';parts[key]=strToU8(strFromU8(parts[key]).replace('<a14:m><m:oMath/></a14:m>',''));
 const result=await readPptx(zipSync(parts),{allowPartial:true});assert.equal(result.status,'parsed');assert.ok(!result.model.slides[0].nodes.some(n=>n.kind==='image'));assert.ok(result.diagnostics.some(d=>d.code==='PPTX_OBJECT_PLACEHOLDER'));
});
test('unknown alternate branches and image fill shapes are not treated as a supported fallback',async()=>{
 for(const [from,to] of [['</mc:AlternateContent>','<mc:Other/></mc:AlternateContent>'],['<a:stretch>','<a:tile/><a:stretch>']]){
  const parts=await fixture(),key='ppt/slides/slide1.xml';parts[key]=strToU8(strFromU8(parts[key]).replace(from,to));
  const result=await readPptx(zipSync(parts),{allowPartial:true});assert.equal(result.status,'parsed');assert.ok(!result.model.slides[0].nodes.some(n=>n.kind==='image'));assert.ok(result.diagnostics.some(d=>d.code==='PPTX_OBJECT_PLACEHOLDER'));
 }
});
