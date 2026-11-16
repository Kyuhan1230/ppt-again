import test from 'node:test';
import assert from 'node:assert/strict';
import PptxGenJS from 'pptxgenjs';
import {unzipSync,zipSync,strFromU8,strToU8} from 'fflate';
import {readPptx} from '../src/html-to-pptx/read-pptx.mjs';
import {exportHtml} from '../src/html-to-pptx/export-html.mjs';
import {publicDiagnostic} from '../src/html-to-pptx/public-diagnostics.mjs';
import {readInput} from '../src/html-to-pptx/input.mjs';
import {createBrowserRuntime} from '../src/html-to-pptx/browser-runtime.mjs';
import {measureInput} from '../src/html-to-pptx/measure.mjs';
import {exportPptx} from '../src/html-to-pptx/export.mjs';
import {chromium} from 'playwright-core';

async function fixture(fontBytes){
 const ppt=new PptxGenJS(),slide=ppt.addSlide();slide.background={color:'FFFFFF'};
 slide.addText('Original font text',{x:1,y:1,w:4,h:1,fontFace:'Arial',fontSize:20});
 const parts=unzipSync(new Uint8Array(await ppt.write({outputType:'uint8array'})));
 const patch=(name,find,replace)=>{const before=strFromU8(parts[name]),after=before.replace(find,replace);assert.notEqual(before,after);parts[name]=strToU8(after);};
 patch('ppt/presentation.xml','</p:presentation>','<p:embeddedFontLst><p:embeddedFont><p:font typeface="Arial"/><p:regular r:id="fontQA"/></p:embeddedFont></p:embeddedFontLst></p:presentation>');
 patch('ppt/_rels/presentation.xml.rels','</Relationships>','<Relationship Id="fontQA" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/font" Target="fonts/test.fntdata"/></Relationships>');
 patch('[Content_Types].xml','</Types>','<Default Extension="fntdata" ContentType="application/x-fontdata"/></Types>');
 parts['ppt/fonts/test.fntdata']=fontBytes;return zipSync(parts);
}
test('partial original-font import retains font names when embedded EOT compression is unsupported',async()=>{
 const eot=Buffer.alloc(82);eot.writeUInt32LE(eot.length,0);eot.writeUInt32LE(0x20002,8);eot.writeUInt16LE(0x504c,34);
 const bytes=await fixture(eot);const strict=await readPptx(bytes);assert.equal(strict.status,'failed');
 const result=await readPptx(bytes,{fontPolicy:'original',allowPartial:true});assert.equal(result.status,'parsed',JSON.stringify(result.diagnostics));
 assert.equal(result.model.slides[0].nodes[0].style.fontFamily,'Arial');
 assert.ok(!result.model.assets.some(asset=>asset.mime.startsWith('font/')));
 assert.match(exportHtml(result.model).html,/Original font text/);
 const warning=result.diagnostics.find(d=>d.reasonKey==='PPTX_EMBEDDED_FONT_LOCAL_ONLY');assert.ok(warning);assert.equal(warning.severity,'warning');assert.equal(publicDiagnostic(warning).reasonKey,'PPTX_EMBEDDED_FONT_LOCAL_ONLY');
});
test('partial original-font import still rejects invalid font headers',async()=>{
 const result=await readPptx(await fixture(new Uint8Array([1,2,3,4])),{fontPolicy:'original',allowPartial:true});
 assert.equal(result.status,'failed');assert.ok(result.diagnostics.some(d=>d.code==='FONT_INVALID'));
});
test('installed full face names render correctly and remain native text on PPT export',async()=>{
 const ppt=new PptxGenJS(),slide=ppt.addSlide();slide.background={color:'FFFFFF'};
 slide.addText('Bold face name',{x:1,y:1,w:5,h:1,fontFace:'Arial Bold',fontSize:24,objectName:'full-face'});
 const parsed=await readPptx(new Uint8Array(await ppt.write({outputType:'uint8array'})),{fontPolicy:'original'});assert.equal(parsed.status,'parsed',JSON.stringify(parsed.diagnostics));
 const html=exportHtml(parsed.model).html;assert.match(html,/@font-face\{font-family:"Arial Bold";src:local\("Arial Bold"\)/);
 const browser=await chromium.launch({channel:'msedge',headless:true});try{
  const page=await browser.newPage();await page.setContent(html);await page.evaluate(()=>document.fonts.ready);
  const width=await page.locator('[data-source-id="full-face"] span span').last().evaluate(el=>{const r=document.createRange();r.selectNodeContents(el);return r.getBoundingClientRect().width;});
  const reference=await page.evaluate(()=>{const e=document.createElement('span');e.textContent='Bold face name';e.style.cssText='position:absolute;font:700 32px Arial;letter-spacing:0';document.body.append(e);return e.getBoundingClientRect().width;});
  assert.ok(Math.abs(width-reference)<.2,JSON.stringify({width,reference}));
 }finally{await browser.close();}
 const input=await readInput({html});assert.ok(input.fontFaces.some(face=>face.family==='Arial Bold'&&face.selfLocal));
 const runtime=await createBrowserRuntime({capture:false});try{
  const measured=await measureInput(input,runtime),result=await exportPptx(measured);
  assert.equal(result.status,'generated',JSON.stringify(result.diagnostics));assert.match(strFromU8(unzipSync(result.pptxBytes)['ppt/slides/slide1.xml']),/typeface="Arial Bold"/);
 }finally{await runtime.close();}
});
test('a local face that aliases a different font is not treated as the named native font',async()=>{
 const html='<style>@font-face{font-family:"Arial Bold";src:local("Times New Roman");font-weight:400;font-style:normal}</style><section data-slide style="width:800px;height:450px"><p style="font:24px \'Arial Bold\'">text</p></section>';
 const input=await readInput({html});assert.equal(input.fontFaces[0].selfLocal,false);
 const runtime=await createBrowserRuntime({capture:false});try{
  const measured=await measureInput(input,runtime);assert.ok(measured.diagnostics.some(d=>d.code==='FONT_SOURCE_UNSUPPORTED'),JSON.stringify(measured.diagnostics));
 }finally{await runtime.close();}
});
test('generated local face aliases stay inside the font-face input limit',async()=>{
 const ppt=new PptxGenJS(),slide=ppt.addSlide();slide.background={color:'FFFFFF'};slide.addText('X',{x:1,y:1,w:4,h:1,fontFace:'Arial',fontSize:16});
 const parsed=await readPptx(new Uint8Array(await ppt.write({outputType:'uint8array'}))),node=parsed.model.slides[0].nodes[0];
 node.content.runs=Array.from({length:40},(_,i)=>({text:'X',style:{...node.content.runs[0].style,fontFamily:`Font${i} Bold`}}));node.content.text='X'.repeat(40);
 const input=await readInput({html:exportHtml(parsed.model).html});assert.equal(input.fontFaces.length,32);assert.ok(input.fontFaces.every(face=>face.selfLocal));
});
