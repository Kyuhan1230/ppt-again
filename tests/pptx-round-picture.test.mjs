import test from 'node:test';import assert from 'node:assert/strict';import PptxGenJS from 'pptxgenjs';
import {unzipSync,zipSync,strFromU8,strToU8} from 'fflate';
import {readPptx} from '../src/html-to-pptx/read-pptx.mjs';import {exportHtml} from '../src/html-to-pptx/export-html.mjs';
import {readInput} from '../src/html-to-pptx/input.mjs';import {measureInput} from '../src/html-to-pptx/measure.mjs';import {createBrowserRuntime} from '../src/html-to-pptx/browser-runtime.mjs';import {exportPptx} from '../src/html-to-pptx/export.mjs';
const png='iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAYAAABytg0kAAAAF0lEQVR4nGP4z8Dwn+E/QwMjmP7P8B8APeIHeyMDVGUAAAAASUVORK5CYII=';
const pt=(x,y)=>`<a:pt x="${x}" y="${y}"/>`;
const customEllipse=`<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="l" t="t" r="r" b="b"/><a:pathLst><a:path w="1000" h="1000"><a:moveTo>${pt(500,0)}</a:moveTo><a:cubicBezTo>${pt(776,0)}${pt(1000,224)}${pt(1000,500)}</a:cubicBezTo><a:cubicBezTo>${pt(1000,776)}${pt(776,1000)}${pt(500,1000)}</a:cubicBezTo><a:cubicBezTo>${pt(224,1000)}${pt(0,776)}${pt(0,500)}</a:cubicBezTo><a:cubicBezTo>${pt(0,224)}${pt(224,0)}${pt(500,0)}</a:cubicBezTo><a:close/></a:path></a:pathLst></a:custGeom>`;
async function fixture(geometry){
 const ppt=new PptxGenJS(),slide=ppt.addSlide();slide.background={color:'FFFFFF'};slide.addImage({data:`data:image/png;base64,${png}`,x:1,y:1,w:2,h:3,objectName:'photo'});
 const parts=unzipSync(new Uint8Array(await ppt.write({outputType:'uint8array'}))),key='ppt/slides/slide1.xml';
 const xml=strFromU8(parts[key]);assert.match(xml,/<a:prstGeom prst="rect">/);const changed=xml.replace(/<a:prstGeom prst="rect">[\s\S]*?<\/a:prstGeom>/,geometry);parts[key]=strToU8(changed);return zipSync(parts);
}
test('ellipse picture keeps its shape through HTML and back to PowerPoint',async()=>{
 const parsed=await readPptx(await fixture('<a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>'));assert.equal(parsed.status,'parsed',JSON.stringify(parsed.diagnostics));
 assert.equal(parsed.model.slides[0].nodes.find(n=>n.id==='photo').content.rounding,true);
 const html=exportHtml(parsed.model).html;assert.match(html,/data-pptx-rounded-image/);assert.match(html,/border-radius:50%/);
 const runtime=await createBrowserRuntime({capture:false});try{
  const measured=await measureInput(await readInput({html}),runtime),ppt=await exportPptx(measured);
  assert.equal(ppt.status,'generated',JSON.stringify(ppt.diagnostics));assert.match(strFromU8(unzipSync(ppt.pptxBytes)['ppt/slides/slide1.xml']),/<a:prstGeom prst="ellipse"/);
 }finally{await runtime.close();}
});
test('a near-elliptical custom picture mask is approximated as an ellipse in partial mode',async()=>{
 const parsed=await readPptx(await fixture(customEllipse),{allowPartial:true});assert.equal(parsed.status,'parsed',JSON.stringify(parsed.diagnostics));
 assert.equal(parsed.model.slides[0].nodes.find(n=>n.id==='photo').content.rounding,true);assert.ok(parsed.diagnostics.some(d=>d.code==='PPTX_IMAGE_APPEARANCE_APPROXIMATED'));
 const square=customEllipse.replace(/<a:moveTo>[\s\S]*?<a:close\/>/,`<a:moveTo>${pt(0,0)}</a:moveTo><a:lnTo>${pt(1000,0)}</a:lnTo><a:lnTo>${pt(1000,1000)}</a:lnTo><a:lnTo>${pt(0,1000)}</a:lnTo><a:close/>`);
 const other=await readPptx(await fixture(square),{allowPartial:true});assert.equal(other.model.slides[0].nodes.find(n=>n.id==='photo').content.rounding,undefined);
});
test('a picture outline with default width remains visible and editable on PPT export',async()=>{
 const parts=unzipSync(await fixture('<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>')),key='ppt/slides/slide1.xml',source=strFromU8(parts[key]);
 const changed=source.replace('</p:spPr>','<a:ln><a:solidFill><a:srgbClr val="7F7F7F"/></a:solidFill></a:ln></p:spPr>');assert.notEqual(changed,source);parts[key]=strToU8(changed);
 const parsed=await readPptx(zipSync(parts));assert.equal(parsed.status,'parsed',JSON.stringify(parsed.diagnostics));
 const outline=parsed.model.slides[0].nodes.find(n=>n.id==='photo::outline');assert.equal(outline?.kind,'shape');assert.equal(outline.style.stroke,'#7f7f7f');assert.equal(outline.style.strokeWidth,1);
 const html=exportHtml(parsed.model).html;assert.match(html,/data-source-id="photo::outline"/);
 const runtime=await createBrowserRuntime({capture:false});try{
  const measured=await measureInput(await readInput({html}),runtime),ppt=await exportPptx(measured);assert.equal(ppt.status,'generated',JSON.stringify(ppt.diagnostics));
  const xml=strFromU8(unzipSync(ppt.pptxBytes)['ppt/slides/slide1.xml']);assert.match(xml,/<p:pic>/);assert.match(xml,/name="photo::outline"/);assert.match(xml,/val="7F7F7F"/);
 }finally{await runtime.close();}
});
test('unsupported picture outline patterns are reported instead of drawn as solid',async()=>{
 const parts=unzipSync(await fixture('<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>')),key='ppt/slides/slide1.xml',source=strFromU8(parts[key]);
 parts[key]=strToU8(source.replace('</p:spPr>','<a:ln w="12700"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:prstDash val="dash"/></a:ln></p:spPr>'));
 assert.equal((await readPptx(zipSync(parts))).status,'failed');
 const partial=await readPptx(zipSync(parts),{allowPartial:true});assert.equal(partial.status,'parsed');assert.ok(partial.diagnostics.some(d=>d.code==='PPTX_IMAGE_APPEARANCE_APPROXIMATED'));
 assert.ok(!partial.model.slides[0].nodes.some(n=>n.id==='photo::outline'));
});
