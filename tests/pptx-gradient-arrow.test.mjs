import test from 'node:test';
import assert from 'node:assert/strict';
import PptxGenJS from 'pptxgenjs';
import {unzipSync,zipSync,strFromU8,strToU8} from 'fflate';
import {readPptx} from '../src/html-to-pptx/read-pptx.mjs';
import {exportHtml} from '../src/html-to-pptx/export-html.mjs';
import {readInput} from '../src/html-to-pptx/input.mjs';
import {measureInput} from '../src/html-to-pptx/measure.mjs';
import {createBrowserRuntime} from '../src/html-to-pptx/browser-runtime.mjs';

async function fixture(){
 const p=new PptxGenJS(),s=p.addSlide();s.background={color:'FFFFFF'};
 for(const [name,rotate] of [['horizontal',0],['vertical',90]])s.addShape(p.ShapeType.rightArrow,{x:1,y:rotate?3:1,w:4,h:.5,rotate,fill:{color:'0070C0'},line:{transparency:100},objectName:name});
 const parts=unzipSync(new Uint8Array(await p.write({outputType:'uint8array'}))),key='ppt/slides/slide1.xml';
 let xml=strFromU8(parts[key]);const old='<a:solidFill><a:srgbClr val="0070C0"/></a:solidFill>';
 assert.equal(xml.split(old).length-1,2);
 xml=xml.replaceAll(old,'<a:gradFill flip="none" rotWithShape="1"><a:gsLst><a:gs pos="0"><a:srgbClr val="0070C0"/></a:gs><a:gs pos="100000"><a:srgbClr val="FF0000"/></a:gs></a:gsLst><a:lin ang="0" scaled="1"/><a:tileRect/></a:gradFill>');
 assert.equal(xml.split('<a:solidFill><a:srgbClr val="333333"><a:alpha val="0"/></a:srgbClr></a:solidFill>').length-1,2);
 xml=xml.replaceAll('<a:solidFill><a:srgbClr val="333333"><a:alpha val="0"/></a:srgbClr></a:solidFill>','<a:noFill/>');
 parts[key]=strToU8(xml);return parts;
}
test('gradient arrows retain two editable vector endpoints and survive HTML reinput',async()=>{
 const parsed=await readPptx(zipSync(await fixture()),{allowPartial:true});assert.equal(parsed.status,'parsed',JSON.stringify(parsed.diagnostics));
 for(const name of ['horizontal','vertical']){
  const node=parsed.model.slides[0].nodes.find(n=>n.id===name);assert.equal(node?.kind,'image',name);
  assert.equal(node.content.editableVector.shape,'rightArrow');assert.deepEqual([node.content.editableVector.gradient.from,node.content.editableVector.gradient.to],['#0070c0','#ff0000']);
  assert.ok(!parsed.diagnostics.some(d=>d.code==='PPTX_OBJECT_PLACEHOLDER'&&d.objectName===name));
 }
 const out=exportHtml(parsed.model);assert.equal(out.status,'generated');assert.equal((out.html.match(/<linearGradient\b/g)??[]).length,2);
 const input=await readInput({html:out.html});assert.ok(!input.diagnostics.some(d=>d.severity==='error'),JSON.stringify(input.diagnostics));
 const runtime=await createBrowserRuntime({capture:false});try{
  const measured=await measureInput(input,runtime);assert.ok(!measured.diagnostics.some(d=>d.severity==='error'),JSON.stringify(measured.diagnostics));
  assert.equal(measured.diagnostics.filter(d=>d.code==='PPTX_SHAPE_IMAGE_APPROXIMATED').length,2);
 }finally{await runtime.close();}
});
test('unsupported gradient angles and external SVG references are rejected',async()=>{
 const parts=await fixture(),key='ppt/slides/slide1.xml';parts[key]=strToU8(strFromU8(parts[key]).replace('ang="0"','ang="12345"'));
 const parsed=await readPptx(zipSync(parts),{allowPartial:true});assert.equal(parsed.status,'parsed');
 assert.ok(parsed.diagnostics.some(d=>d.code==='PPTX_OBJECT_PLACEHOLDER'&&d.objectName==='horizontal'));
 const out=exportHtml(parsed.model);assert.equal(out.status,'generated');
 const tampered=out.html.replace(/fill="url\(#[^)]+\)"/,'fill="url(https://example.invalid/gradient)"');assert.notEqual(tampered,out.html);
 assert.ok((await readInput({html:tampered})).diagnostics.some(d=>d.code==='INPUT_INVALID'));
});
