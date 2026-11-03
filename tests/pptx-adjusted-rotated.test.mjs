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
 const ppt=new PptxGenJS(),slide=ppt.addSlide();slide.background={color:'FFFFFF'};
 slide.addShape(ppt.ShapeType.downArrow,{x:1,y:1,w:1,h:2,rotate:30,fill:{color:'00B050'},line:{color:'00B050',width:1},objectName:'adjusted-arrow'});
 slide.addShape(ppt.ShapeType.ellipse,{x:3,y:1,w:1,h:2,rotate:25,fill:{transparency:100,color:'FFFFFF'},line:{color:'FF0000',width:1},objectName:'colored-ellipse'});
 const parts=unzipSync(new Uint8Array(await ppt.write({outputType:'uint8array'}))),key='ppt/slides/slide1.xml';let xml=strFromU8(parts[key]);
 assert.match(xml,/<a:prstGeom prst="downArrow"><a:avLst><\/a:avLst><\/a:prstGeom>/);
 xml=xml.replace('<a:prstGeom prst="downArrow"><a:avLst></a:avLst></a:prstGeom>','<a:prstGeom prst="downArrow"><a:avLst><a:gd name="adj1" fmla="val 34830"/><a:gd name="adj2" fmla="val 72450"/></a:avLst></a:prstGeom>');
 xml=xml.replace('<a:srgbClr val="FF0000"/>','<a:schemeClr val="accent2"><a:lumMod val="60000"/><a:lumOff val="40000"/></a:schemeClr>');
 xml=xml.replace('<a:solidFill><a:srgbClr val="FFFFFF"><a:alpha val="0"/></a:srgbClr></a:solidFill>','<a:noFill/>');parts[key]=strToU8(xml);return parts;
}
test('adjusted rotated arrow and transformed theme ellipse remain editable SVG elements',async()=>{
 const bytes=zipSync(await fixture());assert.equal((await readPptx(bytes)).status,'failed');
 const r=await readPptx(bytes,{allowPartial:true});assert.equal(r.status,'parsed',JSON.stringify(r.diagnostics));
 for(const id of ['adjusted-arrow','colored-ellipse']){
  const node=r.model.slides[0].nodes.find(n=>n.id===id);assert.equal(node?.kind,'image',id);assert.ok(node.content.editableVector,id);
  assert.ok(!r.diagnostics.some(d=>d.code==='PPTX_OBJECT_PLACEHOLDER'&&d.objectName===id),id);
 }
 const arrow=r.model.slides[0].nodes.find(n=>n.id==='adjusted-arrow');assert.equal(arrow.content.editableVector.adj1,34830);assert.equal(arrow.content.editableVector.adj2,72450);
 const out=exportHtml(r.model);assert.equal(out.status,'generated');assert.match(out.html,/<svg[^>]+data-source-id="adjusted-arrow"/);assert.match(out.html,/<path/);assert.match(out.html,/<ellipse/);
 const edited=out.html.replace('fill="#00b050"','fill="#123456"');assert.notEqual(edited,out.html);
 const runtime=await createBrowserRuntime({capture:false});try{
  const measured=await measureInput(await readInput({html:edited}),runtime);
  assert.ok(measured.diagnostics.some(d=>d.code==='PPTX_SHAPE_IMAGE_APPROXIMATED'&&d.severity==='warning'));
  assert.ok(!measured.diagnostics.some(d=>d.severity==='error'),JSON.stringify(measured.diagnostics));
  const editedNode=measured.model.slides[0].nodes.find(n=>n.id==='adjusted-arrow');
  const bytes=measured.model.assets.find(a=>a.id===editedNode.content.assetId).bytes;
  assert.match(Buffer.from(bytes).toString('utf8'),/fill="#123456"/);
 }finally{await runtime.close();}
});
test('unsupported arrow adjustment or transparent effects remain explicitly unsupported',async()=>{
 for(const mode of ['formula','translucent','oversized','duplicate']){
  const parts=await fixture(),key='ppt/slides/slide1.xml';let xml=strFromU8(parts[key]);
  if(mode==='formula')xml=xml.replace('name="adj1" fmla="val 34830"','name="adj1" fmla="*/ 1 2 3"');
  if(mode==='translucent')xml=xml.replace('<a:srgbClr val="00B050"/>','<a:srgbClr val="00B050"><a:alpha val="50000"/></a:srgbClr>');
  if(mode==='oversized')xml=xml.replace('name="adj1" fmla="val 34830"','name="adj1" fmla="val 999999999999999999999"');
  if(mode==='duplicate')xml=xml.replace('<a:gd name="adj1" fmla="val 34830"/>','<a:gd name="adj1" fmla="val 34830"/><a:gd name="adj1" fmla="val 1"/>');
  parts[key]=strToU8(xml);const r=await readPptx(zipSync(parts),{allowPartial:true});
  assert.equal(r.status,'parsed');assert.ok(r.diagnostics.some(d=>d.code==='PPTX_OBJECT_PLACEHOLDER'&&d.objectName==='adjusted-arrow'),mode);
 }
});

test('unrotated down arrows are also editable SVG and preserve themed fills',async()=>{
 const ppt=new PptxGenJS(),slide=ppt.addSlide();slide.background={color:'FFFFFF'};
 slide.addShape(ppt.ShapeType.downArrow,{x:1,y:1,w:1,h:2,fill:{color:'AABBCC'},line:{color:'000000',width:1},objectName:'down-arrow'});
 const parts=unzipSync(new Uint8Array(await ppt.write({outputType:'uint8array'}))),key='ppt/slides/slide1.xml';
 const xml=strFromU8(parts[key]);assert.match(xml,/<a:srgbClr val="AABBCC"\/>/);
 parts[key]=strToU8(xml.replace('<a:srgbClr val="AABBCC"/>','<a:schemeClr val="bg1"><a:lumMod val="95000"/></a:schemeClr>'));
 const result=await readPptx(zipSync(parts),{allowPartial:true});assert.equal(result.status,'parsed',JSON.stringify(result.diagnostics));
 const arrow=result.model.slides[0].nodes.find(n=>n.id==='down-arrow');assert.equal(arrow?.kind,'image');
 assert.equal(arrow.content.editableVector.shape,'downArrow');assert.equal(arrow.content.editableVector.angle,0);
 assert.match(exportHtml(result.model).html,/<svg[^>]+data-source-id="down-arrow"/);
 assert.ok(!result.diagnostics.some(d=>d.code==='PPTX_OBJECT_PLACEHOLDER'&&d.objectName==='down-arrow'));
});

test('editable vector rejects injected markup, external references and unbounded numbers',async()=>{
 const parsed=await readPptx(zipSync(await fixture()),{allowPartial:true}),out=exportHtml(parsed.model);assert.equal(out.status,'generated');
 const mutations=[
  html=>html.replace('data-preset-vector=""','data-preset-vector="" onload="alert(1)"'),
  html=>html.replace('<g transform=', '<image href="https://example.invalid/image.png"/><g transform='),
  html=>html.replace('fill="#00b050"','fill="url(https://example.invalid/x)"'),
  html=>html.replace(/viewBox="0 0 [^ ]+ [^"]+"/,'viewBox="0 0 1e999 1"'),
  html=>html.replace('rotate(30)','rotate(1e999)'),
  html=>html.replace(/<path d="M [^ ]+ /,'<path d="M 100001 '),
 ];
 for(const mutate of mutations){
  const altered=mutate(out.html);assert.notEqual(altered,out.html);
  const input=await readInput({html:altered});assert.ok(input.diagnostics.some(d=>d.code==='INPUT_INVALID'&&d.severity==='error'),JSON.stringify(input.diagnostics));
 }
 const model=structuredClone(parsed.model);model.slides[0].nodes.find(n=>n.id==='adjusted-arrow').content.editableVector.fill='#123456" onload="alert(1)';
 assert.equal(exportHtml(model).status,'failed');
});
