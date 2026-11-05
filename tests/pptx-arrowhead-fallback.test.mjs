import test from 'node:test';
import assert from 'node:assert/strict';
import PptxGenJS from 'pptxgenjs';
import {zipSync,unzipSync,strFromU8,strToU8} from 'fflate';
import {readPptx} from '../src/html-to-pptx/read-pptx.mjs';
import {exportHtml} from '../src/html-to-pptx/export-html.mjs';
import {publicDiagnostic} from '../src/html-to-pptx/public-diagnostics.mjs';
import {readInput} from '../src/html-to-pptx/input.mjs';import {measureInput} from '../src/html-to-pptx/measure.mjs';import {createBrowserRuntime} from '../src/html-to-pptx/browser-runtime.mjs';import {exportPptx} from '../src/html-to-pptx/export.mjs';

async function fixture({dash=false,shape='straightConnector1',width=9525}={}){
 const ppt=new PptxGenJS(),slide=ppt.addSlide();slide.background={color:'FFFFFF'};
 const parts=unzipSync(new Uint8Array(await ppt.write({outputType:'uint8array'}))),file='ppt/slides/slide1.xml',xml=strFromU8(parts[file]);
 const object=`<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="42" name="arrow"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr><p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="0"/></a:xfrm><a:prstGeom prst="${shape}"><a:avLst/></a:prstGeom><a:ln w="${width}"><a:solidFill><a:srgbClr val="112233"/></a:solidFill>${dash?'<a:prstDash val="dash"/>':''}<a:tailEnd type="triangle" w="sm" len="sm"/></a:ln></p:spPr></p:cxnSp>`;
 parts[file]=strToU8(xml.replace('</p:spTree>',object+'</p:spTree>'));return zipSync(parts);
}
test('partial mode retains a visible editable SVG arrowhead and preserves the thin line',async()=>{
 const bytes=await fixture({width:3175});assert.equal((await readPptx(bytes)).status,'failed');
 const result=await readPptx(bytes,{allowPartial:true});assert.equal(result.status,'parsed',JSON.stringify(result.diagnostics));
 const line=result.model.slides[0].nodes.find(n=>n.id==='arrow');assert.equal(line?.kind,'line');
 assert.deepEqual(line.style.endpoints,{start:{x:96,y:96},end:{x:288,y:96}});
 assert.equal(line.style.stroke.toLowerCase(),'#112233');assert.ok(Math.abs(line.style.strokeWidth-1/3)<1e-9);
 const head=result.model.slides[0].nodes.find(n=>n.id==='arrow::arrow-tail');assert.equal(head?.content.editableVector?.shape,'triangle');
 const html=exportHtml(result.model).html;assert.match(html,/<line\b/);assert.match(html,/data-preset-vector/);assert.ok(!result.diagnostics.some(d=>d.reasonKey==='PPTX_ARROWHEAD_OMITTED'));
 const warning=publicDiagnostic(result.diagnostics.find(d=>d.code==='PPTX_SHAPE_IMAGE_APPROXIMATED'));assert.match(warning.message,/SVG|그림/);
 assert.ok(!result.diagnostics.some(d=>d.code==='PPTX_OBJECT_PLACEHOLDER'&&d.objectName==='arrow'));
 const runtime=await createBrowserRuntime({capture:false});try{
  const measured=await measureInput(await readInput({html}),runtime),out=await exportPptx(measured);assert.equal(out.status,'generated',JSON.stringify(out.diagnostics));
  const xml=strFromU8(unzipSync(out.pptxBytes)['ppt/slides/slide1.xml']);assert.match(xml,/<a:ln w="3175"/);
  const edited=html.replace('stroke-width:1px;shape-rendering:crispEdges','stroke-width:2px;shape-rendering:crispEdges');assert.notEqual(edited,html);
  const editedOut=await exportPptx(await measureInput(await readInput({html:edited}),runtime));assert.equal(editedOut.status,'generated',JSON.stringify(editedOut.diagnostics));
  assert.match(strFromU8(unzipSync(editedOut.pptxBytes)['ppt/slides/slide1.xml']),/<a:ln w="19050"/);
 }finally{await runtime.close();}
});
test('both triangle ends follow a reflected diagonal connector',async()=>{
 const parts=unzipSync(await fixture()),key='ppt/slides/slide1.xml';let xml=strFromU8(parts[key]);
 xml=xml.replace('<a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="0"/>','<a:xfrm flipH="1" flipV="1"><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/>').replace('<a:tailEnd','<a:headEnd type="triangle" w="sm" len="sm"/><a:tailEnd');
 parts[key]=strToU8(xml);const parsed=await readPptx(zipSync(parts),{allowPartial:true});assert.equal(parsed.status,'parsed',JSON.stringify(parsed.diagnostics));
 const nodes=parsed.model.slides[0].nodes,line=nodes.find(n=>n.id==='arrow');
 for(const [end,tip] of [['head',line.style.endpoints.start],['tail',line.style.endpoints.end]]){
  const n=nodes.find(node=>node.id===`arrow::arrow-${end}`),v=n.content.editableVector,a=v.angle*Math.PI/180;
  assert.ok(Math.abs(n.boxCssPx.x+n.boxCssPx.w/2+Math.cos(a)*v.width/2-tip.x)<1e-7);
  assert.ok(Math.abs(n.boxCssPx.y+n.boxCssPx.h/2+Math.sin(a)*v.width/2-tip.y)<1e-7);
 }
});
test('unsupported line patterns and connector geometry are still placeholders',async()=>{
 for(const options of [{dash:true},{shape:'bentConnector4'}]){
  const result=await readPptx(await fixture(options),{allowPartial:true});assert.equal(result.status,'parsed');
  assert.ok(!result.model.slides[0].nodes.some(n=>n.id==='arrow'));
  assert.ok(result.diagnostics.some(d=>d.code==='PPTX_OBJECT_PLACEHOLDER'&&d.objectName==='arrow'));
 }
});
