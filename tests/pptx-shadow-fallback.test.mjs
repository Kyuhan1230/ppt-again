import test from 'node:test';
import assert from 'node:assert/strict';
import PptxGenJS from 'pptxgenjs';
import {zipSync,unzipSync,strFromU8,strToU8} from 'fflate';
import {readPptx} from '../src/html-to-pptx/read-pptx.mjs';
import {exportHtml} from '../src/html-to-pptx/export-html.mjs';
import {publicDiagnostic} from '../src/html-to-pptx/public-diagnostics.mjs';

async function fixture(effect='<a:outerShdw blurRad="190500" dist="38100" dir="5400000" algn="t" rotWithShape="0"><a:prstClr val="black"><a:alpha val="40000"/></a:prstClr></a:outerShdw>'){
 const ppt=new PptxGenJS(),slide=ppt.addSlide();slide.background={color:'FFFFFF'};
 slide.addText('Process step',{x:1,y:1,w:3,h:1,fontFace:'Arial',fontSize:18,fill:{color:'DDDDDD'},line:{color:'333333',width:1},objectName:'step'});
 slide.addShape(ppt.ShapeType.roundRect,{x:1,y:3,w:3,h:2,fill:{color:'FFFFFF'},line:{transparency:100},objectName:'panel'});
 const parts=unzipSync(new Uint8Array(await ppt.write({outputType:'uint8array'}))),file='ppt/slides/slide1.xml';let xml=strFromU8(parts[file]),changed=0;
 xml=xml.replace(/<p:sp>[\s\S]*?<\/p:sp>/g,shape=>{if(!/name="(?:step|panel)"/.test(shape))return shape;changed++;return shape.replace('</p:spPr>',`<a:effectLst>${effect}</a:effectLst></p:spPr>`);});
 assert.equal(changed,2);parts[file]=strToU8(xml);return zipSync(parts);
}
test('partial mode approximates a simple outer shadow without losing editable text',async()=>{
 const bytes=await fixture();const strict=await readPptx(bytes);assert.equal(strict.status,'failed');
 const partial=await readPptx(bytes,{allowPartial:true});assert.equal(partial.status,'parsed',JSON.stringify(partial.diagnostics));
 const nodes=partial.model.slides[0].nodes,step=nodes.find(n=>n.id==='step'),panel=nodes.find(n=>n.id==='panel');
 assert.equal(step?.kind,'text');assert.equal(step?.content.text,'Process step');assert.equal(step?.style.fill.toLowerCase(),'#dddddd');
 assert.equal(panel?.kind,'shape');assert.ok(panel?.style.radius>0);const html=exportHtml(partial.model).html;assert.match(html,/Process step/);assert.match(html,/box-shadow:/);
 assert.ok(Math.abs(step.style.shadow.y-4)<.01);assert.equal(step.style.shadow.blur,20);assert.equal(step.style.shadow.alpha,.4);
 assert.equal(partial.diagnostics.filter(d=>d.reasonKey==='PPTX_SHADOW_APPROXIMATED'&&d.severity==='warning').length,2);
 const warning=publicDiagnostic(partial.diagnostics.find(d=>d.reasonKey==='PPTX_SHADOW_APPROXIMATED'));assert.equal(warning.code,'PPTX_SHADOW_APPROXIMATED');assert.match(warning.message,/그림자/);
 assert.ok(!partial.diagnostics.some(d=>d.code==='PPTX_OBJECT_PLACEHOLDER'&&['step','panel'].includes(d.objectName)));
});
test('an unsupported transformed shadow is omitted with an explicit warning',async()=>{
 const effect='<a:outerShdw blurRad="190500" dist="38100" dir="5400000" kx="600000"><a:prstClr val="black"><a:alpha val="40000"/></a:prstClr></a:outerShdw>';
 const result=await readPptx(await fixture(effect),{allowPartial:true});assert.equal(result.status,'parsed');
 assert.ok(result.diagnostics.some(d=>d.reasonKey==='PPTX_SHADOW_OMITTED'));assert.equal(result.model.slides[0].nodes.find(n=>n.id==='step').style.shadow,undefined);
});
test('other shape effects remain explicit placeholders in partial mode',async()=>{
 for(const effect of [
  '<a:glow rad="190500"><a:srgbClr val="FF0000"/></a:glow>',
  '<a:outerShdw blurRad="10000"/><a:glow rad="190500"><a:srgbClr val="FF0000"/></a:glow>'
 ]){
  const result=await readPptx(await fixture(effect),{allowPartial:true});
  assert.equal(result.status,'parsed');assert.ok(!result.model.slides[0].nodes.some(n=>n.id==='step'||n.id==='panel'));
  assert.equal(result.diagnostics.filter(d=>d.code==='PPTX_OBJECT_PLACEHOLDER'&&d.reasonKey==='PPTX_SHAPE_EFFECT_UNSUPPORTED').length,2);
 }
});
