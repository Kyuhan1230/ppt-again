import test from 'node:test';
import assert from 'node:assert/strict';
import PptxGenJS from 'pptxgenjs';
import {unzipSync,zipSync,strFromU8,strToU8} from 'fflate';
import {readPptx} from '../src/html-to-pptx/read-pptx.mjs';
import {exportHtml} from '../src/html-to-pptx/export-html.mjs';

test('theme-colored text boxes preserve both shape and text in HTML',async()=>{
  const p=new PptxGenJS();const s=p.addSlide();s.background={color:'FFFFFF'};
  s.addText('Flow step',{x:1,y:1,w:2,h:1,fontFace:'Arial',fontSize:16,shape:p.ShapeType.roundRect,fill:{color:'123456'},line:{color:'654321',width:1},objectName:'flow-step'});
  const parts=unzipSync(new Uint8Array(await p.write({outputType:'uint8array'})));
  const slide='ppt/slides/slide1.xml';let xml=strFromU8(parts[slide]);
  xml=xml.replace('<a:srgbClr val="123456"/>','<a:schemeClr val="bg1"><a:lumMod val="75000"/></a:schemeClr>').replace('<a:srgbClr val="654321"/>','<a:sysClr val="windowText" lastClr="000000"/>');
  parts[slide]=strToU8(xml);
  const result=await readPptx(zipSync(parts),{allowPartial:true});
  assert.equal(result.status,'parsed',JSON.stringify(result.diagnostics));
  assert.ok(!result.diagnostics.some(d=>d.nodeId==='flow-step'&&d.code==='PPTX_OBJECT_PLACEHOLDER'),JSON.stringify(result.diagnostics));
  const html=exportHtml(result.model).html;
  assert.match(html,/Flow step/);
  assert.match(html,/background:#bfbfbf/i);
  assert.match(html,/outline:[^;]*#000000/i);
});

test('simple inherited theme fill and line styles are preserved',async()=>{
  const p=new PptxGenJS();const s=p.addSlide();s.background={color:'FFFFFF'};
  s.addShape(p.ShapeType.rect,{x:1,y:1,w:2,h:1,fill:{color:'123456'},line:{color:'654321',width:1},objectName:'inherited'});
  const parts=unzipSync(new Uint8Array(await p.write({outputType:'uint8array'})));
  const slide='ppt/slides/slide1.xml';let xml=strFromU8(parts[slide]);
  xml=xml.replace('<a:solidFill><a:srgbClr val="123456"/></a:solidFill>','').replace(/<a:ln[^>]*>[\s\S]*?<\/a:ln>/,'');
  xml=xml.replace('</p:sp>','<p:style><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef><a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef><a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef></p:style></p:sp>');
  parts[slide]=strToU8(xml);
  const result=await readPptx(zipSync(parts),{allowPartial:true});
  assert.equal(result.status,'parsed',JSON.stringify(result.diagnostics));
  assert.ok(!result.diagnostics.some(d=>d.nodeId==='inherited'&&d.code==='PPTX_OBJECT_PLACEHOLDER'),JSON.stringify(result.diagnostics));
  const node=result.model.slides[0].nodes.find(n=>n.id==='inherited');
  assert.ok(node);assert.notEqual(node.style.fill,'transparent');assert.ok(node.style.strokeWidth>0);assert.equal(node.style.fill,node.style.stroke);
});

test('compound lines, non-flat caps and non-center alignment remain unsupported',async()=>{
  const p=new PptxGenJS();const s=p.addSlide();s.background={color:'FFFFFF'};
  s.addShape(p.ShapeType.rect,{x:1,y:1,w:2,h:1,fill:{color:'123456'},line:{color:'654321',width:1},objectName:'styled-line'});
  const base=unzipSync(new Uint8Array(await p.write({outputType:'uint8array'}))),slide='ppt/slides/slide1.xml';
  let xml=strFromU8(base[slide]);xml=xml.replace(/<a:ln[^>]*>[\s\S]*?<\/a:ln>/,'').replace('</p:sp>','<p:style><a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef></p:style></p:sp>');base[slide]=strToU8(xml);
  for(const [from,to] of [['cmpd="sng"','cmpd="dbl"'],['cap="flat"','cap="rnd"'],['algn="ctr"','algn="in"']]){
    const parts={...base},theme='ppt/theme/theme1.xml';parts[theme]=strToU8(strFromU8(parts[theme]).replace(from,to));
    const strict=await readPptx(zipSync(parts));assert.equal(strict.status,'failed',to);
    const partial=await readPptx(zipSync(parts),{allowPartial:true});assert.equal(partial.status,'parsed');
    assert.ok(partial.diagnostics.some(d=>d.nodeId==='styled-line'&&d.reasonKey==='PPTX_INHERITED_SHAPE_STYLE_UNSUPPORTED'),to);
  }
});

test('malformed explicit RGB is not accepted as a partial conversion',async()=>{
  const p=new PptxGenJS();const s=p.addSlide();s.background={color:'FFFFFF'};
  s.addShape(p.ShapeType.rect,{x:1,y:1,w:2,h:1,fill:{color:'123456'}});
  const parts=unzipSync(new Uint8Array(await p.write({outputType:'uint8array'})));
  const slide='ppt/slides/slide1.xml';parts[slide]=strToU8(strFromU8(parts[slide]).replace('val="123456"','val="not-rgb"'));
  const result=await readPptx(zipSync(parts),{allowPartial:true});
  assert.equal(result.status,'failed');assert.ok(result.diagnostics.some(d=>d.severity==='error'));
});
