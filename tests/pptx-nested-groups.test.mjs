import test from 'node:test';
import assert from 'node:assert/strict';
import PptxGenJS from 'pptxgenjs';
import {unzipSync,zipSync,strFromU8,strToU8} from 'fflate';
import {readPptx} from '../src/html-to-pptx/read-pptx.mjs';
import {exportHtml} from '../src/html-to-pptx/export-html.mjs';
import {publicDiagnostic} from '../src/html-to-pptx/public-diagnostics.mjs';

test('nested translated groups preserve the child and compose positions',async()=>{
  const ppt=new PptxGenJS(),slide=ppt.addSlide();
  slide.background={color:'FFFFFF'};
  slide.addShape(ppt.ShapeType.rect,{x:1,y:1,w:1,h:1,fill:{color:'123456'},line:{transparency:100},objectName:'child'});
  const parts=unzipSync(new Uint8Array(await ppt.write({outputType:'uint8array'})));
  let xml=strFromU8(parts['ppt/slides/slide1.xml']);
  const shape=xml.match(/<p:sp>[\s\S]*?<\/p:sp>/)[0];
  const group=(id,name,x,y,content)=>`<p:grpSp><p:nvGrpSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="${x*914400}" y="${y*914400}"/><a:ext cx="914400" cy="914400"/><a:chOff x="914400" y="914400"/><a:chExt cx="914400" cy="914400"/></a:xfrm></p:grpSpPr>${content}</p:grpSp>`;
  xml=xml.replace(shape,group(900,'outer',3,2,group(901,'inner',2,3,shape)));
  parts['ppt/slides/slide1.xml']=strToU8(xml);
  const result=await readPptx(zipSync(parts),{allowPartial:true});
  assert.equal(result.status,'parsed',JSON.stringify(result.diagnostics));
  const nodes=result.model.slides[0].nodes;
  const child=nodes.find(n=>n.id==='child');
  assert.ok(child,'nested child must survive partial conversion');
  assert.deepEqual(child.boxCssPx,{x:384,y:384,w:96,h:96});
  assert.equal(child.parentId,'inner');
  assert.equal(nodes.find(n=>n.id==='inner').parentId,'outer');
  assert.equal(result.diagnostics.some(d=>d.reasonKey==='PPTX_GROUP_TRANSFORM_UNSUPPORTED'),false);
  const html=exportHtml(result.model);assert.equal(html.status,'generated',JSON.stringify(html.diagnostics));
  assert.match(html.html,/data-source-id="child"/);
});

test('excessively nested groups are bounded and reported instead of recursing indefinitely',async()=>{
  const ppt=new PptxGenJS(),slide=ppt.addSlide();
  slide.background={color:'FFFFFF'};
  slide.addShape(ppt.ShapeType.rect,{x:1,y:1,w:1,h:1,fill:{color:'123456'},line:{transparency:100},objectName:'deep-child'});
  const parts=unzipSync(new Uint8Array(await ppt.write({outputType:'uint8array'})));
  const xml=strFromU8(parts['ppt/slides/slide1.xml']),shape=xml.match(/<p:sp>[\s\S]*?<\/p:sp>/)[0];
  let wrapped=shape;
  for(let i=0;i<33;i++)wrapped=`<p:grpSp><p:nvGrpSpPr><p:cNvPr id="${900+i}" name="group-${i}"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="914400" cy="914400"/><a:chOff x="914400" y="914400"/><a:chExt cx="914400" cy="914400"/></a:xfrm></p:grpSpPr>${wrapped}</p:grpSp>`;
  parts['ppt/slides/slide1.xml']=strToU8(xml.replace(shape,wrapped));
  const result=await readPptx(zipSync(parts),{allowPartial:true});
  assert.equal(result.status,'parsed',JSON.stringify(result.diagnostics));
  assert.equal(result.model.slides[0].nodes.some(n=>n.id==='deep-child'),false);
  assert.ok(result.diagnostics.some(d=>d.code==='PPTX_OBJECT_PLACEHOLDER'&&d.message.includes('중첩 깊이')));
});

test('partial conversion preserves scaled group geometry and reports style approximation',async()=>{
  const ppt=new PptxGenJS(),slide=ppt.addSlide();
  slide.background={color:'FFFFFF'};
  slide.addShape(ppt.ShapeType.rect,{x:1,y:1,w:1,h:1,fill:{color:'123456'},line:{transparency:100},objectName:'scaled-child'});
  const parts=unzipSync(new Uint8Array(await ppt.write({outputType:'uint8array'})));
  const xml=strFromU8(parts['ppt/slides/slide1.xml']),shape=xml.match(/<p:sp>[\s\S]*?<\/p:sp>/)[0];
  const group=`<p:grpSp><p:nvGrpSpPr><p:cNvPr id="900" name="scaled-group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="2743200" y="1828800"/><a:ext cx="1828800" cy="914400"/><a:chOff x="914400" y="914400"/><a:chExt cx="914400" cy="914400"/></a:xfrm></p:grpSpPr>${shape}</p:grpSp>`;
  parts['ppt/slides/slide1.xml']=strToU8(xml.replace(shape,group));
  const bytes=zipSync(parts);
  assert.equal((await readPptx(bytes)).status,'failed','strict mode must continue to reject approximated groups');
  const result=await readPptx(bytes,{allowPartial:true});
  assert.equal(result.status,'parsed',JSON.stringify(result.diagnostics));
  const child=result.model.slides[0].nodes.find(n=>n.id==='scaled-child');
  assert.ok(child,'scaled child must not be replaced by a placeholder');
  assert.deepEqual(child.boxCssPx,{x:288,y:192,w:192,h:96});
  const warning=result.diagnostics.find(d=>d.code==='PPTX_GROUP_SCALE_APPROXIMATED'&&d.severity==='warning');
  assert.ok(warning);
  const publicWarning=publicDiagnostic(warning,result.model.slides);
  assert.equal(publicWarning.reasonKey,'PPTX_GROUP_SCALE_APPROXIMATED');
  assert.match(publicWarning.message,/글자 크기와 선 굵기/);
  assert.equal(result.diagnostics.some(d=>d.code==='PPTX_OBJECT_PLACEHOLDER'&&d.objectName==='scaled-group'),false);
});

test('a scaled group transforms table rows and columns as well as its position',async()=>{
  const ppt=new PptxGenJS(),slide=ppt.addSlide();slide.background={color:'FFFFFF'};
  slide.addTable([['A','B']],{x:.25,y:.25,w:3,h:1,fontFace:'Arial',fontSize:16,objectName:'scaled-table'});
  const parts=unzipSync(new Uint8Array(await ppt.write({outputType:'uint8array'}))),file='ppt/slides/slide1.xml';
  const xml=strFromU8(parts[file]),table=xml.match(/<p:graphicFrame>[\s\S]*?<\/p:graphicFrame>/)[0];
  const group=`<p:grpSp><p:nvGrpSpPr><p:cNvPr id="900" name="scaled-group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="1371600"/><a:chOff x="0" y="0"/><a:chExt cx="914400" cy="914400"/></a:xfrm></p:grpSpPr>${table}</p:grpSp>`;
  parts[file]=strToU8(xml.replace(table,group));const result=await readPptx(zipSync(parts),{allowPartial:true});
  assert.equal(result.status,'parsed',JSON.stringify(result.diagnostics));const node=result.model.slides[0].nodes.find(n=>n.id==='scaled-table');assert.ok(node);
  assert.deepEqual(node.boxCssPx,{x:48,y:36,w:576,h:144});assert.deepEqual(node.content.columnWidthsCssPx,[288,288]);assert.deepEqual(node.content.rowHeightsCssPx,[144]);
  assert.equal(node.content.cells[0].style.fontSize,64/3,'font appearance remains approximated and is reported');
  assert.ok(result.diagnostics.some(d=>d.reasonKey==='PPTX_GROUP_SCALE_APPROXIMATED'));
});
