import test from 'node:test';
import assert from 'node:assert/strict';
import PptxGenJS from 'pptxgenjs';
import {unzipSync,zipSync,strFromU8,strToU8} from 'fflate';
import {readPptx} from '../src/html-to-pptx/read-pptx.mjs';

test('line inherits theme paint while retaining its explicit width',async()=>{
  const ppt=new PptxGenJS(),slide=ppt.addSlide();slide.background={color:'FFFFFF'};
  slide.addShape(ppt.ShapeType.line,{x:1,y:1,w:2,h:0,line:{color:'112233',width:1.5},objectName:'threshold'});
  const parts=unzipSync(new Uint8Array(await ppt.write({outputType:'uint8array'})));
  const xml=strFromU8(parts['ppt/slides/slide1.xml']);
  const line=xml.match(/<p:sp>[\s\S]*?<\/p:sp>/)[0];
  const changed=line.replace('<a:solidFill><a:srgbClr val="112233"/></a:solidFill>','').replace('</p:spPr>','</p:spPr><p:style><a:lnRef idx="1"><a:schemeClr val="dk1"/></a:lnRef><a:fillRef idx="0"><a:schemeClr val="dk1"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="dk1"/></a:effectRef></p:style>');
  assert.notEqual(changed,line);
  parts['ppt/slides/slide1.xml']=strToU8(xml.replace(line,changed));
  const result=await readPptx(zipSync(parts),{allowPartial:true});
  assert.equal(result.status,'parsed',JSON.stringify(result.diagnostics));
  const node=result.model.slides[0].nodes.find(n=>n.id==='threshold');
  assert.equal(node?.kind,'line');
  assert.match(node.style.stroke,/^#000000$/i);
  assert.equal(node.style.strokeWidth,2);
});

test('an explicit line color does not discard its theme-inherited width',async()=>{
  const ppt=new PptxGenJS(),slide=ppt.addSlide();slide.background={color:'FFFFFF'};
  slide.addShape(ppt.ShapeType.line,{x:1,y:1,w:2,h:0,line:{color:'112233',width:1.5},objectName:'explicit-paint'});
  const parts=unzipSync(new Uint8Array(await ppt.write({outputType:'uint8array'}))),slideFile='ppt/slides/slide1.xml';
  const xml=strFromU8(parts[slideFile]),shape=xml.match(/<p:sp>[\s\S]*?<\/p:sp>/)[0];
  const changed=shape.replace(/<a:ln\s+w="[^"]*"/,'<a:ln').replace('</p:spPr>','</p:spPr><p:style><a:lnRef idx="1"><a:schemeClr val="dk1"/></a:lnRef><a:fillRef idx="0"><a:schemeClr val="dk1"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="dk1"/></a:effectRef></p:style>');
  assert.notEqual(changed,shape);parts[slideFile]=strToU8(xml.replace(shape,changed));
  const themeFile='ppt/theme/theme1.xml',theme=strFromU8(parts[themeFile]);
  const updatedTheme=theme.replace(/(<a:lnStyleLst>\s*<a:ln\s+w=")[^"]+/, '$112700');assert.notEqual(updatedTheme,theme);parts[themeFile]=strToU8(updatedTheme);
  const result=await readPptx(zipSync(parts),{allowPartial:true});assert.equal(result.status,'parsed',JSON.stringify(result.diagnostics));
  const node=result.model.slides[0].nodes.find(n=>n.id==='explicit-paint');assert.equal(node.style.stroke,'#112233');assert.ok(Math.abs(node.style.strokeWidth-4/3)<1e-8);
});
