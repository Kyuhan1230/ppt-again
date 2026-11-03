import test from 'node:test';
import assert from 'node:assert/strict';
import PptxGenJS from 'pptxgenjs';
import {unzipSync,zipSync,strFromU8,strToU8} from 'fflate';
import {readPptx} from '../src/html-to-pptx/read-pptx.mjs';
import {exportHtml} from '../src/html-to-pptx/export-html.mjs';
import {readInput} from '../src/html-to-pptx/input.mjs';

test('right arrows inherit theme paints and flipped default bent arrows retain their path',async()=>{
 const p=new PptxGenJS(),s=p.addSlide();s.background={color:'FFFFFF'};
 s.addShape(p.ShapeType.rightArrow,{x:1,y:1,w:2,h:1,fill:{color:'123456'},line:{color:'654321',width:1},objectName:'theme-arrow'});
 s.addShape(p.ShapeType.bentArrow,{x:4,y:1,w:1,h:1.2,fill:{color:'C00000'},line:{transparency:100},objectName:'bent-arrow'});
 const parts=unzipSync(new Uint8Array(await p.write({outputType:'uint8array'}))),key='ppt/slides/slide1.xml';let xml=strFromU8(parts[key]);
 xml=xml.replace('<a:solidFill><a:srgbClr val="123456"/></a:solidFill>','').replace(/<a:ln[^>]*>[^]*?<\/a:ln>/,'');
 xml=xml.replace('</p:sp>','<p:style><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef><a:lnRef idx="1"><a:schemeClr val="accent1"><a:shade val="15000"/></a:schemeClr></a:lnRef><a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef></p:style></p:sp>');
 assert.match(xml,/<a:srgbClr val="333333"><a:alpha val="0"\/><\/a:srgbClr>/);
 xml=xml.replace('<a:solidFill><a:srgbClr val="333333"><a:alpha val="0"/></a:srgbClr></a:solidFill>','<a:noFill/>');
 xml=xml.replace('<a:off x="3657600" y="914400"/>','<a:off x="3657600" y="914400"/>').replace(/<a:xfrm>(?=<a:off x="3657600")/,'<a:xfrm flipV="1">');
 parts[key]=strToU8(xml);const parsed=await readPptx(zipSync(parts),{allowPartial:true});assert.equal(parsed.status,'parsed',JSON.stringify(parsed.diagnostics));
 for(const name of ['theme-arrow','bent-arrow'])assert.ok(!parsed.diagnostics.some(d=>d.code==='PPTX_OBJECT_PLACEHOLDER'&&d.objectName===name),name);
 const right=parsed.model.slides[0].nodes.find(n=>n.id==='theme-arrow'),bent=parsed.model.slides[0].nodes.find(n=>n.id==='bent-arrow');
 assert.equal(right?.content.editableVector.shape,'rightArrow');assert.match(right.content.editableVector.fill,/^#[a-f\d]{6}$/);assert.ok(right.content.editableVector.strokeWidth>0);
 assert.equal(right.content.editableVector.fill,'#4472c4');assert.equal(right.content.editableVector.stroke,'#172c51');
 assert.equal(bent?.content.editableVector.shape,'bentArrow');assert.equal(bent.content.editableVector.flipV,true);
 const output=exportHtml(parsed.model);assert.equal(output.status,'generated');assert.match(output.html,/scale\(1 -1\)/);assert.match(output.html,/ A /);
 const input=await readInput({html:output.html});assert.ok(!input.diagnostics.some(d=>d.severity==='error'),JSON.stringify(input.diagnostics));
});
