import test from 'node:test';
import assert from 'node:assert/strict';
import PptxGenJS from 'pptxgenjs';
import {zipSync,unzipSync,strFromU8,strToU8} from 'fflate';
import {readPptx} from '../src/html-to-pptx/read-pptx.mjs';

async function document(line){
 const p=new PptxGenJS(),s=p.addSlide();s.background={color:'FFFFFF'};
 s.addText('Visible title',{x:1,y:1,w:6,h:1,fontFace:'Arial',fontSize:20,objectName:'title'});
 const parts=unzipSync(new Uint8Array(await p.write({outputType:'uint8array'}))),key='ppt/slides/slide1.xml',xml=strFromU8(parts[key]);
 const updated=xml.replace('</a:rPr>',`${line}</a:rPr>`);assert.notEqual(updated,xml);parts[key]=strToU8(updated);return zipSync(parts);
}
for(const [label,line] of [
 ['zero width','<a:ln w="0"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln>'],
 ['fully transparent','<a:ln w="12700"><a:solidFill><a:srgbClr val="FF0000"><a:alpha val="0"/></a:srgbClr></a:solidFill></a:ln>']
])test(`${label} text outline is visually absent and preserves the text`,async()=>{
 const result=await readPptx(await document(line),{allowPartial:true});assert.equal(result.status,'parsed',JSON.stringify(result.diagnostics));
 const node=result.model.slides[0].nodes.find(n=>n.id==='title');assert.equal(node?.content.text,'Visible title');assert.ok(!node.style.outline);
 assert.ok(!result.diagnostics.some(d=>d.objectName==='title'&&d.code==='PPTX_OBJECT_PLACEHOLDER'));
});
test('partly transparent visible outline remains a representation limit',async()=>{
 const result=await readPptx(await document('<a:ln w="12700"><a:solidFill><a:srgbClr val="FF0000"><a:alpha val="50000"/></a:srgbClr></a:solidFill></a:ln>'),{allowPartial:true});
 assert.ok(result.diagnostics.some(d=>d.objectName==='title'&&d.code==='PPTX_OBJECT_PLACEHOLDER'));
});
