import test from 'node:test';
import assert from 'node:assert/strict';
import PptxGenJS from 'pptxgenjs';
import {unzipSync,zipSync,strFromU8,strToU8} from 'fflate';
import {readPptx} from '../src/html-to-pptx/read-pptx.mjs';

async function fixture(color){
 const ppt=new PptxGenJS(),slide=ppt.addSlide();slide.background={color:'FFFFFF'};
 slide.addShape(ppt.ShapeType.rect,{x:1,y:1,w:2,h:1,fill:{color:'123456'},line:{transparency:100},objectName:'preset-shape'});
 const parts=unzipSync(new Uint8Array(await ppt.write({outputType:'uint8array'})));
 const key='ppt/slides/slide1.xml',xml=strFromU8(parts[key]);
 const changed=xml.replace('<a:srgbClr val="123456"/>',`<a:prstClr val="${color}"><a:lumMod val="75000"/><a:lumOff val="25000"/></a:prstClr>`);
 assert.notEqual(xml,changed);parts[key]=strToU8(changed);return zipSync(parts);
}
test('PowerPoint preset black with luminance resolves to explicit RGB',async()=>{
 const result=await readPptx(await fixture('black'));assert.equal(result.status,'parsed',JSON.stringify(result.diagnostics));
 assert.equal(result.model.slides[0].nodes.find(n=>n.id==='preset-shape').style.fill.toLowerCase(),'#404040');
});
test('unknown preset colors are diagnosed instead of guessed',async()=>{
 const result=await readPptx(await fixture('unknownPreset'));assert.equal(result.status,'failed');
 assert.ok(result.diagnostics.some(d=>d.reasonKey==='PPTX_THEME_COLOR_UNSUPPORTED'));
});
