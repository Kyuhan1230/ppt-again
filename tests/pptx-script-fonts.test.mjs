import test from 'node:test';
import assert from 'node:assert/strict';
import PptxGenJS from 'pptxgenjs';
import {unzipSync,zipSync,strFromU8,strToU8} from 'fflate';
import {readPptx} from '../src/html-to-pptx/read-pptx.mjs';
import {exportHtml} from '../src/html-to-pptx/export-html.mjs';

async function fixture(){
 const ppt=new PptxGenJS(),slide=ppt.addSlide();slide.background={color:'FFFFFF'};
 slide.addText('한글 ABC',{x:1,y:1,w:5,h:1,fontFace:'Arial',fontSize:20,objectName:'mixed'});
 const parts=unzipSync(new Uint8Array(await ppt.write({outputType:'uint8array'})));
 const patch=(file,fn)=>{const source=strFromU8(parts[file]),changed=fn(source);assert.notEqual(source,changed);parts[file]=strToU8(changed);};
 patch('ppt/slides/slide1.xml',s=>s.replace(/<a:latin\b[^>]*\/>/,'<a:latin typeface="+mn-lt"/>').replace(/<a:ea\b[^>]*\/>/,'<a:ea typeface="Nanum Gothic"/>'));
 patch('ppt/theme/theme1.xml',s=>s.replace(/(<a:minorFont>\s*<a:latin\b[^>]*typeface=")[^"]*/, '$1Arial'));
 return zipSync(parts);
}
test('original import resolves theme Latin and East Asian fonts within mixed text',async()=>{
 const result=await readPptx(await fixture(),{fontPolicy:'original',allowPartial:true});assert.equal(result.status,'parsed',JSON.stringify(result.diagnostics));
 const node=result.model.slides[0].nodes.find(n=>n.id==='mixed');assert.ok(node);assert.equal(node.content.text,'한글 ABC');
 const runs=node.content.runs.filter(r=>r.text.trim());assert.equal(runs.find(r=>r.text.includes('한글')).style.fontFamily,'Nanum Gothic');
 assert.equal(runs.find(r=>r.text.includes('ABC')).style.fontFamily,'Arial');
 assert.match(exportHtml(result.model).html,/한글/);
});
test('an explicit replacement font applies to both script ranges',async()=>{
 const result=await readPptx(await fixture(),{fontPolicy:'pretendard',allowPartial:true});assert.equal(result.status,'parsed');
 const node=result.model.slides[0].nodes.find(n=>n.id==='mixed');assert.ok(node.content.runs.every(r=>r.style.fontFamily==='Pretendard'));
});
