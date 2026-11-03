import test from 'node:test';
import assert from 'node:assert/strict';
import PptxGenJS from 'pptxgenjs';
import {unzipSync,zipSync,strFromU8,strToU8} from 'fflate';
import {readPptx} from '../src/html-to-pptx/read-pptx.mjs';
import {exportHtml} from '../src/html-to-pptx/export-html.mjs';

test('import retains superscript and each paragraph alignment',async()=>{
  const p=new PptxGenJS(),s=p.addSlide();s.background={color:'FFFFFF'};
  s.addText('Main*\nNote',{x:1,y:1,w:6,h:1,fontFace:'Arial',fontSize:16,objectName:'description'});
  const parts=unzipSync(new Uint8Array(await p.write({outputType:'uint8array'}))),file='ppt/slides/slide1.xml';let xml=strFromU8(parts[file]);
  xml=xml.replace('<a:rPr ','<a:rPr baseline="30000" ');
  let n=0;xml=xml.replace(/<a:pPr([^>]*)>/g,(all,attrs)=>++n===2?`<a:pPr${attrs} algn="r">`:all);parts[file]=strToU8(xml);
  assert.equal((await readPptx(zipSync(parts))).status,'parsed');
  const result=await readPptx(zipSync(parts),{allowPartial:true});assert.equal(result.status,'parsed',JSON.stringify(result.diagnostics));
  const node=result.model.slides[0].nodes.find(x=>x.id==='description');assert.equal(node?.content.text,'Main*\nNote');
  assert.equal(node.content.runs.find(r=>r.text.includes('Main')).style.baseline,30000);
  assert.deepEqual(node.content.paragraphs.map(p=>p.align),['left','right']);
  assert.ok(!result.diagnostics.some(d=>d.reasonKey==='PPTX_PARAGRAPH_ALIGNMENT_APPROXIMATED'));
  assert.match(exportHtml(result.model).html,/text-align:right;/);
  assert.match(exportHtml(result.model).html,/Main\*/);assert.match(exportHtml(result.model).html,/Note/);
});

test('explicit no-wrap text preserves line breaks without wrapping long lines',async()=>{
  const p=new PptxGenJS(),s=p.addSlide();s.background={color:'FFFFFF'};
  s.addText('Unwrapped long title\nSecond line',{x:1,y:1,w:1,h:1,fontFace:'Arial',fontSize:16,objectName:'unwrapped'});
  const parts=unzipSync(new Uint8Array(await p.write({outputType:'uint8array'}))),file='ppt/slides/slide1.xml';
  parts[file]=strToU8(strFromU8(parts[file]).replace(/<a:bodyPr[^>]*>/,tag=>tag.replace(/\s+wrap="[^"]*"/,'').replace('<a:bodyPr','<a:bodyPr wrap="none"')));
  const result=await readPptx(zipSync(parts));assert.equal(result.status,'parsed',JSON.stringify(result.diagnostics));
  const node=result.model.slides[0].nodes.find(x=>x.id==='unwrapped');assert.equal(node.style.whiteSpace,'pre');assert.equal(node.content.text,'Unwrapped long title\nSecond line');
  assert.match(exportHtml(result.model).html,/white-space:pre;/);
});
