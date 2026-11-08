import test from 'node:test';import assert from 'node:assert/strict';import PptxGenJS from 'pptxgenjs';
import {zipSync,unzipSync,strFromU8,strToU8} from 'fflate';import {readPptx} from '../src/html-to-pptx/read-pptx.mjs';
import {publicDiagnostic} from '../src/html-to-pptx/public-diagnostics.mjs';
import {exportHtml} from '../src/html-to-pptx/export-html.mjs';import {readInput} from '../src/html-to-pptx/input.mjs';
import {measureInput} from '../src/html-to-pptx/measure.mjs';import {createBrowserRuntime} from '../src/html-to-pptx/browser-runtime.mjs';import {exportPptx} from '../src/html-to-pptx/export.mjs';
async function fixture(change){const p=new PptxGenJS(),s=p.addSlide();s.background={color:'FFFFFF'};s.addText('First',{x:1,y:1,w:5,h:2,fontFace:'Arial',fontSize:20,objectName:'content'});const parts=unzipSync(new Uint8Array(await p.write({outputType:'uint8array'}))),key='ppt/slides/slide1.xml',source=strFromU8(parts[key]),updated=change(source);assert.notEqual(updated,source);parts[key]=strToU8(updated);return zipSync(parts);}
const read=b=>readPptx(b,{allowPartial:true,fontPolicy:'malgun-gothic'});const text=r=>r.model.slides[0].nodes.find(n=>n.id==='content')?.content.text;
test('unused custom tab stops do not hide a paragraph without tabs',async()=>{
 const b=await fixture(x=>x.replace('</a:pPr>','<a:tabLst><a:tab pos="266700" algn="l"/></a:tabLst></a:pPr>'));const r=await read(b);
 assert.equal(r.status,'parsed');assert.equal(text(r),'First');assert.ok(!r.diagnostics.some(d=>d.objectName==='content'&&d.code==='PPTX_OBJECT_PLACEHOLDER'));
});
test('tabbed text is retained with an explicit placement approximation',async()=>{
 const b=await fixture(x=>x.replace('</a:pPr>','<a:tabLst><a:tab pos="266700" algn="l"/></a:tabLst></a:pPr>').replace('>First</a:t>','>First\tSecond</a:t>'));
 const strict=await readPptx(b,{fontPolicy:'malgun-gothic'});assert.equal(strict.status,'failed');const r=await read(b);assert.equal(text(r),'First\tSecond');assert.ok(r.diagnostics.some(d=>d.reasonKey==='PPTX_TEXT_TAB_APPROXIMATED'));
});
test('highlighted text retains its color in HTML and PowerPoint',async()=>{
 const b=await fixture(x=>x.replace('</a:rPr>','<a:highlight><a:srgbClr val="FFFF00"/></a:highlight></a:rPr>'));
 const r=await readPptx(b,{fontPolicy:'malgun-gothic'});assert.equal(r.status,'parsed',JSON.stringify(r.diagnostics));assert.equal(text(r),'First');
 assert.equal(r.model.slides[0].nodes.find(n=>n.id==='content').content.runs[0].style.highlight,'#ffff00');
 const html=exportHtml(r.model);assert.match(html.html,/data-pptx-highlight/);assert.match(html.html,/background-color:#ffff00/);
 const runtime=await createBrowserRuntime({capture:false});try{
  const measured=await measureInput(await readInput({html:html.html}),runtime),ppt=await exportPptx(measured);
  assert.equal(ppt.status,'generated',JSON.stringify(ppt.diagnostics));assert.match(strFromU8(unzipSync(ppt.pptxBytes)['ppt/slides/slide1.xml']),/<a:highlight><a:srgbClr val="FFFF00"\/>/);
 }finally{await runtime.close();}
});
test('a highlighted table-cell run keeps its color on a PPT round trip',async()=>{
 const ppt=new PptxGenJS(),slide=ppt.addSlide();slide.background={color:'FFFFFF'};
 slide.addTable([['Marked']],{x:1,y:1,w:3,h:1,fontFace:'Arial',fontSize:20,objectName:'table'});
 const parts=unzipSync(new Uint8Array(await ppt.write({outputType:'uint8array'}))),key='ppt/slides/slide1.xml',source=strFromU8(parts[key]);
 const changed=source.replace('</a:rPr>','<a:highlight><a:srgbClr val="FFFF00"/></a:highlight></a:rPr>');assert.notEqual(source,changed);parts[key]=strToU8(changed);
 const result=await readPptx(zipSync(parts));assert.equal(result.status,'parsed',JSON.stringify(result.diagnostics));
 const html=exportHtml(result.model).html;assert.match(html,/data-pptx-highlight/);
 const runtime=await createBrowserRuntime({capture:false});try{
  const measured=await measureInput(await readInput({html}),runtime),out=await exportPptx(measured);
  assert.equal(out.status,'generated',JSON.stringify(out.diagnostics));assert.match(strFromU8(unzipSync(out.pptxBytes)['ppt/slides/slide1.xml']),/<a:highlight><a:srgbClr val="FFFF00"\/>/);
 }finally{await runtime.close();}
});
test('consecutive Arabic-numbered paragraphs keep their text and native numbering',async()=>{
 const b=await fixture(x=>x.replace(/<a:p>[\s\S]*?<\/a:p>/,p=>{const first=p.replace('<a:buNone/>','').replace('</a:pPr>','<a:buAutoNum type="arabicPeriod" startAt="2"/></a:pPr>'),second=p.replace('>First</a:t>','>Second</a:t>').replace('<a:buNone/>','').replace('</a:pPr>','<a:buAutoNum type="arabicPeriod" startAt="2"/></a:pPr>');return first+second;}));
 const r=await readPptx(b,{fontPolicy:'malgun-gothic'});assert.equal(r.status,'parsed',JSON.stringify(r.diagnostics));assert.equal(text(r),'2. First\n3. Second');
 assert.deepEqual(r.model.slides[0].nodes.find(n=>n.id==='content').content.paragraphs.map(p=>p.bullet.start),[2,3]);
});
test('a changed native numbering start begins a new PowerPoint list',async()=>{
 const b=await fixture(x=>x.replace(/<a:p>[\s\S]*?<\/a:p>/,p=>{const first=p.replace('<a:buNone/>','').replace('</a:pPr>','<a:buAutoNum type="arabicPeriod" startAt="3"/></a:pPr>'),second=p.replace('>First</a:t>','>Second</a:t>').replace('<a:buNone/>','').replace('</a:pPr>','<a:buAutoNum type="arabicPeriod"/></a:pPr>');return first+second;}));
 const r=await readPptx(b,{fontPolicy:'malgun-gothic'});assert.equal(r.status,'parsed',JSON.stringify(r.diagnostics));assert.equal(text(r),'3. First\n1. Second');
});
test('text approximation reasons explain the actual limitation publicly',()=>{
 for(const [reason,word] of [['PPTX_AUTONUM_APPROXIMATED','번호'],['PPTX_TEXT_TAB_APPROXIMATED','탭'],['PPTX_TEXT_HIGHLIGHT_APPROXIMATED','강조']]){
  const d=publicDiagnostic({code:'PARAGRAPH_SPACING_APPROXIMATED',severity:'warning',reasonKey:reason});assert.equal(d.reasonKey,reason);assert.match(d.message,new RegExp(word));assert.match(d.reason,new RegExp(word));
 }
});
test('an explicit tab node remains in the partial text output',async()=>{
 const b=await fixture(x=>x.replace('</a:r>','</a:r><a:tab/>'));
 const r=await read(b);assert.equal(text(r),'First\t');assert.ok(r.diagnostics.some(d=>d.reasonKey==='PPTX_TEXT_TAB_APPROXIMATED'));
});
