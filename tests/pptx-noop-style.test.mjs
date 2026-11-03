import test from 'node:test';
import assert from 'node:assert/strict';
import PptxGenJS from 'pptxgenjs';
import {unzipSync,zipSync,strFromU8,strToU8} from 'fflate';
import {readPptx} from '../src/html-to-pptx/read-pptx.mjs';
import {exportHtml} from '../src/html-to-pptx/export-html.mjs';
const text='Synthetic plain text / 기본 문구';
async function fixture(mutate){const p=new PptxGenJS();p.layout='LAYOUT_WIDE';const s=p.addSlide();s.background={color:'FFFFFF'};s.addText(text,{x:1,y:1,w:7,h:1,fontFace:'Arial',fontSize:24,objectName:'simple-text'});const parts=unzipSync(new Uint8Array(await p.write({outputType:'uint8array'})));const source=strFromU8(parts['ppt/slides/slide1.xml']),modified=mutate(source);assert.notEqual(modified,source,'fixture mutation must apply');parts['ppt/slides/slide1.xml']=strToU8(modified);return zipSync(parts);}
const flags=x=>x.replace(/<a:rPr([^>]*)>/,(_,attrs)=>'<a:rPr'+attrs.replace(/\s+(?:strike|cap|baseline|u)="[^"]*"/g,'')+' strike="noStrike" cap="none" baseline="0" u="none">');
for(const [name,mutate] of [
 ['empty effect list',x=>x.replace(/<a:rPr([^>]*)>/,'<a:rPr$1><a:effectLst/>')],
 ['empty effect list and default flags',x=>flags(x).replace(/<a:rPr([^>]*)>/,'<a:rPr$1><a:effectLst/>')],
 ['empty tab list',x=>x.replace(/<a:pPr([^>]*)>/,'<a:pPr$1><a:tabLst/>')],
 ['unused tab stops',x=>x.replace(/<a:pPr([^>]*)>/,'<a:pPr$1><a:tabLst><a:tab pos="914400" algn="l"/></a:tabLst>')]
])test(`${name} preserves text in strict and partial imports`,async()=>{const b=await fixture(mutate);for(const allowPartial of [false,true]){const r=await readPptx(b,{allowPartial});assert.equal(r.status,'parsed',JSON.stringify(r.diagnostics));assert.equal(r.model.slides[0].nodes.length,1);assert.equal(r.model.slides[0].nodes[0].content.text,text);assert.ok(!r.diagnostics.some(d=>d.code==='PPTX_OBJECT_PLACEHOLDER'||d.reasonKey==='PPTX_TEXT_STYLE_UNSUPPORTED'));assert.match(exportHtml(r.model).html,/Synthetic plain text/);}});
for(const [name,mutate] of [
 ['nonempty shadow effect',x=>x.replace(/<a:rPr([^>]*)>/,'<a:rPr$1><a:effectLst><a:outerShdw blurRad="0" dist="12700" dir="0"><a:srgbClr val="000000"/></a:outerShdw></a:effectLst>')]
])test(`${name} remains unsupported rather than silently removed`,async()=>{const b=await fixture(mutate);const strict=await readPptx(b);assert.equal(strict.status,'failed');assert.ok(strict.diagnostics.some(d=>d.reasonKey==='PPTX_TEXT_STYLE_UNSUPPORTED'&&d.severity==='error'));const partial=await readPptx(b,{allowPartial:true});assert.equal(partial.status,'parsed');assert.ok(partial.diagnostics.some(d=>d.code==='PPTX_OBJECT_PLACEHOLDER'));assert.match(exportHtml(partial.model).html,/변환하지 못한 객체/);});
