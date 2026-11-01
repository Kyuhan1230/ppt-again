import test from 'node:test';
import assert from 'node:assert/strict';
import {unzipSync,strFromU8} from 'fflate';
import {readInput} from '../src/html-to-pptx/input.mjs';
import {measureInput} from '../src/html-to-pptx/measure.mjs';
import {createBrowserRuntime} from '../src/html-to-pptx/browser-runtime.mjs';
import {exportPptx} from '../src/html-to-pptx/export.mjs';
import {installedNativeAliases} from '../src/html-to-pptx/font-profiles.mjs';

test('unavailable text font uses embedded Pretendard while an available font stays unchanged',async()=>{
 const html='<section data-slide style="width:800px;height:600px"><p style="font:24px Arial">Available</p><p style="font:24px NoSuchPptAgainFont">한글 대체 글꼴</p></section>';
 const runtime=await createBrowserRuntime({capture:false});try{const measured=await measureInput(await readInput({html}),runtime);assert.equal(measured.diagnostics.some(d=>d.severity==='error'),false,JSON.stringify(measured.diagnostics));assert.ok(measured.diagnostics.some(d=>d.code==='FONT_SUBSTITUTED'&&d.severity==='warning'));const nodes=measured.model.slides[0].nodes.filter(n=>n.kind==='text');assert.equal(nodes.find(n=>n.content.text==='Available').style.fontFamily,'Arial');assert.equal(nodes.find(n=>n.content.text==='한글 대체 글꼴').style.fontFamily,'PPTAgainFallback');const output=await exportPptx(measured,{fontEmbedding:'auto'});assert.equal(output.status,'generated',JSON.stringify(output.diagnostics));const zip=unzipSync(output.pptxBytes);assert.ok(Object.keys(zip).some(k=>k.startsWith('ppt/fonts/')));assert.match(strFromU8(zip['ppt/slides/slide1.xml']),/Pretendard/);}finally{await runtime.close();}
});

test('fallback includes a bold face when needed and keeps bold text editable',async()=>{
 const html='<section data-slide style="width:800px;height:600px"><p style="font:700 24px NoSuchPptAgainFont">굵은 글자</p></section>';
 const runtime=await createBrowserRuntime({capture:false});try{const measured=await measureInput(await readInput({html}),runtime);assert.equal(measured.diagnostics.some(d=>d.severity==='error'),false,JSON.stringify(measured.diagnostics));assert.deepEqual(measured.model.fonts.map(f=>f.weight),[700]);const result=await exportPptx(measured,{fontEmbedding:'auto'});assert.equal(result.status,'generated',JSON.stringify(result.diagnostics));const zip=unzipSync(result.pptxBytes);assert.match(strFromU8(zip['ppt/presentation.xml']),/<p:bold\b/);assert.match(strFromU8(zip['ppt/slides/slide1.xml']),/b="1"/);}finally{await runtime.close();}
});

test('fallback does not claim success when Pretendard lacks a character',async()=>{
 const html='<section data-slide style="width:800px;height:600px"><p style="font:24px NoSuchPptAgainFont">Missing &#x10FFFF;</p></section>';
 const runtime=await createBrowserRuntime({capture:false});try{const measured=await measureInput(await readInput({html}),runtime);const result=await exportPptx(measured,{fontEmbedding:'auto'});assert.equal(result.status,'failed');assert.ok(result.diagnostics.some(d=>d.code==='FONT_GLYPH_MISSING'));}finally{await runtime.close();}
});

test('verified installed font aliases are preserved instead of being replaced',async t=>{
 if(!installedNativeAliases().some(a=>a.family==='Pretendard')){t.skip('Verified Pretendard face is not installed');return;}
 const html='<section data-slide style="width:800px;height:600px"><p style="font:24px Pretendard">기존 글꼴</p></section>';
 const runtime=await createBrowserRuntime({capture:false});try{const measured=await measureInput(await readInput({html}),runtime);assert.equal(measured.diagnostics.some(d=>d.code==='FONT_SUBSTITUTED'),false,JSON.stringify(measured.diagnostics));assert.equal(measured.model.slides[0].nodes.find(n=>n.kind==='text').style.fontFamily,'Pretendard');}finally{await runtime.close();}
});
