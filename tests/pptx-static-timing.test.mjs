import test from 'node:test';
import assert from 'node:assert/strict';
import PptxGenJS from 'pptxgenjs';
import {unzipSync,zipSync,strFromU8,strToU8} from 'fflate';
import {readPptx} from '../src/html-to-pptx/read-pptx.mjs';
import {exportHtml} from '../src/html-to-pptx/export-html.mjs';
import {publicDiagnostic} from '../src/html-to-pptx/public-diagnostics.mjs';
import {readInput} from '../src/html-to-pptx/input.mjs';
import {measureInput} from '../src/html-to-pptx/measure.mjs';
import {createBrowserRuntime} from '../src/html-to-pptx/browser-runtime.mjs';

test('static slides retain content and report omitted transitions without an invented on-slide marker',async()=>{
 const ppt=new PptxGenJS(),slide=ppt.addSlide();slide.background={color:'FFFFFF'};
 slide.addText('Visible content',{x:1,y:1,w:4,h:1,fontFace:'Arial',fontSize:18});
 const parts=unzipSync(new Uint8Array(await ppt.write({outputType:'uint8array'}))),key='ppt/slides/slide1.xml';
 const xml=strFromU8(parts[key]);parts[key]=strToU8(xml.replace('</p:sld>','<p:transition><p:fade/></p:transition></p:sld>'));assert.notEqual(xml,strFromU8(parts[key]));
 assert.equal((await readPptx(zipSync(parts))).status,'failed');
 const result=await readPptx(zipSync(parts),{allowPartial:true});assert.equal(result.status,'parsed',JSON.stringify(result.diagnostics));
 assert.match(exportHtml(result.model).html,/Visible content/);assert.doesNotMatch(exportHtml(result.model).html,/data-placeholder-marker/);
 const warning=result.diagnostics.find(d=>d.code==='PPTX_ANIMATION_OMITTED');assert.equal(warning?.severity,'warning');
 assert.match(publicDiagnostic(warning).message,/애니메이션|전환/);
 const html=exportHtml(result.model).html;assert.match(html,/data-animation-omitted/);
 const runtime=await createBrowserRuntime({capture:false});try{
  const measured=await measureInput(await readInput({html}),runtime);
  assert.ok(measured.diagnostics.some(d=>d.code==='PPTX_ANIMATION_OMITTED'&&d.severity==='warning'));
 }finally{await runtime.close();}
});
