import test from 'node:test';
import assert from 'node:assert/strict';
import PptxGenJS from 'pptxgenjs';
import {unzipSync,zipSync,strFromU8,strToU8} from 'fflate';
import {DOMParser} from '@xmldom/xmldom';
import {readPptx} from '../src/html-to-pptx/read-pptx.mjs';
import {exportHtml} from '../src/html-to-pptx/export-html.mjs';
import {readInput} from '../src/html-to-pptx/input.mjs';
import {measureInput} from '../src/html-to-pptx/measure.mjs';
import {createBrowserRuntime} from '../src/html-to-pptx/browser-runtime.mjs';
import {exportPptx} from '../src/html-to-pptx/export.mjs';
test('slide fade timing survives PPTX to HTML and back to editable PPTX',async()=>{
 const ppt=new PptxGenJS(),s=ppt.addSlide();s.addText('Transition test',{x:1,y:1,w:3,h:1,fontFace:'Arial',fontSize:18});
 const parts=unzipSync(new Uint8Array(await ppt.write({outputType:'uint8array'}))),key='ppt/slides/slide1.xml';
 parts[key]=strToU8(strFromU8(parts[key]).replace('</p:sld>','<p:transition xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" p14:dur="1250" advClick="0" advTm="5000"><p:fade/></p:transition></p:sld>'));
 const parsed=await readPptx(zipSync(parts),{allowPartial:true}),html=exportHtml(parsed.model).html;
 const rt=await createBrowserRuntime({capture:false});try{
 const m=await measureInput(await readInput({html}),rt),out=await exportPptx(m);assert.equal(out.status,'generated',JSON.stringify(out.diagnostics));
 const d=new DOMParser().parseFromString(strFromU8(unzipSync(out.pptxBytes)[key]),'application/xml');const t=d.getElementsByTagNameNS('http://schemas.openxmlformats.org/presentationml/2006/main','transition')[0];
 assert.ok(t,'transition is present');assert.equal(t.getElementsByTagNameNS('*','fade').length,1);assert.equal(t.getAttributeNS('http://schemas.microsoft.com/office/powerpoint/2010/main','dur'),'1250');assert.equal(t.getAttribute('advClick'),'0');assert.equal(t.getAttribute('advTm'),'5000');
 }finally{await rt.close();}
});
