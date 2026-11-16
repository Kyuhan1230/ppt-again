import test from 'node:test';import assert from 'node:assert/strict';import PptxGenJS from 'pptxgenjs';import {unzipSync,strFromU8} from 'fflate';import {DOMParser} from '@xmldom/xmldom';
import {readPptx} from '../src/html-to-pptx/read-pptx.mjs';import {exportHtml} from '../src/html-to-pptx/export-html.mjs';import {readInput} from '../src/html-to-pptx/input.mjs';import {measureInput} from '../src/html-to-pptx/measure.mjs';import {createBrowserRuntime} from '../src/html-to-pptx/browser-runtime.mjs';import {exportPptx} from '../src/html-to-pptx/export.mjs';
test('separate text paragraphs preserve alignment and spacing through HTML',async()=>{
 const ppt=new PptxGenJS(),s=ppt.addSlide();s.addText([{text:'First paragraph',options:{align:'left',breakLine:true}},{text:'Right footnote',options:{align:'right',paraSpaceBefore:8}}],{x:1,y:1,w:7,h:2,fontFace:'Arial',fontSize:18,objectName:'paragraph-box'});
 const parsed=await readPptx(new Uint8Array(await ppt.write({outputType:'uint8array'})),{allowPartial:true}),html=exportHtml(parsed.model).html,rt=await createBrowserRuntime({capture:false});
 try{const m=await measureInput(await readInput({html}),rt),out=await exportPptx(m);assert.equal(out.status,'generated',JSON.stringify(out.diagnostics));
 const doc=new DOMParser().parseFromString(strFromU8(unzipSync(out.pptxBytes)['ppt/slides/slide1.xml']),'application/xml'),ps=[...doc.getElementsByTagNameNS('*','p')];
 const foot=ps.find(p=>p.textContent.includes('Right footnote'));assert.ok(foot);const props=foot.getElementsByTagNameNS('*','pPr')[0];assert.equal(props.getAttribute('algn'),'r');const sp=props.getElementsByTagNameNS('*','spcBef')[0];assert.ok(sp);assert.equal(sp.getElementsByTagNameNS('*','spcPts')[0].getAttribute('val'),'800');
 }finally{await rt.close();}
});
