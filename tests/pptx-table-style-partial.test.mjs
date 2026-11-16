import test from 'node:test';
import assert from 'node:assert/strict';
import PptxGenJS from 'pptxgenjs';
import {unzipSync,zipSync,strFromU8,strToU8} from 'fflate';
import {readPptx} from '../src/html-to-pptx/read-pptx.mjs';
import {exportHtml} from '../src/html-to-pptx/export-html.mjs';
import {publicDiagnostic} from '../src/html-to-pptx/public-diagnostics.mjs';
import {DOMParser} from '@xmldom/xmldom';
import {readNativeTable} from '../src/html-to-pptx/read-table.mjs';

test('partial import preserves table text when a border style cannot be reproduced',async()=>{
 const ppt=new PptxGenJS(),slide=ppt.addSlide();slide.background={color:'FFFFFF'};
 slide.addTable([['Version','0.1'],['Check List','next']],{x:1,y:1,w:5,h:1,fontFace:'Arial',fontSize:16,objectName:'version-table'});
 const parts=unzipSync(new Uint8Array(await ppt.write({outputType:'uint8array'})));
 const key='ppt/slides/slide1.xml',xml=strFromU8(parts[key]);
 const changed=xml.replace(/<a:lnL\b[^>]*><a:noFill\/><\/a:lnL>/,'<a:lnL w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:prstDash val="dash"/></a:lnL>');
 assert.notEqual(changed,xml);parts[key]=strToU8(changed);const bytes=zipSync(parts);
 const strict=await readPptx(bytes);assert.equal(strict.status,'failed');
 const partial=await readPptx(bytes,{allowPartial:true});assert.equal(partial.status,'parsed',JSON.stringify(partial.diagnostics));
 const table=partial.model.slides[0].nodes.find(n=>n.kind==='table');assert.ok(table);assert.equal(table.content.cells[0].text,'Version');
 assert.match(exportHtml(partial.model).html,/Check List/);
 const warning=partial.diagnostics.find(d=>d.reasonKey==='PPTX_TABLE_STYLE_APPROXIMATED');assert.ok(warning);assert.equal(warning.severity,'warning');
 assert.equal(publicDiagnostic(warning).reasonKey,'PPTX_TABLE_STYLE_APPROXIMATED');
});
test('unsupported SmartArt remains an explicit recoverable error in partial mode',()=>{
 const frame=new DOMParser().parseFromString('<p:graphicFrame xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram"/></a:graphic></p:graphicFrame>','application/xml').documentElement;
 const result=readNativeTable(frame,{slideId:'slide-1',sourceId:'smartart',boxCssPx:{x:0,y:0,w:10,h:10},allowPartial:true});
 assert.equal(result.node,null);assert.equal(result.diagnostics[0].severity,'error');assert.equal(result.diagnostics[0].recoverable,true);
});
