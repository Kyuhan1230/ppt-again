import test from 'node:test';
import assert from 'node:assert/strict';
import PptxGenJS from 'pptxgenjs';
import {unzipSync,zipSync,strFromU8,strToU8} from 'fflate';
import {createHtmlServer} from '../src/html-to-pptx/server.mjs';
import {projectDiagnostics} from '../web/html-to-pptx/diagnostic-view.js';

test('failed PPTX analysis retains slide and distinct same-named objects in public diagnostics',async()=>{
 const ppt=new PptxGenJS(),slide=ppt.addSlide();slide.background={color:'FFFFFF'};
 for(const [text,y] of [['First\nSecond',1],['Third\nFourth',3]])slide.addText(text,{x:1,y,w:4,h:2,fontFace:'Arial',fontSize:16,objectName:'same-name'});
 const parts=unzipSync(new Uint8Array(await ppt.write({outputType:'uint8array'}))),key='ppt/slides/slide1.xml';
 const original=strFromU8(parts[key]),updated=original.replaceAll('<a:pPr indent="0" marL="0">','<a:pPr indent="0" marL="0"><a:spcBef><a:spcPct val="100000"/></a:spcBef>');
 assert.notEqual(updated,original);parts[key]=strToU8(updated);
 parts['ppt/embeddings/unused.xlsx']=zipSync({'[Content_Types].xml':strToU8('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>')});
 const app=await createHtmlServer({port:0});try{
  const session=await fetch(`${app.url}/api/session`),cookie=session.headers.get('set-cookie').split(';')[0],{csrfToken}=await session.json();
  const headers={'content-type':'application/json',cookie,'x-csrf-token':csrfToken,origin:app.url};
  const response=await fetch(`${app.url}/api/analyze`,{method:'POST',headers,body:JSON.stringify({sourceType:'pptx',fileName:'synthetic.pptx',pptxBase64:Buffer.from(zipSync(parts)).toString('base64'),pptxOptions:{fontPolicy:'original',allowPartial:true}})});
  const job=await response.json();assert.equal(response.status,422,JSON.stringify(job));assert.equal(job.status,'failed');
  const warnings=job.diagnostics.filter(d=>d.reasonKey==='PPTX_PARAGRAPH_SPACING_APPROXIMATED');assert.equal(warnings.length,2,JSON.stringify(job.diagnostics));
  assert.deepEqual(warnings.map(d=>[d.slideNumber,d.objectName,d.objectType]),[[1,'same-name','text'],[1,'same-name','text']]);
  assert.notEqual(warnings[0].objectIndex,warnings[1].objectIndex);
  const category=projectDiagnostics(job.diagnostics).categories.find(c=>c.reasonKey==='PPTX_PARAGRAPH_SPACING_APPROXIMATED');assert.equal(category.count,2);assert.equal(category.rows.length,2);
 }finally{await app.close();}
});
