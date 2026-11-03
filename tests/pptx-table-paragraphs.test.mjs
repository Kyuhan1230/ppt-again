import test from 'node:test';
import assert from 'node:assert/strict';
import PptxGenJS from 'pptxgenjs';
import {unzipSync,zipSync,strFromU8,strToU8} from 'fflate';
import {chromium} from 'playwright-core';
import {readPptx} from '../src/html-to-pptx/read-pptx.mjs';
import {exportHtml} from '../src/html-to-pptx/export-html.mjs';
import {readInput} from '../src/html-to-pptx/input.mjs';
import {createBrowserRuntime} from '../src/html-to-pptx/browser-runtime.mjs';
import {measureInput} from '../src/html-to-pptx/measure.mjs';
import {exportPptx} from '../src/html-to-pptx/export.mjs';

test('table paragraphs preserve hanging indents and remain available for PPT export',async()=>{
 const ppt=new PptxGenJS(),slide=ppt.addSlide();slide.background={color:'FFFFFF'};
 slide.addTable([['Alpha beta gamma delta epsilon zeta eta theta iota kappa lambda\nSecond paragraph']],{x:1,y:1,w:2,h:1.5,fontFace:'Arial',fontSize:10,objectName:'table'});
 const parts=unzipSync(new Uint8Array(await ppt.write({outputType:'uint8array'}))),key='ppt/slides/slide1.xml',xml=strFromU8(parts[key]);
 const patched=xml.replace('<a:pPr indent="0" marL="0"><a:buNone/></a:pPr>','<a:pPr marL="171450" indent="-171450"><a:buFont typeface="Arial"/><a:buChar char="•"/></a:pPr>');
 assert.notEqual(patched,xml);parts[key]=strToU8(patched);
 const parsed=await readPptx(zipSync(parts),{allowPartial:true});assert.equal(parsed.status,'parsed',JSON.stringify(parsed.diagnostics));
 const cell=parsed.model.slides[0].nodes.find(n=>n.id==='table').content.cells[0];
 assert.equal(cell.paragraphs.length,2);assert.equal(cell.paragraphs[0].leftPx,18);assert.equal(cell.paragraphs[0].indentPx,-18);
 assert.equal(cell.paragraphs[0].hangingPunct,true);
 const html=exportHtml(parsed.model).html,browser=await chromium.launch({channel:'msedge',headless:true});
 try{
  const page=await browser.newPage({viewport:{width:1200,height:800}});await page.setContent(html);
  const paragraph=page.locator('td [data-pptx-paragraph]').first();assert.equal(await page.locator('td [data-pptx-paragraph]').count(),2);
  const starts=await paragraph.evaluate(p=>{
   const rows=new Map();for(const span of p.querySelectorAll('span:not([data-pptx-list-marker])'))for(const node of span.childNodes){if(node.nodeType!==Node.TEXT_NODE)continue;
    for(let i=0;i<node.length;i++){if(/\s/.test(node.data[i]))continue;const range=document.createRange();range.setStart(node,i);range.setEnd(node,i+1);const r=range.getBoundingClientRect(),y=Math.round(r.top*2)/2;if(!rows.has(y))rows.set(y,r.left);}
   }return [...rows.values()];
  });
  assert.ok(starts.length>=2,JSON.stringify(starts));assert.ok(Math.abs(starts[0]-starts[1])<1,JSON.stringify(starts));
 }finally{await browser.close();}
 const runtime=await createBrowserRuntime({capture:false});try{
  const measured=await measureInput(await readInput({html}),runtime),output=await exportPptx(measured);
  assert.equal(output.status,'generated',JSON.stringify(output.diagnostics));
  assert.equal(measured.model.slides[0].nodes.find(n=>n.id==='table').content.cells[0].text,cell.text);
  assert.ok(measured.diagnostics.some(d=>d.code==='PARAGRAPH_SPACING_APPROXIMATED'),JSON.stringify(measured.diagnostics));
 }finally{await runtime.close();}
});

test('a hanging final period can stay on the preceding line and explicit opt-out is respected',async()=>{
 const ppt=new PptxGenJS(),slide=ppt.addSlide();slide.background={color:'FFFFFF'};
 slide.addTable([['AVATAR.']],{x:1,y:1,w:2,h:1,fontFace:'Arial',fontSize:10,objectName:'table'});
 const parts=unzipSync(new Uint8Array(await ppt.write({outputType:'uint8array'}))),key='ppt/slides/slide1.xml',xml=strFromU8(parts[key]);
 const original='<a:pPr indent="0" marL="0"><a:buNone/></a:pPr>',paragraph='<a:pPr marL="171450" indent="-171450" hangingPunct="1"><a:buFont typeface="Arial"/><a:buChar char="•"/></a:pPr>';
 assert.ok(xml.includes(original));parts[key]=strToU8(xml.replace(original,paragraph));
 const parsed=await readPptx(zipSync(parts),{allowPartial:true});assert.equal(parsed.status,'parsed',JSON.stringify(parsed.diagnostics));
 const table=parsed.model.slides[0].nodes.find(n=>n.id==='table'),cell=table.content.cells[0];assert.equal(cell.paragraphs[0].hangingPunct,true);
 const browser=await chromium.launch({channel:'msedge',headless:true});try{
  const page=await browser.newPage({viewport:{width:1200,height:800}});
  const widths=await page.evaluate(()=>{const ctx=document.createElement('canvas').getContext('2d');ctx.font='13.3333333333px Arial';return [ctx.measureText('AVATAR').width,ctx.measureText('.').width];});
  const width=18+widths[0]+widths[1]/2;table.boxCssPx.w=width;table.content.columnWidthsCssPx=[width];
  cell.style.padding={top:0,right:0,bottom:0,left:0};for(const side of ['top','right','bottom','left'])cell.style.borders[side]={width:0,style:'none',color:'transparent'};
  await page.setContent(exportHtml(parsed.model).html);const paragraphEl=page.locator('td [data-pptx-paragraph]');
  const info=await paragraphEl.evaluate(p=>{
   const period=p.querySelector('[data-pptx-hanging-punct]'),text=period?.previousSibling;
   const last=document.createRange(),mark=document.createRange();last.setStart(text,text.length-1);last.setEnd(text,text.length);mark.selectNodeContents(period);
   return {height:p.getBoundingClientRect().height,text:p.innerText,periods:p.querySelectorAll('[data-pptx-hanging-punct]').length,gap:mark.getBoundingClientRect().left-last.getBoundingClientRect().right};
  });
  assert.equal(info.periods,1);assert.ok(info.height<20,JSON.stringify(info));assert.equal(info.text,'• AVATAR.');
  assert.ok(Math.abs(info.gap)<1,JSON.stringify(info));
 }finally{await browser.close();}
 table.boxCssPx.w=192;table.content.columnWidthsCssPx=[192];
 const generated=exportHtml(parsed.model).html,pattern=/(<span data-pptx-hanging-punct[^>]*>)\.(<\/span>)/;
 const edited=[['• AVATAR!',generated.replace(pattern,'$1!$2')],['• AVATAR. Next',generated.replace(pattern,'$1.$2 Next')]];
 const runtime=await createBrowserRuntime({capture:false});try{
  for(const [expected,html] of edited){assert.notEqual(html,generated);const measured=await measureInput(await readInput({html}),runtime),output=await exportPptx(measured);
   assert.equal(output.status,'generated',JSON.stringify(output.diagnostics));assert.equal(measured.model.slides[0].nodes.find(n=>n.id==='table').content.cells[0].text,expected);
  }
 }finally{await runtime.close();}
 parts[key]=strToU8(xml.replace(original,paragraph.replace('hangingPunct="1"','hangingPunct="0"')));
 const disabled=await readPptx(zipSync(parts),{allowPartial:true}),disabledCell=disabled.model.slides[0].nodes.find(n=>n.id==='table').content.cells[0];
 assert.equal(disabledCell.paragraphs[0].hangingPunct,false);assert.doesNotMatch(exportHtml(disabled.model).html,/data-pptx-hanging-punct/);
});

test('malformed table paragraph coordinates and spacing are rejected',async()=>{
 const ppt=new PptxGenJS(),slide=ppt.addSlide();slide.background={color:'FFFFFF'};slide.addTable([['Text']],{x:1,y:1,w:4,h:1,fontFace:'Arial',fontSize:10,objectName:'table'});
 const base=unzipSync(new Uint8Array(await ppt.write({outputType:'uint8array'}))),key='ppt/slides/slide1.xml',xml=strFromU8(base[key]);
 const original='<a:pPr indent="0" marL="0"><a:buNone/></a:pPr>';assert.ok(xml.includes(original));
 const cases=['<a:pPr indent="0" marL="invalid"><a:buNone/></a:pPr>','<a:pPr indent="invalid" marL="0"><a:buNone/></a:pPr>','<a:pPr indent="0" marL="1.5"><a:buNone/></a:pPr>','<a:pPr><a:lnSpc><a:spcPct val="invalid"/></a:lnSpc><a:buNone/></a:pPr>','<a:pPr><a:lnSpc><a:spcPct val="99999999"/></a:lnSpc><a:buNone/></a:pPr>','<a:pPr><a:lnSpc/><a:buNone/></a:pPr>','<a:pPr><a:lnSpc><a:spcPct val="100000"/><a:spcPts val="1200"/></a:lnSpc><a:buNone/></a:pPr>'];
 for(const candidate of cases){const parts={...base,[key]:strToU8(xml.replace(original,candidate))},result=await readPptx(zipSync(parts),{allowPartial:true});assert.equal(result.status,'failed',candidate);assert.ok(result.diagnostics.some(d=>d.blocking),candidate);}
 const inherited=xml.replace('<a:lstStyle/>','<a:lstStyle><a:lvl1pPr><a:lnSpc><a:spcPct val="not-a-number"/></a:lnSpc></a:lvl1pPr></a:lstStyle>');assert.notEqual(inherited,xml);
 const invalidInherited=await readPptx(zipSync({...base,[key]:strToU8(inherited)}),{allowPartial:true});assert.equal(invalidInherited.status,'failed');assert.ok(invalidInherited.diagnostics.some(d=>d.blocking));
 const badBefore=xml.replace('<a:lstStyle/>','<a:lstStyle><a:lvl1pPr><a:spcBef><a:spcPts val="invalid"/></a:spcBef></a:lvl1pPr></a:lstStyle>');assert.notEqual(badBefore,xml);
 const invalidBefore=await readPptx(zipSync({...base,[key]:strToU8(badBefore)}),{allowPartial:true});assert.equal(invalidBefore.status,'failed');assert.ok(invalidBefore.diagnostics.some(d=>d.blocking));
});

test('percentage strings and OOXML integer line spacing render identically',async()=>{
 const ppt=new PptxGenJS(),slide=ppt.addSlide();slide.background={color:'FFFFFF'};slide.addTable([['Text']],{x:1,y:1,w:4,h:1,fontFace:'Arial',fontSize:10,objectName:'table'});
 const base=unzipSync(new Uint8Array(await ppt.write({outputType:'uint8array'}))),key='ppt/slides/slide1.xml',xml=strFromU8(base[key]);
 const original='<a:pPr indent="0" marL="0"><a:buNone/></a:pPr>';assert.ok(xml.includes(original));
 for(const [value,expected] of [['100%','1.2'],['200%','2.4'],['100000','1.2']]){
  const replacement=`<a:pPr marL="88900" indent="-88900"><a:lnSpc><a:spcPct val="${value}"/></a:lnSpc><a:buNone/></a:pPr>`;
  const result=await readPptx(zipSync({...base,[key]:strToU8(xml.replace(original,replacement))}),{allowPartial:true});assert.equal(result.status,'parsed',JSON.stringify(result.diagnostics));
  assert.equal(result.model.slides[0].nodes.find(n=>n.id==='table').content.cells[0].paragraphs[0].lineHeight,expected);
 }
});

test('table paragraph spacing inherited from its own list style is preserved',async()=>{
 const ppt=new PptxGenJS(),slide=ppt.addSlide();slide.background={color:'FFFFFF'};
 slide.addTable([['First\nSecond']],{x:1,y:1,w:4,h:2,fontFace:'Arial',fontSize:10,objectName:'table'});
 const parts=unzipSync(new Uint8Array(await ppt.write({outputType:'uint8array'}))),key='ppt/slides/slide1.xml',xml=strFromU8(parts[key]);
 const inherited='<a:lstStyle><a:lvl1pPr><a:spcBef><a:spcPts val="1200"/></a:spcBef><a:spcAft><a:spcPts val="600"/></a:spcAft></a:lvl1pPr></a:lstStyle>';
 const patched=xml.replace('<a:lstStyle/>',inherited);assert.notEqual(patched,xml);parts[key]=strToU8(patched);
 const result=await readPptx(zipSync(parts),{allowPartial:true});assert.equal(result.status,'parsed',JSON.stringify(result.diagnostics));
 const paragraphs=result.model.slides[0].nodes.find(n=>n.id==='table').content.cells[0].paragraphs;
 assert.equal(paragraphs.length,2);
 assert.deepEqual(paragraphs.map(p=>[p.beforePx,p.afterPx]),[[0,8],[16,0]]);
});

test('empty table paragraphs do not introduce extra text when read back',async()=>{
 const ppt=new PptxGenJS(),slide=ppt.addSlide();slide.background={color:'FFFFFF'};
 slide.addTable([['Visible','']],{x:1,y:1,w:5,h:1,fontFace:'Arial',fontSize:10,objectName:'table'});
 const parts=unzipSync(new Uint8Array(await ppt.write({outputType:'uint8array'}))),key='ppt/slides/slide1.xml';let xml=strFromU8(parts[key]);
 const end='<a:p><a:pPr marL="88900" indent="-88900"><a:buNone/></a:pPr><a:endParaRPr sz="1000"><a:latin typeface="Arial"/></a:endParaRPr></a:p>';
 xml=xml.replace('</a:txBody>',`${end}</a:txBody>`);
 let cells=0;xml=xml.replace(/<a:tc>[\s\S]*?<\/a:tc>/g,cell=>cells++===1?cell.replace('<a:pPr indent="0" marL="0">','<a:pPr indent="-88900" marL="88900">'):cell);
 parts[key]=strToU8(xml);const parsed=await readPptx(zipSync(parts),{allowPartial:true});assert.equal(parsed.status,'parsed',JSON.stringify(parsed.diagnostics));
 const source=parsed.model.slides[0].nodes.find(n=>n.id==='table').content.cells;assert.deepEqual(source.map(c=>c.text),['Visible\n','']);
 const html=exportHtml(parsed.model).html;assert.ok((html.match(/data-pptx-empty/g)??[]).length>=2);
 const runtime=await createBrowserRuntime({capture:false});try{
  const measured=await measureInput(await readInput({html}),runtime),output=await exportPptx(measured);
  assert.equal(output.status,'generated',JSON.stringify(output.diagnostics));
  assert.deepEqual(measured.model.slides[0].nodes.find(n=>n.id==='table').content.cells.map(c=>c.text),source.map(c=>c.text));
 }finally{await runtime.close();}
});

test('a manual line break remains inside its table paragraph',async()=>{
 const ppt=new PptxGenJS(),slide=ppt.addSlide();slide.background={color:'FFFFFF'};
 slide.addTable([['FirstLine\nSecond']],{x:1,y:1,w:4,h:1,fontFace:'Arial',fontSize:10,objectName:'table'});
 const parts=unzipSync(new Uint8Array(await ppt.write({outputType:'uint8array'}))),key='ppt/slides/slide1.xml';let xml=strFromU8(parts[key]);
 const start='<a:t>FirstLine</a:t></a:r>';
 assert.ok(xml.includes(start));xml=xml.replace(start,'<a:t>First</a:t></a:r><a:br/><a:r><a:rPr sz="1000"><a:latin typeface="Arial"/></a:rPr><a:t>Line</a:t></a:r>');
 xml=xml.replace('<a:pPr indent="0" marL="0"><a:buNone/></a:pPr>','<a:pPr marL="171450" indent="-171450"><a:buFont typeface="Arial"/><a:buChar char="•"/></a:pPr>');
 parts[key]=strToU8(xml);const parsed=await readPptx(zipSync(parts),{allowPartial:true});assert.equal(parsed.status,'parsed',JSON.stringify(parsed.diagnostics));
 const source=parsed.model.slides[0].nodes.find(n=>n.id==='table').content.cells[0];
 assert.equal(source.paragraphs.length,2);assert.equal(source.text,'• First\nLine\nSecond');assert.ok(source.paragraphs[0].runs.some(run=>run.text==='\n'));
 const runtime=await createBrowserRuntime({capture:false});try{
  const measured=await measureInput(await readInput({html:exportHtml(parsed.model).html}),runtime),output=await exportPptx(measured);
  assert.equal(output.status,'generated',JSON.stringify(output.diagnostics));assert.equal(measured.model.slides[0].nodes.find(n=>n.id==='table').content.cells[0].text,source.text);
 }finally{await runtime.close();}
});

test('leading and middle empty table paragraphs retain their lines',async()=>{
 const ppt=new PptxGenJS(),slide=ppt.addSlide();slide.background={color:'FFFFFF'};
 slide.addTable([['Alpha\nBeta']],{x:1,y:1,w:4,h:2,fontFace:'Arial',fontSize:10,objectName:'table'});
 const parts=unzipSync(new Uint8Array(await ppt.write({outputType:'uint8array'}))),key='ppt/slides/slide1.xml';let xml=strFromU8(parts[key]);
 const empty='<a:p><a:pPr marL="88900" indent="-88900"><a:buNone/></a:pPr><a:endParaRPr sz="1000"><a:latin typeface="Arial"/></a:endParaRPr></a:p>';
 let count=0;xml=xml.replace(/<a:txBody>[\s\S]*?<\/a:txBody>/,body=>body.replace(/<a:p>[\s\S]*?<\/a:p>/g,para=>{count++;return empty+para;}));assert.equal(count,2);
 parts[key]=strToU8(xml);const parsed=await readPptx(zipSync(parts),{allowPartial:true});assert.equal(parsed.status,'parsed',JSON.stringify(parsed.diagnostics));
 const cell=parsed.model.slides[0].nodes.find(n=>n.id==='table').content.cells[0];assert.equal(cell.text,'\nAlpha\n\nBeta');assert.equal(cell.paragraphs.length,4);
 const html=exportHtml(parsed.model).html,browser=await chromium.launch({channel:'msedge',headless:true});try{
  const page=await browser.newPage();await page.setContent(html);
  const heights=await page.locator('td [data-pptx-paragraph]').evaluateAll(paragraphs=>paragraphs.map(p=>p.getBoundingClientRect().height));
  assert.ok(heights[0]>10&&heights[2]>10,JSON.stringify(heights));
 }finally{await browser.close();}
 const runtime=await createBrowserRuntime({capture:false});try{
  const measured=await measureInput(await readInput({html}),runtime),output=await exportPptx(measured);
  assert.equal(output.status,'generated',JSON.stringify(output.diagnostics));assert.equal(measured.model.slides[0].nodes.find(n=>n.id==='table').content.cells[0].text,cell.text);
 }finally{await runtime.close();}
});
