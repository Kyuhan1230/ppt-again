import test from 'node:test';
import assert from 'node:assert/strict';
import PptxGenJS from 'pptxgenjs';
import {unzipSync,zipSync,strFromU8,strToU8} from 'fflate';
import {readPptx} from '../src/html-to-pptx/read-pptx.mjs';
import {exportHtml} from '../src/html-to-pptx/export-html.mjs';
import {publicDiagnostic} from '../src/html-to-pptx/public-diagnostics.mjs';
import {chromium} from 'playwright-core';

test('table cells retain each character bullet and paragraph as editable text',async()=>{
  const p=new PptxGenJS(),s=p.addSlide();s.background={color:'FFFFFF'};
  s.addTable([['First bullet\nSecond bullet','Neighbor']],{x:1,y:1,w:6,h:1,fontFace:'Arial',fontSize:16,objectName:'bullet-table'});
  const parts=unzipSync(new Uint8Array(await p.write({outputType:'uint8array'}))),file='ppt/slides/slide1.xml';
  let xml=strFromU8(parts[file]);let count=0;
  xml=xml.replace(/<a:pPr indent="0" marL="0"><a:buNone\/><\/a:pPr>/g,original=>count++<2?'<a:pPr marL="171450" indent="-171450"><a:buFont typeface="Arial"/><a:buChar char="•"/></a:pPr>':original);
  parts[file]=strToU8(xml);
  const result=await readPptx(zipSync(parts),{allowPartial:true});assert.equal(result.status,'parsed',JSON.stringify(result.diagnostics));
  const table=result.model.slides[0].nodes.find(n=>n.id==='bullet-table');assert.ok(table,JSON.stringify(result.diagnostics));
  assert.equal(table.content.cells[0].text,'• First bullet\n• Second bullet');assert.equal(table.content.cells[1].text,'Neighbor');
  const warning=result.diagnostics.find(d=>d.reasonKey==='PPTX_PARAGRAPH_INDENT_APPROXIMATED');assert.ok(warning);
  const shown=publicDiagnostic(warning);assert.match(shown.message,/들여쓰기/);assert.match(shown.reason,/글머리/);assert.match(shown.message,/글꼴/);assert.match(shown.message,/색/);
  const text=exportHtml(result.model).html.replace(/<[^>]+>/g,'');assert.match(text,/• First bullet/);assert.match(text,/• Second bullet/);assert.match(text,/Neighbor/);
});

test('a bullet uses its own specified font without changing the table text font',async()=>{
  const ppt=new PptxGenJS(),slide=ppt.addSlide();slide.background={color:'FFFFFF'};
  slide.addTable([['Content']],{x:1,y:1,w:3,h:1,fontFace:'Calibri',fontSize:16,objectName:'bullet-font'});
  const parts=unzipSync(new Uint8Array(await ppt.write({outputType:'uint8array'}))),file='ppt/slides/slide1.xml';let xml=strFromU8(parts[file]);
  xml=xml.replace('<a:buNone/>','<a:buFont typeface="Arial"/><a:buChar char="•"/>');parts[file]=strToU8(xml);
  const parsed=await readPptx(zipSync(parts),{allowPartial:true,fontPolicy:'original'});assert.equal(parsed.status,'parsed',JSON.stringify(parsed.diagnostics));
  const runs=parsed.model.slides[0].nodes.find(n=>n.id==='bullet-font').content.cells[0].runs;
  assert.equal(runs.find(r=>r.text==='• ').style.fontFamily,'Arial');assert.equal(runs.find(r=>r.text==='Content').style.fontFamily,'Calibri');
  const replaced=await readPptx(zipSync(parts),{allowPartial:true,fontPolicy:'pretendard'});
  assert.ok(replaced.model.slides[0].nodes.find(n=>n.id==='bullet-font').content.cells[0].runs.every(r=>r.style.fontFamily==='Pretendard'));
});

test('character bullets use their font face without inheriting bold from the text run',async()=>{
  const ppt=new PptxGenJS(),slide=ppt.addSlide();slide.background={color:'FFFFFF'};
  slide.addText('Text',{x:1,y:1,w:3,h:1,fontFace:'Calibri',fontSize:16,bold:true,objectName:'bullet-shape'});
  slide.addTable([['Cell']],{x:1,y:2,w:3,h:1,fontFace:'Calibri',fontSize:16,bold:true,objectName:'bullet-cell'});
  const parts=unzipSync(new Uint8Array(await ppt.write({outputType:'uint8array'}))),file='ppt/slides/slide1.xml',original=strFromU8(parts[file]);let seen=0;
  const xml=original.replace(/<a:buNone\/>/g,()=>`<a:buFont typeface="${seen++?'Arial Bold':'Arial'}"/><a:buChar char="•"/>`);
  assert.equal(seen,2);parts[file]=strToU8(xml);
  const parsed=await readPptx(zipSync(parts),{allowPartial:true,fontPolicy:'original'});assert.equal(parsed.status,'parsed',JSON.stringify(parsed.diagnostics));
  for(const [id,face] of [['bullet-shape','Arial'],['bullet-cell','Arial Bold']]){
    const node=parsed.model.slides[0].nodes.find(n=>n.id===id),runs=node.kind==='table'?node.content.cells[0].runs:node.content.runs;
    const marker=runs.find(r=>r.marker),text=runs.find(r=>r.text==='Text'||r.text==='Cell');
    assert.equal(marker.style.fontFamily,face);assert.equal(marker.style.fontWeight,400);assert.equal(text.style.fontWeight,700);
  }
  const replaced=await readPptx(zipSync(parts),{allowPartial:true,fontPolicy:'pretendard'});
  const marker=replaced.model.slides[0].nodes.find(n=>n.id==='bullet-shape').content.runs.find(r=>r.marker);
  assert.equal(marker.style.fontFamily,'Pretendard');assert.equal(marker.style.fontWeight,400);
  const numbered=await readPptx(zipSync({...parts,[file]:strToU8(original.replace('<a:buNone/>','<a:buAutoNum type="arabicPeriod"/>'))}),{allowPartial:true});
  assert.equal(numbered.model.slides[0].nodes.find(n=>n.id==='bullet-shape').content.runs.find(r=>r.marker).style.fontWeight,700);
});

test('bullet format inherits from its list style, while explicit text-font inheritance takes priority',async()=>{
  const ppt=new PptxGenJS(),slide=ppt.addSlide();slide.background={color:'FFFFFF'};slide.addText('Content',{x:1,y:1,w:3,h:1,fontFace:'Calibri',fontSize:16,objectName:'bullet'});
  const parts=unzipSync(new Uint8Array(await ppt.write({outputType:'uint8array'}))),file='ppt/slides/slide1.xml',source=strFromU8(parts[file]);
  let xml=source.replace('<a:lstStyle/>','<a:lstStyle><a:lvl1pPr><a:buFont typeface="Arial"/></a:lvl1pPr></a:lstStyle>').replace('<a:buNone/>','<a:buClr><a:srgbClr val="FF0000"/></a:buClr><a:buSzPct val="150000"/><a:buChar char="•"/>');
  assert.notEqual(xml,source);parts[file]=strToU8(xml);let parsed=await readPptx(zipSync(parts),{allowPartial:true,fontPolicy:'original'});
  let marker=parsed.model.slides[0].nodes.find(n=>n.id==='bullet').content.runs.find(r=>r.text==='• ');assert.equal(marker.style.fontFamily,'Arial');assert.equal(marker.style.color,'#ff0000');assert.equal(marker.style.fontSize,32);
  xml=xml.replace('<a:buChar char="•"/>','<a:buFontTx/><a:buChar char="•"/>');parts[file]=strToU8(xml);parsed=await readPptx(zipSync(parts),{allowPartial:true,fontPolicy:'original'});
  marker=parsed.model.slides[0].nodes.find(n=>n.id==='bullet').content.runs.find(r=>r.text==='• ');assert.equal(marker.style.fontFamily,'Calibri');
});

test('a different bullet font does not enlarge a table cell line box',async()=>{
  const ppt=new PptxGenJS(),slide=ppt.addSlide();slide.background={color:'FFFFFF'};
  slide.addTable([['First\nSecond\nThird']],{x:1,y:1,w:3,h:1,fontFace:'Arial',fontSize:10,objectName:'bullet-height'});
  const parts=unzipSync(new Uint8Array(await ppt.write({outputType:'uint8array'}))),file='ppt/slides/slide1.xml';
  parts[file]=strToU8(strFromU8(parts[file]).replaceAll('<a:buNone/>','<a:buFont typeface="Times New Roman"/><a:buChar char="•"/>'));
  const parsed=await readPptx(zipSync(parts),{allowPartial:true,fontPolicy:'original'}),html=exportHtml(parsed.model).html,browser=await chromium.launch({channel:'msedge',headless:true});
  try{
    const page=await browser.newPage();await page.setContent(html);const cell=page.locator('[data-source-id="bullet-height"] td');
    const info=await cell.evaluate(el=>{const markers=[...el.querySelectorAll('[data-pptx-list-marker]')],family=getComputedStyle(el).fontFamily,before=el.getBoundingClientRect().height;for(const marker of markers)marker.style.fontFamily=family;return {family,count:markers.length,before,after:el.getBoundingClientRect().height};});
    assert.equal(info.count,3);assert.match(info.family,/Arial/i);assert.ok(Math.abs(info.before-info.after)<.1,JSON.stringify(info));
  }finally{await browser.close();}
});

test('an empty list paragraph adds no visible bullet',async()=>{
  const p=new PptxGenJS(),s=p.addSlide();s.background={color:'FFFFFF'};
  s.addTable([['Visible','Neighbor']],{x:1,y:1,w:6,h:1,fontFace:'Arial',fontSize:16,objectName:'table'});
  const parts=unzipSync(new Uint8Array(await p.write({outputType:'uint8array'}))),file='ppt/slides/slide1.xml';let xml=strFromU8(parts[file]);
  xml=xml.replace('</a:txBody>','<a:p><a:pPr marL="88900" indent="-88900"><a:buChar char="•"/></a:pPr><a:endParaRPr sz="1000"><a:latin typeface="Arial"/></a:endParaRPr></a:p></a:txBody>');
  parts[file]=strToU8(xml);const result=await readPptx(zipSync(parts),{allowPartial:true});assert.equal(result.status,'parsed',JSON.stringify(result.diagnostics));
  const table=result.model.slides[0].nodes.find(n=>n.id==='table');assert.ok(table);assert.equal(table.content.cells[0].text,'Visible\n');
});

test('first-row text inherits its table style color when no run color is specified',async()=>{
  const p=new PptxGenJS(),s=p.addSlide();s.background={color:'FFFFFF'};
  s.addTable([['Header','Explicit'],['Body','Other']],{x:1,y:1,w:6,h:2,fontFace:'Arial',fontSize:16,fill:{color:'777777'},objectName:'styled-table'});
  const parts=unzipSync(new Uint8Array(await p.write({outputType:'uint8array'}))),file='ppt/slides/slide1.xml';let xml=strFromU8(parts[file]);
  xml=xml.replace('<a:tblPr/>','<a:tblPr firstRow="1"><a:tableStyleId>{TEST-STYLE}</a:tableStyleId></a:tblPr>');
  xml=xml.replace(/<a:tc>[\s\S]*?<\/a:tc>/,cell=>cell.replace('<a:solidFill><a:srgbClr val="000000"/></a:solidFill>',''));
  parts[file]=strToU8(xml);
  parts['ppt/tableStyles.xml']=strToU8('<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:tblStyle styleId="{TEST-STYLE}"><a:wholeTbl><a:tcTxStyle><a:schemeClr val="dk1"/></a:tcTxStyle></a:wholeTbl><a:firstRow><a:tcTxStyle><a:schemeClr val="lt1"/></a:tcTxStyle></a:firstRow></a:tblStyle></a:tblStyleLst>');
  const result=await readPptx(zipSync(parts));assert.equal(result.status,'parsed',JSON.stringify(result.diagnostics));
  const table=result.model.slides[0].nodes.find(n=>n.id==='styled-table');assert.ok(table);
  assert.equal(table.content.cells[0].runs[0].style.color,'#ffffff');assert.equal(table.content.cells[1].runs[0].style.color,'#000000');
});

test('unused unknown inherited table color does not override explicit run color',async()=>{
  const p=new PptxGenJS(),s=p.addSlide();s.background={color:'FFFFFF'};
  s.addTable([['Explicit']],{x:1,y:1,w:3,h:1,fontFace:'Arial',fontSize:16,fill:{color:'777777'},objectName:'styled-table'});
  const parts=unzipSync(new Uint8Array(await p.write({outputType:'uint8array'}))),file='ppt/slides/slide1.xml';
  parts[file]=strToU8(strFromU8(parts[file]).replace('<a:tblPr/>','<a:tblPr firstRow="1"><a:tableStyleId>{TEST-STYLE}</a:tableStyleId></a:tblPr>'));
  parts['ppt/tableStyles.xml']=strToU8('<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:tblStyle styleId="{TEST-STYLE}"><a:firstRow><a:tcTxStyle><a:schemeClr val="unknown"/></a:tcTxStyle></a:firstRow></a:tblStyle></a:tblStyleLst>');
  const result=await readPptx(zipSync(parts));assert.equal(result.status,'parsed',JSON.stringify(result.diagnostics));
  assert.equal(result.model.slides[0].nodes.find(n=>n.id==='styled-table').content.cells[0].runs[0].style.color,'#000000');
});

test('first-row color applies outside first-column intersection and warns on ambiguous corner',async()=>{
  const p=new PptxGenJS(),s=p.addSlide();s.background={color:'FFFFFF'};
  s.addTable([['Corner','Header','Explicit'],['First','Body','Other']],{x:1,y:1,w:6,h:2,fontFace:'Arial',fontSize:16,fill:{color:'777777'},objectName:'mixed-table'});
  const parts=unzipSync(new Uint8Array(await p.write({outputType:'uint8array'}))),file='ppt/slides/slide1.xml';let xml=strFromU8(parts[file]);
  xml=xml.replace('<a:tblPr/>','<a:tblPr firstRow="1" firstCol="1"><a:tableStyleId>{TEST-STYLE}</a:tableStyleId></a:tblPr>');let seen=0;
  xml=xml.replace(/<a:tc>[\s\S]*?<\/a:tc>/g,cell=>seen++<2?cell.replace('<a:solidFill><a:srgbClr val="000000"/></a:solidFill>',''):cell);parts[file]=strToU8(xml);
  parts['ppt/tableStyles.xml']=strToU8('<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:tblStyle styleId="{TEST-STYLE}"><a:firstRow><a:tcTxStyle><a:schemeClr val="lt1"/></a:tcTxStyle></a:firstRow><a:firstCol><a:tcTxStyle><a:srgbClr val="FF0000"/></a:tcTxStyle></a:firstCol></a:tblStyle></a:tblStyleLst>');
  const r=await readPptx(zipSync(parts),{allowPartial:true});assert.equal(r.status,'parsed',JSON.stringify(r.diagnostics));const cells=r.model.slides[0].nodes.find(n=>n.id==='mixed-table').content.cells;
  assert.equal(cells[1].runs[0].style.color,'#ffffff');assert.equal(cells[2].runs[0].style.color,'#000000');
  const warning=r.diagnostics.find(d=>d.reasonKey==='PPTX_TABLE_TEXT_COLOR_APPROXIMATED');assert.ok(warning);
  const shown=publicDiagnostic(warning);assert.equal(shown.reasonKey,'PPTX_TABLE_TEXT_COLOR_APPROXIMATED');assert.match(shown.message,/표/);assert.match(shown.reason,/색/);
});
