import test from 'node:test';
import assert from 'node:assert/strict';
import PptxGenJS from 'pptxgenjs';
import {readPptx} from '../src/html-to-pptx/read-pptx.mjs';
import {exportHtml} from '../src/html-to-pptx/export-html.mjs';
import {readInput} from '../src/html-to-pptx/input.mjs';
import {measureInput} from '../src/html-to-pptx/measure.mjs';
import {createBrowserRuntime} from '../src/html-to-pptx/browser-runtime.mjs';
import {exportPptx} from '../src/html-to-pptx/export.mjs';
import {chromium} from 'playwright-core';
import {unzipSync,strFromU8} from 'fflate';

test('a table that grows to fit cell text can be returned to PPT at its displayed size',async()=>{
  const ppt=new PptxGenJS();ppt.layout='LAYOUT_WIDE';const slide=ppt.addSlide();slide.background={color:'FFFFFF'};
  slide.addTable([['first\nsecond\nthird\nfourth\nfifth']],{x:1,y:1,w:4,h:.3,fontFace:'Arial',fontSize:20,objectName:'Growing table'});
  const parsed=await readPptx(new Uint8Array(await ppt.write({outputType:'uint8array'})));
  assert.equal(parsed.status,'parsed',JSON.stringify(parsed.diagnostics));
  const sourceHeight=parsed.model.slides[0].nodes.find(n=>n.kind==='table').boxCssPx.h;
  const html=exportHtml(parsed.model);assert.equal(html.status,'generated');
  const runtime=await createBrowserRuntime({capture:false});
  try {
    const measured=await measureInput(await readInput({html:html.html}),runtime);
    assert.ok(measured.diagnostics.some(d=>d.code==='TABLE_HEIGHT_EXPANDED'&&d.severity==='warning'),JSON.stringify(measured.diagnostics));
    const table=measured.model.slides[0].nodes.find(n=>n.kind==='table');
    assert.ok(table.boxCssPx.h>sourceHeight);
    assert.ok(Math.abs(table.content.rowHeightsCssPx.reduce((a,b)=>a+b,0)-table.boxCssPx.h)<.05);
    const output=await exportPptx(measured);
    assert.equal(output.status,'generated',JSON.stringify(output.diagnostics));
    assert.ok(output.diagnostics.some(d=>d.code==='TABLE_HEIGHT_EXPANDED'));
  } finally {await runtime.close();}
});

test('thin table borders do not add extra height to every content-sized row',async()=>{
  const ppt=new PptxGenJS(),slide=ppt.addSlide();slide.background={color:'FFFFFF'};
  slide.addTable([['First'],['Second']],{x:1,y:1,w:4,h:1,fontFace:'Arial',fontSize:12,objectName:'thin-border-table'});
  const parsed=await readPptx(new Uint8Array(await ppt.write({outputType:'uint8array'})),{allowPartial:true});assert.equal(parsed.status,'parsed',JSON.stringify(parsed.diagnostics));
  const table=parsed.model.slides[0].nodes.find(n=>n.id==='thin-border-table');
  table.content.rowHeightsCssPx=[10,10];table.boxCssPx.h=20;
  for(const cell of table.content.cells){
    Object.assign(cell.style,{fontSize:16,lineHeight:'1.2',padding:{top:4.8,right:4.8,bottom:4.8,left:4.8}});
    for(const run of cell.runs)run.style.fontSize=16;
    for(const side of ['top','right','bottom','left'])cell.style.borders[side]={width:1/3,style:'solid',color:'#000000'};
  }
  const html=exportHtml(parsed.model);assert.equal(html.status,'generated',JSON.stringify(html.diagnostics));
  assert.match(html.html,/data-pptx-native-border-top/);
  const browser=await chromium.launch({channel:'msedge',headless:true});try{
    const page=await browser.newPage();await page.setContent(html.html);
    const heights=await page.locator('[data-source-id="thin-border-table"] tr').evaluateAll(rows=>rows.map(r=>r.getBoundingClientRect().height));
    assert.ok(heights.every(h=>Math.abs(h-28.8)<.1),JSON.stringify(heights));
    assert.equal(await page.locator('[data-source-id="thin-border-table"] td').first().evaluate(e=>getComputedStyle(e).borderTopColor),'rgba(0, 0, 0, 0.5)');
    for(const cell of table.content.cells){cell.style.borders.left.width=1;cell.style.borders.right.width=1;}
    await page.setContent(exportHtml(parsed.model).html);
    const sidePadding=await page.locator('[data-source-id="thin-border-table"] td').first().evaluate(e=>[parseFloat(getComputedStyle(e).paddingLeft),parseFloat(getComputedStyle(e).paddingRight)]);
    assert.deepEqual(sidePadding,[4.8,4.8]);
  }finally{await browser.close();}
  const runtime=await createBrowserRuntime({capture:false});try{
    const measured=await measureInput(await readInput({html:html.html}),runtime);
    assert.equal(measured.diagnostics.filter(d=>d.severity==='error').length,0,JSON.stringify(measured.diagnostics));
    const cell=measured.model.slides[0].nodes.find(n=>n.id==='thin-border-table').content.cells[0];
    assert.ok(Math.abs(cell.style.padding.top-4.8)<.01);assert.ok(Math.abs(cell.style.padding.bottom-4.8)<.01);
    assert.ok(Math.abs(cell.style.padding.left-4.8)<.01);assert.ok(Math.abs(cell.style.padding.right-4.8)<.01);
    assert.ok(Math.abs(cell.style.borders.top.width-1/3)<.0001);assert.equal(cell.style.borders.top.color,'#000000');
    const out=await exportPptx(measured);assert.equal(out.status,'generated',JSON.stringify(out.diagnostics));assert.match(strFromU8(unzipSync(out.pptxBytes)['ppt/slides/slide1.xml']),/<a:lnT w="3175"/);
    const edited=html.html.replace('border-top:0.3333333333px solid rgba(0,0,0,0.5)','border-top:0.3333333333px solid rgb(255,0,0)');assert.notEqual(edited,html.html);
    const changed=await measureInput(await readInput({html:edited}),runtime),changedCell=changed.model.slides[0].nodes.find(n=>n.id==='thin-border-table').content.cells[0];
    assert.ok(Math.abs(changedCell.style.borders.top.width-1/3)<.0001);assert.match(changedCell.style.borders.top.color,/255, 0, 0/);
    const invalid=html.html.replace('data-pptx-native-border-top="0.3333333333 #000000"','data-pptx-native-border-top="5000 #000000"');assert.notEqual(invalid,html.html);
    const rejected=await measureInput(await readInput({html:invalid}),runtime);assert.ok(rejected.diagnostics.some(d=>d.code==='INPUT_INVALID'&&d.severity==='error'));
  }finally{await runtime.close();}
});
