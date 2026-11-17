import test from 'node:test';
import assert from 'node:assert/strict';
import {readInput} from '../src/html-to-pptx/input.mjs';
import {createBrowserRuntime} from '../src/html-to-pptx/browser-runtime.mjs';
import {inspectInput} from '../src/html-to-pptx/html-regions.mjs';
import {measureInput} from '../src/html-to-pptx/measure.mjs';
import {measureSelectedInput} from '../src/html-to-pptx/html-selection.mjs';
import {exportPptx} from '../src/html-to-pptx/export.mjs';
import {inspectPptxIndependently} from './inspect-pptx.mjs';
import {unzipSync,strFromU8} from 'fflate';
import {DOMParser} from '@xmldom/xmldom';
import {strToU8} from 'fflate';
import {normalizeSlideSizes} from '../src/html-to-pptx/slide-size-policy.mjs';
import {readPptx} from '../src/html-to-pptx/read-pptx.mjs';
import {exportHtml} from '../src/html-to-pptx/export-html.mjs';
import PptxGenJS from 'pptxgenjs';
import {zipSync} from 'fflate';
import {addNativeChart} from '../src/html-to-pptx/write-chart.mjs';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import {execFile} from 'node:child_process';
import {promisify} from 'node:util';
import {chromium} from 'playwright-core';
import {createHtmlServer} from '../src/html-to-pptx/server.mjs';

const doc=body=>`<!doctype html><html><body style="margin:0">${body}</body></html>`;

test('UI and CLI explain mismatched source PPT dimensions and report physical scale and margins when overridden',async t=>{
  const ppt=new PptxGenJS();ppt.defineLayout({name:'fixture',width:10,height:5.625});ppt.layout='fixture';
  ppt.addSlide().addText('Original',{x:1,y:1,w:3,h:.5,fontFace:'Arial',fontSize:18});
  const source=await readPptx(new Uint8Array(await ppt.write({outputType:'uint8array'})));
  const original=exportHtml(source.model).html,html=original.replace('width:960px;height:540px;','width:960px;height:720px;');assert.notEqual(html,original);
  const dir=await fs.mkdtemp(path.join(os.tmpdir(),'ppt-size-mismatch-')),input=path.join(dir,'source.html'),output=path.join(dir,'out.pptx');
  t.after(async()=>{await fs.unlink(input).catch(()=>{});await fs.unlink(output).catch(()=>{});await fs.rmdir(dir).catch(()=>{});});
  await fs.writeFile(input,html);
  const run=async(...args)=>{try{return (await promisify(execFile)(process.execPath,['src/html-to-pptx/cli.mjs','convert',input,'--output',output,...args,'--json'],{timeout:120000,maxBuffer:1024*1024})).stdout;}catch(error){assert.ok(error.stdout,error.message);return error.stdout;}};
  const failed=JSON.parse(await run());assert.equal(failed.status,'failed');
  assert.ok(failed.diagnostics.some(d=>d.severity==='error'&&/출력 크기/.test(d.message)),JSON.stringify(failed));
  const result=JSON.parse(await run('--slide-size','16:9'));assert.equal(result.status,'generated',JSON.stringify(result));
  assert.ok(result.diagnostics.some(d=>d.severity==='warning'&&/원본/.test(d.message)),JSON.stringify(result));
  assert.equal(result.sizePolicy.originalAspectMismatch,true);
  assert.ok(Math.abs(result.sizePolicy.slides[0].physicalScale-1)<1e-6,JSON.stringify(result.sizePolicy));
  assert.ok(Math.abs(result.sizePolicy.slides[0].offsetXEmu-160*9525)<20,JSON.stringify(result.sizePolicy));
  const app=await createHtmlServer({port:0}),browser=await chromium.launch({channel:'msedge',headless:true});
  t.after(async()=>{await browser.close();await app.close();});
  const page=await browser.newPage({viewport:{width:1280,height:900}});page.setDefaultTimeout(120000);
  await page.goto(app.url);await page.locator('#html-options').evaluate(el=>el.open=true);await page.locator('#html-file').setInputFiles(input);
  await page.waitForFunction(()=>['analyzed','failed','selection-required'].includes(document.querySelector('#status-badge').dataset.state));
  assert.equal(await page.locator('#status-badge').getAttribute('data-state'),'failed');
  assert.match(await page.locator('#diagnostic-list').textContent(),/출력 크기/);
  await page.locator('#slide-size-policy').selectOption('wide');await page.locator('#analyze-button').click();
  await page.waitForFunction(()=>['analyzed','failed','selection-required'].includes(document.querySelector('#status-badge').dataset.state));
  assert.equal(await page.locator('#status-badge').getAttribute('data-state'),'analyzed',await page.locator('#status-message').textContent());
  assert.match(await page.locator('#diagnostic-list').textContent(),/원본/);
  assert.match(await page.locator('#slide-size-report').textContent(),/원본 PPT 크기 정보와 장표 비율이 다름/);
  assert.match(await page.locator('#slide-size-list').textContent(),/100%.*1\.67인치/);
});

test('PPT with a one-EMU fractional canvas dimension round-trips without rejecting browser pixel rounding',async()=>{
  const ppt=new PptxGenJS();ppt.defineLayout({name:'fixture',width:10,height:7.5});ppt.layout='fixture';
  ppt.addSlide().addText('Fractional canvas',{x:1,y:1,w:3,h:.5,fontFace:'Arial',fontSize:18});
  const parts=unzipSync(new Uint8Array(await ppt.write({outputType:'uint8array'}))),path='ppt/presentation.xml';
  const xml=strFromU8(parts[path]),patched=xml.replace('cx="9144000"','cx="9144001"');assert.notEqual(patched,xml);parts[path]=strToU8(patched);
  const source=await readPptx(zipSync(parts));assert.equal(source.status,'parsed',JSON.stringify(source.diagnostics));
  const html=exportHtml(source.model);assert.equal(html.status,'generated',JSON.stringify(html.diagnostics));
  const runtime=await createBrowserRuntime({capture:false,checkLimits:true});
  try{
    const measured=await measureInput(await readInput({html:html.html}),runtime);
    const output=await exportPptx(measured);assert.equal(output.status,'generated',JSON.stringify(output.diagnostics));
    const actual=await inspectPptxIndependently(output.pptxBytes);assert.deepEqual([actual.widthEmu,actual.heightEmu],[9144001,6858000]);
  }finally{await runtime.close();}
});

test('resizing PPT-origin HTML scales absolute paragraph spacing and bullet size with physical font size',async()=>{
  const ppt=new PptxGenJS();ppt.defineLayout({name:'fixture',width:10,height:5.625});ppt.layout='fixture';
  ppt.addSlide().addText('Bullet item',{x:1,y:1,w:3,h:.5,fontFace:'Arial',fontSize:15,objectName:'bullet-text'});
  const parts=unzipSync(new Uint8Array(await ppt.write({outputType:'uint8array'}))),path='ppt/slides/slide1.xml';
  const xml=strFromU8(parts[path]);
  const modified=xml.replace(/<a:pPr([^>]*?)(?:\/>|>[\s\S]*?<\/a:pPr>)/,'<a:pPr$1><a:spcBef><a:spcPts val="200"/></a:spcBef><a:buSzPts val="800"/><a:buChar char="•"/></a:pPr>');
  assert.notEqual(modified,xml);parts[path]=strToU8(modified);
  const source=await readPptx(zipSync(parts));assert.equal(source.status,'parsed',JSON.stringify(source.diagnostics));
  const outputHtml=exportHtml(source.model);assert.equal(outputHtml.status,'generated',JSON.stringify(outputHtml.diagnostics));
  const smaller=outputHtml.html.replace('width:960px;height:540px;','width:480px;height:270px;');assert.notEqual(smaller,outputHtml.html);
  const runtime=await createBrowserRuntime({capture:false,checkLimits:true});
  try{
    const measured=await measureInput(await readInput({html:smaller}),runtime);
    const output=await exportPptx(measured);assert.equal(output.status,'generated',JSON.stringify(output.diagnostics));
    const data=unzipSync(output.pptxBytes),outXml=new DOMParser().parseFromString(strFromU8(data[path]),'application/xml');
    const text=[...outXml.getElementsByTagNameNS('*','sp')].find(n=>n.textContent.includes('Bullet item'));assert.ok(text);
    assert.equal(Number(text.getElementsByTagNameNS('*','rPr')[0]?.getAttribute('sz')),3000,'15pt text must become 30pt');
    assert.equal(Number(text.getElementsByTagNameNS('*','spcBef')[0]?.getElementsByTagNameNS('*','spcPts')[0]?.getAttribute('val')),400);
    assert.equal(Number(text.getElementsByTagNameNS('*','buSzPts')[0]?.getAttribute('val')),1600);
  }finally{await runtime.close();}
});

test('an extreme valid target either exports the scaled coordinates or reports an explicit error without mutating the input',async()=>{
  const html=doc('<section data-slide style="position:relative;width:10px;height:10px;background:white"><p id="tiny" style="position:absolute;left:1px;top:1px;width:8px;height:8px;margin:0;font:1px/1px Arial">A</p></section>');
  const runtime=await createBrowserRuntime({capture:false,checkLimits:true});
  try{
    const original=await measureInput(await readInput({html}),runtime);assert.ok(original.model,JSON.stringify(original.diagnostics));
    const saved=structuredClone({model:original.model,measurements:original.measurements});
    const sizePolicy={mode:'contain',targetWidthEmu:51206400,targetHeightEmu:51206400};
    const fitted=normalizeSlideSizes(original.model,original.measurements,sizePolicy);
    assert.deepEqual({model:original.model,measurements:original.measurements},saved);
    const output=await exportPptx({...original,...fitted});
    if(output.status==='generated'){
      const actual=await inspectPptxIndependently(output.pptxBytes),text=actual.slides[0].objects.find(o=>o.name==='tiny');
      assert.deepEqual([actual.widthEmu,actual.heightEmu],[51206400,51206400]);assert.ok(text);
      assert.ok(Math.abs(text.boxEmu.x-537.6*9525)<=30,JSON.stringify(text));
      assert.ok(Math.abs(text.runs[0].fontSizePt-403.2)<.1,JSON.stringify(text));
    }else{
      assert.equal(output.pptxBytes,null);
      assert.ok(output.diagnostics.some(d=>d.severity==='error'),JSON.stringify(output.diagnostics));
    }
  }finally{await runtime.close();}
});

test('resizing imported PPT retains nested groups, native table and chart geometry and paragraph units in PPT XML',async()=>{
  const ppt=new PptxGenJS();ppt.defineLayout({name:'fixture',width:10,height:7.5});ppt.layout='fixture';
  const slide=ppt.addSlide();slide.background={color:'FF0000'};
  slide.addText('Paragraph',{x:.5,y:.5,w:3,h:.7,fontSize:18,objectName:'text-shape'});
  slide.addShape(ppt.ShapeType.rect,{x:.5,y:2,w:1,h:1,fill:{color:'335577'},objectName:'group-child'});
  slide.addShape(ppt.ShapeType.line,{x:.5,y:1.5,w:2,h:0,line:{color:'112233',width:2},objectName:'line'});
  slide.addTable([['Label','Value'],['A','42']],{x:3,y:2,w:3,h:1,fontFace:'Arial',fontSize:14,objectName:'native-table'});
  addNativeChart(slide,{id:'native-chart',kind:'chart',boxCssPx:{x:700,y:100,w:360,h:250},style:{fontFamily:'Arial',fontSize:16},content:{type:'bar',categories:['A','B'],series:[{name:'Series',values:[10,20],color:'#1976d2'}],title:'Native chart',legend:true,axes:{}}},{pointScale:.75,resolveFont:()=>({fontFace:'Arial'})});
  const parts=unzipSync(new Uint8Array(await ppt.write({outputType:'uint8array'}))),path='ppt/slides/slide1.xml';
  let xml=strFromU8(parts[path]);
  const child=xml.match(/<p:sp>(?:(?!<\/p:sp>)[\s\S])*?name="group-child"[\s\S]*?<\/p:sp>/)?.[0];assert.ok(child);
  const group=(id,name,contents)=>`<p:grpSp><p:nvGrpSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="457200" y="1828800"/><a:ext cx="914400" cy="914400"/><a:chOff x="457200" y="1828800"/><a:chExt cx="914400" cy="914400"/></a:xfrm></p:grpSpPr>${contents}</p:grpSp>`;
  xml=xml.replace(child,group(900,'outer',group(901,'inner',child)));
  const title=xml.match(/<p:sp>(?:(?!<\/p:sp>)[\s\S])*?name="text-shape"[\s\S]*?<\/p:sp>/)?.[0];assert.ok(title);
  const paragraph=title.replace(/<a:pPr([^>]*?)(?:\/>|>[\s\S]*?<\/a:pPr>)/,'<a:pPr$1><a:lnSpc><a:spcPct val="100000"/></a:lnSpc><a:spcBef><a:spcPts val="200"/></a:spcBef></a:pPr>');
  assert.notEqual(paragraph,title,'fixture needs explicit absolute and relative paragraph spacing');xml=xml.replace(title,paragraph);parts[path]=strToU8(xml);
  const fixture=zipSync(parts),beforePpt=await inspectPptxIndependently(fixture);
  const parsed=await readPptx(fixture,{allowPartial:true});assert.equal(parsed.status,'parsed',JSON.stringify(parsed.diagnostics));
  assert.ok(parsed.model.slides[0].nodes.some(n=>n.kind==='table'),JSON.stringify(parsed.diagnostics));
  const html=exportHtml(parsed.model);assert.equal(html.status,'generated',JSON.stringify(html.diagnostics));
  const mixedHtml=html.html.replace('</body>','<section data-slide data-ppt-width-emu="9144000" data-ppt-height-emu="6858000" style="position:relative;width:960px;height:540px;background:blue"><p data-source-id="second-text" data-kind="text" style="position:absolute;left:10px;top:10px;width:200px;height:40px;margin:0">Second size</p></section></body>');
  assert.notEqual(mixedHtml,html.html);
  const runtime=await createBrowserRuntime({capture:false,checkLimits:true});
  try{
    const sizePolicy={mode:'contain',targetWidthEmu:1600*9525,targetHeightEmu:900*9525};
    const measured=await measureInput(await readInput({html:mixedHtml,sizePolicy}),runtime);
    assert.ok(measured.model,JSON.stringify(measured.diagnostics));
    assert.equal(measured.model.slides.length,2);
    assert.ok(measured.sizePolicy,JSON.stringify(measured.diagnostics));
    assert.notEqual(measured.sizePolicy.slides[0].scale,measured.sizePolicy.slides[1].scale);
    const prior=structuredClone(measured.model);
    const result=await exportPptx(measured);assert.equal(result.status,'generated',JSON.stringify(result.diagnostics));
    assert.deepEqual(measured.model,prior,'export must not mutate the normalized source model');
    const out=strFromU8(unzipSync(result.pptxBytes)[path]),outDoc=new DOMParser().parseFromString(out,'application/xml');
    const groups=[...outDoc.getElementsByTagNameNS('*','grpSp')],outer=groups.find(g=>g.getElementsByTagNameNS('*','cNvPr')[0]?.getAttribute('name')==='outer');
    assert.ok(outer&&[...outer.getElementsByTagNameNS('*','cNvPr')].some(n=>n.getAttribute('name')==='group-child'),'nested native groups must retain their child');
    assert.equal(outDoc.getElementsByTagNameNS('*','tbl').length,1,`table must remain native: ${JSON.stringify({nodes:measured.model.slides[0].nodes.map(n=>[n.id,n.kind]),diagnostics:measured.diagnostics})}`);
    assert.match(out,/native-chart/,'chart must remain native');
    const titleNode=[...outDoc.getElementsByTagNameNS('*','sp')].find(s=>s.textContent.includes('Paragraph'));assert.ok(titleNode);
    const spacing=titleNode.getElementsByTagNameNS('*','spcBef')[0]?.getElementsByTagNameNS('*','spcPts')[0];
    const relative=titleNode.getElementsByTagNameNS('*','lnSpc')[0]?.getElementsByTagNameNS('*','spcPct')[0];
    assert.equal(Number(spacing?.getAttribute('val')),250,'absolute paragraph spacing must scale by 1.25');
    assert.equal(Number(relative?.getAttribute('val')),100000,'relative paragraph spacing must not scale');
    const independently=await inspectPptxIndependently(result.pptxBytes),objects=independently.slides[0].objects;
    const line=objects.find(o=>o.name==='line');assert.ok(line,JSON.stringify(objects));
    assert.ok(Math.abs(line.boxEmu.x-260*9525)<=20,JSON.stringify(line));
    for(const name of ['outer','native-table','native-chart']){
      const original=beforePpt.slides[0].objects.find(o=>o.name===name),actual=objects.find(o=>o.name===name);
      assert.ok(original?.boxEmu&&actual?.boxEmu,JSON.stringify({name,original,actual}));
      for(const key of ['x','y','w','h']){
        const expected=original.boxEmu[key]*1.25+(key==='x'?200*9525:0);
        assert.ok(Math.abs(actual.boxEmu[key]-expected)<=30,JSON.stringify({name,key,expected,actual:actual.boxEmu[key]}));
      }
    }
    const sourceTable=beforePpt.slides[0].objects.find(o=>o.name==='native-table').table;
    const outputTable=objects.find(o=>o.name==='native-table').table;
    assert.deepEqual(outputTable.rows.map(r=>r.cells),sourceTable.rows.map(r=>r.cells));
    for(let i=0;i<sourceTable.columns.length;i++)assert.ok(Math.abs(outputTable.columns[i]-sourceTable.columns[i]*1.25)<=30,JSON.stringify(outputTable));
    for(let i=0;i<sourceTable.rows.length;i++)assert.ok(Math.abs(outputTable.rows[i].height-sourceTable.rows[i].height*1.25)<=30,JSON.stringify(outputTable));
  }finally{await runtime.close();}
});

test('a transformed responsive root is offered as unsupported and cannot be selected or directly converted',async()=>{
  const html=doc('<section style="width:100%;height:auto;transform:scale(.8);transform-origin:top left;background:white"><p style="height:300px;margin:0">Visible</p></section>');
  const runtime=await createBrowserRuntime({capture:false,checkLimits:true});
  try{
    const input=await readInput({html}),viewport={width:800,height:600},inspection=await inspectInput(input,runtime,{viewport});
    const region=inspection.regions.find(r=>r.bounds?.width===640&&r.bounds?.height===240);assert.ok(region,JSON.stringify(inspection.regions));
    const choice=inspection.choices.find(c=>c.scope==='elements'&&c.regionIds.length===1&&c.regionIds[0]===region.id);assert.ok(choice,JSON.stringify(inspection.choices));
    assert.equal(choice.convertible,false,JSON.stringify(choice));
    assert.ok(choice.reasons.includes('root-appearance-not-supported'),JSON.stringify(choice));
    await assert.rejects(()=>measureSelectedInput(input,runtime,{viewport,snapshotId:inspection.snapshotId,choiceId:choice.id}),{code:'REGION_UNSUPPORTED'});
    const direct=await measureInput(await readInput({html,slideSelector:'section'}),runtime,{viewport});
    assert.ok(direct.diagnostics.some(d=>d.code==='UNSUPPORTED_OBJECT'&&d.severity==='error'),JSON.stringify(direct.diagnostics));
  }finally{await runtime.close();}
});

test('a round-tripped PPT keeps its original EMU by default and retains image crop when fitted to a new size',async()=>{
  const ppt=new PptxGenJS();ppt.defineLayout({name:'fixture',width:10,height:7.5});ppt.layout='fixture';
  const slide=ppt.addSlide();slide.addText('Source',{x:1,y:1,w:3,h:.5,fontSize:20,objectName:'source-title'});
  slide.addShape(ppt.ShapeType.line,{x:1,y:3,w:3,h:0,line:{color:'FF0000',width:2},objectName:'source-line'});
  slide.addImage({x:4,y:1,w:2,h:2,objectName:'source-image',data:'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+j1ioAAAAASUVORK5CYII='});
  const parts=unzipSync(new Uint8Array(await ppt.write({outputType:'uint8array'}))),path='ppt/slides/slide1.xml';
  const xml=strFromU8(parts[path]),cropped=xml.replace('</a:blip>','</a:blip><a:srcRect l="10000" t="5000" r="15000" b="5000"/>');
  assert.notEqual(cropped,xml);parts[path]=strToU8(cropped);
  const parsed=await readPptx(zipSync(parts));assert.equal(parsed.status,'parsed',JSON.stringify(parsed.diagnostics));
  const html=exportHtml(parsed.model);assert.equal(html.status,'generated',JSON.stringify(html.diagnostics));
  const runtime=await createBrowserRuntime({capture:false,checkLimits:true});
  try{
    const original=await measureInput(await readInput({html:html.html}),runtime);
    assert.equal(original.sizePolicy.source,'original',JSON.stringify(original.sizePolicy));
    assert.deepEqual(original.model.pptxSizeEmu,{width:9144000,height:6858000});
    const ordinary=await exportPptx(original);assert.equal(ordinary.status,'generated',JSON.stringify(ordinary.diagnostics));
    const physical=await inspectPptxIndependently(ordinary.pptxBytes);assert.deepEqual([physical.widthEmu,physical.heightEmu],[9144000,6858000]);
    const target={mode:'contain',targetWidthEmu:1600*9525,targetHeightEmu:900*9525};
    const resized=await measureInput(await readInput({html:html.html,sizePolicy:target}),runtime);
    const output=await exportPptx(resized);assert.equal(output.status,'generated',JSON.stringify(output.diagnostics));
    const converted=await inspectPptxIndependently(output.pptxBytes);
    assert.deepEqual([converted.widthEmu,converted.heightEmu],[target.targetWidthEmu,target.targetHeightEmu]);
    const title=converted.slides[0].objects.find(o=>o.name==='source-title'),line=converted.slides[0].objects.find(o=>o.name==='source-line');
    assert.ok(title&&line,JSON.stringify(converted.slides[0].objects));
    assert.ok(Math.abs(title.boxEmu.x-320*9525)<=20,JSON.stringify(title));
    assert.ok(Math.abs(title.runs[0].fontSizePt-25)<.05,JSON.stringify(title.runs));
    assert.ok(Math.abs(line.boxEmu.x-320*9525)<=20,JSON.stringify(line));
    assert.ok(Math.abs(line.boxEmu.y-360*9525)<=20,JSON.stringify(line));
    const pptXml=strFromU8(unzipSync(output.pptxBytes)[path]);
    const parsedXml=new DOMParser().parseFromString(pptXml,'application/xml');
    const crops=[...parsedXml.getElementsByTagNameNS('*','srcRect')];
    assert.ok(crops.some(c=>['l','t','r','b'].every((key,i)=>Number(c.getAttribute(key))===[10000,5000,15000,5000][i])),pptXml);
  }finally{await runtime.close();}
});

test('mixed-size gradient slide background keeps its original captured pixels when fitted and exported',async()=>{
  const html=doc('<section data-slide style="position:relative;width:800px;height:600px;background:linear-gradient(90deg,red,blue)"><p style="margin:0">First</p></section><section data-slide style="position:relative;width:960px;height:540px;background:linear-gradient(90deg,red,blue)"><p style="margin:0">Second</p></section>');
  const sizePolicy={mode:'contain',targetWidthEmu:1600*9525,targetHeightEmu:900*9525};
  const runtime=await createBrowserRuntime({capture:false,checkLimits:true});
  try{
    const measured=await measureInput(await readInput({html,sizePolicy}),runtime);
    assert.ok(measured.model,JSON.stringify(measured.diagnostics));
    const slide=measured.model.slides[0],background=slide.fitBackground?.background,capture=background?.capture;
    assert.ok(capture,JSON.stringify(slide));
    assert.ok(Math.abs(capture.width-slide.fitBackground.boxCssPx.w*capture.renderScale)<=1,JSON.stringify(capture));
    assert.ok(Math.abs(capture.height-slide.fitBackground.boxCssPx.h*capture.renderScale)<=1,JSON.stringify(capture));
    const asset=measured.model.assets.find(a=>a.id===background.assetId);assert.ok(asset);
    const bytes=Buffer.from(asset.bytes);assert.equal(bytes.readUInt32BE(16),capture.width);assert.equal(bytes.readUInt32BE(20),capture.height);
    const exported=await exportPptx(measured);assert.equal(exported.status,'generated',JSON.stringify(exported.diagnostics));
    const zip=unzipSync(exported.pptxBytes);
    assert.ok(Object.entries(zip).some(([name,data])=>name.startsWith('ppt/media/')&&Buffer.from(data).equals(bytes)),'PPT must embed the unmodified captured PNG');
    const ppt=await inspectPptxIndependently(exported.pptxBytes);
    const image=ppt.slides[0].objects.find(o=>o.kind==='image');assert.ok(image,JSON.stringify(ppt.slides[0].objects));
    assert.ok(Math.abs(image.boxEmu.x-200*9525)<=20,JSON.stringify(image));
    assert.ok(Math.abs(image.boxEmu.w-1200*9525)<=20,JSON.stringify(image));
  }finally{await runtime.close();}
});

test('shared source styles and boxes are transformed independently for slides with different scales',()=>{
  const sharedStyle={fontSize:20,lineHeight:'30px',padding:{top:4,right:4,bottom:4,left:4}};
  const sharedBox={x:100,y:100,w:100,h:50};
  const make=(id,width,height)=>({id,widthCssPx:width,heightCssPx:height,background:{color:'#ffffff'},nodes:[{id:`text-${id}`,kind:'text',boxCssPx:sharedBox,style:sharedStyle,content:{text:id,runs:[{text:id,style:sharedStyle}]}}]});
  const model={slides:[make('one',800,600),make('two',960,540)]};
  const policy={mode:'contain',targetWidthEmu:1600*9525,targetHeightEmu:900*9525,marginColor:'#ffffff'};
  const result=normalizeSlideSizes(model,[],policy),first=result.model.slides[0].nodes[0],second=result.model.slides[1].nodes[0];
  assert.equal(sharedStyle.fontSize,20,'normalization must not mutate the source model');
  assert.deepEqual(sharedBox,{x:100,y:100,w:100,h:50});
  assert.ok(Math.abs(first.style.fontSize-30)<1e-6,JSON.stringify(first.style));
  assert.ok(Math.abs(second.style.fontSize-20*5/3)<1e-6,JSON.stringify(second.style));
  assert.ok(Math.abs(first.boxCssPx.x-350)<1e-6,JSON.stringify(first.boxCssPx));
  assert.ok(Math.abs(second.boxCssPx.x-100*5/3)<1e-6,JSON.stringify(second.boxCssPx));
});

test('an explicitly selectable 100-percent width and automatic height region is convertible at the observed viewport',async()=>{
  const html=doc('<section style="width:100%;height:auto;background:white"><p style="margin:0;height:300px">Page content</p></section>');
  const runtime=await createBrowserRuntime({capture:false,checkLimits:true});
  try{
    const input=await readInput({html});
    const viewport={width:800,height:600};
    const inspection=await inspectInput(input,runtime,{viewport});
    const region=inspection.regions.find(r=>r.bounds?.width===800&&r.bounds?.height===300);
    assert.ok(region,JSON.stringify(inspection.regions));
    const choice=inspection.choices.find(c=>c.scope==='elements'&&c.regionIds.length===1&&c.regionIds[0]===region.id);
    assert.ok(choice,JSON.stringify(inspection.choices));
    assert.equal(choice.convertible,true,JSON.stringify(choice));
    const measured=await measureSelectedInput(input,runtime,{viewport,snapshotId:inspection.snapshotId,choiceId:choice.id});
    assert.equal(measured.model?.slides.length,1,JSON.stringify(measured.diagnostics));
    assert.deepEqual([measured.sizePolicy?.slides?.[0]?.sourceWidth,measured.sizePolicy?.slides?.[0]?.sourceHeight],[800,300]);
    const text=measured.model.slides[0].nodes.find(n=>n.kind==='text'&&n.content?.text.includes('Page content'));
    assert.ok(text,'the explicitly selected responsive region must contain its text');
    assert.ok(measured.measurements[0].objects.some(o=>o.nodeId===text.id&&o.lines?.length));
  }finally{await runtime.close();}
});

test('an explicitly selected automatic-height document remains one slide and covers its observed content',async()=>{
  const html=doc('<main style="position:relative;width:100%;height:1200px;background:white"><p style="position:absolute;top:1100px;margin:0">Near the document end</p></main>');
  const runtime=await createBrowserRuntime({capture:false,checkLimits:true});
  try{
    const input=await readInput({html}),viewport={width:800,height:600};
    const inspection=await inspectInput(input,runtime,{viewport});
    const choice=inspection.choices.find(c=>c.scope==='document');assert.ok(choice);
    assert.equal(choice.convertible,true,JSON.stringify(choice));
    assert.ok(choice.bounds.height>=1200,JSON.stringify(choice.bounds));
    const measured=await measureSelectedInput(input,runtime,{viewport,snapshotId:inspection.snapshotId,choiceId:choice.id});
    assert.equal(measured.model?.slides.length,1,JSON.stringify(measured.diagnostics));
    assert.ok(measured.sizePolicy.slides[0].sourceHeight>=1200,JSON.stringify(measured.sizePolicy));
    assert.ok(JSON.stringify(measured.model.slides[0]).includes('Near the document end'));
  }finally{await runtime.close();}
});

test('responsive content is measured on the chosen viewport and stale viewport or changed assets are rejected',async()=>{
  const html='<!doctype html><html><head><style>@media(min-width:900px){.detail{display:none}}</style></head><body style="margin:0"><section style="position:relative;width:100%;height:auto;background:white"><p style="margin:0;height:300px">Visible <span class="detail">at selected width</span></p><img src="icon.svg" style="position:absolute;left:500px;top:20px;width:30px;height:30px"></section></body></html>';
  const asset=color=>({path:'icon.svg',bytes:strToU8(`<svg xmlns="http://www.w3.org/2000/svg" width="30" height="30"><rect width="30" height="30" fill="${color}"/></svg>`)});
  const runtime=await createBrowserRuntime({capture:false,checkLimits:true});
  try{
    const viewport={width:800,height:600};
    const input=await readInput({html,assets:[asset('red')]});
    const inspection=await inspectInput(input,runtime,{viewport});
    const region=inspection.regions.find(r=>r.bounds?.width===800&&r.bounds?.height===300);assert.ok(region);
    const choice=inspection.choices.find(c=>c.scope==='elements'&&c.regionIds.length===1&&c.regionIds[0]===region.id);assert.ok(choice?.convertible,JSON.stringify(choice));
    const options={viewport,snapshotId:inspection.snapshotId,choiceId:choice.id};
    const measured=await measureSelectedInput(input,runtime,options);
    assert.ok(JSON.stringify(measured.model).includes('at selected width'),JSON.stringify(measured.diagnostics));
    assert.deepEqual([measured.sizePolicy.slides[0].sourceWidth,measured.sizePolicy.slides[0].sourceHeight],[800,300]);
    await assert.rejects(()=>measureSelectedInput(input,runtime,{...options,viewport:{width:1000,height:600}}),{code:'STALE_INSPECTION'});
    const changed=await readInput({html,assets:[asset('blue')]});
    await assert.rejects(()=>measureSelectedInput(changed,runtime,options),{code:'STALE_INSPECTION'});
  }finally{await runtime.close();}
});

test('differently sized similar slides are offered together without being automatically selected',async()=>{
  const html=doc('<section style="position:relative;width:800px;height:600px;background:white"><p style="position:absolute;left:80px;top:60px;margin:0">First</p></section><section style="position:relative;width:960px;height:540px;background:white"><p style="position:absolute;left:96px;top:54px;margin:0">Second</p></section>');
  const input=await readInput({html});
  const runtime=await createBrowserRuntime({capture:false,checkLimits:true});
  try{
    const inspection=await inspectInput(input,runtime,{viewport:{width:1280,height:960}});
    const group=inspection.choices.find(c=>c.evidence.includes('repeated-sibling')&&c.regionIds.length===2);
    assert.ok(group,JSON.stringify(inspection.choices));
    assert.equal(group.convertible,true,JSON.stringify(group));
    const unselected=await measureInput(input,runtime);
    assert.ok(unselected.diagnostics.some(d=>d.code==='SELECTION_REQUIRED'&&d.severity==='error'),JSON.stringify(unselected.diagnostics));
    assert.equal(unselected.model,null,'a repeated group is only a candidate until the user selects it');
  }finally{await runtime.close();}
});

test('mixed slide sizes fit the explicit 1600 by 900 PPT while transforming text and its measured lines together',async()=>{
  const html=doc('<section data-slide style="position:relative;width:800px;height:600px;background:red"><p id="first-text" style="position:absolute;left:100px;top:100px;width:200px;height:60px;margin:0;font:20px/30px Arial">First slide</p></section><section data-slide style="position:relative;width:960px;height:540px;background:blue"><p id="second-text" style="position:absolute;left:96px;top:54px;width:200px;height:60px;margin:0;font:20px/30px Arial">Second slide</p></section>');
  const policy={mode:'contain',targetWidthEmu:1600*9525,targetHeightEmu:900*9525,marginColor:'#ffffff'};
  const runtime=await createBrowserRuntime({capture:false,checkLimits:true});
  try{
    const measured=await measureInput(await readInput({html,sizePolicy:policy}),runtime);
    assert.ok(measured.model,JSON.stringify(measured.diagnostics));
    assert.equal(measured.model.slides.length,2);
    assert.deepEqual(measured.model.slides.map(s=>[s.widthCssPx,s.heightCssPx]),[[1600,900],[1600,900]]);
    assert.equal(measured.sizePolicy?.mode,'contain');
    assert.deepEqual([measured.sizePolicy?.targetWidthEmu,measured.sizePolicy?.targetHeightEmu],[1600*9525,900*9525]);
    assert.equal(measured.sizePolicy?.marginColor?.toLowerCase(),'#ffffff');
    const first=measured.model.slides[0].nodes.find(n=>n.id==='first-text');assert.ok(first);
    assert.ok(Math.abs(first.boxCssPx.x-350)<0.1,JSON.stringify(first.boxCssPx));
    assert.ok(Math.abs(first.boxCssPx.y-150)<0.1,JSON.stringify(first.boxCssPx));
    assert.ok(Math.abs(first.style.fontSize-30)<0.1,JSON.stringify(first.style));
    const object=measured.measurements[0].objects.find(o=>o.nodeId==='first-text');assert.ok(object?.lines?.length&&Number.isFinite(object.baseline),JSON.stringify(object));
    assert.ok(Math.abs(object.boxCssPx.x-first.boxCssPx.x)<0.1);
    assert.ok(object.lines[0].x>=first.boxCssPx.x-0.1&&object.lines[0].x<=first.boxCssPx.x+first.boxCssPx.w+0.1);
    assert.ok(object.baseline>=first.boxCssPx.y-0.1&&object.baseline<=first.boxCssPx.y+first.boxCssPx.h+0.1);
    const exported=await exportPptx(measured);assert.equal(exported.status,'generated',JSON.stringify(exported.diagnostics));
    const ppt=await inspectPptxIndependently(exported.pptxBytes);
    assert.deepEqual([ppt.widthEmu,ppt.heightEmu],[policy.targetWidthEmu,policy.targetHeightEmu]);
    const native=ppt.slides[0].objects.find(o=>o.name==='first-text');assert.ok(native,JSON.stringify(ppt.slides[0].objects));
    assert.ok(Math.abs(native.boxEmu.x-350*9525)<=20,JSON.stringify(native));
    assert.ok(Math.abs(native.boxEmu.y-150*9525)<=20,JSON.stringify(native));
    assert.ok(Math.abs(native.runs[0].fontSizePt-22.5)<0.05,JSON.stringify(native.runs));
    const redBackground=ppt.slides[0].objects.find(o=>o.kind==='shape'&&o.fillColor==='FF0000');
    assert.ok(redBackground,'the red source background must remain inside the fitted slide');
    assert.ok(redBackground.order<native.order,'the source background must stay behind content');
    assert.ok(Math.abs(redBackground.boxEmu.x-200*9525)<=20,JSON.stringify(redBackground));
    assert.ok(Math.abs(redBackground.boxEmu.w-1200*9525)<=20,JSON.stringify(redBackground));
    assert.ok(Math.abs(redBackground.boxEmu.h-900*9525)<=20,JSON.stringify(redBackground));
    const xml=strFromU8(unzipSync(exported.pptxBytes)['ppt/slides/slide1.xml']);
    const xmlDoc=new DOMParser().parseFromString(xml,'application/xml');
    const slideBg=[...xmlDoc.getElementsByTagNameNS('*','bg')][0];
    const canvasColor=slideBg&&[...slideBg.getElementsByTagNameNS('*','srgbClr')][0]?.getAttribute('val');
    assert.equal(canvasColor,'FFFFFF','the white margin must not be filled with the red source background');
  }finally{await runtime.close();}
});
