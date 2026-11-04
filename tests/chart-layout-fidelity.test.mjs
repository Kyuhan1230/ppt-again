import test from 'node:test';
import assert from 'node:assert/strict';
import PptxGenJS from 'pptxgenjs';
import {chromium} from 'playwright-core';
import {unzipSync,strFromU8} from 'fflate';
import {renderChartSvg,validateChart} from '../src/html-to-pptx/charts.mjs';
import {addNativeChart} from '../src/html-to-pptx/write-chart.mjs';
import {readPptx} from '../src/html-to-pptx/read-pptx.mjs';
import {exportHtml} from '../src/html-to-pptx/export-html.mjs';
import {readInput} from '../src/html-to-pptx/input.mjs';
import {measureInput} from '../src/html-to-pptx/measure.mjs';
import {createBrowserRuntime} from '../src/html-to-pptx/browser-runtime.mjs';
import {exportPptx} from '../src/html-to-pptx/export.mjs';

const chart = (gapWidth = 33, multiple = false) => ({
  type:'bar', categories:['01월','02월','03월','04월'],
  series:[{name:'First',values:[5500,7500,9500,12500],color:'#495ed6'},...(multiple?[{name:'Second',values:[6000,7000,9000,11000],color:'#164e63'}]:[])],
  title:'',legend:false,axes:{min:0,max:15000,majorUnit:5000},gapWidth,
});
function bars(content) {
  const svg=renderChartSvg(content,{width:500,height:180,fontFamily:'Malgun Gothic',fontSize:13});
  return [...svg.matchAll(/<rect data-series-index="(\d+)" data-point-index="(\d+)" x="([\d.]+)" y="([\d.]+)" width="([\d.]+)" height="([\d.]+)"/g)].map(m=>({series:+m[1],point:+m[2],x:+m[3],y:+m[4],width:+m[5],height:+m[6]}));
}

test('automatic plot area matches the independently measured PowerPoint example',()=>{
  const first=bars(chart())[0];
  // The 500 x 180 reference was measured from a PowerPoint slide export.
  assert.ok(Math.abs(first.x-71)<3,`bar x: ${first.x}`);
  assert.ok(Math.abs(first.width-80)<3,`bar width: ${first.width}`);
  assert.ok(Math.abs(first.y+first.height-139)<3,`baseline: ${first.y+first.height}`);
});

test('PowerPoint uses less vertical space for Latin category labels',()=>{
  const first=bars({...chart(),categories:['A','B','C','D']})[0];
  assert.ok(Math.abs(first.y+first.height-144)<3,`baseline: ${first.y+first.height}`);
});

test('negative values cross the horizontal axis at zero and use compact tick labels',()=>{
  const content={...chart(),categories:['A','B','C','D'],series:[{name:'First',values:[-100,0,0,100],color:'#495ed6'}],axes:{min:-100,max:100,majorUnit:200/3}};
  const svg=renderChartSvg(content,{width:500,height:180,fontFamily:'Malgun Gothic',fontSize:13});
  assert.match(svg,/>-33\.33333<\/text>/);
  const bar=bars(content)[0],path=svg.match(/<path d="M[\d.]+ [\d.]+ V[\d.]+ M[\d.]+ ([\d.]+) H[\d.]+"/);
  assert.ok(path,'vertical and horizontal axes must be positioned independently');
  assert.ok(Math.abs(Number(path[1])-bar.y)<.01,'negative bar begins at the zero axis');
});

test('titles and a single legend are positioned near the measured PowerPoint locations',()=>{
  const options={width:500,height:180,fontFamily:'Malgun Gothic',fontSize:13},base={...chart(),categories:['A','B','C','D']};
  const title=renderChartSvg({...base,title:'Revenue'},options).match(/<text x="250" y="([\d.]+)" text-anchor="middle">Revenue<\/text>/);
  assert.ok(title);assert.ok(Math.abs(Number(title[1])-27)<3,`title y: ${title[1]}`);
  const legend=renderChartSvg({...base,series:[{...base.series[0],name:'Series'}],legend:true},options).match(/<rect x="([\d.]+)" y="([\d.]+)" width="[\d.]+" height="[\d.]+" fill="#495ed6"\/>/);
  assert.ok(legend);assert.ok(Math.abs(Number(legend[1])-230)<5,`legend x: ${legend[1]}`);
  assert.ok(Math.abs(Number(legend[2])-156)<3,`legend y: ${legend[2]}`);
  const titles=renderChartSvg({...base,axes:{...base.axes,xTitle:'Month',yTitle:'Amount'}},options);
  const vertical=titles.match(/<text x="([\d.]+)"[^>]*>Amount<\/text>/),horizontal=titles.match(/<text x="[\d.]+" y="([\d.]+)"[^>]*>Month<\/text>/);
  assert.ok(vertical&&horizontal);assert.ok(Math.abs(Number(vertical[1])-29)<3);assert.ok(Math.abs(Number(horizontal[1])-158)<4);
});

test('gap width is a percentage of an individual bar, also for several series',()=>{
  for(const gap of [0,33,50,150,500])for(const multiple of [false,true]){
    const rects=bars(chart(gap,multiple));
    const first=rects.find(r=>r.series===0&&r.point===0),next=rects.find(r=>r.series===0&&r.point===1);
    const seriesCount=multiple?2:1;
    assert.ok(Math.abs((next.x-first.x-first.width*seriesCount)/first.width-gap/100)<1e-5,`${gap}: ${JSON.stringify(rects)}`);
  }
});

test('zero-gap bars do not show antialiased seams in the browser',async()=>{
  const browser=await chromium.launch({channel:'msedge',headless:true});
  try{
    const page=await browser.newPage({viewport:{width:600,height:220},deviceScaleFactor:1});
    await page.setContent(`<style>body{margin:0;background:white}</style>${renderChartSvg(chart(0),{width:500,height:180,fontFamily:'Malgun Gothic',fontSize:13})}`);
    const svg=page.locator('svg'),bar=page.locator('rect[data-series-index="0"][data-point-index="0"]');
    const [frame,box,png]=await Promise.all([svg.boundingBox(),bar.boundingBox(),svg.screenshot()]);
    const x=Math.floor(box.x+box.width-frame.x),y=Math.floor(box.y+box.height-frame.y-5);
    const pixel=await page.evaluate(async({data,x,y})=>{const image=new Image();image.src=`data:image/png;base64,${data}`;await image.decode();const canvas=document.createElement('canvas');canvas.width=image.width;canvas.height=image.height;const ctx=canvas.getContext('2d');ctx.drawImage(image,0,0);return [...ctx.getImageData(x,y,1,1).data];},{data:png.toString('base64'),x,y});
    assert.deepEqual(pixel,[73,94,214,255]);
  }finally{await browser.close();}
});

test('explicit bar spacing survives native PowerPoint export and import',async()=>{
  for(const gap of [0,50,150,500]){
    const ppt=new PptxGenJS(),slide=ppt.addSlide();slide.background={color:'FFFFFF'};
    addNativeChart(slide,{id:'spacing',kind:'chart',boxCssPx:{x:40,y:40,w:500,h:180},style:{fontFamily:'Arial',fontSize:13},content:chart(gap)},{pointScale:.75,resolveFont:()=>({fontFace:'Arial'})});
    const bytes=new Uint8Array(await ppt.write({outputType:'uint8array'})),parts=unzipSync(bytes),chartPath=Object.keys(parts).find(name=>/^ppt\/charts\/chart\d+\.xml$/.test(name)),xml=strFromU8(parts[chartPath]);
    assert.match(xml,new RegExp(`<c:gapWidth val="${gap}"\\s*/>`));
    const result=await readPptx(bytes,{allowPartial:true});assert.equal(result.status,'parsed',JSON.stringify(result.diagnostics));
    assert.equal(result.model.slides[0].nodes.find(n=>n.kind==='chart').content.gapWidth,gap);
  }
});

test('bar spacing and chart data survive a PPTX to HTML to PPTX round trip',async()=>{
  const ppt=new PptxGenJS(),slide=ppt.addSlide();slide.background={color:'FFFFFF'};
  addNativeChart(slide,{id:'round-trip',kind:'chart',boxCssPx:{x:40,y:40,w:500,h:180},style:{fontFamily:'Arial',fontSize:13},content:chart(50)},{pointScale:.75,resolveFont:()=>({fontFace:'Arial'})});
  const parsed=await readPptx(new Uint8Array(await ppt.write({outputType:'uint8array'})),{allowPartial:true});
  assert.equal(parsed.status,'parsed',JSON.stringify(parsed.diagnostics));
  const input=await readInput({html:exportHtml(parsed.model).html});
  const runtime=await createBrowserRuntime({capture:false});
  try{
    const measured=await measureInput(input,runtime),output=await exportPptx(measured);
    assert.equal(output.status,'generated',JSON.stringify(output.diagnostics));
    const roundTrip=await readPptx(output.pptxBytes,{allowPartial:true});
    assert.equal(roundTrip.status,'parsed',JSON.stringify(roundTrip.diagnostics));
    const result=roundTrip.model.slides[0].nodes.find(node=>node.kind==='chart');
    assert.equal(result.content.gapWidth,50);
    assert.deepEqual(result.content.series[0].values,[5500,7500,9500,12500]);
  }finally{await runtime.close();}
});

test('bar spacing is bounded and cannot be applied to other chart types',()=>{
  for(const gap of [-1,501,2.5,'50'])assert.equal(validateChart(chart(gap)).valid,false);
  assert.equal(validateChart({...chart(),type:'line'}).valid,false);
});
