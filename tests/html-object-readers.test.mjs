import test, {before, after} from 'node:test';
import assert from 'node:assert/strict';
import PptxGenJS from 'pptxgenjs';
import {unzipSync, strFromU8} from 'fflate';
import {DOMParser} from '@xmldom/xmldom';
import {readInput} from '../src/html-to-pptx/input.mjs';
import {measureInput} from '../src/html-to-pptx/measure.mjs';
import {createBrowserRuntime} from '../src/html-to-pptx/browser-runtime.mjs';
import {exportPptx} from '../src/html-to-pptx/export.mjs';
import {readPptx} from '../src/html-to-pptx/read-pptx.mjs';
import {publicDiagnostic} from '../src/html-to-pptx/public-diagnostics.mjs';
import {exportHtml} from '../src/html-to-pptx/export-html.mjs';
import {inspectPptxIndependently} from './inspect-pptx.mjs';

let runtime;
before(async()=>{runtime=await createBrowserRuntime({capture:false});});
after(async()=>{await runtime?.close();});

const imageBytes=Uint8Array.from(Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+aO7sAAAAASUVORK5CYII=','base64'));
const assets=[{path:'chart.png',bytes:imageBytes}];
const patternedSvg=new TextEncoder().encode('<svg xmlns="http://www.w3.org/2000/svg" width="100" height="100"><rect width="100" height="100" fill="blue"/><rect width="100" height="10" fill="red"/></svg>');
const transparentSvg=new TextEncoder().encode('<svg xmlns="http://www.w3.org/2000/svg" width="100" height="100"><circle cx="50" cy="50" r="20" fill="blue"/></svg>');
const slide=body=>`<section data-slide style="width:800px;height:600px;position:relative;background:white;font:20px Arial">${body}</section>`;
const measure=async(html,extra={})=>measureInput(await readInput({html,...extra}),runtime);
const errors=result=>result.diagnostics.filter(d=>d.severity==='error');
const detail=result=>JSON.stringify(result.diagnostics);

test('independent CSS border sides become editable flat lines in paint order',async()=>{
  const result=await measure(slide(`<div id="panel" style="position:absolute;left:10px;top:20px;width:120px;height:70px;box-sizing:border-box;background:#ddeeff;border-top:4px solid #123456;border-bottom:2px solid rgba(20,40,60,.5)"></div>
    <div id="uniform" style="position:absolute;left:60px;top:18px;width:40px;height:30px;border:3px solid #010203;background:white"></div>
    <p id="caption" style="position:absolute;left:180px;top:20px;margin:0;border-left:5px solid #445566;background:#eee;padding-left:8px">Caption</p>`));
  assert.deepEqual(errors(result),[]);
  const nodes=result.model.slides[0].nodes;
  const edges=nodes.filter(n=>n.kind==='line');
  assert.deepEqual(edges.map(n=>[n.style.strokeWidth,n.style.stroke,n.style.endpoints]),[
    [4,'rgb(18, 52, 86)',{start:{x:10,y:22},end:{x:130,y:22}}],
    [2,'rgba(20, 40, 60, 0.5)',{start:{x:10,y:89},end:{x:130,y:89}}],
    [5,'rgb(68, 85, 102)',{start:{x:182.5,y:20},end:{x:182.5,y:43}}],
  ]);
  assert.deepEqual(nodes.map(n=>n.id),['panel',edges[0].id,edges[1].id,'uniform','caption::background',edges[2].id,'caption']);
  assert.equal(nodes.find(n=>n.id==='uniform').style.strokeWidth,3);
  assert.equal(nodes.find(n=>n.id==='caption').style.padding.left,13);
  const {xml,output}=await generated(result);
  assert.equal((xml.match(/prst="line"/g)??[]).length,3);
  assert.match(xml,new RegExp(`<a:ln w="${Math.round(4*output.profile.pointScale*12700)}"[^>]*>`));
  assert.match(xml,/<a:srgbClr val="14283C"><a:alpha val="50000"\/>/i);
});

test('adjacent unequal borders trim line ends and paint separate corner triangles',async()=>{
  const result=await measure(slide('<div id="joined" style="width:120px;height:80px;box-sizing:border-box;border-top:4px solid red;border-left:8px solid blue"></div>'));
  assert.deepEqual(errors(result),[]);
  const nodes=result.model.slides[0].nodes;
  assert.deepEqual(nodes.filter(n=>n.kind==='line').map(n=>n.boxCssPx),[{x:8,y:2,w:112,h:0},{x:4,y:4,w:0,h:76}]);
  const corners=nodes.filter(n=>n.style.shape==='rightTriangle');
  assert.deepEqual(corners.map(n=>[n.boxCssPx,n.style.fill,n.style.flipH,n.style.flipV]),[
    [{x:0,y:0,w:8,h:4},'rgb(255, 0, 0)',true,true],
    [{x:0,y:0,w:8,h:4},'rgb(0, 0, 255)',false,false],
  ]);
  const {xml}=await generated(result);
  assert.equal((xml.match(/prst="rtTriangle"/g)??[]).length,2);
  assert.match(xml,/<a:xfrm flipV="1" flipH="1"|<a:xfrm flipH="1" flipV="1"/);
});

test('an unsupported asymmetric border reports the failing side and keeps the other side',async()=>{
  const result=await measure(slide('<div id="mixed" style="width:120px;height:80px;border-top:2px dashed red;border-bottom:4px solid blue"></div>'));
  const warnings=result.diagnostics.map(d=>publicDiagnostic(d,result.model.slides));
  assert.equal(warnings.length,1);
  assert.equal(warnings[0].severity,'warning');
  assert.equal(warnings[0].reasonKey,'HTML_BORDER_STYLE_UNSUPPORTED');
  assert.match(warnings[0].message,/mixed.*위쪽.*실선/);
  assert.deepEqual(result.model.slides[0].nodes.filter(n=>n.kind==='line').map(n=>n.id),['mixed::border-bottom']);
  const {xml}=await generated(result);
  assert.match(xml,/name="mixed::border-bottom"/);
  assert.doesNotMatch(xml,/name="mixed::border-top"/);
});

test('a uniformly bordered CSS circle uses an editable ellipse with separate text',async()=>{
  const result=await measure(slide('<div style="width:30px;height:30px;display:flex;align-items:center;justify-content:center;background:#223344;border:1px solid #abcdef;border-radius:50%">1</div>'));
  assert.deepEqual(errors(result),[]);
  const {xml}=await generated(result);
  assert.match(xml,/prst="ellipse"/);
  assert.match(xml,/<a:t>1<\/a:t>/);
});
async function generated(result){
  const output=await exportPptx(result);
  assert.equal(output.status,'generated',JSON.stringify(output.diagnostics));
  return {output,xml:strFromU8(unzipSync(output.pptxBytes)['ppt/slides/slide1.xml'])};
}
async function stoppedOrPreserved(result,verify){
  const output=await exportPptx(result);
  if(errors(result).length){assert.equal(output.status,'failed');return;}
  assert.equal(output.status,'generated',JSON.stringify(output.diagnostics));
  const xml=strFromU8(unzipSync(output.pptxBytes)['ppt/slides/slide1.xml']);
  await verify(result,xml,output);
}
function cropRect(xml){
  const doc=new DOMParser().parseFromString(xml,'application/xml');
  const crop=doc.getElementsByTagNameNS('*','srcRect')[0];
  assert.ok(crop,'the PPT image must specify a crop');
  return Object.fromEntries(['l','t','r','b'].map(side=>[side,Number(crop.getAttribute(side)||0)]));
}
async function imagePixels(output,xFraction=0.5,yFraction=0.5){
  const zip=unzipSync(output.pptxBytes);
  const images=Object.entries(zip).filter(([path])=>/^ppt\/media\/.*\.png$/i.test(path));
  assert.ok(images.length,'raster replacement must embed an actual PNG image');
  const session=await runtime.render(await readInput({html:slide('')}));
  try{
    return await Promise.all(images.map(async([,bytes])=>session.page.evaluate(async ({url,xFraction,yFraction})=>{
      const image=new Image();image.src=url;await image.decode();
      const canvas=document.createElement('canvas');canvas.width=image.naturalWidth;canvas.height=image.naturalHeight;
      const context=canvas.getContext('2d');context.drawImage(image,0,0);
      const x=Math.min(canvas.width-1,Math.floor(canvas.width*xFraction)),y=Math.min(canvas.height-1,Math.floor(canvas.height*yFraction));
      return [...context.getImageData(x,y,1,1).data];
    },{url:`data:image/png;base64,${Buffer.from(bytes).toString('base64')}`,xFraction,yFraction})));
  }finally{await session.context.close();}
}

test('background, styled text, and image are each read once and remain native in PPT',async()=>{
  const html=slide('<div style="position:absolute;left:20px;top:20px;width:360px;height:220px;background:lightblue"><p style="margin:0">Revenue <b>growth</b></p><img src="chart.png" style="width:40px;height:40px"></div>');
  const result=await measure(html,{assets});
  assert.deepEqual(errors(result),[],detail(result));
  const nodes=result.model.slides[0].nodes;
  assert.equal(nodes.filter(n=>n.kind==='shape').length,1,JSON.stringify(nodes));
  assert.equal(nodes.filter(n=>n.kind==='text').length,1,JSON.stringify(nodes));
  assert.equal(nodes.filter(n=>n.kind==='image').length,1,JSON.stringify(nodes));
  const text=nodes.find(n=>n.kind==='text');
  assert.equal(text.content.text,'Revenue growth');
  assert.ok(text.content.runs.some(r=>r.text.includes('growth')&&Number(r.style.fontWeight)>=700));
  const {output}=await generated(result),ppt=await inspectPptxIndependently(output.pptxBytes);
  const objects=ppt.slides[0].objects;
  assert.equal(objects.filter(o=>o.kind==='shape').length,1);
  assert.equal(objects.filter(o=>o.kind==='text').length,1);
  assert.equal(objects.filter(o=>o.kind==='image').length,1);
  assert.ok(objects.find(o=>o.kind==='text').runs.some(r=>r.text.includes('growth')&&r.bold));
});

test('native table cell text is not duplicated as a separate text object',async()=>{
  const html=slide('<table style="position:absolute;left:20px;top:20px;width:400px;table-layout:fixed;border-collapse:collapse"><tr><td>Unique cell</td><td>Second cell</td></tr></table>');
  const result=await measure(html);
  assert.deepEqual(errors(result),[],detail(result));
  const nodes=result.model.slides[0].nodes;
  assert.equal(nodes.filter(n=>n.kind==='table').length,1);
  assert.ok(!nodes.some(n=>n.kind==='text'&&n.content.text.includes('Unique cell')));
  const {output}=await generated(result),ppt=await inspectPptxIndependently(output.pptxBytes);
  assert.equal(ppt.nativeTableCount,1);
  assert.deepEqual(ppt.slides[0].objects.find(o=>o.kind==='table').table.rows[0].cells,['Unique cell','Second cell']);
  assert.ok(!ppt.slides[0].objects.some(o=>o.kind==='text'&&o.text.includes('Unique cell')));
});

test('an ordinary chart picture remains an image and missing chart data is not invented',async()=>{
  const picture=await measure(slide('<img alt="quarterly chart" src="chart.png" style="position:absolute;left:20px;top:20px;width:100px;height:100px">'),{assets});
  assert.deepEqual(errors(picture),[],detail(picture));
  assert.equal(picture.model.slides[0].nodes.filter(n=>n.kind==='image').length,1);
  assert.equal(picture.model.slides[0].nodes.filter(n=>n.kind==='chart').length,0);
  const {output}=await generated(picture),ppt=await inspectPptxIndependently(output.pptxBytes);
  assert.equal(ppt.pictureCount,1);assert.equal(ppt.nativeChartCount,0);
  const noData=await readInput({html:slide('<figure data-chart-id="missing" style="width:300px;height:200px"></figure>')});
  assert.ok(errors(noData).some(d=>d.code==='CHART_DATA_REQUIRED'),detail(noData));
});

test('overlapping ordinary objects keep their visual order and absolute positions',async()=>{
  const result=await measure(slide('<div style="position:absolute;left:20px;top:30px;width:160px;height:90px;background:#0000ff;z-index:1"></div><div style="position:absolute;left:70px;top:50px;width:160px;height:90px;background:#ff0000;z-index:2"></div>'));
  assert.deepEqual(errors(result),[],detail(result));
  const nodes=result.model.slides[0].nodes.filter(n=>n.kind==='shape');
  assert.equal(nodes.length,2);
  assert.ok(nodes[0].zOrder<nodes[1].zOrder);
  assert.deepEqual(nodes.map(n=>[n.boxCssPx.x,n.boxCssPx.y,n.boxCssPx.w,n.boxCssPx.h]),[[20,30,160,90],[70,50,160,90]]);
  const {output}=await generated(result),ppt=await inspectPptxIndependently(output.pptxBytes);
  const shapes=ppt.slides[0].objects.filter(o=>o.kind==='shape');
  assert.deepEqual(shapes.map(s=>s.fillColor),['0000FF','FF0000']);
  assert.ok(shapes[0].boxEmu.x<shapes[1].boxEmu.x);
});

test('a PPT shape crossing the slide edge survives a PPT to HTML to PPT round trip',async()=>{
  const ppt=new PptxGenJS();ppt.layout='LAYOUT_WIDE';
  const sourceSlide=ppt.addSlide();sourceSlide.background={color:'FFFFFF'};
  sourceSlide.addShape('rect',{x:12.8,y:1,w:1.2,h:0.8,fill:{color:'FF0000'},line:{color:'FF0000'},objectName:'edge-shape'});
  const sourceBytes=new Uint8Array(await ppt.write({outputType:'uint8array'}));
  const before=await inspectPptxIndependently(sourceBytes),sourceShape=before.slides[0].objects.find(o=>o.name==='edge-shape');
  assert.ok(sourceShape&&sourceShape.boxEmu.x+sourceShape.boxEmu.w>before.widthEmu,'source shape must extend past the slide');
  const parsed=await readPptx(sourceBytes);assert.equal(parsed.status,'parsed',JSON.stringify(parsed.diagnostics));
  const html=exportHtml(parsed.model);assert.equal(html.status,'generated',JSON.stringify(html.diagnostics));
  assert.match(html.html,/overflow\s*:\s*hidden/i,'generated slide clips to its own boundary');
  const result=await measure(html.html);
  assert.deepEqual(errors(result),[],detail(result));
  const {output}=await generated(result),after=await inspectPptxIndependently(output.pptxBytes);
  const actual=after.slides[0].objects.find(o=>o.name==='edge-shape');assert.ok(actual,'the original editable shape must survive');
  for(const key of ['x','w'])assert.ok(Math.abs(sourceShape.boxEmu[key]/before.widthEmu-actual.boxEmu[key]/after.widthEmu)<0.001,key);
  assert.ok(actual.boxEmu.x+actual.boxEmu.w>after.widthEmu,'the slide itself clips the protruding shape');
});

test('supported SVG blending over a solid slide remains on the SVG renderer path',async()=>{
  const png='data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jRZkAAAAASUVORK5CYII=';
  const html=`<section data-slide style="width:800px;height:600px;background:rgb(19,18,37)"><div style="mix-blend-mode:lighten"><svg width="200" height="100" viewBox="0 0 200 100" style="overflow:hidden"><image x="50" y="0" width="100" height="100" transform="rotate(45 100 50)" href="${png}"/></svg></div></section>`;
  const result=await measure(html),{output}=await generated(result);
  const ppt=await inspectPptxIndependently(output.pptxBytes);
  assert.ok(ppt.pictureCount>0,'the established SVG renderer must preserve the artwork');
});

test('a right-only border is represented or explicitly blocks export',async()=>{
  const result=await measure(slide('<div style="position:absolute;left:100px;top:100px;width:180px;height:80px;border-right:8px solid red"></div>'));
  await stoppedOrPreserved(result,async(_result,xml,output)=>{
    assert.match(xml,/FF0000/i,'the visible red border must survive in the exported PPT');
    const ppt=await inspectPptxIndependently(output.pptxBytes),unit=ppt.widthEmu/800;
    const rightEdge=ppt.slides[0].objects.some(o=>o.boxEmu&&(
      o.fillColor==='FF0000'&&Math.abs(o.boxEmu.x/unit-280)<1&&Math.abs(o.boxEmu.w/unit-8)<1&&Math.abs(o.boxEmu.h/unit-80)<1||
      o.kind==='line'&&Math.abs(o.boxEmu.x/unit-284)<1&&Math.abs(o.boxEmu.h/unit-80)<1&&Math.abs(o.lineWidthEmu/unit-8)<1
    ));
    assert.ok(rightEdge,'only the 8px right border should be rendered at the correct location');
  });
});

test('a shape clipped by an ancestor is not silently extended in PPT',async()=>{
  const result=await measure(slide('<div style="position:absolute;left:100px;top:100px;width:100px;height:50px;overflow:hidden"><div style="position:absolute;left:80px;top:0;width:100px;height:40px;background:red"></div></div>'));
  await stoppedOrPreserved(result,async(_result,_xml,output)=>{
    const ppt=await inspectPptxIndependently(output.pptxBytes),shape=ppt.slides[0].objects.find(o=>o.fillColor==='FF0000');
    assert.ok(shape,'visible red portion must be represented');
    assert.ok(shape.boxEmu.w/ppt.widthEmu<=20/800+0.002,'the 20px visible portion must not become the original 100px');
  });
});

test('inline overflow does not clip an inline-block child',async()=>{
  const html=slide('<span style="overflow:hidden;width:20px;height:20px"><span style="display:inline-block;width:100px;height:40px;background:red"></span></span>');
  const result=await measure(html);assert.deepEqual(errors(result),[],detail(result));
  const {output}=await generated(result),ppt=await inspectPptxIndependently(output.pptxBytes);
  const shape=ppt.slides[0].objects.find(o=>o.fillColor==='FF0000');assert.ok(shape);
  assert.ok(Math.abs(shape.boxEmu.w/ppt.widthEmu-100/800)<0.002,'inline overflow must not truncate the visible child');
});

test('an image clipped by an ancestor is not silently extended in PPT',async()=>{
  const result=await measure(slide('<div style="position:absolute;left:100px;top:100px;width:20px;height:20px;overflow:hidden"><img src="pattern.svg" style="position:absolute;left:0;top:0;width:100px;height:100px"></div>'),{assets:[{path:'pattern.svg',bytes:patternedSvg}]});
  await stoppedOrPreserved(result,async(_result,xml,output)=>{
    const ppt=await inspectPptxIndependently(output.pptxBytes),picture=ppt.slides[0].objects.find(o=>o.kind==='image');
    assert.ok(picture,'visible image portion must be represented');
    assert.ok(picture.boxEmu.w/ppt.widthEmu<=20/800+0.002,'the visible image must fit the 20px clip');
    const crop=cropRect(xml);
    assert.ok(Math.abs(crop.r-80000)<100&&Math.abs(crop.b-80000)<100&&crop.l===0&&crop.t===0,'the right and bottom 80% must be cropped');
  });
});

test('object-fit cover does not silently stretch the whole source image',async()=>{
  const result=await measure(slide('<img src="pattern.svg" style="position:absolute;left:20px;top:20px;width:200px;height:80px;object-fit:cover">'),{assets:[{path:'pattern.svg',bytes:patternedSvg}]});
  await stoppedOrPreserved(result,async(_result,xml)=>{
    const crop=cropRect(xml);
    assert.ok(Math.abs(crop.t-30000)<100&&Math.abs(crop.b-30000)<100&&crop.l===0&&crop.r===0,'centered cover must crop 30% from top and bottom');
  });
});

test('object-position top changes where cover crops the source',async()=>{
  const html=slide('<img src="pattern.svg" style="position:absolute;left:20px;top:20px;width:200px;height:80px;object-fit:cover;object-position:top">');
  const result=await measure(html,{assets:[{path:'pattern.svg',bytes:patternedSvg}]});
  await stoppedOrPreserved(result,async(_result,xml)=>{
    const crop=cropRect(xml);
    assert.ok(Math.abs(crop.b-60000)<100&&crop.t===0&&crop.l===0&&crop.r===0,'top alignment must crop 60% from the bottom');
  });
});

test('cover and contain preserve square images in square boxes',async()=>{
  const html=slide('<img src="pattern.svg" style="position:absolute;left:20px;top:20px;width:100px;height:100px;object-fit:cover"><img src="pattern.svg" style="position:absolute;left:150px;top:20px;width:100px;height:100px;object-fit:contain">');
  const result=await measure(html,{assets:[{path:'pattern.svg',bytes:patternedSvg}]});
  assert.deepEqual(errors(result),[],detail(result));
  const {output}=await generated(result),ppt=await inspectPptxIndependently(output.pptxBytes);
  assert.equal(ppt.pictureCount,2);
  assert.ok(ppt.slides[0].objects.filter(o=>o.kind==='image').every(o=>Math.abs(o.boxEmu.w-o.boxEmu.h)<2));
});

test('object-fit does not fill image padding with artwork',async()=>{
  const html=slide('<img src="pattern.svg" style="position:absolute;left:100px;top:100px;width:80px;height:80px;padding:10px;object-fit:cover">');
  const result=await measure(html,{assets:[{path:'pattern.svg',bytes:patternedSvg}]});
  await stoppedOrPreserved(result,async(_result,_xml,output)=>{
    const ppt=await inspectPptxIndependently(output.pptxBytes),unit=ppt.widthEmu/800;
    const pictures=ppt.slides[0].objects.filter(o=>o.kind==='image');assert.equal(pictures.length,1);
    for(const [key,value] of Object.entries({x:110,y:110,w:80,h:80}))assert.ok(Math.abs(pictures[0].boxEmu[key]/unit-value)<1,key);
  });
});

test('a transparent image keeps its CSS background underneath it',async()=>{
  const html=slide('<img src="transparent.svg" style="position:absolute;left:100px;top:100px;width:100px;height:100px;background:red">');
  const result=await measure(html,{assets:[{path:'transparent.svg',bytes:transparentSvg}]});
  await stoppedOrPreserved(result,async(_result,_xml,output)=>{
    const ppt=await inspectPptxIndependently(output.pptxBytes),objects=ppt.slides[0].objects;
    const backgrounds=objects.filter(o=>o.kind==='shape'&&o.fillColor==='FF0000'),pictures=objects.filter(o=>o.kind==='image');
    assert.equal(backgrounds.length,1,'the red background must be preserved as one editable shape');
    assert.equal(pictures.length,1,'the transparent artwork must remain a single image');
    assert.ok(backgrounds[0].order<pictures[0].order,'the background must be behind the image');
    assert.deepEqual(backgrounds[0].boxEmu,pictures[0].boxEmu,'the background must cover the same image box');
  });
});

test('an outline is represented at its outer position or explicitly blocks export',async()=>{
  const result=await measure(slide('<div style="position:absolute;left:100px;top:100px;width:100px;height:60px;background:red;outline:5px solid blue"></div>'));
  await stoppedOrPreserved(result,async(_result,xml,output)=>{
    assert.match(xml,/0000FF/i,'the blue outline must survive in the exported PPT');
    const ppt=await inspectPptxIndependently(output.pptxBytes),unit=ppt.widthEmu/800;
    assert.ok(ppt.slides[0].objects.some(o=>o.boxEmu&&Math.abs(o.lineWidthEmu/unit-5)<1&&Math.abs(o.boxEmu.x/unit-97.5)<1&&Math.abs(o.boxEmu.y/unit-97.5)<1&&Math.abs(o.boxEmu.w/unit-105)<1&&Math.abs(o.boxEmu.h/unit-65)<1),
      'an outline must lie outside the original box, rather than replacing its border');
  });
});

test('content-box background does not silently fill its padding',async()=>{
  const result=await measure(slide('<div style="position:absolute;left:100px;top:100px;width:100px;height:60px;padding:10px;background:red;background-clip:content-box"></div>'));
  await stoppedOrPreserved(result,async(_result,_xml,output)=>{
    const ppt=await inspectPptxIndependently(output.pptxBytes),unit=ppt.widthEmu/800;
    const shapes=ppt.slides[0].objects.filter(o=>o.fillColor==='FF0000');
    assert.equal(shapes.length,1,'the colored content area should appear once');
    const box=shapes[0].boxEmu;
    for(const [key,value] of Object.entries({x:110,y:110,w:100,h:60}))assert.ok(Math.abs(box[key]/unit-value)<1,key);
  });
});

test('a single rounded corner is not silently exported as four rounded corners',async()=>{
  const result=await measure(slide('<div style="position:absolute;left:100px;top:100px;width:100px;height:60px;background:red;border-radius:20px 0 0 0"></div>'));
  await stoppedOrPreserved(result,async(_result,xml,output)=>{
    const doc=new DOMParser().parseFromString(xml,'application/xml');
    const custom=[...doc.getElementsByTagNameNS('*','custGeom')];
    if(custom.length){
      const curves=custom.flatMap(node=>[...node.getElementsByTagNameNS('*','quadBezTo'),...node.getElementsByTagNameNS('*','cubicBezTo'),...node.getElementsByTagNameNS('*','arcTo')]);
      assert.equal(curves.length,1,'only the top-left corner is rounded');
    }else{
      assert.match(xml,/<p:pic\b/,'a raster replacement or a custom one-corner path is required');
      assert.ok(output.diagnostics.some(d=>d.severity==='warning'),'raster replacement must be reported');
      const left=await imagePixels(output,.02,.02),right=await imagePixels(output,.98,.02);
      assert.ok(left.some(([r,g,b,a])=>a<30||r>240&&g>240&&b>240),'the top-left corner must stay clear');
      assert.ok(right.some(([r,g,b,a])=>r>200&&g<50&&b<50&&a>220),'the top-right corner must remain square and red');
    }
  });
});

test('CSS shadow is represented or explicitly blocks export',async()=>{
  const result=await measure(slide('<div style="position:absolute;left:100px;top:100px;width:100px;height:60px;background:red;box-shadow:20px 0 0 black"></div>'));
  await stoppedOrPreserved(result,async(_result,xml)=>assert.match(xml,/<a:outerShdw\b/,'the visible CSS shadow must survive in PPT'));
});

test('multiply blending is represented or explicitly blocks export',async()=>{
  const result=await measure(slide('<div style="position:absolute;left:100px;top:100px;width:100px;height:80px;background:blue"></div><div style="position:absolute;left:100px;top:100px;width:100px;height:80px;background:red;mix-blend-mode:multiply"></div>'));
  await stoppedOrPreserved(result,async(_result,xml,output)=>{
    assert.match(xml,/<p:pic\b/,'unsupported blend needs an explicit visual representation');
    assert.ok(output.diagnostics.some(d=>d.severity==='warning'),'raster replacement must be reported');
    const pixels=await imagePixels(output);
    assert.ok(pixels.some(([r,g,b,a])=>r<30&&g<30&&b<30&&a>220),'red multiplied over blue should produce opaque black');
  });
});

test('CSS mask is represented or explicitly blocks export',async()=>{
  const result=await measure(slide('<div style="position:absolute;left:100px;top:100px;width:100px;height:80px;background:red;mask-image:linear-gradient(black,transparent)"></div>'));
  await stoppedOrPreserved(result,async(_result,xml,output)=>{
    assert.match(xml,/<p:pic\b/,'the visible mask must survive in PPT');
    assert.ok(output.diagnostics.some(d=>d.severity==='warning'),'raster replacement must be reported');
    const pixels=await imagePixels(output);
    assert.ok(pixels.some(([r,g,b,a])=>r>200&&(a>40&&a<220&&g<40&&b<40||a>=220&&g>40&&g<220&&Math.abs(g-b)<10)),
      'the center must preserve the partial alpha or the corresponding blended color');
  });
});
