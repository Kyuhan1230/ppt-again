import test, {after} from 'node:test';
import assert from 'node:assert/strict';
import {unzipSync, zipSync, strFromU8, strToU8} from 'fflate';
import PptxGenJS from 'pptxgenjs';
import {DOMParser} from '@xmldom/xmldom';
import {readInput} from '../src/html-to-pptx/input.mjs';
import {measureInput} from '../src/html-to-pptx/measure.mjs';
import {createBrowserRuntime} from '../src/html-to-pptx/browser-runtime.mjs';
import {exportHtml} from '../src/html-to-pptx/export-html.mjs';
import {exportPptx} from '../src/html-to-pptx/export.mjs';
import {readPptx} from '../src/html-to-pptx/read-pptx.mjs';
import {resolveAsset,sha256} from '../src/html-to-pptx/assets.mjs';
import {addNativeChart} from '../src/html-to-pptx/write-chart.mjs';
import {publicDiagnostic} from '../src/html-to-pptx/public-diagnostics.mjs';

let runtime;
async function browser(){return runtime??=await createBrowserRuntime({capture:false});}
after(async()=>{await runtime?.close();});

const slide=body=>`<section data-slide style="position:relative;width:320px;height:180px;background:white">${body}</section>`;
const errors=result=>result.diagnostics.filter(d=>d.severity==='error');
async function measured(html){return measureInput(await readInput({html}),await browser());}
async function slidePixels(html,points){
  const input=await readInput({html});assert.deepEqual(errors(input),[],JSON.stringify(input.diagnostics));
  const session=await (await browser()).render(input);
  try{
    const shot=await session.page.locator('[data-slide]').first().screenshot({animations:'disabled',scale:'css'});
    return await session.page.evaluate(async ({data,points})=>{
      const image=new Image();image.src='data:image/png;base64,'+data;await image.decode();
      const canvas=document.createElement('canvas');canvas.width=image.width;canvas.height=image.height;
      const ctx=canvas.getContext('2d');ctx.drawImage(image,0,0);
      return points.map(([x,y])=>[...ctx.getImageData(x,y,1,1).data]);
    },{data:Buffer.from(shot).toString('base64'),points});
  }finally{await session.context.close();}
}
function assertPixelsClose(actual,expected){
  assert.equal(actual.length,expected.length);
  for(let i=0;i<actual.length;i++)for(let channel=0;channel<4;channel++)
    assert.ok(Math.abs(actual[i][channel]-expected[i][channel])<=8,`pixel ${i} channel ${channel}: got ${actual[i]}, expected ${expected[i]}`);
}

for(const [label,html,points,target] of [
  ['slide', '<section data-slide style="position:relative;width:320px;height:180px;background:linear-gradient(90deg,red,blue)"></section>', [[50,90],[270,90]],m=>m.slides[0].background],
  ['shape', slide('<div style="position:absolute;left:40px;top:40px;width:240px;height:100px;background:linear-gradient(90deg,red,blue)"></div>'), [[70,90],[250,90]],m=>m.slides[0].nodes.find(n=>n.kind==='shape')?.style],
  ['slide translucent over color', '<section data-slide style="position:relative;width:320px;height:180px;background-color:#31a864;background-image:linear-gradient(90deg,rgba(255,0,0,.4),rgba(0,0,255,.4))"></section>', [[50,90],[270,90]],m=>m.slides[0].background],
  ['rounded shape translucent over color', slide('<div style="position:absolute;left:40px;top:40px;width:240px;height:100px;box-sizing:border-box;border:2px solid black;border-radius:30px;background-color:#31a864;background-image:linear-gradient(90deg,rgba(255,0,0,.4),rgba(0,0,255,.4))"></div>'), [[42,42],[70,90],[250,90]],m=>m.slides[0].nodes.find(n=>n.kind==='shape')?.style],
  ['fractional-size shape', slide('<div style="position:absolute;left:40px;top:40px;width:239.5px;height:99.25px;background:linear-gradient(90deg,red,blue)"></div>'), [[70,90],[250,90]],m=>m.slides[0].nodes.find(n=>n.kind==='shape')?.style]
])test(`${label} gradient background remains visible in HTML and remains available to PPT`,async()=>{
  const input=await measured(html);assert.deepEqual(errors(input),[],JSON.stringify(input.diagnostics));
  assert.ok(target(input.model)?.assetId,'the measured model must carry captured background artwork');
  const ppt=await exportPptx(input);assert.equal(ppt.status,'generated',JSON.stringify(ppt.diagnostics));
  assert.ok(!ppt.diagnostics.some(d=>d.code==='BACKGROUND_JPEG_UNVERIFIED'),'PNG artwork must not get a JPEG warning');
  const xml=strFromU8(unzipSync(ppt.pptxBytes)['ppt/slides/slide1.xml']);
  assert.match(xml,/<p:pic>/,'PPT must contain the rendered background picture');
  const output=exportHtml(input.model);assert.equal(output.status,'generated',JSON.stringify(output.diagnostics));
  assert.ok(!output.diagnostics.some(d=>d.code==='BACKGROUND_JPEG_UNVERIFIED'),'PNG artwork must not get a JPEG warning');
  const originalPixels=await slidePixels(html,points),outputPixels=await slidePixels(output.html,points);
  assertPixelsClose(outputPixels,originalPixels);
  const roundtrip=await measured(output.html);assert.deepEqual(errors(roundtrip),[],JSON.stringify(roundtrip.diagnostics));
  const roundtripPpt=await exportPptx(roundtrip);assert.equal(roundtripPpt.status,'generated',JSON.stringify(roundtripPpt.diagnostics));
  const roundtripXml=strFromU8(unzipSync(roundtripPpt.pptxBytes)['ppt/slides/slide1.xml']);
  assert.match(roundtripXml,/<p:pic>/,'reimported PPT must still contain the background artwork');
});

async function simpleModel(){
  const ppt=new PptxGenJS();ppt.layout='LAYOUT_WIDE';
  const page=ppt.addSlide();page.background={color:'FFFFFF'};
  page.addShape(ppt.ShapeType.rect,{x:1,y:1,w:2,h:1,fill:{color:'FF0000'},line:{transparency:100},objectName:'background-target'});
  const parsed=await readPptx(new Uint8Array(await ppt.write({outputType:'uint8array'})));
  assert.equal(parsed.status,'parsed',JSON.stringify(parsed.diagnostics));return parsed.model;
}
const validPng='data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADElEQVR4nGOQTDkBAAHfAUZ76qvRAAAAAElFTkSuQmCC';
const transparentArt='data:image/svg+xml;base64,'+Buffer.from('<svg xmlns="http://www.w3.org/2000/svg" width="100" height="100"><circle cx="50" cy="50" r="30" fill="blue"/></svg>').toString('base64');
const unsafeSvg={
  'svg-import':'<svg xmlns="http://www.w3.org/2000/svg"><style>@import "https://example.invalid/x.css";</style></svg>',
  'svg-escaped-import':String.raw`<svg xmlns="http://www.w3.org/2000/svg"><style>@\69 mport "https://example.invalid/x.css";</style></svg>`,
  'svg-namespaced-href':'<svg xmlns="http://www.w3.org/2000/svg" xmlns:q="http://www.w3.org/1999/xlink"><image q:href="https://example.invalid/x.png"/></svg>',
  'svg-external-base':'<svg xmlns="http://www.w3.org/2000/svg" xml:base="https://example.invalid/"><defs><rect id="x" width="1" height="1"/></defs><use href="#x"/></svg>'
};
for(const target of ['slide','shape'])test(`${target} background accepts SVG references to its own gradient and shapes`,async()=>{
  const svg='<svg xmlns="http://www.w3.org/2000/svg" xmlns:q="http://www.w3.org/1999/xlink" width="100" height="100"><defs><linearGradient id="g"><stop stop-color="red"/><stop offset="1" stop-color="blue"/></linearGradient><rect id="r" width="100" height="100" fill="url(#g)"/></defs><use q:href="#r"/></svg>';
  const model=await simpleModel(),asset=resolveAsset('data:image/svg+xml;base64,'+Buffer.from(svg).toString('base64'));model.assets.push(asset);
  if(target==='slide')model.slides[0].background.assetId=asset.id;else model.slides[0].nodes.find(node=>node.kind==='shape').style.assetId=asset.id;
  const html=exportHtml(model);assert.equal(html.status,'generated',JSON.stringify(html.diagnostics));
  const slideId=model.slides[0].id,ppt=await exportPptx({model,environment:{fonts:[]},diagnostics:[],measurements:[{slideId,assetsReady:true,fontsReady:true,layoutStable:true,objects:[]}]});
  assert.equal(ppt.status,'generated',JSON.stringify(ppt.diagnostics));
});
for(const target of ['slide','shape'])test(`${target} background artwork preserves its underlying color through transparent pixels`,async()=>{
  const model=await simpleModel(),art=resolveAsset(transparentArt);model.assets.push(art);
  if(target==='slide'){model.slides[0].background.color='#ff0000';model.slides[0].background.assetId=art.id;}
  else {const shape=model.slides[0].nodes.find(node=>node.kind==='shape');shape.style.fill='#ff0000';shape.style.assetId=art.id;}
  const output=exportHtml(model);assert.equal(output.status,'generated',JSON.stringify(output.diagnostics));
  const reference=target==='slide'
    ?`<section data-slide style="width:1280px;height:720px;background-color:red;background-image:url('${transparentArt}');background-size:100% 100%"></section>`
    :`<section data-slide style="position:relative;width:1280px;height:720px;background:white"><div style="position:absolute;left:96px;top:96px;width:192px;height:96px;background-color:red;background-image:url('${transparentArt}');background-size:100% 100%"></div></section>`;
  const points=target==='slide'?[[10,10],[640,360]]:[[100,100],[192,144]];
  assertPixelsClose(await slidePixels(output.html,points),await slidePixels(reference,points));
});

test('transparent background artwork preserves the shape fill opacity in both HTML and PPT',async()=>{
  const model=await simpleModel(),art=resolveAsset(transparentArt);model.assets.push(art);
  const shape=model.slides[0].nodes.find(node=>node.kind==='shape');
  shape.style.fill='#ff0000';shape.style.fillOpacity=.25;shape.style.assetId=art.id;
  const html=exportHtml(model);assert.equal(html.status,'generated',JSON.stringify(html.diagnostics));
  const reference=`<section data-slide style="position:relative;width:1280px;height:720px;background:white"><div style="position:absolute;left:96px;top:96px;width:192px;height:96px;background-color:rgba(255,0,0,.25);background-image:url('${transparentArt}');background-size:100% 100%"></div></section>`;
  assertPixelsClose(await slidePixels(html.html,[[100,100],[192,144]]),await slidePixels(reference,[[100,100],[192,144]]));
  const slideId=model.slides[0].id;
  const ppt=await exportPptx({model,environment:{fonts:[]},diagnostics:[],measurements:[{slideId,assetsReady:true,fontsReady:true,layoutStable:true,objects:[]}]});
  assert.equal(ppt.status,'generated',JSON.stringify(ppt.diagnostics));
  const doc=new DOMParser().parseFromString(strFromU8(unzipSync(ppt.pptxBytes)['ppt/slides/slide1.xml']),'application/xml');
  const colors=[...doc.getElementsByTagNameNS('*','srgbClr')].filter(color=>color.getAttribute('val')==='FF0000');
  assert.ok(colors.some(color=>[...color.getElementsByTagNameNS('*','alpha')].some(alpha=>Number(alpha.getAttribute('val'))===25000)),'the native red fill must retain 25% opacity');
});

for(const property of ['fill','stroke'])test(`ordinary shape ${property} opacity matches its PPT color in HTML`,async()=>{
  const model=await simpleModel(),shape=model.slides[0].nodes.find(node=>node.kind==='shape');
  shape.style.fill=property==='fill'?'#ff0000':'transparent';shape.style.fillOpacity=.25;
  shape.style.stroke=property==='stroke'?'#0000ff':'transparent';shape.style.strokeWidth=property==='stroke'?8:0;shape.style.strokeOpacity=.25;
  const html=exportHtml(model);assert.equal(html.status,'generated',JSON.stringify(html.diagnostics));
  const css=property==='fill'?'background:rgba(255,0,0,.25)':'border:8px solid rgba(0,0,255,.25);box-sizing:border-box';
  const reference=`<section data-slide style="position:relative;width:1280px;height:720px;background:white"><div style="position:absolute;left:96px;top:96px;width:192px;height:96px;${css}"></div></section>`;
  const point=property==='fill'?[192,144]:[100,144];
  assertPixelsClose(await slidePixels(html.html,[point]),await slidePixels(reference,[point]));
  const slideId=model.slides[0].id,ppt=await exportPptx({model,environment:{fonts:[]},diagnostics:[],measurements:[{slideId,assetsReady:true,fontsReady:true,layoutStable:true,objects:[]}]});
  assert.equal(ppt.status,'generated',JSON.stringify(ppt.diagnostics));
  const doc=new DOMParser().parseFromString(strFromU8(unzipSync(ppt.pptxBytes)['ppt/slides/slide1.xml']),'application/xml');
  const rgb=property==='fill'?'FF0000':'0000FF';
  assert.ok([...doc.getElementsByTagNameNS('*','srgbClr')].some(color=>color.getAttribute('val')===rgb&&[...color.getElementsByTagNameNS('*','alpha')].some(alpha=>Number(alpha.getAttribute('val'))===25000)),'PPT color must retain 25% opacity');
});

for(const property of ['fill','stroke'])for(const explicit of [false,true])test(`${explicit?'explicit model':'original PPT'} text-shape ${property} opacity is retained in HTML`,async()=>{
  const ppt=new PptxGenJS();ppt.layout='LAYOUT_WIDE';const page=ppt.addSlide();page.background={color:'FFFFFF'};
  page.addText('X',{x:1,y:1,w:2,h:1,fontSize:12,objectName:'original-text-shape',fill:{color:'FF0000',transparency:property==='fill'?75:100},line:{color:'0000FF',width:6,transparency:property==='stroke'?75:100}});
  const parsed=await readPptx(new Uint8Array(await ppt.write({outputType:'uint8array'})));
  assert.equal(parsed.status,'parsed',JSON.stringify(parsed.diagnostics));
  const original=parsed.model.slides[0].nodes.find(node=>node.id==='original-text-shape');assert.ok(original,'the source text-shape must remain one node');
  if(explicit){original.style.fill=property==='fill'?'#ff0000':'transparent';original.style.fillOpacity=.25;original.style.stroke=property==='stroke'?'#0000ff':'transparent';original.style.strokeOpacity=.25;}
  const html=exportHtml(parsed.model);assert.equal(html.status,'generated',JSON.stringify(html.diagnostics));
  const css=property==='fill'?'background:rgba(255,0,0,.25)':'border:8px solid rgba(0,0,255,.25);box-sizing:border-box';
  const reference=`<section data-slide style="position:relative;width:1280px;height:720px;background:white;transform:scale(min(1,calc((100vw - 32px)/1280px)));transform-origin:0 0"><div style="position:absolute;left:96px;top:96px;width:192px;height:96px;${css}"></div></section>`;
  const point=property==='fill'?[192,170]:[94,144];
  assertPixelsClose(await slidePixels(html.html,[point]),await slidePixels(reference,[point]));
});

test('PPT preserves or explicitly rejects the fill behind an uncomposited transparent shape background',async()=>{
  const model=await simpleModel(),art=resolveAsset(transparentArt);model.assets.push(art);
  const shape=model.slides[0].nodes.find(node=>node.kind==='shape');shape.style.fill='#ff0000';shape.style.assetId=art.id;
  const slideId=model.slides[0].id;
  const output=await exportPptx({model,environment:{fonts:[]},diagnostics:[],measurements:[{slideId,assetsReady:true,fontsReady:true,layoutStable:true,objects:[]}]});
  if(output.status==='failed'){assert.ok(output.diagnostics.some(d=>d.severity==='error'));return;}
  const doc=new DOMParser().parseFromString(strFromU8(unzipSync(output.pptxBytes)['ppt/slides/slide1.xml']),'application/xml');
  const tree=doc.getElementsByTagNameNS('*','spTree')[0],drawables=[...tree.childNodes].filter(node=>node.nodeType===1&&['sp','pic'].includes(node.localName));
  const red=drawables.findIndex(node=>{
    if(node.localName!=='sp')return false;
    const pr=[...node.childNodes].find(child=>child.nodeType===1&&child.localName==='spPr');
    const fill=[...(pr?.childNodes??[])].find(child=>child.nodeType===1&&child.localName==='solidFill');
    const color=fill?.getElementsByTagNameNS('*','srgbClr')[0];
    return color?.getAttribute('val')==='FF0000'&&!Array.from(color.getElementsByTagNameNS('*','alpha')).some(alpha=>Number(alpha.getAttribute('val'))===0);
  });
  const picture=drawables.findIndex(node=>node.localName==='pic');
  assert.ok(red>=0&&picture>red,'the opaque red underlay must be behind the transparent background image');
});

for(const unsupported of ['group-opacity','rounded-shape','ellipse'])
  test(`PPT explicitly rejects uncomposited transparent art with ${unsupported}`,async()=>{
    const model=await simpleModel(),art=resolveAsset(transparentArt);model.assets.push(art);
    const shape=model.slides[0].nodes.find(node=>node.kind==='shape');shape.style.fill='#ff0000';shape.style.assetId=art.id;
    if(unsupported==='group-opacity')shape.style.opacity=.5;
    if(unsupported==='rounded-shape')shape.style.radius=20;
    if(unsupported==='ellipse')shape.style.shape='ellipse';
    const slideId=model.slides[0].id;
    const output=await exportPptx({model,environment:{fonts:[]},diagnostics:[],measurements:[{slideId,assetsReady:true,fontsReady:true,layoutStable:true,objects:[]}]});
    assert.equal(output.status,'failed',`${unsupported} must not silently change composite appearance`);
    assert.ok(output.diagnostics.some(d=>d.severity==='error'));
  });

test('PPT does not add an extra fill under already composited shape artwork',async()=>{
  const model=await simpleModel(),art=resolveAsset(validPng);model.assets.push(art);
  const shape=model.slides[0].nodes.find(node=>node.kind==='shape');shape.style.fill='#ff0000';shape.style.assetId=art.id;
  shape.boxCssPx.w=1;shape.boxCssPx.h=1;
  shape.style.capture={width:1,height:1,renderScale:1,content:'computed-background-only',sourceChildrenCopied:0};
  const slideId=model.slides[0].id;
  const output=await exportPptx({model,environment:{fonts:[]},diagnostics:[],measurements:[{slideId,assetsReady:true,fontsReady:true,layoutStable:true,objects:[]}]});
  assert.equal(output.status,'generated',JSON.stringify(output.diagnostics));
  const artEntries=output.objectMap.filter(item=>item.derivedFrom===shape.id);
  assert.equal(artEntries.length,1,'the capture is already composited and needs just one derived picture');
});

test('captured shape artwork does not apply fill opacity a second time',async()=>{
  const html=slide('<div style="position:absolute;left:40px;top:40px;width:240px;height:100px;background-color:rgba(255,0,0,.5);background-image:linear-gradient(90deg,rgba(0,0,255,.5),rgba(0,0,255,.5))"></div>');
  const input=await measured(html);assert.deepEqual(errors(input),[],JSON.stringify(input.diagnostics));
  const shape=input.model.slides[0].nodes.find(node=>node.kind==='shape'&&node.style.capture);
  assert.ok(shape,'the background must be captured as composited artwork');shape.style.fillOpacity=.5;
  const output=exportHtml(input.model);assert.equal(output.status,'generated',JSON.stringify(output.diagnostics));
  assertPixelsClose(await slidePixels(output.html,[[100,90],[200,90]]),await slidePixels(html,[[100,90],[200,90]]));
  const ppt=await exportPptx(input);assert.equal(ppt.status,'generated',JSON.stringify(ppt.diagnostics));
  assert.equal(ppt.objectMap.filter(item=>item.derivedFrom===shape.id).length,1,'the composited capture requires just one picture');
});

test('captured translucent slide background is not composed twice in PPT',async()=>{
  const html='<section data-slide style="position:relative;width:320px;height:180px;background-color:rgba(255,0,0,.5);background-image:linear-gradient(90deg,rgba(0,0,255,.5),rgba(0,0,255,.5))"></section>';
  const input=await measured(html);assert.deepEqual(errors(input),[],JSON.stringify(input.diagnostics));
  assert.equal(input.model.slides[0].background.capture?.content,'computed-background-only');
  const output=await exportPptx(input);assert.equal(output.status,'generated',JSON.stringify(output.diagnostics));
  const xml=strFromU8(unzipSync(output.pptxBytes)['ppt/slides/slide1.xml']);
  const doc=new DOMParser().parseFromString(xml,'application/xml');
  const bg=doc.getElementsByTagNameNS('*','bg')[0],color=bg?.getElementsByTagNameNS('*','srgbClr')[0];
  assert.notEqual(color?.getAttribute('val'),'FF0000','captured red background must not be painted a second time under its own PNG');
  assert.match(xml,/<p:pic>/,'the captured background remains an actual picture');
});

test('a noncaptured slide color is preserved beneath transparent PPT background art',async()=>{
  const model=await simpleModel(),art=resolveAsset(transparentArt);model.assets.push(art);
  model.slides[0].background.color='#ff0000';model.slides[0].background.assetId=art.id;
  const slideId=model.slides[0].id;
  const output=await exportPptx({model,environment:{fonts:[]},diagnostics:[],measurements:[{slideId,assetsReady:true,fontsReady:true,layoutStable:true,objects:[]}]});
  assert.equal(output.status,'generated',JSON.stringify(output.diagnostics));
  const doc=new DOMParser().parseFromString(strFromU8(unzipSync(output.pptxBytes)['ppt/slides/slide1.xml']),'application/xml');
  const bg=doc.getElementsByTagNameNS('*','bg')[0],color=bg?.getElementsByTagNameNS('*','srgbClr')[0];
  assert.equal(color?.getAttribute('val'),'FF0000','uncaptured base color must remain visible through transparent art');
});

for(const target of ['slide','shape'])test(`a structurally valid JPEG remains usable as ${target} background artwork`,async()=>{
  const model=await simpleModel(),rt=await browser(),session=await rt.render(await readInput({html:slide('')}));
  let jpeg;
  try{jpeg=await session.page.evaluate(()=>{const c=document.createElement('canvas');c.width=8;c.height=8;const g=c.getContext('2d');g.fillStyle='#ab2233';g.fillRect(0,0,8,8);return c.toDataURL('image/jpeg',.9);});}
  finally{await session.context.close();}
  const asset=resolveAsset(jpeg);model.assets.push(asset);
  const shape=model.slides[0].nodes.find(node=>node.kind==='shape');
  if(target==='slide')model.slides[0].background.assetId=asset.id;else shape.style.assetId=asset.id;
  const output=exportHtml(model);assert.equal(output.status,'generated',JSON.stringify(output.diagnostics));
  assert.ok(output.html.includes('data:image/jpeg;base64,'),'the valid JPEG must be preserved in HTML');
  const diagnostic=output.diagnostics.find(d=>d.code==='BACKGROUND_JPEG_UNVERIFIED');
  assert.equal(diagnostic?.severity,'warning','JPEG pixels cannot be completely verified synchronously');
  assert.equal(diagnostic.slideId,model.slides[0].id);
  assert.equal(diagnostic.nodeId,target==='shape'?shape.id:null);
  assert.ok(publicDiagnostic(diagnostic).message,'the limitation has a safe public explanation');
  const pixel=(await slidePixels(output.html,[target==='slide'?[10,10]:[100,100]]))[0];
  assert.ok(pixel[0]>120&&pixel[0]>pixel[1]*2,`the image should actually render: ${pixel}`);
  const slideId=model.slides[0].id;
  const ppt=await exportPptx({model,environment:{fonts:[]},diagnostics:[],measurements:[{slideId,assetsReady:true,fontsReady:true,layoutStable:true,objects:[]}]});
  assert.equal(ppt.status,'generated',JSON.stringify(ppt.diagnostics));
  const pptDiagnostic=ppt.diagnostics.find(d=>d.code==='BACKGROUND_JPEG_UNVERIFIED');
  assert.equal(pptDiagnostic?.severity,'warning','PPT output must disclose the same JPEG verification limit');
  assert.equal(pptDiagnostic.slideId,model.slides[0].id);
  assert.equal(pptDiagnostic.nodeId,target==='shape'?shape.id:null);
});

for(const target of ['slide','shape'])for(const invalid of ['missing','bad-bytes','truncated-png','truncated-jpeg','malformed-svg','svg-import','svg-escaped-import','svg-namespaced-href','svg-external-base','bad-hash','non-image','malicious-mime'])
  test(`${target} artwork rejects ${invalid} instead of silently emitting incomplete HTML`,async()=>{
    const model=await simpleModel(),asset=resolveAsset(validPng);
    model.assets.push(asset);
    if(invalid==='bad-bytes')asset.bytes=new Uint8Array([1,2,3,4]);
    if(invalid==='truncated-png'){asset.bytes=new Uint8Array([137,80,78,71,13,10,26,10]);asset.sha256=sha256(asset.bytes);asset.id=`asset-${asset.sha256}`;}
    if(invalid==='truncated-jpeg'){asset.mime='image/jpeg';asset.bytes=new Uint8Array([255,216,255]);asset.sha256=sha256(asset.bytes);asset.id=`asset-${asset.sha256}`;}
    if(invalid==='malformed-svg'){asset.mime='image/svg+xml';asset.bytes=new TextEncoder().encode('<svg xmlns="http://www.w3.org/2000/svg"><g>');asset.sha256=sha256(asset.bytes);asset.id=`asset-${asset.sha256}`;}
    if(Object.hasOwn(unsafeSvg,invalid)){asset.mime='image/svg+xml';asset.bytes=new TextEncoder().encode(unsafeSvg[invalid]);asset.sha256=sha256(asset.bytes);asset.id=`asset-${asset.sha256}`;}
    if(invalid==='bad-hash')asset.sha256='0'.repeat(64);
    if(invalid==='non-image')asset.mime='text/css';
    if(invalid==='malicious-mime')asset.mime='image/png;" onerror="alert(1)';
    const reference=invalid==='missing'?'asset-missing':asset.id;
    if(target==='slide')model.slides[0].background.assetId=reference;
    else model.slides[0].nodes.find(node=>node.kind==='shape').style.assetId=reference;
    const output=exportHtml(model);
    assert.equal(output.status,'failed',`invalid ${target} artwork must not succeed without notice: ${JSON.stringify(output.diagnostics)}`);
    assert.equal(output.html,null);
    assert.ok(output.diagnostics.some(d=>d.severity==='error'));
  });

for(const target of ['slide','shape'])for(const invalid of ['wrong-size','not-png'])
  test(`${target} background rejects ${invalid} capture metadata`,async()=>{
    const model=await simpleModel(),asset=resolveAsset(invalid==='not-png'?transparentArt:validPng);model.assets.push(asset);
    const background=target==='slide'?model.slides[0].background:model.slides[0].nodes.find(node=>node.kind==='shape').style;
    background.assetId=asset.id;background.capture={width:invalid==='wrong-size'?2:100,height:1,renderScale:1,content:'computed-background-only',sourceChildrenCopied:0};
    const output=exportHtml(model);assert.equal(output.status,'failed','invalid capture metadata must not silently change compositing');
    assert.ok(output.diagnostics.some(d=>d.severity==='error'));
  });

for(const target of ['slide','shape'])for(const direction of ['HTML','PPT'])
  test(`${direction} ${target} background rejects capture pixels that disagree with its box and scale`,async()=>{
    const model=await simpleModel(),asset=resolveAsset(validPng);model.assets.push(asset);
    const background=target==='slide'?model.slides[0].background:model.slides[0].nodes.find(node=>node.kind==='shape').style;
    background.assetId=asset.id;background.capture={width:1,height:1,renderScale:1,content:'computed-background-only',sourceChildrenCopied:0};
    let output;
    if(direction==='HTML')output=exportHtml(model);
    else {const slideId=model.slides[0].id;output=await exportPptx({model,environment:{fonts:[]},diagnostics:[],measurements:[{slideId,assetsReady:true,fontsReady:true,layoutStable:true,objects:[]}]});}
    assert.equal(output.status,'failed','one captured pixel cannot represent a hundreds-of-pixels box at scale 1');
    assert.ok(output.diagnostics.some(d=>d.severity==='error'));
  });

for(const target of ['slide','shape'])for(const invalid of ['missing','bad-bytes','truncated-png','truncated-jpeg','malformed-svg','svg-import','svg-escaped-import','svg-namespaced-href','svg-external-base','bad-hash','non-image','malicious-mime','wrong-capture'])
  test(`PPT ${target} background rejects ${invalid} instead of emitting corrupt artwork`,async()=>{
    const model=await simpleModel(),asset=resolveAsset(validPng);model.assets.push(asset);
    if(invalid==='bad-bytes')asset.bytes=new Uint8Array([1,2,3,4]);
    if(invalid==='truncated-png'){asset.bytes=new Uint8Array([137,80,78,71,13,10,26,10]);asset.sha256=sha256(asset.bytes);asset.id=`asset-${asset.sha256}`;}
    if(invalid==='truncated-jpeg'){asset.mime='image/jpeg';asset.bytes=new Uint8Array([255,216,255]);asset.sha256=sha256(asset.bytes);asset.id=`asset-${asset.sha256}`;}
    if(invalid==='malformed-svg'){asset.mime='image/svg+xml';asset.bytes=new TextEncoder().encode('<svg xmlns="http://www.w3.org/2000/svg"><g>');asset.sha256=sha256(asset.bytes);asset.id=`asset-${asset.sha256}`;}
    if(Object.hasOwn(unsafeSvg,invalid)){asset.mime='image/svg+xml';asset.bytes=new TextEncoder().encode(unsafeSvg[invalid]);asset.sha256=sha256(asset.bytes);asset.id=`asset-${asset.sha256}`;}
    if(invalid==='bad-hash')asset.sha256='0'.repeat(64);
    if(invalid==='non-image')asset.mime='text/css';
    if(invalid==='malicious-mime')asset.mime='image/png;" onerror="alert(1)';
    const background=target==='slide'?model.slides[0].background:model.slides[0].nodes.find(node=>node.kind==='shape').style;
    background.assetId=invalid==='missing'?'asset-missing':asset.id;
    if(invalid==='wrong-capture')background.capture={width:2,height:1,renderScale:1,content:'computed-background-only',sourceChildrenCopied:0};
    const slideId=model.slides[0].id;
    const output=await exportPptx({model,environment:{fonts:[]},diagnostics:[],measurements:[{slideId,assetsReady:true,fontsReady:true,layoutStable:true,objects:[]}]});
    assert.equal(output.status,'failed',`invalid ${target} artwork must not be reported as success: ${JSON.stringify(output.diagnostics)}`);
    assert.equal(output.pptxBytes,null);assert.ok(output.diagnostics.some(d=>d.severity==='error'));
  });

async function mixedSource(){
  const ppt=new PptxGenJS();ppt.layout='LAYOUT_WIDE';
  const page=ppt.addSlide();page.background={color:'FFFFFF'};
  page.addText([{text:'Current ',options:{bold:false}},{text:'growth',options:{bold:true}}],{x:.5,y:.5,w:3,h:.7,fontFace:'Arial',fontSize:18,shape:'roundRect',rectRadius:.2,fill:{color:'DCEEFF'},line:{color:'123456',width:1},isTextBox:false,objectName:'text-shape'});
  page.addShape(ppt.ShapeType.line,{x:.5,y:1.5,w:2,h:0,line:{color:'112233',width:2},objectName:'line'});
  page.addImage({data:validPng,x:4,y:.5,w:1,h:1,objectName:'chart-picture'});
  page.addShape(ppt.ShapeType.rect,{x:.5,y:2,w:1,h:1,fill:{color:'335577'},line:{transparency:100},objectName:'group-child'});
  page.addShape(ppt.ShapeType.rect,{x:2,y:2,w:.5,h:.5,fill:{color:'557799'},line:{transparency:100},objectName:'group-child::art'});
  page.addTable([['Label','Value'],['A','42']],{x:3,y:2,w:3,h:1,fontFace:'Arial',fontSize:14,objectName:'native-table'});
  addNativeChart(page,{id:'native-chart',kind:'chart',boxCssPx:{x:700,y:100,w:360,h:250},style:{fontFamily:'Arial',fontSize:16},content:{type:'bar',categories:['A','B'],series:[{name:'Series',values:[10,20],color:'#1976d2'}],title:'Native chart',legend:true,axes:{}}},{pointScale:.75,resolveFont:()=>({fontFace:'Arial'})});
  const parts=unzipSync(new Uint8Array(await ppt.write({outputType:'uint8array'}))),file='ppt/slides/slide1.xml';
  let xml=strFromU8(parts[file]);
  const child=xml.match(/<p:sp>(?:(?!<\/p:sp>)[\s\S])*?name="group-child"[\s\S]*?<\/p:sp>/)?.[0];
  assert.ok(child,'the fixture must include the grouped shape');
  const group=(id,name,contents)=>`<p:grpSp><p:nvGrpSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="457200" y="1828800"/><a:ext cx="914400" cy="914400"/><a:chOff x="457200" y="1828800"/><a:chExt cx="914400" cy="914400"/></a:xfrm></p:grpSpPr>${contents}</p:grpSp>`;
  xml=xml.replace(child,group(900,'outer',group(901,'inner',child)));
  const before=xml;xml=xml.replace('</a:blip>','</a:blip><a:srcRect l="10000" t="5000" r="15000" b="5000"/>');
  assert.notEqual(xml,before,'the fixture must contain a cropped image');
  xml=xml.replace('</p:sld>','<p:transition xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" p14:dur="1250" advClick="0" advTm="5000"><p:fade/></p:transition></p:sld>');
  parts[file]=strToU8(xml);return zipSync(parts);
}

test('mixed HTML and PPT writers preserve native order, grouping, crop, charts and text identity',async()=>{
  const parsed=await readPptx(await mixedSource(),{allowPartial:true});assert.equal(parsed.status,'parsed',JSON.stringify(parsed.diagnostics));
  const html=exportHtml(parsed.model);assert.equal(html.status,'generated',JSON.stringify(html.diagnostics));
  const textIndex=html.html.indexOf('data-source-id="text-shape"'),lineIndex=html.html.indexOf('data-source-id="line"'),pictureIndex=html.html.indexOf('data-source-id="chart-picture"');
  assert.ok(textIndex>=0&&textIndex<lineIndex&&lineIndex<pictureIndex,'HTML preserves the original object order');
  assert.match(html.html,/data-pptx-origin=/,'the text shape keeps its source relationship');
  const input=await measured(html.html);assert.deepEqual(errors(input),[],JSON.stringify(input.diagnostics));
  const target=input.model.slides[0].nodes.find(n=>n.id==='group-child');assert.equal(target.parentId,'inner');
  const bg=resolveAsset(validPng);if(!input.model.assets.some(asset=>asset.id===bg.id))input.model.assets.push(bg);target.style.assetId=bg.id;
  const output=await exportPptx(input);assert.equal(output.status,'generated',JSON.stringify(output.diagnostics));
  const map=output.objectMap,derived=map.find(item=>item.derivedFrom==='group-child');
  assert.ok(derived,'the background image is registered as derived art');
  assert.notEqual(derived.pptObjectName,'group-child::art','derived names cannot collide with user objects');
  assert.equal(map.filter(item=>item.sourceNodeId==='text-shape').length,1);
  assert.ok(map.findIndex(item=>item.sourceNodeId==='text-shape')<map.findIndex(item=>item.sourceNodeId==='line'));
  const xml=strFromU8(unzipSync(output.pptxBytes)['ppt/slides/slide1.xml']);
  const doc=new DOMParser().parseFromString(xml,'application/xml');
  const groups=[...doc.getElementsByTagNameNS('*','grpSp')],inner=groups.find(g=>g.getElementsByTagNameNS('*','cNvPr')[0]?.getAttribute('name')==='inner');
  assert.ok(inner,'nested group remains an editable group');
  assert.ok([...inner.getElementsByTagNameNS('*','cNvPr')].some(n=>n.getAttribute('name')===derived.pptObjectName),'derived background remains inside its original group');
  const text=[...doc.getElementsByTagNameNS('*','sp')].filter(node=>node.textContent.includes('Current growth'));
  assert.equal(text.length,1,'the source text shape remains one native object');
  const bold=[...text[0].getElementsByTagNameNS('*','r')].find(run=>run.textContent.includes('growth'));
  assert.equal(bold?.getElementsByTagNameNS('*','rPr')[0]?.getAttribute('b'),'1');
  const crops=[...doc.getElementsByTagNameNS('*','srcRect')];
  assert.ok(crops.some(c=>Number(c.getAttribute('l'))===10000&&Number(c.getAttribute('r'))===15000),'image crop remains attached to the picture');
  assert.equal(doc.getElementsByTagNameNS('*','tbl').length,1,'table remains native');
  assert.match(xml,/native-chart/,'chart remains present as its own object');
  const transition=doc.getElementsByTagNameNS('*','transition')[0];
  assert.equal(transition?.getElementsByTagNameNS('*','fade').length,1,'slide transition stays intact');
  assert.equal(transition.getAttributeNS('http://schemas.microsoft.com/office/powerpoint/2010/main','dur'),'1250');
});
