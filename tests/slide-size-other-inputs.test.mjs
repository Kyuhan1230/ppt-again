import test from 'node:test';
import assert from 'node:assert/strict';
import {readInput} from '../src/html-to-pptx/input.mjs';
import {createBrowserRuntime} from '../src/html-to-pptx/browser-runtime.mjs';
import {measureInput} from '../src/html-to-pptx/measure.mjs';
import {exportPptx} from '../src/html-to-pptx/export.mjs';
import {inspectPptxIndependently} from './inspect-pptx.mjs';
import {unzipSync,strFromU8} from 'fflate';
import {inflateSync} from 'node:zlib';

const policy={mode:'contain',targetWidthEmu:1600*9525,targetHeightEmu:900*9525};
const verify=async(html,profile)=>{
  const input=await readInput({html,sizePolicy:policy}),runtime=await createBrowserRuntime({capture:profile==='svg-slides-v1',checkLimits:true});
  try{
    if(profile)assert.equal(input.profile,profile,JSON.stringify(input.diagnostics));
    const measured=await measureInput(input,runtime);
    assert.ok(!measured.diagnostics.some(d=>d.severity==='error'),JSON.stringify(measured.diagnostics));
    if(profile==='svg-slides-v1'){
      assert.deepEqual(measured.environment.viewport,{width:1280,height:960});
      assert.equal(measured.environment.previewKind,'source');assert.equal(measured.previews.length,2);
      assert.ok(!measured.model.slides[0].nodes.find(n=>n.kind==='text')?.style.color.includes('255, 0, 0'));
    }
    assert.deepEqual(measured.model.slides.map(s=>[s.widthCssPx,s.heightCssPx]),[[1600,900],[1600,900]]);
    assert.equal(measured.sizePolicy.slides[0].offsetX,200);
    const exported=await exportPptx(measured);assert.equal(exported.status,'generated',JSON.stringify(exported.diagnostics));
    const ppt=await inspectPptxIndependently(exported.pptxBytes);assert.equal(ppt.slides.length,2);
    assert.deepEqual([ppt.widthEmu,ppt.heightEmu],[policy.targetWidthEmu,policy.targetHeightEmu]);
    assert.ok(JSON.stringify(ppt.slides[0]).includes('First'));
    assert.ok(JSON.stringify(ppt.slides[1]).includes('Second'));
  }finally{await runtime.close();}
};

test('mixed SVG viewBoxes become uniformly sized PPT slides',async()=>{
  const svg=(w,h,text)=>`<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${w} ${h}"><rect x="0" y="0" width="${w}" height="${h}" fill="white"/><text x="100" y="100" font-family="Arial" font-size="20">${text}</text></svg>`;
  await verify(`<!doctype html><html><head><style>@media(max-height:800px){text{fill:red!important}}</style></head><body>${svg(800,600,'First')}${svg(960,540,'Second')}</body></html>`,'svg-slides-v1');
});

test('mixed validated embedded pages become uniformly sized PPT slides',async()=>{
  const page=(w,h,text)=>`<section style="position:relative;width:${w}px;height:${h}px;background:white"><p style="position:absolute;left:100px;top:100px;width:200px;height:60px;margin:0;font:20px/30px Arial">${text}</p></section>`;
  const data={pages:[{html:page(800,600,'First'),source:'a'},{html:page(960,540,'Second'),source:'a'}],sources:{a:{styles:''}}};
  await verify(`<!doctype html><html><body><script type="application/json">${JSON.stringify(data)}</script></body></html>`);
});

test('SVG root background is preserved and a painted ancestor behind transparency is rejected',async()=>{
  const svg='<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 800 600" style="width:800px;height:600px"><rect x="100" y="100" width="100" height="100" fill="blue"/></svg>';
  const runtime=await createBrowserRuntime({capture:true,checkLimits:true});
  try{
    const measured=await measureInput(await readInput({html:`<!doctype html><html><head><style>.wrap svg{background:rgb(255,0,0)}</style></head><body><div class="wrap">${svg}</div></body></html>`}),runtime);
    assert.ok(!measured.diagnostics.some(d=>d.severity==='error'),JSON.stringify(measured.diagnostics));
    assert.equal(measured.model.slides[0].background.color,'rgb(255, 0, 0)');
    const result=await exportPptx(measured);assert.equal(result.status,'generated',JSON.stringify(result.diagnostics));
    const xml=strFromU8(unzipSync(result.pptxBytes)['ppt/slides/slide1.xml']);
    assert.match(xml,/FF0000/);
    const png=Buffer.from(measured.previews[0].bytes),chunks=[];assert.ok([2,6].includes(png[25]));
    for(let i=8;i<png.length;){const length=png.readUInt32BE(i),type=png.toString('ascii',i+4,i+8);if(type==='IDAT')chunks.push(png.subarray(i+8,i+8+length));i+=length+12;}
    assert.deepEqual([...inflateSync(Buffer.concat(chunks)).subarray(1,4)],[255,0,0],'the preview must retain the original CSS background');
    const rejected=await measureInput(await readInput({html:`<!doctype html><html><body style="background:red">${svg}</body></html>`}),runtime);
    assert.ok(rejected.diagnostics.some(d=>d.code==='UNSUPPORTED_OBJECT'&&d.severity==='error'&&d.message.includes('부모 배경')),JSON.stringify(rejected.diagnostics));
  }finally{await runtime.close();}
});
