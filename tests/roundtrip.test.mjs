import test,{before,after} from 'node:test';
import assert from 'node:assert/strict';
import {readInput} from '../src/html-to-pptx/input.mjs';
import {measureInput} from '../src/html-to-pptx/measure.mjs';
import {createBrowserRuntime} from '../src/html-to-pptx/browser-runtime.mjs';
import {exportPptx} from '../src/html-to-pptx/export.mjs';
import {inspectPptxIndependently} from './inspect-pptx.mjs';
import PptxGenJS from 'pptxgenjs';
import {unzipSync,zipSync,strFromU8,strToU8} from 'fflate';
let runtime;
before(async()=>{runtime=await createBrowserRuntime({capture:false});});
after(async()=>runtime?.close());
const source=`<section data-slide style="width:1600px;height:900px;position:relative;background:white"><p id="title" style="position:absolute;left:160px;top:90px;width:800px;height:120px;margin:0;font:40px/50px Arial;white-space:pre-wrap">Quarterly plan</p><div id="block" style="position:absolute;left:320px;top:360px;width:480px;height:180px;background:#164e63"></div></section>`;
const measure=async html=>measureInput(await readInput({html}),runtime);

// Catches replacing an imported 10 x 5.625 inch canvas with the default long edge.
test('PPT export preserves source physical size and normalized positions',async()=>{
  const m=await measure(source);m.model.pptxSizeEmu={width:9144000,height:5143500};
  const result=await exportPptx(m);assert.equal(result.status,'generated',JSON.stringify(result.diagnostics));
  const a=await inspectPptxIndependently(result.pptxBytes);
  assert.deepEqual([a.widthEmu,a.heightEmu],[9144000,5143500]);
  assert.deepEqual(a.slides[0].objects.find(o=>o.name==='title').boxEmu,{x:914400,y:514350,w:4572000,h:685800});
});

test('invalid source dimensions cannot silently stretch or replace the canvas',async()=>{
  for(const size of [{width:0,height:5143500},{width:9144000,height:9144000},{width:NaN,height:5143500}]){
    const m=await measure(source);m.model.pptxSizeEmu=size;
    const r=await exportPptx(m);assert.equal(r.status,'failed');assert.equal(r.pptxBytes,null);
  }
});

async function roundtripApi(){
  const reader=await import('../src/html-to-pptx/read-pptx.mjs').catch(()=>null);
  const writer=await import('../src/html-to-pptx/export-html.mjs').catch(()=>null);
  assert.equal(typeof reader?.readPptx,'function','PPT input must expose the shared deck model');
  assert.equal(typeof writer?.exportHtml,'function','HTML output must be reimportable');
  return {...reader,...writer};
}
const tinyPng='data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+j1ioAAAAASUVORK5CYII=';
async function nativeFixture(width=10,height=5.625){
  const p=new PptxGenJS();p.defineLayout({name:'fixture',width,height});p.layout='fixture';
  const s=p.addSlide();s.background={color:'FFFFFF'};
  s.addText([{text:'Current ',options:{bold:false}},{text:'plan',options:{bold:true,color:'164E63'}}],{x:1,y:.5,w:5,h:1,fontFace:'Arial',fontSize:24,margin:0,breakLine:false,paraSpaceAfter:0,objectName:'heading'});
  s.addShape(p.ShapeType.rect,{x:2,y:2,w:3,h:1,fill:{color:'164E63'},line:{color:'164E63',transparency:100},objectName:'block'});
  s.addShape(p.ShapeType.line,{x:1,y:3.5,w:4,h:0,line:{color:'164E63',width:2},objectName:'rule'});
  s.addImage({x:7,y:2,w:1,h:1,data:tinyPng,objectName:'picture'});
  return new Uint8Array(await p.write({outputType:'uint8array'}));
}

// Catches SVG-only reverse output, stale embedded text and viewport-based coordinates.
test('PPT to HTML to PPT reads visible edits and retains untouched native objects',async()=>{
  const {readPptx,exportHtml}=await roundtripApi();
  const bytes=await nativeFixture(),read=await readPptx(bytes);
  assert.equal(read.status,'parsed',JSON.stringify(read.diagnostics));
  const out=exportHtml(read.model);assert.equal(out.status,'generated',JSON.stringify(out.diagnostics));
  assert.ok(!out.html.includes('foreignObject'));
  const edited=out.html.replace('Current ','Updated ').replace('left:192px','left:216px');
  assert.notEqual(edited,out.html);
  const m=await measure(edited),result=await exportPptx(m);
  assert.equal(result.status,'generated',JSON.stringify(result.diagnostics));
  const a=await inspectPptxIndependently(result.pptxBytes);
  assert.deepEqual([a.widthEmu,a.heightEmu],[9144000,5143500]);
  const objects=a.slides[0].objects;
  assert.deepEqual(objects.map(o=>o.name),['heading','block','rule','picture']);
  assert.equal(objects[0].text,'Updated plan');assert.equal(objects[0].runs.at(-1).bold,true);
  assert.equal(objects[1].boxEmu.x,2057400);
  assert.equal(a.nativeTextCount,1);assert.equal(a.nativeShapeCount,1);assert.equal(a.nativeLineCount,1);assert.equal(a.pictureCount,1);
  const before=await inspectPptxIndependently(bytes);
  for(const i of [0,2,3])assert.deepEqual(objects[i].boxEmu,before.slides[0].objects[i].boxEmu);
});

test('three round trips retain current edits and untouched coordinates for independent canvas ratios',async()=>{
  const {readPptx,exportHtml}=await roundtripApi();
  for(const [width,height] of [[10,5.625],[10,7.5],[8,10]]){
    let bytes=await nativeFixture(width,height);const initial=await inspectPptxIndependently(bytes);
    for(let cycle=0;cycle<3;cycle++){
      const read=await readPptx(bytes);assert.equal(read.status,'parsed',JSON.stringify(read.diagnostics));
      const currentLabel=cycle===0?'Current ':`Revision ${cycle} `;
      const html=exportHtml(read.model).html.replace(currentLabel,`Revision ${cycle+1} `).replace(`left:${192+cycle*12}px`,`left:${192+(cycle+1)*12}px`);
      const result=await exportPptx(await measure(html));assert.equal(result.status,'generated',JSON.stringify(result.diagnostics));
      bytes=result.pptxBytes;const a=await inspectPptxIndependently(bytes);
      assert.deepEqual([a.widthEmu,a.heightEmu],[initial.widthEmu,initial.heightEmu]);
      const expected=structuredClone(initial.slides[0].objects);
      expected[0].text=`Revision ${cycle+1} plan`;expected[1].boxEmu.x+=114300*(cycle+1);
      assert.deepEqual(a.slides[0].objects.map(o=>[o.name,o.kind,o.text,o.boxEmu]),expected.map(o=>[o.name,o.kind,o.text,o.boxEmu]));
    }
  }
});

test('normalized preview does not clip the canvas to its responsive viewer wrapper',async()=>{
  const {readPptx,exportHtml}=await roundtripApi();const html=exportHtml((await readPptx(await nativeFixture())).model).html;
  const input=await readInput({html}),m=await measureInput(input,{...runtime,capture:true});
  assert.equal(m.previews.length,1,JSON.stringify(m.diagnostics));
  const session=await runtime.render(input);
  try{
    const pixel=await session.page.evaluate(async src=>{const i=new Image();i.src=src;await i.decode();const c=document.createElement('canvas');c.width=i.width;c.height=i.height;const x=c.getContext('2d');x.drawImage(i,0,0);return [...x.getImageData(1200,680,1,1).data];},`data:image/png;base64,${Buffer.from(m.previews[0].bytes).toString('base64')}`);
    assert.deepEqual(pixel,[255,255,255,255],'The lower right of the slide must be visible');
  }finally{await session.context.close();}
});

test('PPT text margins survive the round trip without swapping top and left',async()=>{
  const {readPptx,exportHtml}=await roundtripApi(),p=new PptxGenJS();p.layout='LAYOUT_WIDE';const s=p.addSlide();
  s.background={color:'FFFFFF'};
  s.addText('Margin check',{x:1,y:1,w:5,h:2,fontFace:'Arial',fontSize:24,margin:[6,12,18,24],objectName:'margins'});
  const bytes=new Uint8Array(await p.write({outputType:'uint8array'}));
  const read=await readPptx(bytes);assert.equal(read.status,'parsed',JSON.stringify(read.diagnostics));
  const html=exportHtml(read.model),out=await exportPptx(await measure(html.html));
  assert.equal(out.status,'generated',JSON.stringify(out.diagnostics));
  const actual=(await inspectPptxIndependently(out.pptxBytes)).slides[0].objects[0].bodyPr;
  assert.deepEqual([actual.tIns,actual.rIns,actual.bIns,actual.lIns],['76200','152400','228600','304800']);
});

test('unsupported master artwork cannot disappear from a successful reverse conversion',async()=>{
  const {readPptx}=await roundtripApi(),parts=unzipSync(await nativeFixture());
  const slide=strFromU8(parts['ppt/slides/slide1.xml']),shape=slide.match(/<p:sp>.*?<\/p:sp>/s)[0];
  parts['ppt/slideMasters/slideMaster1.xml']=strToU8(strFromU8(parts['ppt/slideMasters/slideMaster1.xml']).replace('</p:spTree>',shape+'</p:spTree>'));
  const r=await readPptx(zipSync(parts));assert.equal(r.status,'failed');assert.ok(r.diagnostics.some(d=>d.code==='UNSUPPORTED_OBJECT'));
});

test('scaled PPT groups cannot silently preserve boxes while losing text scale',async()=>{
  const {readPptx}=await roundtripApi(),parts=unzipSync(await nativeFixture());let xml=strFromU8(parts['ppt/slides/slide1.xml']);
  const shape=xml.match(/<p:sp>.*?<\/p:sp>/s)[0];
  const group='<p:grpSp><p:nvGrpSpPr><p:cNvPr id="999" name="Scaled group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="9144000" cy="5143500"/><a:chOff x="0" y="0"/><a:chExt cx="4572000" cy="2571750"/></a:xfrm></p:grpSpPr>'+shape+'</p:grpSp>';
  parts['ppt/slides/slide1.xml']=strToU8(xml.replace(shape,group));
  const r=await readPptx(zipSync(parts));assert.equal(r.status,'failed');assert.equal(r.model,null);assert.ok(r.diagnostics.some(d=>d.code==='UNSUPPORTED_OBJECT'));
});

test('PPT parsing rejects external references, entities and expanded-size overflow',async()=>{
  const {readPptx}=await roundtripApi(),original=unzipSync(await nativeFixture());
  for(const mode of ['external','entity','expanded']){
    const parts={...original};
    if(mode==='external')parts['ppt/slides/_rels/slide1.xml.rels']=strToU8(strFromU8(parts['ppt/slides/_rels/slide1.xml.rels']).replace('</Relationships>','<Relationship Id="unsafe" Type="image" Target="https://example.invalid/private" TargetMode="External"/></Relationships>'));
    if(mode==='entity')parts['ppt/presentation.xml']=strToU8('<!DOCTYPE p [<!ENTITY unsafe "payload">]>'+strFromU8(parts['ppt/presentation.xml']));
    if(mode==='expanded')parts['ppt/media/oversize.png']=new Uint8Array(17*1024*1024);
    const r=await readPptx(zipSync(parts));assert.equal(r.status,'failed',mode);
  }
});

test('generated HTML rejects stale duplicate project data and changed aspect metadata',async()=>{
  const {readPptx,exportHtml}=await roundtripApi(),html=exportHtml((await readPptx(await nativeFixture())).model).html;
  for(const bad of [html.replace('</body>','<script id="project-data" type="application/json">{}</script></body>'),html.replace('data-ppt-height-emu="5143500"','data-ppt-height-emu="9144000"')]){
    const result=await exportPptx(await measure(bad));assert.equal(result.status,'failed');
  }
});

test('visible edits inside unsupported generated shape content cannot be discarded',async()=>{
  const {readPptx,exportHtml}=await roundtripApi();const html=exportHtml((await readPptx(await nativeFixture())).model).html;
  const edited=html.replace(/(<div data-source-id="block"[^>]*>)/,'$1<span style="font:24px Arial">Current added text</span>');
  const result=await exportPptx(await measure(edited));assert.equal(result.status,'failed');
});

test('PPT unsupported text rotation, strike and theme fill are diagnosed rather than lost',async()=>{
  const {readPptx}=await roundtripApi(),original=unzipSync(await nativeFixture());
  for(const mode of ['rotation','strike','theme']){
    const parts={...original};let xml=strFromU8(parts['ppt/slides/slide1.xml']);
    if(mode==='rotation')xml=xml.replace(/<a:bodyPr[^>]*>/,s=>s.replace(/ rot="[^"]*"/g,'').replace('>',' rot="5400000">'));
    if(mode==='strike')xml=xml.replace('<a:rPr ','<a:rPr strike="sngStrike" ');
    if(mode==='theme')xml=xml.replace(/<p:spPr>([\s\S]*?)<\/p:spPr>/,(_,s)=>'<p:spPr>'+s.replace(/<a:noFill\s*\/>|<a:solidFill>[\s\S]*?<\/a:solidFill>/g,'')+'</p:spPr><p:style><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef></p:style>');
    parts['ppt/slides/slide1.xml']=strToU8(xml);const r=await readPptx(zipSync(parts));assert.equal(r.status,'failed',mode);
  }
});

test('a stored ZIP member cannot understate the allocation size in its directory',async()=>{
  const {readPptx}=await roundtripApi(),parts=unzipSync(await nativeFixture());parts['oversize.bin']=new Uint8Array(17*1024*1024);
  const archive=zipSync(parts,{level:0}),view=new DataView(archive.buffer,archive.byteOffset,archive.byteLength);
  for(let i=0;i<archive.length-46;i++)if(view.getUint32(i,true)===0x02014b50){const length=view.getUint16(i+28,true);if(strFromU8(archive.subarray(i+46,i+46+length))==='oversize.bin'){view.setUint32(i+24,0,true);break;}}
  const r=await readPptx(archive);assert.equal(r.status,'failed');
});

test('fractional current HTML coordinates retain precision beyond CSSOM serialization',async()=>{
  const {readPptx,exportHtml}=await roundtripApi(),p=new PptxGenJS();p.layout='LAYOUT_WIDE';
  const s=p.addSlide();s.background={color:'FFFFFF'};
  s.addShape(p.ShapeType.rect,{x:11.6754829,y:1.8372736,w:.5738291,h:2.7283918,objectName:'precise',fill:{color:'164E63'},line:{transparency:100}});
  const bytes=new Uint8Array(await p.write({outputType:'uint8array'})),before=await inspectPptxIndependently(bytes);
  const read=await readPptx(bytes);assert.equal(read.status,'parsed',JSON.stringify(read.diagnostics));
  const html=exportHtml(read.model).html,r=await exportPptx(await measure(html));
  assert.equal(r.status,'generated',JSON.stringify(r.diagnostics));const after=await inspectPptxIndependently(r.pptxBytes);
  for(const key of ['x','y','w','h'])assert.ok(Math.abs(before.slides[0].objects[0].boxEmu[key]-after.slides[0].objects[0].boxEmu[key])<=1,key);
});

test('round-tripped ellipses do not contain rectangle adjustments that Office rejects',async()=>{
  const {readPptx,exportHtml}=await roundtripApi(),p=new PptxGenJS();p.layout='LAYOUT_WIDE';
  const s=p.addSlide();s.background={color:'FFFFFF'};
  s.addShape(p.ShapeType.ellipse,{x:1,y:1,w:1,h:1,fill:{color:'164E63'},line:{transparency:100},objectName:'circle'});
  const bytes=new Uint8Array(await p.write({outputType:'uint8array'})),read=await readPptx(bytes);
  assert.equal(read.status,'parsed',JSON.stringify(read.diagnostics));
  const r=await exportPptx(await measure(exportHtml(read.model).html));
  assert.equal(r.status,'generated',JSON.stringify(r.diagnostics));
  const xml=strFromU8(unzipSync(r.pptxBytes)['ppt/slides/slide1.xml']);
  assert.match(xml,/<a:prstGeom prst="ellipse"><a:avLst\s*\/>|<a:prstGeom prst="ellipse"><a:avLst><\/a:avLst>/);
});

test('solid text outlines remain editable text through HTML and PPT',async()=>{
  const {readPptx,exportHtml}=await roundtripApi(),p=new PptxGenJS();p.layout='LAYOUT_WIDE';
  const s=p.addSlide();s.background={color:'FFFFFF'};
  s.addText('Outlined text',{x:1,y:1,w:5,h:1,fontFace:'Arial',fontSize:28,margin:0,outline:{color:'79CFFF',size:2},objectName:'outline'});
  const read=await readPptx(new Uint8Array(await p.write({outputType:'uint8array'})));assert.equal(read.status,'parsed',JSON.stringify(read.diagnostics));
  const result=await exportPptx(await measure(exportHtml(read.model).html));assert.equal(result.status,'generated',JSON.stringify(result.diagnostics));
  const audit=await inspectPptxIndependently(result.pptxBytes);assert.equal(audit.nativeTextCount,1);assert.equal(audit.pictureCount,0);assert.equal(audit.slides[0].objects[0].text,'Outlined text');
  const xml=strFromU8(unzipSync(result.pptxBytes)['ppt/slides/slide1.xml']);assert.match(xml,/<a:ln w="25400"><a:solidFill><a:srgbClr val="79CFFF"/);
});

test('translucent text outlines cannot silently become opaque',async()=>{
  const html=source.replace('font:40px/50px Arial','font:40px/50px Arial;-webkit-text-stroke:2px rgba(121,207,255,.5)');
  const result=await exportPptx(await measure(html));assert.equal(result.status,'failed');assert.equal(result.pptxBytes,null);
});
