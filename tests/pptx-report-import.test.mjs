import test, {before, after} from 'node:test';
import assert from 'node:assert/strict';
import PptxGenJS from 'pptxgenjs';
import {unzipSync, zipSync, strFromU8, strToU8} from 'fflate';
import {readPptx} from '../src/html-to-pptx/read-pptx.mjs';
import {exportHtml} from '../src/html-to-pptx/export-html.mjs';
import {readInput} from '../src/html-to-pptx/input.mjs';
import {measureInput} from '../src/html-to-pptx/measure.mjs';
import {createBrowserRuntime} from '../src/html-to-pptx/browser-runtime.mjs';
import {exportPptx} from '../src/html-to-pptx/export.mjs';
import {runConversion} from '../src/html-to-pptx/job-runner.mjs';

const image='data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+j1ioAAAAASUVORK5CYII=';
let runtime;
before(async()=>{runtime=await createBrowserRuntime({capture:false});});
after(async()=>runtime?.close());

async function fixture(picture=image){
  const ppt=new PptxGenJS();ppt.layout='LAYOUT_WIDE';
  const slide=ppt.addSlide();slide.background={color:'FFFFFF'};
  slide.addText('First paragraph\nSecond paragraph',{x:1,y:1,w:5,h:2,fontFace:'Arial',fontSize:20,margin:0,objectName:'report'});
  slide.addImage({data:picture,x:7,y:1,w:2,h:1,objectName:'picture'});
  return unzipSync(new Uint8Array(await ppt.write({outputType:'uint8array'})));
}
function update(parts,name,replace){parts[name]=strToU8(replace(strFromU8(parts[name])));}

test('theme background uses the inherited color map and becomes an explicit HTML background',async()=>{
  const parts=await fixture();
  update(parts,'ppt/slides/slide1.xml',xml=>xml.replace(/<p:bg>[\s\S]*?<\/p:bg>/,'<p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>'));
  update(parts,'ppt/slideMasters/slideMaster1.xml',xml=>xml.replace('bg1="lt1"','bg1="accent1"'));
  update(parts,'ppt/theme/theme1.xml',xml=>xml.replace(/<a:accent1>[\s\S]*?<\/a:accent1>/,'<a:accent1><a:srgbClr val="123456"/></a:accent1>'));
  const result=await readPptx(zipSync(parts));assert.equal(result.status,'parsed',JSON.stringify(result.diagnostics));
  assert.equal(result.model.slides[0].background.color.toUpperCase(),'#123456');
  assert.match(exportHtml(result.model).html,/background:#123456/i);
});

test('theme background follows the master relationship even when the theme has a nonstandard path',async()=>{
  const parts=await fixture();parts['ppt/slideMasters/theme/theme2.xml']=parts['ppt/theme/theme1.xml'];
  update(parts,'ppt/slideMasters/_rels/slideMaster1.xml.rels',xml=>xml.replace('../theme/theme1.xml','theme/theme2.xml'));
  update(parts,'ppt/slides/slide1.xml',xml=>xml.replace(/<p:bg>[\s\S]*?<\/p:bg>/,'<p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>'));
  const result=await readPptx(zipSync(parts));assert.equal(result.status,'parsed',JSON.stringify(result.diagnostics));
  assert.equal(result.model.slides[0].background.color.toUpperCase(),'#FFFFFF');
});

test('unsupported complex theme backgrounds still stop conversion',async()=>{
  const parts=await fixture();
  update(parts,'ppt/slides/slide1.xml',xml=>xml.replace(/<p:bg>[\s\S]*?<\/p:bg>/,'<p:bg><p:bgRef idx="1003"><a:schemeClr val="bg1"/></p:bgRef></p:bg>'));
  const result=await readPptx(zipSync(parts));assert.equal(result.status,'failed');assert.ok(result.diagnostics.some(d=>d.severity==='error'));
});

test('paragraph spacing is reported as an approximation while current text remains editable',async()=>{
  const parts=await fixture();
  update(parts,'ppt/slides/slide1.xml',xml=>xml.replace(/<a:pPr([^>]*)>/,'<a:pPr$1><a:spcAft><a:spcPts val="500"/></a:spcAft>'));
  const result=await readPptx(zipSync(parts));assert.equal(result.status,'parsed',JSON.stringify(result.diagnostics));
  assert.ok(result.diagnostics.some(d=>d.severity==='warning'&&d.code==='PARAGRAPH_SPACING_APPROXIMATED'));
  const html=exportHtml(result.model).html;assert.match(html,/First paragraph/);assert.match(html,/Second paragraph/);
});

test('a missing local font warns during PPTX import without blocking the editable HTML',async()=>{
  const parts=await fixture();
  update(parts,'ppt/slides/slide1.xml',xml=>xml.replaceAll('Arial','PptAgainDefinitelyMissingFont_74f3'));
  const {measured}=await runConversion({operation:'analyze-pptx',input:zipSync(parts)});
  assert.ok(measured.model,JSON.stringify(measured.diagnostics));
  assert.ok(!measured.diagnostics.some(d=>d.severity==='error'),JSON.stringify(measured.diagnostics));
  assert.ok(measured.diagnostics.some(d=>d.code==='FONT_PREVIEW_SUBSTITUTED'&&d.severity==='warning'));
  assert.equal(measured.previews.length,1);
  assert.match(exportHtml(measured.model).html,/PptAgainDefinitelyMissingFont_74f3/);
});

test('the HTML preview displays only the requested region of a cropped image',async()=>{
  const setup=await runtime.render(await readInput({html:'<section data-slide style="width:800px;height:600px"></section>'}));
  let stripes;try{stripes=await setup.page.evaluate(()=>{const c=document.createElement('canvas');c.width=100;c.height=20;const g=c.getContext('2d');g.fillStyle='#ff0000';g.fillRect(0,0,20,20);g.fillStyle='#00ff00';g.fillRect(20,0,60,20);g.fillStyle='#0000ff';g.fillRect(80,0,20,20);return c.toDataURL('image/png');});}finally{await setup.context.close();}
  const parts=await fixture(stripes);
  update(parts,'ppt/slides/slide1.xml',xml=>xml.replace('</a:blip>  <a:stretch>','</a:blip><a:srcRect l="20000" r="20000"/><a:stretch>'));
  const result=await readPptx(zipSync(parts));assert.equal(result.status,'parsed',JSON.stringify(result.diagnostics));
  const preview=await runtime.render(await readInput({html:exportHtml(result.model).html}));
  try{
    const screenshot=await preview.page.locator('[data-kind="image"]').screenshot();
    const colors=await preview.page.evaluate(async data=>{const img=new Image();img.src=data;await img.decode();const c=document.createElement('canvas');c.width=img.width;c.height=img.height;const g=c.getContext('2d');g.drawImage(img,0,0);return [0.15,0.5,0.85].map(x=>[...g.getImageData(Math.floor(c.width*x),Math.floor(c.height/2),1,1).data]);},`data:image/png;base64,${screenshot.toString('base64')}`);
    assert.deepEqual(colors,[[0,255,0,255],[0,255,0,255],[0,255,0,255]]);
  }finally{await preview.context.close();}
});

test('a crop that removes the whole image is rejected',async()=>{
  const parts=await fixture();
  update(parts,'ppt/slides/slide1.xml',xml=>xml.replace('</a:blip>  <a:stretch>','</a:blip><a:srcRect l="90000" r="20000"/><a:stretch>'));
  const result=await readPptx(zipSync(parts));assert.equal(result.status,'failed');assert.ok(result.diagnostics.some(d=>d.severity==='error'));
});

test('cropped images retain their crop through HTML edits and a new PPTX',async()=>{
  const parts=await fixture();
  update(parts,'ppt/slides/slide1.xml',xml=>xml.replace('</a:blip>  <a:stretch>','</a:blip><a:srcRect l="10000" t="20000" r="15000" b="5000"/><a:stretch>'));
  const result=await readPptx(zipSync(parts));assert.equal(result.status,'parsed',JSON.stringify(result.diagnostics));
  const html=exportHtml(result.model).html;
  assert.match(html,/object-view-box:inset\(20% 15% 5% 10%\)/);
  const measured=await measureInput(await readInput({html:html.replace('First paragraph','Updated paragraph')}),runtime);
  const output=await exportPptx(measured);assert.equal(output.status,'generated',JSON.stringify(output.diagnostics));
  const slide=strFromU8(unzipSync(output.pptxBytes)['ppt/slides/slide1.xml']);
  const crop=slide.match(/<a:srcRect\b[^>]*\/>/)?.[0];assert.ok(crop);
  for(const [side,expected] of Object.entries({l:10000,t:20000,r:15000,b:5000})){
    const value=Number(new RegExp(`${side}="(\\d+)"`).exec(crop)?.[1]);assert.ok(Math.abs(value-expected)<=1,`${side}: ${value}`);
  }
  assert.match(slide,/Updated paragraph/);
});
