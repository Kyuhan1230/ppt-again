import test from 'node:test';
import assert from 'node:assert/strict';
import PptxGenJS from 'pptxgenjs';
import {unzipSync,zipSync,strFromU8,strToU8} from 'fflate';
import {readPptx} from '../src/html-to-pptx/read-pptx.mjs';
import {exportHtml} from '../src/html-to-pptx/export-html.mjs';
import {publicDiagnostic} from '../src/html-to-pptx/public-diagnostics.mjs';
import {readInput} from '../src/html-to-pptx/input.mjs';
import {createBrowserRuntime} from '../src/html-to-pptx/browser-runtime.mjs';
import {measureInput} from '../src/html-to-pptx/measure.mjs';

const rgba=Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAYAAABytg0kAAAAF0lEQVR4nGP4z8Dwn+E/QwMjmP7P8B8APeIHeyMDVGUAAAAASUVORK5CYII=','base64');
const rgb=Buffer.from('iVBORw0KGgoAAAANSUhEUgAAABQAAAAUCAIAAAAC64paAAAAJUlEQVR4nGP8z0A+YKJAL8Oo5pERYIwMDP9Hmp+ZRjWTBigKMAB8+QIlC4gOEgAAAABJRU5ErkJggg==','base64');
async function fixture(bytes=rgba,stops='<a:srgbClr val="000000"/><a:srgbClr val="FFFFFF"/>'){
 const ppt=new PptxGenJS(),slide=ppt.addSlide();slide.background={color:'FFFFFF'};
 slide.addImage({data:'data:image/png;base64,'+bytes.toString('base64'),x:1,y:1,w:2,h:2,objectName:'duotone-picture'});
 const parts=unzipSync(new Uint8Array(await ppt.write({outputType:'uint8array'}))),key='ppt/slides/slide1.xml';
 const source=strFromU8(parts[key]),modified=source.replace('</a:blip>',`<a:duotone>${stops}</a:duotone></a:blip>`);
 assert.notEqual(modified,source);parts[key]=strToU8(modified);return parts;
}
test('duotone is baked into pixels without losing source alpha and is diagnosed on HTML reimport',async()=>{
 const bytes=zipSync(await fixture());assert.equal((await readPptx(bytes)).status,'failed');
 const result=await readPptx(bytes,{allowPartial:true});assert.equal(result.status,'parsed',JSON.stringify(result.diagnostics));
 const image=result.model.slides[0].nodes.find(n=>n.id==='duotone-picture');assert.ok(image);
 const warning=result.diagnostics.find(d=>d.code==='PPTX_IMAGE_DUOTONE_BAKED');assert.equal(warning?.severity,'warning');assert.match(publicDiagnostic(warning).message,/편집/);
 const asset=result.model.assets.find(a=>a.id===image.content.assetId);assert.ok(asset);assert.notDeepEqual(Buffer.from(asset.bytes),rgba);
 const html=exportHtml(result.model).html;assert.match(html,/data-duotone-baked/);
 const runtime=await createBrowserRuntime({capture:false});
 try{
  const session=await runtime.render(await readInput({html}));
  try {
   const pixels=await session.page.evaluate(async url=>{
    const img=new Image();img.src=url;await img.decode();const c=document.createElement('canvas');c.width=2;c.height=2;const ctx=c.getContext('2d');ctx.drawImage(img,0,0);return [...ctx.getImageData(0,0,2,2).data];
   },`data:image/png;base64,${Buffer.from(asset.bytes).toString('base64')}`);
   assert.deepEqual(pixels.slice(0,4),[54,54,54,255]);
   assert.ok(Math.abs(pixels[4]-182)<=1&&pixels[7]===128);
   assert.equal(pixels[11],0);assert.deepEqual(pixels.slice(12,16),[255,255,255,255]);
  }finally{await session.context.close();}
  const measured=await measureInput(await readInput({html}),runtime);
  assert.ok(measured.diagnostics.some(d=>d.code==='PPTX_IMAGE_DUOTONE_BAKED'&&d.severity==='warning'));
 }finally{await runtime.close();}
});
test('safe but unsupported colors or PNG encoding remain explicit approximations',async()=>{
 for(const [data,stops] of [[rgba,'<a:prstClr val="unknown"/><a:srgbClr val="FFFFFF"/>'],[rgb,'<a:srgbClr val="000000"/><a:srgbClr val="FFFFFF"/>']]){
  const result=await readPptx(zipSync(await fixture(data,stops)),{allowPartial:true});assert.equal(result.status,'parsed',JSON.stringify(result.diagnostics));
  assert.ok(result.diagnostics.some(d=>d.code==='PPTX_IMAGE_APPEARANCE_APPROXIMATED'));assert.ok(!result.diagnostics.some(d=>d.code==='PPTX_IMAGE_DUOTONE_BAKED'));
 }
});
test('malformed PNG data is blocked rather than decoded as a duotone',async()=>{
 const parts=await fixture(),media=Object.keys(parts).find(k=>/^ppt\/media\/.*\.png$/.test(k));
 const corrupted=Buffer.from(parts[media]);corrupted[corrupted.length-1]^=1;parts[media]=corrupted;
 const result=await readPptx(zipSync(parts),{allowPartial:true});assert.equal(result.status,'failed');assert.ok(result.diagnostics.some(d=>d.blocking));
});
