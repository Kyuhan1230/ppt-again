import test from 'node:test';
import assert from 'node:assert/strict';
import PptxGenJS from 'pptxgenjs';
import {unzipSync,zipSync,strFromU8,strToU8} from 'fflate';
import {readPptx} from '../src/html-to-pptx/read-pptx.mjs';
import {exportHtml} from '../src/html-to-pptx/export-html.mjs';
import {publicDiagnostic} from '../src/html-to-pptx/public-diagnostics.mjs';

const png=Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+j1ioAAAAASUVORK5CYII=','base64');
const jpeg=Buffer.from('/9j/4AAQSkZJRgABAQEAYABgAAD/2wBDAAMCAgMCAgMDAwMEAwMEBQgFBQQEBQoHBwYIDAoMDAsKCwsNDhIQDQ4RDgsLEBYQERMUFRUVDA8XGBYUGBIUFRT/2wBDAQMEBAUEBQkFBQkUDQsNFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBT/wAARCAACAAIDASIAAhEBAxEB/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAAAgEDAwIEAwUFBAQAAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3ODk6Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uHi4+Tl5ufo6erx8vP09fb3+Pn6/8QAHwEAAwEBAQEBAQEBAQAAAAAAAAECAwQFBgcICQoL/8QAtREAAgECBAQDBAcFBAQAAQJ3AAECAxEEBSExBhJBUQdhcRMiMoEIFEKRobHBCSMzUvAVYnLRChYkNOEl8RcYGRomJygpKjU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6goOEhYaHiImKkpOUlZaXmJmaoqOkpaanqKmqsrO0tba3uLm6wsPExcbHyMnK0tPU1dbX2Nna4uPk5ebn6Onq8vP09fb3+Pn6/9oADAMBAAIRAxEAPwD8qqKKKAP/2Q==','base64');
async function fixture(){
  const ppt=new PptxGenJS(),slide=ppt.addSlide();slide.background={color:'FFFFFF'};
  slide.addImage({data:'data:image/png;base64,'+png.toString('base64'),x:1,y:1,w:1,h:1});
  const parts=unzipSync(new Uint8Array(await ppt.write({outputType:'uint8array'})));
  const media=Object.keys(parts).find(k=>/^ppt\/media\/.*\.png$/.test(k));assert.ok(media);
  return {parts,media};
}
test('PNG-named PPTX image containing JPEG uses the validated JPEG media type',async()=>{
  const {parts,media}=await fixture();parts[media]=jpeg;
  const result=await readPptx(zipSync(parts),{allowPartial:true});
  assert.equal(result.status,'parsed',JSON.stringify(result.diagnostics));
  const image=result.model.assets.find(a=>a.mime==='image/jpeg');assert.ok(image);
  assert.deepEqual(Buffer.from(image.bytes),jpeg);
  assert.match(exportHtml(result.model).html,/data:image\/jpeg;base64,/);
});
test('JPEG-named PPTX image containing PNG uses the validated PNG media type',async()=>{
  const {parts,media}=await fixture(),renamed=media.replace(/\.png$/,'.jpg');
  parts[renamed]=parts[media];delete parts[media];
  const name='ppt/slides/_rels/slide1.xml.rels';
  parts[name]=strToU8(strFromU8(parts[name]).replace(/(\.\.\/media\/[^" ]+)\.png/g,'$1.jpg'));
  const result=await readPptx(zipSync(parts),{allowPartial:true});
  assert.equal(result.status,'parsed',JSON.stringify(result.diagnostics));
  assert.deepEqual(Buffer.from(result.model.assets.find(a=>a.mime==='image/png').bytes),png);
  assert.match(exportHtml(result.model).html,/data:image\/png;base64,/);
});
test('unknown bytes and oversized images still block conversion',async()=>{
  const oversized=Buffer.alloc(16*1024*1024+1,0);png.copy(oversized,0,0,8);
  for(const bytes of [Buffer.from('<svg onload="alert(1)"></svg>'),oversized]){
    const {parts,media}=await fixture();parts[media]=bytes;
    const result=await readPptx(zipSync(parts),{allowPartial:true});
    assert.equal(result.status,'failed');assert.equal(result.model,null);
    assert.ok(result.diagnostics.some(d=>d.severity==='error'&&d.blocking));
  }
});
test('external image relationships remain forbidden',async()=>{
  const {parts}=await fixture(),name='ppt/slides/_rels/slide1.xml.rels';
  const rels=strFromU8(parts[name]);
  parts[name]=strToU8(rels.replace(/Target="\.\.\/media\/[^\"]+"/, 'Target="https://example.com/image.png" TargetMode="External"'));
  assert.notEqual(strFromU8(parts[name]),rels);
  const result=await readPptx(zipSync(parts),{allowPartial:true});
  assert.equal(result.status,'failed');assert.ok(result.diagnostics.some(d=>d.blocking));
});
test('partial import keeps a cropped picture when its custom mask and shadow cannot be rendered',async()=>{
  const {parts}=await fixture(),slide='ppt/slides/slide1.xml';
  let xml=strFromU8(parts[slide]);
  xml=xml.replace('<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>','<a:custGeom><a:avLst/><a:pathLst/></a:custGeom><a:effectLst><a:outerShdw blurRad="50800"/></a:effectLst>');
  xml=xml.replace('<a:stretch><a:fillRect/></a:stretch>','<a:srcRect l="10000"/><a:stretch><a:fillRect/></a:stretch>');
  parts[slide]=strToU8(xml);const bytes=zipSync(parts);
  assert.equal((await readPptx(bytes)).status,'failed');
  const result=await readPptx(bytes,{allowPartial:true});
  assert.equal(result.status,'parsed',JSON.stringify(result.diagnostics));
  const image=result.model.slides[0].nodes.find(n=>n.kind==='image');assert.ok(image);
  assert.equal(image.content.crop.left,10000);
  assert.ok(!result.diagnostics.some(d=>d.code==='PPTX_OBJECT_PLACEHOLDER'&&d.objectType==='image'));
  const warning=result.diagnostics.find(d=>d.code==='PPTX_IMAGE_APPEARANCE_APPROXIMATED');assert.equal(warning?.severity,'warning');
  const pub=publicDiagnostic(warning);assert.equal(pub.code,'PPTX_IMAGE_APPEARANCE_APPROXIMATED');
  assert.match(pub.message,/마스크|그림자/);
  const html=exportHtml(result.model).html;assert.match(html,/data:image\/png;base64,/);assert.match(html,/object-view-box:inset/);
});
test('image color effects are reported instead of silently omitted',async()=>{
  const {parts}=await fixture(),slide='ppt/slides/slide1.xml';
  const xml=strFromU8(parts[slide]);parts[slide]=strToU8(xml.replace('</a:blip>','<a:grayscl/></a:blip>'));
  const bytes=zipSync(parts);assert.equal((await readPptx(bytes)).status,'failed');
  const result=await readPptx(bytes,{allowPartial:true});assert.equal(result.status,'parsed',JSON.stringify(result.diagnostics));
  assert.ok(result.model.slides[0].nodes.some(n=>n.kind==='image'));
  const warning=result.diagnostics.find(d=>d.code==='PPTX_IMAGE_APPEARANCE_APPROXIMATED');assert.equal(warning?.severity,'warning');
  assert.match(publicDiagnostic(warning).message,/색상/);
});
