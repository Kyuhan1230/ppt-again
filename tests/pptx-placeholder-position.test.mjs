import test from 'node:test';
import assert from 'node:assert/strict';
import PptxGenJS from 'pptxgenjs';
import {unzipSync,zipSync,strFromU8,strToU8} from 'fflate';
import {readPptx} from '../src/html-to-pptx/read-pptx.mjs';

const transform=(x,y,w,h)=>`<a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${w}" cy="${h}"/></a:xfrm>`;
const placeholder=(type,index,xfrm)=>`<p:sp><p:nvSpPr><p:cNvPr id="100" name="layout placeholder"/><p:cNvSpPr/><p:nvPr><p:ph type="${type}" idx="${index}"/></p:nvPr></p:nvSpPr><p:spPr>${xfrm}</p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>`;
const edit=(parts,key,fn)=>{const before=strFromU8(parts[key]),after=fn(before);assert.notEqual(after,before,key);parts[key]=strToU8(after);};
async function fixture({own=false,layoutTransform=transform(914400,1828800,4572000,914400),masterTransform='',duplicate=false,slidePlaceholder=true,nested=false}={}){
 const ppt=new PptxGenJS(),slide=ppt.addSlide();slide.background={color:'FFFFFF'};
 slide.addText('Inherited title',{x:1,y:1,w:5,h:1,fontFace:'Arial',fontSize:20,objectName:'Title 1'});
 const parts=unzipSync(new Uint8Array(await ppt.write({outputType:'uint8array'})));
 edit(parts,'ppt/slides/slide1.xml',xml=>xml.replace(/<p:sp>[\s\S]*?<\/p:sp>/,sp=>{
  if(slidePlaceholder)sp=sp.replace('<p:nvPr></p:nvPr>','<p:nvPr><p:ph type="title" idx="42"/></p:nvPr>');
  return own?sp:sp.replace(/<a:xfrm\b[^>]*>[\s\S]*?<\/a:xfrm>/,'');
 }));
 const layout=placeholder('title',42,layoutTransform);
 edit(parts,'ppt/slideLayouts/slideLayout1.xml',xml=>xml.replace('</p:spTree>',(nested?`<p:extLst><p:ext uri="test">${layout}</p:ext></p:extLst>`:layout+(duplicate?layout:''))+'</p:spTree>'));
 if(masterTransform)edit(parts,'ppt/slideMasters/slideMaster1.xml',xml=>xml.replace('</p:spTree>',placeholder('title',7,masterTransform)+'</p:spTree>'));
 return zipSync(parts);
}
const read=bytes=>readPptx(bytes,{fontPolicy:'original',allowPartial:true});

test('slide title inherits missing position and extent from its matching layout placeholder',async()=>{
 const result=await read(await fixture());assert.equal(result.status,'parsed',JSON.stringify(result.diagnostics));
 const title=result.model.slides[0].nodes.find(n=>n.id==='Title 1');assert.ok(title,JSON.stringify(result.diagnostics));
 assert.deepEqual(title.boxCssPx,{x:96,y:192,w:480,h:96});assert.equal(title.content.text,'Inherited title');
 assert.ok(!result.diagnostics.some(d=>d.reasonKey==='PPTX_OBJECT_POSITION_UNSUPPORTED'));
});
test('an explicit slide transform takes precedence over the layout',async()=>{
 const result=await read(await fixture({own:true}));assert.equal(result.status,'parsed',JSON.stringify(result.diagnostics));
 assert.deepEqual(result.model.slides[0].nodes.find(n=>n.id==='Title 1').boxCssPx,{x:96,y:96,w:480,h:96});
});
test('a matching master supplies position when the matching layout has no transform',async()=>{
 const result=await read(await fixture({layoutTransform:'',masterTransform:transform(1828800,2743200,3657600,914400)}));assert.equal(result.status,'parsed',JSON.stringify(result.diagnostics));
 const title=result.model.slides[0].nodes.find(n=>n.id==='Title 1');assert.ok(title,JSON.stringify(result.diagnostics));
 assert.deepEqual(title.boxCssPx,{x:192,y:288,w:384,h:96});
});
test('ambiguous matching placeholders and ordinary shapes do not borrow a position',async()=>{
 for(const options of [{duplicate:true},{slidePlaceholder:false},{nested:true}]){
  const result=await read(await fixture(options));assert.equal(result.status,'parsed',JSON.stringify(result.diagnostics));
  assert.ok(!result.model.slides[0].nodes.some(n=>n.id==='Title 1'));assert.ok(result.diagnostics.some(d=>d.reasonKey==='PPTX_OBJECT_POSITION_UNSUPPORTED'));
 }
});
test('a malformed inherited transform still stops analysis',async()=>{
 const result=await read(await fixture({layoutTransform:'<a:xfrm><a:off x="0" y="0"/></a:xfrm>'}));
 assert.equal(result.status,'failed');assert.equal(result.model,null);assert.ok(result.diagnostics.some(d=>d.reasonKey==='PPTX_INVALID_COORDINATES'&&d.blocking));
});
