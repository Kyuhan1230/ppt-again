import test from 'node:test';
import assert from 'node:assert/strict';
import PptxGenJS from 'pptxgenjs';
import {unzipSync,zipSync,strFromU8,strToU8} from 'fflate';
import {readPptx} from '../src/html-to-pptx/read-pptx.mjs';
import {publicDiagnostic} from '../src/html-to-pptx/public-diagnostics.mjs';

async function fixture(){
 const p=new PptxGenJS(),s=p.addSlide();s.background={color:'FFFFFF'};
 s.addText('Reference text',{x:1,y:1,w:5,h:1,fontFace:'Arial',fontSize:22,objectName:'reference'});
 const parts=unzipSync(new Uint8Array(await p.write({outputType:'uint8array'})));
 const key='ppt/slides/slide1.xml';let xml=strFromU8(parts[key]);
 xml=xml.replace(/(<a:rPr\b[^>]*?)\s+sz="[^"]*"/,'$1');parts[key]=strToU8(xml);
 return parts;
}
function edit(parts,key,change){const x=strFromU8(parts[key]),y=change(x);assert.notEqual(y,x);parts[key]=strToU8(y);}
const node=result=>result.model?.slides[0]?.nodes.find(n=>n.id==='reference');
test('ordinary shape inherits point size from master otherStyle before presentation default',async()=>{
 const parts=await fixture();edit(parts,'ppt/slideMasters/slideMaster1.xml',x=>x.replace(/(<p:otherStyle>[\s\S]*?<a:defRPr\b[^>]*?\bsz=")\d+/, '$12600'));
 const result=await readPptx(zipSync(parts),{allowPartial:true,fontPolicy:'malgun-gothic'});
 assert.equal(result.status,'parsed',JSON.stringify(result.diagnostics));
 assert.equal(node(result)?.content.text,'Reference text');assert.ok(Math.abs(node(result).style.fontSize-26*4/3)<1e-8);
 assert.ok(!result.diagnostics.some(d=>d.objectName==='reference'&&d.reasonKey==='PPTX_INHERITED_TEXT_UNRESOLVED'));
});
test('shape list style has priority over master and direct run size has priority over all',async()=>{
 const parts=await fixture();edit(parts,'ppt/slides/slide1.xml',x=>x.replace('<a:lstStyle/>','<a:lstStyle><a:lvl1pPr><a:defRPr sz="1400"/></a:lvl1pPr></a:lstStyle>'));
 const result=await readPptx(zipSync(parts),{allowPartial:true,fontPolicy:'malgun-gothic'});assert.equal(result.status,'parsed');
 assert.ok(Math.abs(node(result).style.fontSize-14*4/3)<1e-8);
 edit(parts,'ppt/slides/slide1.xml',x=>x.replace('<a:rPr ', '<a:rPr sz="3100" '));
 const explicit=await readPptx(zipSync(parts),{allowPartial:true,fontPolicy:'malgun-gothic'});assert.ok(Math.abs(node(explicit).style.fontSize-31*4/3)<1e-8);
});
const placeholder=(type,index,size)=>`<p:sp><p:nvSpPr><p:cNvPr id="100" name="inherited"/><p:cNvSpPr/><p:nvPr><p:ph type="${type}" idx="${index}"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>${size?`<a:p><a:r><a:rPr sz="${size}"/><a:t>1</a:t></a:r></a:p>`:'<a:p/>'}</p:txBody></p:sp>`;
test('placeholder uses matching layout then matching master type rather than another placeholder',async()=>{
 const parts=await fixture();
 edit(parts,'ppt/slides/slide1.xml',x=>x.replace(/<p:sp>[\s\S]*?<\/p:sp>/,sp=>sp.replace('<p:nvPr></p:nvPr>','<p:nvPr><p:ph type="sldNum" idx="12"/></p:nvPr>')));
 edit(parts,'ppt/slideLayouts/slideLayout1.xml',x=>x.replace('</p:spTree>',placeholder('sldNum',12,1600)+'</p:spTree>'));
 edit(parts,'ppt/slideMasters/slideMaster1.xml',x=>x.replace('</p:spTree>',placeholder('sldNum',4,1200)+placeholder('ftr',3,3000)+'</p:spTree>'));
 let result=await readPptx(zipSync(parts),{allowPartial:true,fontPolicy:'malgun-gothic'});assert.equal(result.status,'parsed');assert.ok(Math.abs(node(result).style.fontSize-16*4/3)<1e-8);
 edit(parts,'ppt/slideLayouts/slideLayout1.xml',x=>x.replace('sz="1600"',''));
 result=await readPptx(zipSync(parts),{allowPartial:true,fontPolicy:'malgun-gothic'});assert.equal(result.status,'parsed');assert.ok(Math.abs(node(result).style.fontSize-12*4/3)<1e-8);
});
test('unresolved point size preserves text only in partial mode with an explicit approximation warning',async()=>{
 const parts=await fixture();for(const key of ['ppt/presentation.xml','ppt/slideMasters/slideMaster1.xml']) edit(parts,key,x=>x.replace(/\s+sz="\d+"/g,''));
 const strict=await readPptx(zipSync(parts),{fontPolicy:'malgun-gothic'});assert.equal(strict.status,'failed');
 const result=await readPptx(zipSync(parts),{allowPartial:true,fontPolicy:'malgun-gothic'});assert.equal(result.status,'parsed');assert.equal(node(result).content.text,'Reference text');
 assert.equal(node(result).style.fontSize,24);const warning=result.diagnostics.find(d=>d.reasonKey==='PPTX_INHERITED_SIZE_APPROXIMATED'&&d.severity==='warning');assert.ok(warning);
 const visible=publicDiagnostic(warning);assert.equal(visible.reasonKey,'PPTX_INHERITED_SIZE_APPROXIMATED');assert.match(visible.message,/18pt/);assert.match(visible.reason,/18pt/);
});
test('placeholder inherits layout text color and explicit run color wins',async()=>{
 const parts=await fixture();
 edit(parts,'ppt/slides/slide1.xml',x=>x.replace(/<p:sp>[\s\S]*?<\/p:sp>/,sp=>sp.replace('<p:nvPr></p:nvPr>','<p:nvPr><p:ph type="title" idx="0"/></p:nvPr>').replace(/<a:solidFill><a:srgbClr val="000000"\/><\/a:solidFill>/,'')));
 const layout='<p:sp><p:nvSpPr><p:cNvPr id="101" name="title"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle><a:lvl1pPr><a:defRPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:defRPr></a:lvl1pPr></a:lstStyle><a:p/></p:txBody></p:sp>';
 edit(parts,'ppt/slideLayouts/slideLayout1.xml',x=>x.replace('</p:spTree>',layout+'</p:spTree>'));
 let result=await readPptx(zipSync(parts),{allowPartial:true,fontPolicy:'malgun-gothic'});assert.equal(result.status,'parsed');assert.equal(node(result).style.color.toLowerCase(),'#ffffff');
 edit(parts,'ppt/slides/slide1.xml',x=>x.replace('<a:rPr ', '<a:rPr ').replace('</a:rPr>','<a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></a:rPr>'));
 result=await readPptx(zipSync(parts),{allowPartial:true,fontPolicy:'malgun-gothic'});assert.equal(node(result).style.color.toLowerCase(),'#00ff00');
});

test('a title placeholder inherits horizontal and vertical alignment from its layout',async()=>{
 const parts=await fixture();
 edit(parts,'ppt/slides/slide1.xml',x=>x.replace(/<p:sp>[\s\S]*?<\/p:sp>/,sp=>sp.replace('<p:nvPr></p:nvPr>','<p:nvPr><p:ph type="ctrTitle"/></p:nvPr>').replace(/\s+anchor="[^"]*"/,'')));
 const layout='<p:sp><p:nvSpPr><p:cNvPr id="101" name="title"/><p:cNvSpPr/><p:nvPr><p:ph type="ctrTitle"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr anchor="b"/><a:lstStyle><a:lvl1pPr algn="ctr"><a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:defRPr sz="2400"/></a:lvl1pPr></a:lstStyle><a:p/></p:txBody></p:sp>';
 edit(parts,'ppt/slideLayouts/slideLayout1.xml',x=>x.replace('</p:spTree>',layout+'</p:spTree>'));
 let result=await readPptx(zipSync(parts),{allowPartial:true,fontPolicy:'malgun-gothic'});assert.equal(result.status,'parsed',JSON.stringify(result.diagnostics));
 assert.equal(node(result).style.align,'center');assert.equal(node(result).style.valign,'bottom');assert.equal(node(result).style.lineHeight,'1.08');
 edit(parts,'ppt/slides/slide1.xml',x=>x.replace('<a:bodyPr ','<a:bodyPr anchor="t" ').replace('<a:pPr ', '<a:pPr algn="r" '));
 result=await readPptx(zipSync(parts),{allowPartial:true,fontPolicy:'malgun-gothic'});assert.equal(node(result).style.align,'right');assert.equal(node(result).style.valign,'top');
});
