import test from 'node:test';
import assert from 'node:assert/strict';
import PptxGenJS from 'pptxgenjs';
import {unzipSync,zipSync,strFromU8,strToU8} from 'fflate';
import {readPptx} from '../src/html-to-pptx/read-pptx.mjs';
import {exportHtml} from '../src/html-to-pptx/export-html.mjs';
import {readInput} from '../src/html-to-pptx/input.mjs';
import {measureInput} from '../src/html-to-pptx/measure.mjs';
import {createBrowserRuntime} from '../src/html-to-pptx/browser-runtime.mjs';
import {exportPptx} from '../src/html-to-pptx/export.mjs';
import {publicDiagnostic} from '../src/html-to-pptx/public-diagnostics.mjs';

async function fixture(){
  const p=new PptxGenJS(),s=p.addSlide();s.background={color:'FFFFFF'};
  s.addShape(p.ShapeType.ellipse,{x:1,y:1,w:2,h:1,rotate:30,fill:{color:'FFFFFF',transparency:100},line:{color:'000000',width:2},objectName:'rotated-ellipse'});
  s.addShape(p.ShapeType.downArrow,{x:4,y:1,w:1,h:2,rotate:60,fill:{color:'FF0000'},line:{color:'FF0000'},objectName:'rotated-arrow'});
  const parts=unzipSync(new Uint8Array(await p.write({outputType:'uint8array'})));
  const xml=strFromU8(parts['ppt/slides/slide1.xml']);
  parts['ppt/slides/slide1.xml']=strToU8(xml.replace('<a:solidFill><a:srgbClr val="FFFFFF"><a:alpha val="0"/></a:srgbClr></a:solidFill>','<a:noFill/>').replaceAll('</p:sp>','<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr/></a:p></p:txBody></p:sp>'));
  return zipSync(parts);
}

test('partial import preserves rotated ellipses and default arrows as warned, reimportable artwork',async()=>{
  const bytes=await fixture();assert.equal((await readPptx(bytes)).status,'failed');
  const result=await readPptx(bytes,{allowPartial:true});assert.equal(result.status,'parsed',JSON.stringify(result.diagnostics));
  for(const id of ['rotated-ellipse','rotated-arrow']){
    const node=result.model.slides[0].nodes.find(n=>n.id===id);assert.equal(node?.kind,'image',id);
    assert.equal(result.model.assets.find(a=>a.id===node.content.assetId)?.mime,'image/svg+xml');
    const warning=result.diagnostics.find(d=>d.code==='PPTX_SHAPE_IMAGE_APPROXIMATED'&&d.nodeId===id&&d.severity==='warning');assert.ok(warning);
    const publicWarning=publicDiagnostic(warning,result.model.slides);
    assert.equal(publicWarning.code,'PPTX_SHAPE_IMAGE_APPROXIMATED');
    assert.equal(publicWarning.reasonKey,'PPTX_SHAPE_IMAGE_APPROXIMATED');
    assert.match(publicWarning.message,/그림.*개별 편집/);
    assert.equal(result.diagnostics.some(d=>d.code==='PPTX_OBJECT_PLACEHOLDER'&&d.objectName===id),false);
  }
  const out=exportHtml(result.model);assert.equal(out.status,'generated',JSON.stringify(out.diagnostics));
  const rt=await createBrowserRuntime({capture:false});try{
    const measured=await measureInput(await readInput({html:out.html}),rt);
    const pptx=await exportPptx(measured);assert.equal(pptx.status,'generated',JSON.stringify(pptx.diagnostics));
    assert.ok(pptx.diagnostics.some(d=>d.severity==='warning'&&d.code==='UNSUPPORTED_OBJECT'));
  }finally{await rt.close();}
});

test('unsupported rotated geometry remains a visible placeholder',async()=>{
  const p=new PptxGenJS(),s=p.addSlide();s.background={color:'FFFFFF'};
  s.addShape(p.ShapeType.star5,{x:1,y:1,w:2,h:1,rotate:30,fill:{color:'FF0000'},objectName:'star'});
  const r=await readPptx(new Uint8Array(await p.write({outputType:'uint8array'})),{allowPartial:true});
  assert.equal(r.status,'parsed',JSON.stringify(r.diagnostics));
  assert.ok(r.diagnostics.some(d=>d.code==='PPTX_OBJECT_PLACEHOLDER'&&d.objectName==='star'));
});

test('simple rotated theme-colored shapes resolve their theme before rendering',async()=>{
  const parts=unzipSync(await fixture());
  const xml=strFromU8(parts['ppt/slides/slide1.xml']);
  parts['ppt/slides/slide1.xml']=strToU8(xml.replaceAll('<a:srgbClr val="FF0000"/>','<a:schemeClr val="tx1"/>'));
  const result=await readPptx(zipSync(parts),{allowPartial:true});
  const arrow=result.model?.slides[0].nodes.find(n=>n.id==='rotated-arrow');
  assert.equal(arrow?.kind,'image');
  const svg=Buffer.from(result.model.assets.find(a=>a.id===arrow.content.assetId).bytes).toString('utf8');
  assert.match(svg,/fill="#000000"/i);
});

test('rotated artwork rejects line cap, compound, alignment and effects it cannot reproduce',async()=>{
  for(const [label,extra] of [['round cap',' cap="rnd"'],['compound line',' cmpd="dbl"'],['inset line',' algn="in"']]){
    const parts=unzipSync(await fixture()),xml=strFromU8(parts['ppt/slides/slide1.xml']);
    assert.match(xml,/<a:ln w="25400">/);
    parts['ppt/slides/slide1.xml']=strToU8(xml.replace('<a:ln w="25400">',`<a:ln w="25400"${extra}>`));
    const result=await readPptx(zipSync(parts),{allowPartial:true});
    assert.equal(result.status,'parsed',`${label}: ${JSON.stringify(result.diagnostics)}`);
    assert.ok(result.diagnostics.some(d=>d.code==='PPTX_OBJECT_PLACEHOLDER'&&d.objectName==='rotated-ellipse'),label);
    assert.equal(result.model.slides[0].nodes.some(n=>n.id==='rotated-ellipse'&&n.kind==='image'),false,label);
    assert.ok(result.model.slides[0].nodes.some(n=>n.id==='rotated-arrow'&&n.kind==='image'),'supported sibling remains');
  }
  const parts=unzipSync(await fixture()),xml=strFromU8(parts['ppt/slides/slide1.xml']);
  parts['ppt/slides/slide1.xml']=strToU8(xml.replace('<a:ln w="25400">','<a:ln w="25400"><a:effectLst><a:outerShdw blurRad="100"/></a:effectLst>'));
  const result=await readPptx(zipSync(parts),{allowPartial:true});
  assert.equal(result.status,'parsed');
  assert.ok(result.diagnostics.some(d=>d.code==='PPTX_OBJECT_PLACEHOLDER'&&d.objectName==='rotated-ellipse'),'effects remain unsupported');
});
