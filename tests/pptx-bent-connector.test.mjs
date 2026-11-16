import test from 'node:test';
import assert from 'node:assert/strict';
import PptxGenJS from 'pptxgenjs';
import {unzipSync, zipSync, strFromU8, strToU8} from 'fflate';
import {readPptx} from '../src/html-to-pptx/read-pptx.mjs';
import {exportHtml} from '../src/html-to-pptx/export-html.mjs';
import {readInput} from '../src/html-to-pptx/input.mjs';
import {measureInput} from '../src/html-to-pptx/measure.mjs';
import {createBrowserRuntime} from '../src/html-to-pptx/browser-runtime.mjs';

async function connector({shape='bentConnector3', adjustment='25000', flip=''}={}) {
  const ppt = new PptxGenJS(); ppt.layout='LAYOUT_WIDE'; ppt.addSlide();
  const parts=unzipSync(new Uint8Array(await ppt.write({outputType:'uint8array'})));
  const xml=strFromU8(parts['ppt/slides/slide1.xml']);
  const object=`<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="42" name="Elbow"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr><p:spPr><a:xfrm ${flip}><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm><a:prstGeom prst="${shape}"><a:avLst><a:gd name="adj1" fmla="val ${adjustment}"/></a:avLst></a:prstGeom><a:ln w="9525"><a:solidFill><a:srgbClr val="112233"/></a:solidFill></a:ln></p:spPr></p:cxnSp>`;
  parts['ppt/slides/slide1.xml']=strToU8(xml.replace('</p:spTree>',object+'</p:spTree>'));
  return zipSync(parts);
}

test('a bent connector keeps its adjustable orthogonal path and stays editable after HTML import',async()=>{
  const parsed=await readPptx(await connector());
  assert.equal(parsed.status,'parsed',JSON.stringify(parsed.diagnostics));
  const html=exportHtml(parsed.model); assert.equal(html.status,'generated',JSON.stringify(html.diagnostics));
  const runtime=await createBrowserRuntime({capture:false});
  try {
    const measured=await measureInput(await readInput({html:html.html}),runtime);
    assert.equal(measured.diagnostics.some(d=>d.severity==='error'),false,JSON.stringify(measured.diagnostics));
    const lines=measured.model.slides[0].nodes.filter(n=>n.kind==='line');
    assert.equal(lines.length,3);
    assert.deepEqual(lines.map(n=>n.style.endpoints),[
      {start:{x:96,y:96},end:{x:144,y:96}},
      {start:{x:144,y:96},end:{x:144,y:192}},
      {start:{x:144,y:192},end:{x:288,y:192}}
    ]);
  } finally {await runtime.close();}
});

test('an unimplemented connector geometry is diagnosed rather than drawn as a straight line',async()=>{
  const parsed=await readPptx(await connector({shape:'bentConnector4'}));
  assert.equal(parsed.status,'failed');
  assert.ok(parsed.diagnostics.some(d=>d.code==='UNSUPPORTED_OBJECT'));
});

test('bent connectors keep their reflected direction',async()=>{
  const parsed=await readPptx(await connector({flip:'flipH="1" flipV="1"'}));
  assert.equal(parsed.status,'parsed',JSON.stringify(parsed.diagnostics));
  const line=parsed.model.slides[0].nodes.find(n=>n.id==='Elbow');
  assert.deepEqual(line.style.connectorPoints,[{x:288,y:192},{x:240,y:192},{x:240,y:96},{x:96,y:96}]);
});
