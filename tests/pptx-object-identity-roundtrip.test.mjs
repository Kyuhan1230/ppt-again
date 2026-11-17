import test,{before,after} from 'node:test';
import assert from 'node:assert/strict';
import PptxGenJS from 'pptxgenjs';
import {DOMParser} from '@xmldom/xmldom';
import {unzipSync,zipSync,strFromU8,strToU8} from 'fflate';
import {readPptx} from '../src/html-to-pptx/read-pptx.mjs';
import {exportHtml} from '../src/html-to-pptx/export-html.mjs';
import {readInput} from '../src/html-to-pptx/input.mjs';
import {measureInput} from '../src/html-to-pptx/measure.mjs';
import {createBrowserRuntime} from '../src/html-to-pptx/browser-runtime.mjs';
import {exportPptx} from '../src/html-to-pptx/export.mjs';

let runtime;
before(async()=>{runtime=await createBrowserRuntime({capture:false});});
after(async()=>runtime?.close());

async function fixture(label='Current ',lineWidth=2){
  const ppt=new PptxGenJS();ppt.layout='LAYOUT_WIDE';
  const slide=ppt.addSlide();slide.background={color:'FFFFFF'};
  slide.addText([{text:label,options:{bold:false}},{text:'plan',options:{bold:true}}],{
    shape:'roundRect',rectRadius:0.2,isTextBox:false,fill:{color:'123456'},line:{color:'ABCDEF',width:lineWidth},
    x:1,y:1,w:3,h:1,fontFace:'Arial',fontSize:18,margin:0,objectName:'text-with-fill'
  });
  return new Uint8Array(await ppt.write({outputType:'uint8array'}));
}
async function nestedFixture(){
  const parts=unzipSync(await fixture()),file='ppt/slides/slide1.xml';
  const xml=strFromU8(parts[file]),shape=xml.match(/<p:sp>[\s\S]*?<\/p:sp>/)?.[0];
  assert.ok(shape);
  const group=(id,name,content)=>`<p:grpSp><p:nvGrpSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="2743200" cy="914400"/><a:chOff x="914400" y="914400"/><a:chExt cx="2743200" cy="914400"/></a:xfrm></p:grpSpPr>${content}</p:grpSp>`;
  parts[file]=strToU8(xml.replace(shape,group(900,'outer',group(901,'inner',shape))));return zipSync(parts);
}
function slideDocument(bytes,number=1){return new DOMParser().parseFromString(strFromU8(unzipSync(bytes)[`ppt/slides/slide${number}.xml`]),'application/xml');}
function nativeShapes(document){return [...document.getElementsByTagNameNS('*','sp')];}
function sourceElement(html){
  const match=html.match(/<div\b(?=[^>]*\bdata-source-id="text-with-fill")[\s\S]*?<\/div>/);
  assert.ok(match,'fixture should contain the original text-shape element');return match[0];
}
function stripOrigin(html){return html.replace(/\sdata-pptx-origin-document=(?:"[^"]*"|'[^']*')/g,'').replace(/\sdata-pptx-origin=(?:"[^"]*"|'[^']*')/g,'');}
function originAttribute(element){
  const match=element.match(/\sdata-pptx-origin=(['"])(.*?)\1/);
  assert.ok(match,'new editable HTML must describe the original object relation');
  const decoded=match[2].replaceAll('&quot;','"').replaceAll('&#39;',"'").replaceAll('&lt;','<').replaceAll('&gt;','>').replaceAll('&amp;','&');
  return {raw:match[0],value:JSON.parse(decoded)};
}
function editOrigin(element,edit){
  const {raw,value}=originAttribute(element);edit(value);
  const encoded=JSON.stringify(value).replaceAll('&','&amp;').replaceAll('"','&quot;').replaceAll('<','&lt;').replaceAll('>','&gt;');
  return element.replace(raw,` data-pptx-origin="${encoded}"`);
}
async function fromPpt(bytes){
  bytes??=await fixture();
  const read=await readPptx(bytes);assert.equal(read.status,'parsed',JSON.stringify(read.diagnostics));
  const html=exportHtml(read.model);assert.equal(html.status,'generated',JSON.stringify(html.diagnostics));
  return {bytes,read,html:html.html};
}
async function fromHtml(html){
  const measured=await measureInput(await readInput({html}),runtime);
  const output=await exportPptx(measured);return {measured,output};
}

test('one PowerPoint text shape with fill and border remains one editable native shape after HTML round trip',async()=>{
  const source=await fromPpt(),original=nativeShapes(slideDocument(source.bytes));
  assert.equal(original.length,1);
  const {output}=await fromHtml(source.html);
  assert.equal(output.status,'generated',JSON.stringify(output.diagnostics));
  const shapes=nativeShapes(slideDocument(output.pptxBytes));
  assert.equal(shapes.length,1,'a single source shape must not become separate background and text shapes');
  assert.match(shapes[0].textContent,/Current plan/);
  assert.equal(shapes[0].getElementsByTagNameNS('*','prstGeom')[0]?.getAttribute('prst'),'roundRect');
  assert.ok(shapes[0].getElementsByTagNameNS('*','solidFill').length>=1,'shape fill remains editable');
  assert.ok(shapes[0].getElementsByTagNameNS('*','ln').length>=1,'shape border remains editable');
  const originalSp=original[0].getElementsByTagNameNS('*','spPr')[0],outputSp=shapes[0].getElementsByTagNameNS('*','spPr')[0];
  assert.equal(outputSp.getElementsByTagNameNS('*','ln')[0]?.getAttribute('w'),originalSp.getElementsByTagNameNS('*','ln')[0]?.getAttribute('w'),'border width must remain unchanged');
  const beforeAdj=originalSp.getElementsByTagNameNS('*','gd')[0]?.getAttribute('fmla'),afterAdj=outputSp.getElementsByTagNameNS('*','gd')[0]?.getAttribute('fmla');
  assert.equal(afterAdj,beforeAdj,'roundRect adjustment must preserve the original radius');
  const bold=[...shapes[0].getElementsByTagNameNS('*','r')].find(run=>run.textContent.includes('plan'));
  assert.equal(bold?.getElementsByTagNameNS('*','rPr')[0]?.getAttribute('b'),'1');
  assert.equal(output.objectMap.filter(item=>item.sourceNodeId==='text-with-fill').length,1);
});

test('thin original fractional PPT border widths survive browser CSS pixel rounding',async()=>{
  for(const width of [.25,.5,1,2]){
    const source=await fromPpt(await fixture('Current ',width)),{output}=await fromHtml(source.html);
    assert.equal(output.status,'generated',`${width}pt: ${JSON.stringify(output.diagnostics)}`);
    const shapes=nativeShapes(slideDocument(output.pptxBytes));assert.equal(shapes.length,1,`${width}pt`);
    const line=shapes[0].getElementsByTagNameNS('*','spPr')[0].getElementsByTagNameNS('*','ln')[0];
    assert.equal(Number(line.getAttribute('w')),Math.round(width*12700),`${width}pt width must stay native`);
  }
});

test('current HTML text, fill, outline, roundness and position are used when restoring the single shape',async()=>{
  const source=await fromPpt();
  const before=sourceElement(source.html);
  const edited=before.replace('Current ','Updated ').replace('background:#123456','background:#cc3300')
    .replace('#abcdef','#00cc44').replace('left:96px','left:120px').replace('border-radius:19.2px','border-radius:24px');
  assert.notEqual(before,edited);
  const {measured,output}=await fromHtml(source.html.replace(before,edited));
  assert.equal(output.status,'generated',JSON.stringify(output.diagnostics));
  const shapes=nativeShapes(slideDocument(output.pptxBytes));assert.equal(shapes.length,1);
  const shape=shapes[0],spPr=shape.getElementsByTagNameNS('*','spPr')[0];
  assert.match(shape.textContent,/Updated plan/);
  assert.equal(spPr.getElementsByTagNameNS('*','solidFill')[0]?.getElementsByTagNameNS('*','srgbClr')[0]?.getAttribute('val'),'CC3300');
  assert.equal(spPr.getElementsByTagNameNS('*','ln')[0]?.getElementsByTagNameNS('*','srgbClr')[0]?.getAttribute('val'),'00CC44');
  assert.equal(Number(spPr.getElementsByTagNameNS('*','ln')[0]?.getAttribute('w')),25400);
  assert.equal(spPr.getElementsByTagNameNS('*','prstGeom')[0]?.getAttribute('prst'),'roundRect');
  const x=Number(spPr.getElementsByTagNameNS('*','off')[0]?.getAttribute('x'));
  assert.ok(Math.abs(x-1143000)<1000,`current HTML x should become 1.25 inches; got ${x}`);
  assert.equal(output.objectMap.filter(item=>item.sourceNodeId==='text-with-fill').length,1);
  assert.ok(measured.model.slides[0].nodes.some(n=>n.kind==='text'&&n.content.runs?.some(r=>r.text.includes('Updated'))));
  const reread=await readPptx(output.pptxBytes);assert.equal(reread.status,'parsed',JSON.stringify(reread.diagnostics));
  const restored=reread.model.slides[0].nodes.find(n=>n.id==='text-with-fill');
  assert.ok(Math.abs(restored.style.radius-24)<.1,`edited radius must remain 24 CSS pixels; got ${restored.style.radius}`);
});

test('removing the current HTML fill does not restore the old source fill',async()=>{
  const source=await fromPpt(),element=sourceElement(source.html);
  const edited=element.replace('background:#123456','background:transparent');assert.notEqual(edited,element);
  const {output}=await fromHtml(source.html.replace(element,edited));
  assert.equal(output.status,'generated',JSON.stringify(output.diagnostics));
  const shapes=nativeShapes(slideDocument(output.pptxBytes));assert.equal(shapes.length,1);
  const spPr=shapes[0].getElementsByTagNameNS('*','spPr')[0];
  const fill=spPr.getElementsByTagNameNS('*','solidFill')[0];
  assert.notEqual(fill?.getElementsByTagNameNS('*','srgbClr')[0]?.getAttribute('val'),'123456');
  assert.match(shapes[0].textContent,/Current plan/);
});

test('an edited text run uses its current font formatting inside the restored shape',async()=>{
  const source=await fromPpt(),element=sourceElement(source.html);
  const edited=element.replace('font-weight:700','font-weight:400').replace('font-size:24px;font-weight:400;font-style:normal;color:#000000;letter-spacing:0px;text-decoration:none;">plan',
    'font-size:28px;font-weight:400;font-style:normal;color:#bb0000;letter-spacing:0px;text-decoration:none;">plan');
  assert.notEqual(edited,element);
  const {output}=await fromHtml(source.html.replace(element,edited));
  assert.equal(output.status,'generated',JSON.stringify(output.diagnostics));
  const shapes=nativeShapes(slideDocument(output.pptxBytes));assert.equal(shapes.length,1);
  const run=[...shapes[0].getElementsByTagNameNS('*','r')].find(r=>r.textContent.includes('plan'));
  assert.ok(run,'the edited text run remains individually editable');
  const properties=run.getElementsByTagNameNS('*','rPr')[0];
  assert.notEqual(properties.getAttribute('b'),'1');
  assert.equal(properties.getAttribute('sz'),'2100','28 CSS pixels corresponds to 21 points');
  assert.equal(properties.getElementsByTagNameNS('*','srgbClr')[0]?.getAttribute('val'),'BB0000');
});

test('legacy v1 and v2 editable HTML without origin relation keep separate text and background',async()=>{
  const source=await fromPpt(),legacy=stripOrigin(source.html);
  for(const version of [1,2]){
    const html=legacy.replace(/data-slidehandoff="[12]"/,`data-slidehandoff="${version}"`);
    const {output}=await fromHtml(html);
    assert.equal(output.status,'generated',`v${version}: ${JSON.stringify(output.diagnostics)}`);
    const shapes=nativeShapes(slideDocument(output.pptxBytes));
    assert.equal(shapes.length,2,`v${version}: old HTML has no trustworthy relation to merge`);
    assert.ok(shapes.some(s=>s.textContent.includes('Current plan')));
  }
});

test('ordinary HTML does not merge a painted background and text merely because they share one element',async()=>{
  const html='<section data-slide style="position:relative;width:800px;height:600px;background:white"><div style="position:absolute;left:80px;top:80px;width:240px;height:80px;background:#123456;font:24px Arial">Ordinary text</div></section>';
  const {output}=await fromHtml(html);assert.equal(output.status,'generated',JSON.stringify(output.diagnostics));
  const shapes=nativeShapes(slideDocument(output.pptxBytes));assert.equal(shapes.length,2);
  assert.ok(shapes.some(s=>s.textContent.includes('Ordinary text')));
});

test('edited decorations that cannot fit one native shape are not silently flattened',async()=>{
  const source=await fromPpt(),element=sourceElement(source.html);
  const edits=[
    ['single-sided border',e=>e.replace('display:flex;','border-right:8px solid #ff0000;display:flex;')],
    ['content-box background and padding',e=>e.replace('padding:0px 0px 0px 0px;','padding:12px 12px 12px 12px;').replace('background:#123456;','background:#123456;background-clip:content-box;')],
    ['moved outline',e=>e.replace('outline-offset:-1.3333333333px;','outline-offset:8px;')],
    ['border and separate outline',e=>e.replace('display:flex;','border:4px solid #ff0000;display:flex;')],
    ['CSS border with no outline',e=>e.replace(/outline:[^;]*;/,'outline:none;').replace('display:flex;','border:10px solid #ff0000;display:flex;')],
    ['unequal corner radii',e=>e.replace('display:flex;','border-top-left-radius:4px;display:flex;')]
  ];
  for(const [label,edit] of edits){
    const changed=edit(element);assert.notEqual(changed,element,label);
    const {output}=await fromHtml(source.html.replace(element,changed));
    assert.equal(output.status,'failed',`${label} cannot be represented faithfully as the current PPT shape`);
    assert.equal(output.pptxBytes,null);
    assert.ok(output.diagnostics.some(d=>d.severity==='error'&&d.code==='UNSUPPORTED_OBJECT'),`${label}: ${JSON.stringify(output.diagnostics)}`);
  }
});

test('current important CSS overrides the old inline outline width on the restored shape',async()=>{
  const source=await fromPpt();
  const html=source.html.replace('</head>','<style>[data-source-id="text-with-fill"]{outline-width:8px!important;outline-offset:-4px!important;background-color:#33aa22!important}</style></head>');
  const {output}=await fromHtml(html);assert.equal(output.status,'generated',JSON.stringify(output.diagnostics));
  const shapes=nativeShapes(slideDocument(output.pptxBytes));assert.equal(shapes.length,1);
  const spPr=shapes[0].getElementsByTagNameNS('*','spPr')[0],line=spPr.getElementsByTagNameNS('*','ln')[0];
  assert.equal(Number(line.getAttribute('w')),76200,'8 CSS pixels of current outline must become 6pt');
  assert.equal(spPr.getElementsByTagNameNS('*','solidFill')[0]?.getElementsByTagNameNS('*','srgbClr')[0]?.getAttribute('val'),'33AA22');
});

test('an important outline override that changes its position is not silently saved at the old position',async()=>{
  const source=await fromPpt();
  const html=source.html.replace('</head>','<style>[data-source-id="text-with-fill"]{outline-width:8px!important}</style></head>');
  const {output}=await fromHtml(html);
  assert.equal(output.status,'failed','a visible outline in a different position must not silently become the old line');
  assert.equal(output.pptxBytes,null);
  assert.ok(output.diagnostics.some(d=>d.severity==='error'&&d.code==='UNSUPPORTED_OBJECT'));
});

test('group opacity on a painted shape with direct text cannot be flattened into independent alpha values',async()=>{
  const source=await fromPpt(),element=sourceElement(source.html);
  const direct=element.replace(/<span data-pptx-text-content[^>]*>[\s\S]*<\/span><\/div>$/,'Current plan</div>')
    .replace('opacity:1;','opacity:0.5;');
  assert.notEqual(direct,element);assert.ok(!direct.includes('data-pptx-text-content'));
  const {output}=await fromHtml(source.html.replace(element,direct));
  assert.equal(output.status,'failed','a translucent composited text-and-background group must not be flattened into separate alpha channels');
  assert.equal(output.pptxBytes,null);
  assert.ok(output.diagnostics.some(d=>d.severity==='error'&&d.code==='UNSUPPORTED_OBJECT'));
});

test('invalid origin metadata is not used to merge otherwise independent content',async()=>{
  const source=await fromPpt(),element=sourceElement(source.html);
  const mutations=[
    ['unsupported relation version',origin=>{origin.version=99;}],
    ['wrong document identity',origin=>{origin.documentId='0'.repeat(64);}],
    ['wrong slide identity',origin=>{origin.slideId='another-slide';}],
    ['wrong object identity',origin=>{origin.objectId='another-object';}],
    ['wrong role',origin=>{origin.role='unrelated-role';}],
    ['wrong parent',origin=>{origin.parentId='another-parent';}]
  ];
  for(const [label,mutation] of mutations){
    const edited=editOrigin(element,mutation),{output}=await fromHtml(source.html.replace(element,edited));
    assert.equal(output.status,'generated',`${label}: ${JSON.stringify(output.diagnostics)}`);
    const shapes=nativeShapes(slideDocument(output.pptxBytes));
    assert.equal(shapes.length,2,`${label} must not merge`);
    assert.ok(shapes.some(shape=>shape.textContent.includes('Current plan')),`${label} must preserve the user's text`);
    assert.ok(output.diagnostics.some(d=>d.code==='PPTX_OBJECT_SPLIT'),`${label} must report the split`);
  }
});

test('malformed relation JSON and a malformed document identity cannot silently merge',async()=>{
  const source=await fromPpt(),element=sourceElement(source.html),origin=originAttribute(element);
  const brokenJson=element.replace(origin.raw,' data-pptx-origin="not-json"');
  const badRoot=source.html.replace(/data-pptx-origin-document="[^"]*"/,'data-pptx-origin-document="not-a-digest"');
  assert.notEqual(badRoot,source.html);
  for(const [label,html] of [['malformed JSON',source.html.replace(element,brokenJson)],['malformed root identity',badRoot]]){
    const {output}=await fromHtml(html);assert.equal(output.status,'generated',`${label}: ${JSON.stringify(output.diagnostics)}`);
    const shapes=nativeShapes(slideDocument(output.pptxBytes));assert.equal(shapes.length,2,label);
    assert.ok(shapes.some(shape=>shape.textContent.includes('Current plan')));
    assert.ok(output.diagnostics.some(d=>d.code==='PPTX_OBJECT_SPLIT'),label);
  }
});

test('an invalid relation cannot bypass safety checks for an asymmetric border or clipped background',async()=>{
  const source=await fromPpt(),element=sourceElement(source.html),origin=originAttribute(element);
  const invalid=element.replace(origin.raw,' data-pptx-origin="not-json"');
  const changes=[
    invalid.replace('display:flex;','border-right:8px solid #ff0000;display:flex;'),
    invalid.replace('padding:0px 0px 0px 0px;','padding:12px 12px 12px 12px;').replace('background:#123456;','background:#123456;background-clip:content-box;')
  ];
  for(const changed of changes){
    const {output}=await fromHtml(source.html.replace(element,changed));
    assert.equal(output.status,'failed',JSON.stringify(output.diagnostics));assert.equal(output.pptxBytes,null);
    assert.ok(output.diagnostics.some(d=>d.code==='PPTX_OBJECT_SPLIT'));
    assert.ok(output.diagnostics.some(d=>d.code==='UNSUPPORTED_OBJECT'&&d.severity==='error'));
  }
});

test('duplicated origin identity on another element does not silently merge the two objects',async()=>{
  const source=await fromPpt(),element=sourceElement(source.html);
  originAttribute(element);
  const copy=element.replace('data-source-id="text-with-fill"','data-source-id="copied-object"')
    .replace('left:96px','left:480px').replace('Current ','Copied ');
  assert.notEqual(copy,element);
  const {output}=await fromHtml(source.html.replace(element,element+copy));
  assert.equal(output.status,'generated',JSON.stringify(output.diagnostics));
  const shapes=nativeShapes(slideDocument(output.pptxBytes)),text=shapes.map(shape=>shape.textContent).join('|');
  assert.match(text,/Current plan/);assert.match(text,/Copied plan/);
  assert.ok(shapes.length>=2,'distinct elements must remain distinct editable PPT objects');
  assert.ok(output.diagnostics.some(d=>d.code==='PPTX_OBJECT_SPLIT'));
});

test('independently moved background and text in another parent stay separate and report the split',async()=>{
  const source=await fromPpt(),element=sourceElement(source.html);
  const claim=originAttribute(element);
  const text=element.replace('data-kind="text"','data-kind="text" data-parent-id="new-parent"')
    .replace('background:#123456','background:transparent').replace(/outline:[^;]*;/,'outline:none;').replace('left:96px','left:120px');
  const group='<div data-source-id="new-parent" data-kind="group" style="position:absolute;left:0;top:0;width:1280px;height:720px"></div>';
  const background=`<div data-source-id="separate-background" data-kind="shape"${claim.raw} style="position:absolute;left:96px;top:96px;width:288px;height:96px;background:#123456"></div>`;
  const {output}=await fromHtml(source.html.replace(element,group+background+text));
  assert.equal(output.status,'generated',JSON.stringify(output.diagnostics));
  const doc=slideDocument(output.pptxBytes),shapes=nativeShapes(doc);
  assert.equal(shapes.length,2,'independent background and text must not be rejoined');
  assert.ok(shapes.some(shape=>shape.textContent.includes('Current plan')));
  assert.ok(output.diagnostics.some(d=>d.code==='PPTX_OBJECT_SPLIT'));
  assert.equal(output.objectMap.find(item=>item.sourceNodeId==='text-with-fill')?.parentId,'new-parent');
});

test('metadata copied from a different original document cannot merge into this document',async()=>{
  const first=await fromPpt(),second=await fromPpt(await fixture('Another '));
  const firstElement=sourceElement(first.html),secondElement=sourceElement(second.html);
  const firstOrigin=originAttribute(firstElement),secondOrigin=originAttribute(secondElement);
  assert.notEqual(firstOrigin.value.documentId,secondOrigin.value.documentId);
  const tampered=firstElement.replace(firstOrigin.raw,secondOrigin.raw);
  const {output}=await fromHtml(first.html.replace(firstElement,tampered));
  assert.equal(output.status,'generated',JSON.stringify(output.diagnostics));
  const shapes=nativeShapes(slideDocument(output.pptxBytes));assert.equal(shapes.length,2);
  assert.ok(shapes.some(s=>s.textContent.includes('Current plan')));
  assert.ok(output.diagnostics.some(d=>d.code==='PPTX_OBJECT_SPLIT'));
});

test('text and background in a nested PPT group stay one object inside the same group',async()=>{
  const source=await fromPpt(await nestedFixture()),{output}=await fromHtml(source.html);
  assert.equal(output.status,'generated',JSON.stringify(output.diagnostics));
  const doc=slideDocument(output.pptxBytes),groups=[...doc.getElementsByTagNameNS('*','grpSp')];
  assert.equal(groups.length,2);
  const outer=groups.find(group=>group.getElementsByTagNameNS('*','cNvPr')[0]?.getAttribute('name')==='outer');
  const inner=groups.find(group=>group.getElementsByTagNameNS('*','cNvPr')[0]?.getAttribute('name')==='inner');
  assert.ok(outer&&inner);assert.equal(inner.parentNode,outer);
  const shapes=[...inner.getElementsByTagNameNS('*','sp')];assert.equal(shapes.length,1);
  assert.match(shapes[0].textContent,/Current plan/);
  const runs=[...shapes[0].getElementsByTagNameNS('*','r')],bold=runs.find(run=>run.textContent.includes('plan'));
  assert.equal(bold?.getElementsByTagNameNS('*','rPr')[0]?.getAttribute('b'),'1');
  const mapping=output.objectMap.filter(item=>item.sourceNodeId==='text-with-fill');assert.equal(mapping.length,1);assert.equal(mapping[0].parentId,'inner');
});

test('same object name on two slides retains separate identities and one text shape per slide',async()=>{
  const ppt=new PptxGenJS();ppt.layout='LAYOUT_WIDE';
  for(const label of ['First slide','Second slide']){
    const slide=ppt.addSlide();slide.background={color:'FFFFFF'};
    slide.addText(label,{shape:'rect',isTextBox:false,x:1,y:1,w:3,h:1,fontFace:'Arial',fontSize:18,margin:0,fill:{color:'123456'},line:{color:'ABCDEF',width:1},objectName:'text-with-fill'});
  }
  const source=await fromPpt(new Uint8Array(await ppt.write({outputType:'uint8array'}))),{output}=await fromHtml(source.html);
  assert.equal(output.status,'generated',JSON.stringify(output.diagnostics));
  for(const [slideNumber,label] of [[1,'First slide'],[2,'Second slide']]){
    const shapes=nativeShapes(slideDocument(output.pptxBytes,slideNumber));assert.equal(shapes.length,1);
    assert.ok(shapes[0].textContent.includes(label));
    assert.equal(output.objectMap.filter(item=>item.sourceNodeId==='text-with-fill'&&item.pptSlideNumber===slideNumber).length,1);
  }
});
