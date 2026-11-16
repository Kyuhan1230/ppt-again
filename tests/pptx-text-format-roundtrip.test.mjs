import test from 'node:test';
import assert from 'node:assert/strict';
import PptxGenJS from 'pptxgenjs';
import {DOMParser} from '@xmldom/xmldom';
import {strFromU8, strToU8, unzipSync, zipSync} from 'fflate';
import {readPptx} from '../src/html-to-pptx/read-pptx.mjs';
import {exportHtml} from '../src/html-to-pptx/export-html.mjs';
import {readInput} from '../src/html-to-pptx/input.mjs';
import {measureInput} from '../src/html-to-pptx/measure.mjs';
import {createBrowserRuntime} from '../src/html-to-pptx/browser-runtime.mjs';
import {exportPptx} from '../src/html-to-pptx/export.mjs';
import {publicDiagnostic} from '../src/html-to-pptx/public-diagnostics.mjs';

const elements = (node, name) => [...(node?.childNodes ?? [])].filter(n => n.nodeType === 1 && n.localName === name);
const child = (node, name) => elements(node, name)[0];
const descendants = (node, name) => [...node.getElementsByTagNameNS('*', name)];
const xmlDocument = bytes => new DOMParser().parseFromString(strFromU8(unzipSync(bytes)['ppt/slides/slide1.xml']), 'application/xml');
const paragraphs = (bytes, objectName, table = false) => {
  const doc = xmlDocument(bytes);
  const object = descendants(doc, table ? 'graphicFrame' : 'sp').find(n => descendants(n, 'cNvPr')[0]?.getAttribute('name') === objectName);
  assert.ok(object, `Missing ${objectName}`);
  const body = table ? descendants(object, 'tc')[0] : child(object, 'txBody');
  return elements(table ? child(body, 'txBody') : body, 'p');
};

let sharedRuntime;
test.after(async()=>{if(sharedRuntime)await sharedRuntime.close();});
async function convertHtml(html) {
  sharedRuntime??=await createBrowserRuntime({capture:false});
  const measured=await measureInput(await readInput({html}),sharedRuntime),result=await exportPptx(measured);
  assert.equal(result.status,'generated',JSON.stringify(result.diagnostics));return result.pptxBytes;
}

test('PPT lists retain native bullets, indents, paragraph spacing and run spacing after an HTML round trip', async () => {
  const ppt = new PptxGenJS(), slide = ppt.addSlide();
  slide.addText('First bullet item with a long sentence\nSecond bullet item', {
    x: 1, y: 1, w: 3, h: 2, objectName: 'bullet-box', fontFace: 'Arial', fontSize: 20,
    margin: 0, bullet: {characterCode: '2022', indent: 18}, lineSpacing: 24,
    paraSpaceAfter: 6, charSpacing: -0.3,
  });
  const original = new Uint8Array(await ppt.write({outputType: 'uint8array'}));
  const parsed = await readPptx(original, {allowPartial: true});
  assert.equal(parsed.status, 'parsed', JSON.stringify(parsed.diagnostics));
  const output = await convertHtml(exportHtml(parsed.model).html);
  const actual = paragraphs(output, 'bullet-box');
  assert.equal(actual.length, 2);
  for (const p of actual) {
    const props = child(p, 'pPr');
    assert.equal(child(props, 'buChar')?.getAttribute('char'), '•');
    assert.equal(props.getAttribute('marL'), '228600');
    assert.equal(props.getAttribute('indent'), '-228600');
    assert.ok(!descendants(p, 't').some(t => t.textContent.includes('•')));
    assert.ok(descendants(p, 'rPr').some(r => r.getAttribute('spc') === '-30'));
  }
  assert.equal(child(child(child(actual[0], 'pPr'), 'spcAft'), 'spcPts')?.getAttribute('val'), '600');
  assert.equal(child(child(child(actual[1], 'pPr'), 'spcAft'), 'spcPts')?.getAttribute('val'), '600');
  assert.equal(child(child(actual[0], 'pPr'), 'buSzPct')?.getAttribute('val'), '100000');
});

test('HTML list markers become native PPT bullets', async () => {
  const html = `<section data-slide style="width:1280px;height:720px;font-family:Arial;font-size:24px">
    <ul style="list-style-type:disc"><li>First item</li><li>Second item</li></ul>
  </section>`;
  const output = await convertHtml(html);
  const doc = xmlDocument(output);
  assert.equal(descendants(doc, 'buChar').filter(n => n.getAttribute('char') === '•').length, 2);
  assert.ok(!descendants(doc, 't').some(t => t.textContent.includes('•')));
});

test('a native HTML list inside a table cell becomes editable PPT bullets', async () => {
  const html=`<section data-slide style="position:relative;width:1280px;height:720px;font:24px Arial"><table id="table-list" style="position:absolute;left:80px;top:80px;width:500px;border-collapse:collapse"><tr><td style="border:1px solid black"><ul style="margin:0;padding-left:24px;line-height:32px"><li style="margin-bottom:12px;letter-spacing:2px">First</li><li>Second</li></ul></td></tr></table></section>`;
  const output=await convertHtml(html),result=paragraphs(output,'table-list',true);assert.equal(result.length,2);
  assert.ok(result.every(p=>child(child(p,'pPr'),'buChar')?.getAttribute('char')==='•'));
  assert.equal(child(child(child(result[0],'pPr'),'spcAft'),'spcPts')?.getAttribute('val'),'900');
  assert.ok(descendants(result[0],'rPr').some(r=>r.getAttribute('spc')==='150'));
  const parsed=await readPptx(output,{allowPartial:true});assert.equal(parsed.status,'parsed',JSON.stringify(parsed.diagnostics));
  assert.match(exportHtml(parsed.model).html,/data-pptx-list-marker/);
});

test('nested table-cell lists retain their numbering and bullet levels', async () => {
  const html=`<section data-slide style="position:relative;width:1280px;height:720px;font:24px Arial"><table id="nested-table" style="position:absolute;left:80px;top:80px;width:500px;border-collapse:collapse"><tr><td><ol start="3" style="margin:0"><li>Parent<ul style="margin:0"><li>Child</li></ul></li><li>Next</li></ol></td></tr></table></section>`;
  const output=await convertHtml(html),result=paragraphs(output,'nested-table',true);assert.equal(result.length,3);
  assert.equal(child(child(result[0],'pPr'),'buAutoNum')?.getAttribute('startAt'),'3');
  assert.equal(child(child(result[1],'pPr'),'buChar')?.getAttribute('char'),'◦');
  assert.equal(child(child(result[2],'pPr'),'buAutoNum')?.getAttribute('startAt'),'3');
  assert.deepEqual(result.map(p=>child(p,'pPr')?.getAttribute('lvl')),['0','1','0']);
});

test('PPT table bullets and paragraph spacing survive HTML round trip', async () => {
  const ppt = new PptxGenJS(), slide = ppt.addSlide();
  slide.addTable([['First\nSecond']], {x: 1, y: 1, w: 4, h: 2, fontFace: 'Arial', fontSize: 16, objectName: 'bullet-table'});
  const parts = unzipSync(new Uint8Array(await ppt.write({outputType: 'uint8array'})));
  const key = 'ppt/slides/slide1.xml', source = strFromU8(parts[key]);
  const original = '<a:pPr indent="0" marL="0"><a:buNone/></a:pPr>';
  const replacement = '<a:pPr marL="171450" indent="-171450"><a:spcAft><a:spcPts val="500"/></a:spcAft><a:buFont typeface="Arial"/><a:buChar char="•"/></a:pPr>';
  assert.ok(source.includes(original));
  parts[key] = strToU8(source.replace(original, replacement));
  const parsed = await readPptx(zipSync(parts), {allowPartial: true});
  assert.equal(parsed.status, 'parsed', JSON.stringify(parsed.diagnostics));
  const output = await convertHtml(exportHtml(parsed.model).html);
  const first = child(paragraphs(output, 'bullet-table', true)[0], 'pPr');
  assert.equal(child(first, 'buChar')?.getAttribute('char'), '•');
  assert.equal(first.getAttribute('marL'), '171450');
  assert.equal(first.getAttribute('indent'), '-171450');
  assert.equal(child(child(first, 'spcAft'), 'spcPts')?.getAttribute('val'), '500');
  assert.ok(!child(first, 'lnSpc'));
});

test('PPT table paragraph marks retain inherited and explicit character spacing', async () => {
  const ppt=new PptxGenJS(),slide=ppt.addSlide();
  slide.addTable([['First\nSecond']],{x:1,y:1,w:4,h:2,objectName:'end-style',fontFace:'Arial',fontSize:20,bold:true});
  const parts=unzipSync(new Uint8Array(await ppt.write({outputType:'uint8array'}))),key='ppt/slides/slide1.xml',source=strFromU8(parts[key]);let index=0;
  const patched=source.replace('<a:pPr indent="0" marL="0"><a:buNone/></a:pPr>','<a:pPr marL="228600" indent="-228600"><a:buChar char="•"/></a:pPr>')
    .replace('<a:rPr lang="en-US" sz="2000" b="1" dirty="0">','<a:rPr lang="en-US" sz="2000" b="1" spc="40" dirty="0">')
    .replace(/<a:endParaRPr\b[^>]*>[\s\S]*?<\/a:endParaRPr>/g,match=>index++===0?'':match.replace('<a:endParaRPr ','<a:endParaRPr spc="70" b="0" '));
  assert.equal(index,2);parts[key]=strToU8(patched);
  const parsed=await readPptx(zipSync(parts));assert.equal(parsed.status,'parsed',JSON.stringify(parsed.diagnostics));
  const output=await convertHtml(exportHtml(parsed.model).html),result=paragraphs(output,'end-style',true);
  assert.ok(!child(result[0],'endParaRPr'));
  assert.equal(child(result[1],'endParaRPr')?.getAttribute('spc'),'70');assert.equal(child(result[1],'endParaRPr')?.getAttribute('b'),'0');
});

test('a single plain PPT paragraph keeps its paragraph-mark character spacing', async () => {
  const ppt=new PptxGenJS(),slide=ppt.addSlide();
  slide.addText('Single',{x:1,y:1,w:3,h:1,objectName:'single-end',fontFace:'Arial',fontSize:20,bold:true,charSpacing:0.4});
  const base=unzipSync(new Uint8Array(await ppt.write({outputType:'uint8array'}))),key='ppt/slides/slide1.xml',source=strFromU8(base[key]);
  for(const end of ['', '<a:endParaRPr spc="70" b="0"/>']){
    const patched=source.replace(/<a:endParaRPr\b[^>]*\/>/,end);assert.notEqual(patched,source);base[key]=strToU8(patched);
    const parsed=await readPptx(zipSync(base));assert.equal(parsed.status,'parsed',JSON.stringify(parsed.diagnostics));
    const output=await convertHtml(exportHtml(parsed.model).html),mark=child(paragraphs(output,'single-end')[0],'endParaRPr');
    if(end){assert.equal(mark?.getAttribute('spc'),'70');assert.equal(mark?.getAttribute('b'),'0');}else assert.ok(!mark);
  }
});

test('PPT list style supplies inherited bullet and character spacing', async () => {
  const ppt = new PptxGenJS(), slide = ppt.addSlide();
  slide.addText('Inherited style', {x: 1, y: 1, w: 4, h: 1, fontFace: 'Arial', fontSize: 20, objectName: 'inherited'});
  const parts = unzipSync(new Uint8Array(await ppt.write({outputType: 'uint8array'})));
  const key = 'ppt/slides/slide1.xml', source = strFromU8(parts[key]);
  const patched = source.replace('<a:lstStyle/>', '<a:lstStyle><a:lvl1pPr marL="228600" indent="-228600"><a:buChar char="•"/><a:defRPr spc="150"/></a:lvl1pPr></a:lstStyle>').replace('<a:buNone/>', '');
  assert.notEqual(patched, source);
  parts[key] = strToU8(patched);
  const parsed = await readPptx(zipSync(parts), {allowPartial: true});
  assert.equal(parsed.status, 'parsed', JSON.stringify(parsed.diagnostics));
  const html = exportHtml(parsed.model).html;
  assert.match(html, /data-pptx-list-marker/);
  assert.match(html, /letter-spacing:2px/);
  const output = await convertHtml(html);
  const p = paragraphs(output, 'inherited')[0];
  assert.equal(child(child(p, 'pPr'), 'buChar')?.getAttribute('char'), '•');
  assert.ok(descendants(p, 'rPr').some(r => r.getAttribute('spc') === '150'));
});

test('PPT bullets keep their follow-text color and font settings', async () => {
  const ppt=new PptxGenJS(),slide=ppt.addSlide();
  slide.addText('Follow text',{x:1,y:1,w:4,h:2,objectName:'follow-text',fontFace:'Arial',fontSize:20,bullet:{characterCode:'2022',indent:18}});
  const parts=unzipSync(new Uint8Array(await ppt.write({outputType:'uint8array'}))),key='ppt/slides/slide1.xml',source=strFromU8(parts[key]);
  const patched=source.replace('<a:buSzPct val="100000"/>','<a:buClrTx/><a:buSzPct val="100000"/><a:buFontTx/>');assert.notEqual(patched,source);parts[key]=strToU8(patched);
  const parsed=await readPptx(zipSync(parts));assert.equal(parsed.status,'parsed',JSON.stringify(parsed.diagnostics));
  const output=await convertHtml(exportHtml(parsed.model).html),props=child(paragraphs(output,'follow-text')[0],'pPr');
  assert.ok(child(props,'buClrTx'));assert.ok(child(props,'buFontTx'));assert.equal(child(props,'buClr'),undefined);assert.equal(child(props,'buFont'),undefined);
});

test('editing only the HTML bullet color overrides follow-text color', async () => {
  const ppt=new PptxGenJS(),slide=ppt.addSlide();
  slide.addText('Text',{x:1,y:1,w:4,h:2,objectName:'edited-color',fontFace:'Arial',fontSize:20,bullet:{characterCode:'2022',indent:18}});
  const parts=unzipSync(new Uint8Array(await ppt.write({outputType:'uint8array'}))),key='ppt/slides/slide1.xml';
  parts[key]=strToU8(strFromU8(parts[key]).replace('<a:buSzPct val="100000"/>','<a:buClrTx/><a:buSzPct val="100000"/><a:buFontTx/>'));
  const parsed=await readPptx(zipSync(parts)),original=exportHtml(parsed.model).html;
  const edited=original.replace(/(<span data-pptx-list-marker="" style="[^"]*?color:)#000000/,'$1#ff0000');assert.notEqual(edited,original);
  const output=await convertHtml(edited),props=child(paragraphs(output,'edited-color')[0],'pPr');
  assert.ok(!child(props,'buClrTx'));assert.equal(child(child(props,'buClr'),'srgbClr')?.getAttribute('val'),'FF0000');assert.ok(child(props,'buFontTx'));
});

test('HTML list paragraph gaps and character spacing remain editable in PPT', async () => {
  const html = `<section data-slide style="width:1280px;height:720px;font-family:Arial;font-size:24px">
    <ul style="list-style-type:square;line-height:32px"><li style="margin-bottom:12px;letter-spacing:2px">First item</li><li>Second item</li></ul>
  </section>`;
  const output = await convertHtml(html),doc=xmlDocument(output);
  const paragraphs=descendants(doc,'p').filter(p=>descendants(p,'t').some(t=>t.textContent==='First item'||t.textContent==='Second item'));
  assert.equal(paragraphs.length,2);
  const first=child(paragraphs[0],'pPr');
  assert.equal(child(first,'buChar')?.getAttribute('char'),'▪');
  assert.equal(child(child(first,'spcAft'),'spcPts')?.getAttribute('val'),'900');
  assert.equal(child(child(first,'lnSpc'),'spcPts')?.getAttribute('val'),'2400');
  assert.ok(descendants(paragraphs[0],'rPr').some(r=>r.getAttribute('spc')==='150'));
  const parsed=await readPptx(output,{allowPartial:true});assert.equal(parsed.status,'parsed',JSON.stringify(parsed.diagnostics));
  const roundtrip=exportHtml(parsed.model).html;
  assert.match(roundtrip,/data-pptx-list-marker/);
  assert.match(roundtrip,/letter-spacing:2px/);
});

test('an explicit HTML before marker becomes a native PPT bullet', async () => {
  const html = `<style>.items{list-style:none}.items li{position:relative;padding-left:24px}.items li::before{content:"•";position:absolute;left:0}</style>
    <section data-slide style="width:1280px;height:720px;font-family:Arial;font-size:24px"><ul class="items"><li>First</li><li>Second</li></ul></section>`;
  const output=await convertHtml(html),doc=xmlDocument(output);
  assert.equal(descendants(doc,'buChar').filter(n=>n.getAttribute('char')==='•').length,2);
});

test('a CSS marker character remains the actual PPT bullet', async () => {
  const html=`<style>li::marker{content:"★ ";color:#c00000}</style><section data-slide style="width:1280px;height:720px;font:28px Arial"><ul id="marker"><li>First</li><li>Second</li></ul></section>`;
  const output=await convertHtml(html),result=paragraphs(output,'marker');
  assert.deepEqual(result.map(p=>child(child(p,'pPr'),'buChar')?.getAttribute('char')),['★','★']);
});

test('HTML list conversion warns when PowerPoint may draw the marker differently', async () => {
  sharedRuntime??=await createBrowserRuntime({capture:false});
  for(const list of ['<ul><li>First</li></ul>','<table style="border-collapse:collapse"><tr><td><ul><li>First</li></ul></td></tr></table>']){
    const html=`<section data-slide style="width:1280px;height:720px;font:28px Arial">${list}</section>`;
    const measured=await measureInput(await readInput({html}),sharedRuntime),output=await exportPptx(measured);
    assert.equal(output.status,'generated',JSON.stringify(output.diagnostics));
    const warning=output.diagnostics.find(d=>d.reasonKey==='HTML_PPTX_BULLET_LAYOUT_APPROXIMATED');
    assert.ok(warning);assert.equal(warning.severity,'warning');assert.match(publicDiagnostic(warning).message,/글머리/);
  }
});

test('an explicitly hidden CSS marker does not reappear in PPT', async () => {
  const html=`<style>li::marker{content:none}</style><section data-slide style="width:1280px;height:720px;font:28px Arial"><ul id="hidden-marker"><li>First</li></ul></section>`;
  const output=await convertHtml(html),props=child(paragraphs(output,'hidden-marker')[0],'pPr');
  assert.ok(child(props,'buNone'));assert.ok(!child(props,'buChar'));
});

test('unsupported multi-character CSS markers are rejected rather than replaced silently', async () => {
  const html=`<style>li::marker{content:"AB "}</style><section data-slide style="width:1280px;height:720px;font:28px Arial"><ul><li>First</li></ul></section>`;
  sharedRuntime??=await createBrowserRuntime({capture:false});
  const measured=await measureInput(await readInput({html}),sharedRuntime),output=await exportPptx(measured);
  assert.equal(output.status,'failed');assert.ok(output.diagnostics.some(d=>d.code==='UNSUPPORTED_OBJECT'&&d.severity==='error'));
});

test('outside HTML bullets with no list padding keep the text position and hanging indent', async () => {
  const html=`<section data-slide style="position:relative;width:1280px;height:720px;font:32px Arial"><ul id="no-padding" style="position:absolute;left:200px;top:120px;margin:0;padding:0;list-style:disc"><li>First item</li><li>Second item</li></ul></section>`;
  const output=await convertHtml(html),doc=xmlDocument(output),shape=descendants(doc,'sp').find(n=>descendants(n,'cNvPr')[0]?.getAttribute('name')==='no-padding');
  const first=child(elements(child(shape,'txBody'),'p')[0],'pPr'),off=descendants(shape,'off')[0];
  const x=Number(off.getAttribute('x')),left=Number(first.getAttribute('marL')),indent=Number(first.getAttribute('indent'));
  assert.ok(indent<0,'Bullet must hang to the left of the text');
  assert.ok(Math.abs(x+left-200*9525)<5,'Text must stay at its original x position');
  assert.ok(x<200*9525,'Text box must leave room for an outside marker');
});

test('HTML marker and spacing edits override the original PPT formatting', async () => {
  const ppt=new PptxGenJS(),slide=ppt.addSlide();
  slide.addText('First\nSecond',{x:1,y:1,w:3,h:2,objectName:'edited',fontFace:'Arial',fontSize:20,margin:0,bullet:{characterCode:'2022',indent:18},paraSpaceAfter:6,charSpacing:-0.3});
  const parsed=await readPptx(new Uint8Array(await ppt.write({outputType:'uint8array'})),{allowPartial:true});
  let html=exportHtml(parsed.model).html;
  assert.match(html,/>• <\/span>/);assert.match(html,/padding-bottom:8px/);assert.match(html,/letter-spacing:-0.4px/);
  html=html.replace('>• </span>','>▪ </span>').replace('padding-bottom:8px','padding-bottom:12px').replaceAll('letter-spacing:-0.4px','letter-spacing:0.8px');
  const output=await convertHtml(html),first=paragraphs(output,'edited')[0],props=child(first,'pPr');
  assert.equal(child(props,'buChar')?.getAttribute('char'),'▪');
  assert.equal(child(child(props,'spcAft'),'spcPts')?.getAttribute('val'),'900');
  assert.ok(descendants(first,'rPr').some(r=>r.getAttribute('spc')==='60'));
});

test('nested HTML lists keep their bullet levels and paragraph order in PPT', async () => {
  const html=`<section data-slide style="width:1280px;height:720px;font-family:Arial;font-size:24px"><ul><li>Parent<ul><li>Child</li></ul></li><li>Next</li></ul></section>`;
  const output=await convertHtml(html),doc=xmlDocument(output);
  const list=descendants(doc,'p').filter(p=>descendants(p,'t').some(t=>['Parent','Child','Next'].includes(t.textContent)));
  assert.deepEqual(list.map(p=>descendants(p,'t').map(t=>t.textContent).join('')),['Parent','Child','Next']);
  assert.deepEqual(list.map(p=>child(p,'pPr')?.getAttribute('lvl')),['0','1','0']);
  assert.ok(list.every(p=>child(child(p,'pPr'),'buChar')));
});

test('HTML ordered lists keep their starting number as native PPT numbering', async () => {
  const html=`<section data-slide style="width:1280px;height:720px;font-family:Arial;font-size:24px"><ol start="3"><li>First</li><li>Second</li></ol></section>`;
  const output=await convertHtml(html),doc=xmlDocument(output),numbers=descendants(doc,'buAutoNum');
  assert.equal(numbers.length,2);assert.equal(numbers[0].getAttribute('type'),'arabicPeriod');
  assert.equal(numbers[0].getAttribute('startAt'),'3');assert.equal(numbers[1].getAttribute('startAt'),'3');
  const parsed=await readPptx(output,{allowPartial:true});assert.equal(parsed.status,'parsed',JSON.stringify(parsed.diagnostics));
  const htmlAgain=exportHtml(parsed.model).html;assert.match(htmlAgain,/>3\. <\/span>/);assert.match(htmlAgain,/>4\. <\/span>/);
});

test('HTML numbering can restart at a different value and then continue', async () => {
  const html=`<section data-slide style="width:1280px;height:720px;font:24px Arial"><ol id="restart" start="3"><li>First</li><li value="5">Restart</li><li>Next</li></ol></section>`;
  const output=await convertHtml(html),result=paragraphs(output,'restart');
  assert.deepEqual(result.map(p=>child(child(p,'pPr'),'buAutoNum')?.getAttribute('startAt')),['3','5','5']);
  const parsed=await readPptx(output);assert.equal(parsed.status,'parsed',JSON.stringify(parsed.diagnostics));
  assert.equal(parsed.model.slides[0].nodes.find(n=>n.id==='restart').content.text,'3. First\n5. Restart\n6. Next');
});

test('an unsupported restart at the same list start cannot silently change numbering', async () => {
  const html=`<section data-slide style="width:1280px;height:720px;font:24px Arial"><ol start="3"><li>First</li><li value="3">Restart</li></ol></section>`;
  sharedRuntime??=await createBrowserRuntime({capture:false});
  const measured=await measureInput(await readInput({html}),sharedRuntime),output=await exportPptx(measured);
  assert.equal(output.status,'failed');assert.ok(output.diagnostics.some(d=>d.severity==='error'));
});

test('nested ordered lists retain the parent counter after a child list', async () => {
  for(const child of ['<ol><li>Child</li></ol>','<ul><li>Child</li></ul>']){
    const html=`<section data-slide style="width:1280px;height:720px;font-family:Arial;font-size:24px"><ol><li>Parent one${child}</li><li>Parent two</li></ol></section>`;
    const output=await convertHtml(html),parsed=await readPptx(output);
    assert.equal(parsed.status,'parsed',JSON.stringify(parsed.diagnostics));
    const text=parsed.model.slides[0].nodes.find(n=>n.kind==='text').content.text;
    assert.match(text,/^1\. Parent one\n.*Child\n2\. Parent two$/);
  }
});

test('editing a numbered marker in generated HTML restarts native PPT numbering', async () => {
  const html='<section data-slide style="width:1280px;height:720px;font-family:Arial;font-size:24px"><ol start="2"><li>First</li><li>Second</li></ol></section>';
  const original=await convertHtml(html),parsed=await readPptx(original),generated=exportHtml(parsed.model).html;
  const edited=generated.replace('>2. </span>','>5. </span>');assert.notEqual(edited,generated);
  const output=await convertHtml(edited),reread=await readPptx(output);
  assert.equal(reread.status,'parsed',JSON.stringify(reread.diagnostics));
  assert.equal(reread.model.slides[0].nodes.find(n=>n.kind==='text').content.text,'5. First\n6. Second');
});

test('PPT automatic numbering remains native after HTML round trip', async () => {
  const ppt=new PptxGenJS(),slide=ppt.addSlide();
  slide.addText('First\nSecond',{x:1,y:1,w:4,h:2,objectName:'numbered',fontFace:'Arial',fontSize:20,margin:0});
  const parts=unzipSync(new Uint8Array(await ppt.write({outputType:'uint8array'}))),key='ppt/slides/slide1.xml',source=strFromU8(parts[key]);let count=0;
  const patched=source.replace(/<a:buNone\/>/g,()=>{count++;return '<a:buAutoNum type="arabicPeriod" startAt="2"/>';});
  assert.equal(count,2);parts[key]=strToU8(patched);
  const parsed=await readPptx(zipSync(parts));assert.equal(parsed.status,'parsed',JSON.stringify(parsed.diagnostics));
  const output=await convertHtml(exportHtml(parsed.model).html),result=paragraphs(output,'numbered');
  assert.deepEqual(result.map(p=>child(child(p,'pPr'),'buAutoNum')?.getAttribute('startAt')),['2','2']);
  const reread=await readPptx(output);assert.equal(reread.status,'parsed',JSON.stringify(reread.diagnostics));
  assert.equal(reread.model.slides[0].nodes.find(n=>n.id==='numbered').content.text,'2. First\n3. Second');
});

test('PPT percentage line and paragraph spacing retain their native units after HTML round trip', async () => {
  const ppt=new PptxGenJS(),slide=ppt.addSlide();
  slide.addText('First\nSecond',{x:1,y:1,w:4,h:2,objectName:'percentage',fontFace:'Arial',fontSize:20,margin:0});
  const parts=unzipSync(new Uint8Array(await ppt.write({outputType:'uint8array'}))),key='ppt/slides/slide1.xml',source=strFromU8(parts[key]);
  const patched=source.replace('<a:buNone/>','<a:lnSpc><a:spcPct val="150000"/></a:lnSpc><a:spcAft><a:spcPct val="50000"/></a:spcAft><a:buNone/>');
  assert.notEqual(source,patched);parts[key]=strToU8(patched);
  const parsed=await readPptx(zipSync(parts),{allowPartial:true});assert.equal(parsed.status,'parsed',JSON.stringify(parsed.diagnostics));
  const output=await convertHtml(exportHtml(parsed.model).html),first=child(paragraphs(output,'percentage')[0],'pPr');
  assert.equal(child(child(first,'lnSpc'),'spcPct')?.getAttribute('val'),'150000');
  assert.equal(child(child(first,'spcAft'),'spcPct')?.getAttribute('val'),'50000');
});

test('a single PPT paragraph retains hidden edge spacing and percentage line spacing', async () => {
  const ppt=new PptxGenJS(),slide=ppt.addSlide();
  slide.addText('Only paragraph',{x:1,y:1,w:4,h:2,objectName:'single-spacing',fontFace:'Arial',fontSize:20,margin:0});
  const parts=unzipSync(new Uint8Array(await ppt.write({outputType:'uint8array'}))),key='ppt/slides/slide1.xml',source=strFromU8(parts[key]);
  const patched=source.replace('<a:buNone/>','<a:lnSpc><a:spcPct val="150000"/></a:lnSpc><a:spcBef><a:spcPts val="300"/></a:spcBef><a:spcAft><a:spcPts val="600"/></a:spcAft><a:buNone/>');
  assert.notEqual(patched,source);parts[key]=strToU8(patched);
  const parsed=await readPptx(zipSync(parts));assert.equal(parsed.status,'parsed',JSON.stringify(parsed.diagnostics));
  const output=await convertHtml(exportHtml(parsed.model).html),props=child(paragraphs(output,'single-spacing')[0],'pPr');
  assert.equal(child(child(props,'lnSpc'),'spcPct')?.getAttribute('val'),'150000');
  assert.equal(child(child(props,'spcBef'),'spcPts')?.getAttribute('val'),'300');
  assert.equal(child(child(props,'spcAft'),'spcPts')?.getAttribute('val'),'600');
  parts[key]=strToU8(source.replace('<a:buNone/>','<a:lnSpc><a:spcPct val="150000"/></a:lnSpc><a:buNone/>'));
  const lineOnly=await readPptx(zipSync(parts));assert.equal(lineOnly.status,'parsed',JSON.stringify(lineOnly.diagnostics));
  const lineOutput=await convertHtml(exportHtml(lineOnly.model).html),lineProps=child(paragraphs(lineOutput,'single-spacing')[0],'pPr');
  assert.equal(child(child(lineProps,'lnSpc'),'spcPct')?.getAttribute('val'),'150000');
});

test('HTML edits replace the original percentage spacing', async () => {
  const ppt=new PptxGenJS(),slide=ppt.addSlide();
  slide.addText('First\nSecond',{x:1,y:1,w:4,h:2,objectName:'percentage-edit',fontFace:'Arial',fontSize:20,margin:0});
  const parts=unzipSync(new Uint8Array(await ppt.write({outputType:'uint8array'}))),key='ppt/slides/slide1.xml',source=strFromU8(parts[key]);
  parts[key]=strToU8(source.replace('<a:buNone/>','<a:spcAft><a:spcPct val="50000"/></a:spcAft><a:buNone/>'));
  const parsed=await readPptx(zipSync(parts),{allowPartial:true}),html=exportHtml(parsed.model).html;
  assert.match(html,/padding-bottom:16px/);
  const output=await convertHtml(html.replace('padding-bottom:16px','padding-bottom:24px'));
  const after=child(child(paragraphs(output,'percentage-edit')[0],'pPr'),'spcAft');
  assert.equal(child(after,'spcPts')?.getAttribute('val'),'1800');assert.equal(child(after,'spcPct'),undefined);
});

test('appending a paragraph does not expose a previously hidden PPT paragraph gap', async () => {
  const ppt=new PptxGenJS(),slide=ppt.addSlide();
  slide.addText('First\nSecond',{x:1,y:1,w:4,h:2,objectName:'spacing-edit',fontFace:'Arial',fontSize:20,margin:0,paraSpaceAfter:6});
  const parsed=await readPptx(new Uint8Array(await ppt.write({outputType:'uint8array'}))),original=exportHtml(parsed.model).html;
  const edited=original.replace('>Second</span></span></span></div>','>Second</span>\n</span><span data-pptx-paragraph style="display:block;line-height:32px;text-align:left;padding:0;text-indent:0">Third</span></span></div>');
  assert.notEqual(edited,original);
  const output=await convertHtml(edited),result=paragraphs(output,'spacing-edit');assert.equal(result.length,3);
  assert.equal(child(child(child(result[1],'pPr'),'spcAft'),'spcPts')?.getAttribute('val'),'0');
});

test('prepending a paragraph does not expose a previously hidden PPT paragraph gap', async () => {
  const ppt=new PptxGenJS(),slide=ppt.addSlide();
  slide.addText('First\nSecond',{x:1,y:1,w:4,h:2,objectName:'spacing-before',fontFace:'Arial',fontSize:20,margin:0,paraSpaceBefore:6});
  const parsed=await readPptx(new Uint8Array(await ppt.write({outputType:'uint8array'}))),original=exportHtml(parsed.model).html;
  const edited=original.replace('><span data-pptx-paragraph','><span data-pptx-paragraph style="display:block;line-height:32px;text-align:left;padding:0;text-indent:0">New\n</span><span data-pptx-paragraph');
  assert.notEqual(edited,original);
  const output=await convertHtml(edited),result=paragraphs(output,'spacing-before');assert.equal(result.length,3);
  assert.equal(child(child(child(result[1],'pPr'),'spcBef'),'spcPts')?.getAttribute('val'),'0');
});

test('removing an HTML marker removes the PPT bullet', async () => {
  const ppt=new PptxGenJS(),slide=ppt.addSlide();
  slide.addText('First',{x:1,y:1,w:3,h:1,objectName:'removed-marker',fontFace:'Arial',fontSize:20,bullet:{characterCode:'2022',indent:18}});
  const parsed=await readPptx(new Uint8Array(await ppt.write({outputType:'uint8array'})),{allowPartial:true}),html=exportHtml(parsed.model).html;
  const edited=html.replace(/<span data-pptx-list-marker=""[^>]*>• <\/span>/,'');assert.notEqual(edited,html);
  const output=await convertHtml(edited),first=child(paragraphs(output,'removed-marker')[0],'pPr');
  assert.equal(child(first,'buChar'),undefined);assert.ok(child(first,'buNone'));
});

test('invalid HTML numbering metadata cannot produce a broken PPT', async () => {
  const html='<section data-slide style="width:1280px;height:720px;font-family:Arial;font-size:24px"><ol start="3"><li>First</li></ol></section>';
  const ppt=await convertHtml(html),parsed=await readPptx(ppt,{allowPartial:true}),generated=exportHtml(parsed.model).html;
  const edited=generated.replace('data-pptx-bullet-start="3"','data-pptx-bullet-start="not-a-number"');assert.notEqual(edited,generated);
  const runtime=await createBrowserRuntime({capture:false});try{
    const measured=await measureInput(await readInput({html:edited}),runtime),output=await exportPptx(measured);
    assert.equal(output.status,'failed');assert.ok(output.diagnostics.some(d=>d.severity==='error'));
  }finally{await runtime.close();}
});

test('a manual line break stays within its PPT bullet paragraph', async () => {
  const ppt=new PptxGenJS(),slide=ppt.addSlide();
  slide.addText('First line\nSecond item',{x:1,y:1,w:4,h:3,objectName:'soft-break',fontFace:'Arial',fontSize:20,margin:0,bullet:{characterCode:'2022',indent:18}});
  const parts=unzipSync(new Uint8Array(await ppt.write({outputType:'uint8array'}))),key='ppt/slides/slide1.xml',source=strFromU8(parts[key]);
  const patched=source.replace('>First line</a:t>','>First line</a:t></a:r><a:br/><a:r><a:rPr sz="2000"><a:latin typeface="Arial"/></a:rPr><a:t>continued</a:t>');
  assert.notEqual(source,patched);parts[key]=strToU8(patched);
  const parsed=await readPptx(zipSync(parts),{allowPartial:true});assert.equal(parsed.status,'parsed',JSON.stringify(parsed.diagnostics));
  const output=await convertHtml(exportHtml(parsed.model).html),result=paragraphs(output,'soft-break');
  assert.equal(result.length,2);assert.equal(elements(result[0],'br').length,1);
  assert.equal(child(child(result[0],'pPr'),'buChar')?.getAttribute('char'),'•');
  assert.equal(descendants(result[0],'t').map(t=>t.textContent).join(''),'First linecontinued');
});

test('PPT first and last paragraph spacing remains enabled after HTML round trip', async () => {
  const ppt=new PptxGenJS(),slide=ppt.addSlide();
  slide.addText([{text:'First',options:{paraSpaceBefore:8,breakLine:true}},{text:'Last',options:{paraSpaceAfter:10}}],{x:1,y:1,w:4,h:2,objectName:'edge-spacing',fontFace:'Arial',fontSize:20,margin:0});
  const parts=unzipSync(new Uint8Array(await ppt.write({outputType:'uint8array'}))),key='ppt/slides/slide1.xml',source=strFromU8(parts[key]);
  const patched=source.replace('<a:bodyPr ','<a:bodyPr spcFirstLastPara="1" ');assert.notEqual(source,patched);parts[key]=strToU8(patched);
  const parsed=await readPptx(zipSync(parts),{allowPartial:true});assert.equal(parsed.status,'parsed',JSON.stringify(parsed.diagnostics));
  const output=await convertHtml(exportHtml(parsed.model).html),doc=xmlDocument(output),shape=descendants(doc,'sp').find(n=>descendants(n,'cNvPr')[0]?.getAttribute('name')==='edge-spacing'),body=child(shape,'txBody'),result=elements(body,'p');
  assert.equal(child(body,'bodyPr').getAttribute('spcFirstLastPara'),'1');
  assert.equal(child(child(child(result[0],'pPr'),'spcBef'),'spcPts')?.getAttribute('val'),'800');
  assert.equal(child(child(child(result[1],'pPr'),'spcAft'),'spcPts')?.getAttribute('val'),'1000');
});

test('HTML paragraphs keep their spacing and run-level character spacing in one PPT text box', async () => {
  const html=`<section data-slide style="width:1280px;height:720px;font-family:Arial;font-size:24px"><div id="paragraph-group" style="position:absolute;left:80px;top:80px;width:500px;line-height:32px"><p style="margin:0 0 12px;letter-spacing:2px">First paragraph</p><p style="margin:0">Second paragraph</p></div></section>`;
  const output=await convertHtml(html),result=paragraphs(output,'paragraph-group');
  assert.equal(result.length,2);assert.equal(child(child(child(result[0],'pPr'),'spcAft'),'spcPts')?.getAttribute('val'),'900');
  assert.ok(descendants(result[0],'rPr').some(r=>r.getAttribute('spc')==='150'));
});

test('HTML last-paragraph margin remains editable as PPT paragraph spacing', async () => {
  const examples=[
    '<ul id="last-margin" style="position:absolute;left:80px;top:80px;margin:0;padding-left:24px"><li>First</li><li style="margin-bottom:12px">Last</li></ul>',
    '<div id="last-margin" style="position:absolute;left:80px;top:80px;width:500px"><p style="margin:0">First</p><p style="margin:0 0 12px">Last</p></div>',
  ];
  for(const example of examples){
    const output=await convertHtml(`<section data-slide style="width:1280px;height:720px;font-family:Arial;font-size:24px;line-height:32px">${example}</section>`);
    const doc=xmlDocument(output),shape=descendants(doc,'sp').find(n=>descendants(n,'cNvPr')[0]?.getAttribute('name')==='last-margin');
    const body=child(shape,'txBody'),last=elements(body,'p').at(-1),props=child(last,'pPr');
    assert.equal(child(body,'bodyPr').getAttribute('spcFirstLastPara'),'1');
    assert.equal(child(child(props,'spcAft'),'spcPts')?.getAttribute('val'),'900');
  }
});
