import test from 'node:test';
import assert from 'node:assert/strict';
import {readInput} from '../src/html-to-pptx/input.mjs';
import {createBrowserRuntime} from '../src/html-to-pptx/browser-runtime.mjs';
import {observePaginationDocument} from '../src/html-to-pptx/html-pagination-observer.mjs';
import {validateDisplayMapping} from '../src/html-to-pptx/html-display-mapping.mjs';
import {adaptMappedPagination} from '../src/html-to-pptx/html-pagination-adapter.mjs';
import {validateSourceCoverage} from '../src/html-to-pptx/html-pagination.mjs';

const documentWith=body=>`<!doctype html><html><body style="margin:0">${body}</body></html>`;
const plain=(text='Hello',extra='')=>documentWith(`<main id="doc" style="width:600px;height:100px"><p style="margin:10px;font:20px/30px Arial;${extra}">${text}</p></main>`);
const unsupported=result=>{assert.equal(result.status,'unsupported',JSON.stringify(result));assert.ok(result.diagnostics.some(d=>d.code==='PAGINATION_UNSUPPORTED'&&d.severity==='error'),JSON.stringify(result));assert.ok(!result.objects&&!result.displayMapping,'unsupported observations must not expose a partial successful model');};

test('a supported paragraph is independently observed without mutating DOM, with real line bounds and baseline',async()=>{
  const viewport={width:800,height:600},runtime=await createBrowserRuntime({capture:false,checkLimits:true});
  try{
    const session=await runtime.render(await readInput({html:plain()}),{viewport});
    try{
      const before=await session.page.evaluate(()=>document.documentElement.outerHTML);
      const actual=await session.page.evaluate(()=>{
        const root=document.querySelector('#doc'),p=root.querySelector('p'),r=root.getBoundingClientRect(),range=document.createRange();range.selectNodeContents(p.firstChild);const box=range.getBoundingClientRect();
        const marker=document.createElement('span');marker.style.cssText='display:inline-block;width:0;height:0;vertical-align:baseline;padding:0;border:0;margin:0';p.append(marker);const baseline=marker.getBoundingClientRect().top;marker.remove();
        const canvas=document.createElement('canvas').getContext('2d');canvas.font='20px Arial';const metrics=canvas.measureText(p.textContent);
        return {line:{x:box.x-r.x,y:box.y-r.y,w:box.width,h:box.height},baseline:baseline-r.y,ascent:metrics.actualBoundingBoxAscent,descent:metrics.actualBoundingBoxDescent};
      });
      const result=await observePaginationDocument({session,selector:'#doc',viewport});
      assert.equal(result.status,'observed',JSON.stringify(result.diagnostics));assert.deepEqual(result.diagnostics,[]);
      assert.equal(result.selectionVerified,false);assert.equal(result.exportReady,false);assert.deepEqual(result.viewport,viewport);
      assert.match(result.fingerprint,/^[a-f0-9]{64}$/);assert.deepEqual(result.sourceLedger.map(s=>s.text),['Hello']);
      assert.equal(validateDisplayMapping(result.displayMapping).valid,true);
      assert.equal(result.objects.length,1);assert.equal(result.lineMeasurements.length,1);
      const line=result.lineMeasurements[0];assert.ok(Number.isFinite(line.baseline));
      for(const key of ['x','y','w','h'])assert.ok(Math.abs(line.boxCssPx[key]-actual.line[key])<.1,JSON.stringify({line,actual}));
      assert.ok(Math.abs(line.baseline-actual.baseline)<.1,JSON.stringify({line,actual}));
      assert.ok(Number.isFinite(actual.ascent)&&Number.isFinite(actual.descent));
      assert.equal(await session.page.evaluate(()=>document.documentElement.outerHTML),before);
    }finally{await session.context.close();}
  }finally{await runtime.close();}
});

test('responsive viewport changes are reflected in measured font size and the fingerprint',async()=>{
  const html=documentWith('<style>@media(max-width:700px){#doc p{font-size:14px!important;line-height:28px!important}}</style><main id="doc" style="width:600px;height:100px"><p style="margin:10px;font:20px/30px Arial">Hello</p></main>');
  const runtime=await createBrowserRuntime({capture:false,checkLimits:true});
  try{
    const results=[];
    for(const viewport of [{width:800,height:600},{width:600,height:600}]){
      const session=await runtime.render(await readInput({html}),{viewport});
      try{const result=await observePaginationDocument({session,selector:'#doc',viewport});assert.equal(result.status,'observed',JSON.stringify(result.diagnostics));results.push(result);}finally{await session.context.close();}
    }
    assert.notEqual(results[0].fingerprint,results[1].fingerprint);
    assert.deepEqual(results.map(r=>r.viewport.width),[800,600]);
    assert.ok(results[0].lineMeasurements[0].boxCssPx.h>results[1].lineMeasurements[0].boxCssPx.h);
  }finally{await runtime.close();}
});

test('Arial baseline matches an independent marker at different font and line heights',async()=>{
  const viewport={width:800,height:600},runtime=await createBrowserRuntime({capture:false,checkLimits:true});
  try{
    for(const [fontSize,lineHeight] of [[14,22],[32,40]]){
      const html=documentWith(`<main id="doc" style="width:600px;height:100px"><p style="margin:10px;font:${fontSize}px/${lineHeight}px Arial">Hello</p></main>`);
      const session=await runtime.render(await readInput({html}),{viewport});
      try{
        const expected=await session.page.evaluate(size=>{const root=document.querySelector('#doc'),p=root.querySelector('p'),rootBox=root.getBoundingClientRect(),range=document.createRange();range.selectNodeContents(p.firstChild);const line=range.getBoundingClientRect();const marker=document.createElement('span');marker.style.cssText='display:inline-block;width:0;height:0;vertical-align:baseline;padding:0;border:0;margin:0';p.append(marker);const baseline=marker.getBoundingClientRect().top-rootBox.y;marker.remove();const context=document.createElement('canvas').getContext('2d');context.font=`${size}px Arial`;const metric=context.measureText(p.textContent);return {baseline,lineY:line.y-rootBox.y,ascent:metric.fontBoundingBoxAscent,descent:metric.fontBoundingBoxDescent};},fontSize);
        const result=await observePaginationDocument({session,selector:'#doc',viewport});assert.equal(result.status,'observed',JSON.stringify(result.diagnostics));
        assert.ok(Number.isFinite(expected.ascent)&&Number.isFinite(expected.descent));
        assert.ok(Math.abs(result.lineMeasurements[0].baseline-expected.baseline)<.1,JSON.stringify({result,expected}));
        assert.ok(Math.abs(result.lineMeasurements[0].boxCssPx.y-expected.lineY)<.1);
      }finally{await session.context.close();}
    }
  }finally{await runtime.close();}
});

test('automatic wrapping preserves original spaces, measured line order and baseline without changing the DOM',async()=>{
  const viewport={width:800,height:600},html=documentWith('<main id="doc" style="width:600px;height:150px"><p style="margin:10px;width:75px;font:20px/30px Arial">Alpha Beta Gamma</p></main>');
  const runtime=await createBrowserRuntime({capture:false,checkLimits:true});
  try{
    const session=await runtime.render(await readInput({html}),{viewport});
    try{
      const before=await session.page.evaluate(()=>document.documentElement.outerHTML);
      const reference=await session.page.evaluate(()=>{const root=document.querySelector('#doc'),p=root.querySelector('p'),base=root.getBoundingClientRect(),node=p.firstChild,out=[];for(let i=0;i<node.length;i++){const range=document.createRange();range.setStart(node,i);range.setEnd(node,i+1);const r=range.getBoundingClientRect();out.push({char:node.data[i],x:r.x-base.x,y:r.y-base.y,w:r.width,h:r.height});}return out;});
      const visible=reference.filter(char=>char.char!==' '),lineYs=[...new Set(visible.map(char=>char.y))];assert.equal(lineYs.length,3);
      assert.deepEqual(reference.filter(char=>char.char===' ').map(char=>char.w),[0,0]);
      const result=await observePaginationDocument({session,selector:'#doc',viewport});
      assert.equal(result.status,'observed',JSON.stringify(result.diagnostics));assert.deepEqual(result.viewport,viewport);
      assert.equal(result.selectionVerified,false);assert.equal(result.exportReady,false);
      assert.equal(result.sourceLedger[0].text,'Alpha Beta Gamma');
      assert.equal(validateDisplayMapping(result.displayMapping).valid,true,JSON.stringify(result.displayMapping));
      assert.equal(result.lineMeasurements.length,3);
      for(let i=0;i<3;i++){
        const actual=result.lineMeasurements[i];assert.ok(Math.abs(actual.boxCssPx.y-lineYs[i])<.1);assert.ok(Number.isFinite(actual.baseline));
      }
      const observations=result.displayMapping.observations,sourceId=result.sourceLedger[0].sourceId;
      assert.deepEqual(observations.filter(o=>o.kind==='suppressed').map(o=>[o.sourceId,o.start,o.end]),[[sourceId,5,6],[sourceId,10,11]]);
      assert.equal(result.displayMapping.tokens.filter(t=>t.kind==='layout-break').length,2);
      assert.equal(result.displayMapping.tokens.map(t=>t.text).join(''),'Alpha\nBeta\nGamma');
      assert.equal(await session.page.evaluate(()=>document.documentElement.outerHTML),before);
    }finally{await session.context.close();}
  }finally{await runtime.close();}
});

test('actual wrapped text preserves semantic spaces from observation through page planning',async()=>{
  const viewport={width:800,height:600},html=documentWith('<main id="doc" style="width:600px;height:100px"><p style="margin:10px;width:75px;font:20px/30px Arial">Alpha Beta Gamma</p></main>');
  const runtime=await createBrowserRuntime({capture:false,checkLimits:true});
  try{
    const session=await runtime.render(await readInput({html}),{viewport});
    try{
      const observed=await observePaginationDocument({session,selector:'#doc',viewport});assert.equal(observed.status,'observed',JSON.stringify(observed.diagnostics));
      assert.equal(observed.objects[0].content.text,'Alpha Beta Gamma');
      const result=adaptMappedPagination({...observed,mode:'paginate',pageHeight:60});
      assert.deepEqual(result.diagnostics,[],JSON.stringify(result.diagnostics));assert.equal(result.exportReady,false);
      const parts=result.pages.flatMap(page=>page.parts);assert.deepEqual(parts.map(part=>part.content.text),['Alpha\nBeta','Gamma'],JSON.stringify({pages:result.pages.map(page=>({sourceStart:page.sourceStart,sourceEnd:page.sourceEnd,parts:page.parts.map(part=>({text:part.content?.text,lines:part.lines}))})),measured:observed.lineMeasurements}));
      const expected=observed.sourceLedger.flatMap(item=>item.ranges.map(range=>({sourceId:item.sourceId,...range})));
      assert.equal(validateSourceCoverage(expected,parts.flatMap(part=>part.sourceRanges)).valid,true);
      assert.equal(parts.flatMap(part=>part.tokens).filter(token=>token.kind==='suppressed-space').length,2);
      assert.equal(parts.flatMap(part=>part.tokens).filter(token=>token.kind==='layout-break').length,2);
      assert.ok(parts.every(part=>part.lines.every(line=>Number.isFinite(line.baseline))));
    }finally{await session.context.close();}
  }finally{await runtime.close();}
});

test('hyphen wraps and break-all wraps preserve every source character in order',async()=>{
  const viewport={width:800,height:600},runtime=await createBrowserRuntime({capture:false,checkLimits:true});
  try{
    for(const [text,extra] of [['Alpha-Beta-Gamma',''],['AlphaBetaGamma','word-break:break-all']]){
      const html=documentWith(`<main id="doc" style="width:600px;height:150px"><p style="margin:10px;width:75px;font:20px/30px Arial;${extra}">${text}</p></main>`);
      const session=await runtime.render(await readInput({html}),{viewport});
      try{
        const result=await observePaginationDocument({session,selector:'#doc',viewport});assert.equal(result.status,'observed',JSON.stringify(result.diagnostics));
        assert.equal(result.sourceLedger[0].text,text);assert.ok(result.lineMeasurements.length>=2);
        assert.equal(result.displayMapping.tokens.map(token=>token.text).join('').replaceAll('\n',''),text);
        assert.equal(result.displayMapping.tokens.filter(token=>token.kind==='layout-break').length,result.lineMeasurements.length-1);
        assert.equal(validateDisplayMapping(result.displayMapping).valid,true);
      }finally{await session.context.close();}
    }
  }finally{await runtime.close();}
});

test('same-style inline text nodes can wrap without losing either source ledger or page ranges',async()=>{
  const viewport={width:800,height:600},html=documentWith('<main id="doc" style="width:600px;height:100px"><p style="margin:10px;width:75px;font:20px/30px Arial">Alpha <span>Beta Gamma</span></p></main>');
  const runtime=await createBrowserRuntime({capture:false,checkLimits:true});
  try{
    const session=await runtime.render(await readInput({html}),{viewport});
    try{
      const before=await session.page.evaluate(()=>document.documentElement.outerHTML);
      const result=await observePaginationDocument({session,selector:'#doc',viewport});assert.equal(result.status,'observed',JSON.stringify(result.diagnostics));
      assert.deepEqual(result.sourceLedger.map(item=>item.text),['Alpha ','Beta Gamma']);assert.equal(new Set(result.sourceLedger.map(item=>item.sourceId)).size,2);
      assert.equal(validateDisplayMapping(result.displayMapping).valid,true);
      assert.equal(result.displayMapping.tokens.map(token=>token.text).join(''),'Alpha\nBeta\nGamma');
      const output=adaptMappedPagination({...result,mode:'paginate',pageHeight:60});assert.deepEqual(output.diagnostics,[],JSON.stringify(output.diagnostics));
      assert.deepEqual(output.pages.flatMap(page=>page.parts).map(part=>part.content.text),['Alpha\nBeta','Gamma']);
      const expected=result.sourceLedger.flatMap(item=>item.ranges.map(range=>({sourceId:item.sourceId,...range})));
      assert.equal(validateSourceCoverage(expected,output.pages.flatMap(page=>page.parts.flatMap(part=>part.sourceRanges))).valid,true);
      assert.equal(await session.page.evaluate(()=>document.documentElement.outerHTML),before);
    }finally{await session.context.close();}
  }finally{await runtime.close();}
});

test('a real BR retains its own source identity and is consumed once at a page boundary',async()=>{
  const viewport={width:800,height:600},html=documentWith('<main id="doc" style="width:600px;height:100px"><p style="margin:10px;width:180px;font:20px/30px Arial">Alpha<br>Beta</p></main>');
  const runtime=await createBrowserRuntime({capture:false,checkLimits:true});
  try{
    const session=await runtime.render(await readInput({html}),{viewport});
    try{
      const before=await session.page.evaluate(()=>document.documentElement.outerHTML);
      const y=await session.page.evaluate(()=>{const root=document.querySelector('#doc').getBoundingClientRect(),nodes=[...document.querySelector('#doc p').childNodes].filter(node=>node.nodeType===Node.TEXT_NODE);return nodes.map(node=>{const r=document.createRange();r.selectNodeContents(node);return r.getBoundingClientRect().y-root.y;});});
      const observed=await observePaginationDocument({session,selector:'#doc',viewport});assert.equal(observed.status,'observed',JSON.stringify(observed.diagnostics));
      assert.deepEqual(observed.sourceLedger.map(item=>[item.kind,item.text]),[['text','Alpha'],['br','\n'],['text','Beta']]);
      const br=observed.sourceLedger.find(item=>item.kind==='br');assert.deepEqual(observed.displayMapping.sourceBreakIds,[br.sourceId]);
      assert.equal(observed.displayMapping.tokens.filter(token=>token.kind==='hard-break').length,1);
      assert.equal(observed.displayMapping.tokens.map(token=>token.text).join(''),'Alpha\nBeta');assert.equal(validateDisplayMapping(observed.displayMapping).valid,true);
      assert.equal(observed.lineMeasurements.length,2);for(let i=0;i<2;i++){assert.ok(Math.abs(observed.lineMeasurements[i].boxCssPx.y-y[i])<.1);assert.ok(Number.isFinite(observed.lineMeasurements[i].baseline));}
      const paged=adaptMappedPagination({...observed,mode:'paginate',pageHeight:30});assert.deepEqual(paged.diagnostics,[],JSON.stringify(paged.diagnostics));
      const parts=paged.pages.flatMap(page=>page.parts);assert.deepEqual(parts.map(part=>part.content.text),['Alpha','Beta']);
      assert.equal(parts.flatMap(part=>part.tokens).find(token=>token.kind==='hard-break')?.handledBy,'page-boundary');
      const expected=observed.sourceLedger.flatMap(item=>item.ranges.map(range=>({sourceId:item.sourceId,...range})));
      assert.equal(validateSourceCoverage(expected,parts.flatMap(part=>part.sourceRanges)).valid,true);
      assert.equal(await session.page.evaluate(()=>document.documentElement.outerHTML),before);
    }finally{await session.context.close();}
  }finally{await runtime.close();}
});

test('unverified empty or hidden BR sequences are not silently converted',async()=>{
  const viewport={width:800,height:600},runtime=await createBrowserRuntime({capture:false,checkLimits:true});
  try{
    for(const content of ['Alpha<br><br>Beta','<br>Beta','Alpha<br>','Alpha<br style="display:none">Beta']){
      const session=await runtime.render(await readInput({html:plain(content)}),{viewport});
      try{unsupported(await observePaginationDocument({session,selector:'#doc',viewport}));}finally{await session.context.close();}
    }
  }finally{await runtime.close();}
});

test('a Korean paragraph uses only an independently confirmed matching system font',async()=>{
  const viewport={width:800,height:600},html=documentWith('<main id="doc" style="width:600px;height:100px"><p style="margin:10px;font:20px/30px \'Malgun Gothic\'">한글</p></main>');
  const runtime=await createBrowserRuntime({capture:false,checkLimits:true});
  try{
    const session=await runtime.render(await readInput({html}),{viewport});
    try{
      const cdp=await session.context.newCDPSession(session.page);await cdp.send('DOM.enable');await cdp.send('CSS.enable');
      const root=(await cdp.send('DOM.getDocument')).root.nodeId,nodeId=(await cdp.send('DOM.querySelector',{nodeId:root,selector:'#doc p'})).nodeId;
      const fonts=(await cdp.send('CSS.getPlatformFontsForNode',{nodeId})).fonts;await cdp.detach();
      const actual=await session.page.evaluate(()=>{const root=document.querySelector('#doc'),p=root.querySelector('p'),base=root.getBoundingClientRect(),range=document.createRange();range.selectNodeContents(p.firstChild);const box=range.getBoundingClientRect(),marker=document.createElement('span');marker.style.cssText='display:inline-block;width:0;height:0;vertical-align:baseline;padding:0;border:0;margin:0';p.append(marker);const baseline=marker.getBoundingClientRect().top-base.y;marker.remove();const context=document.createElement('canvas').getContext('2d');context.font=getComputedStyle(p).font;const metric=context.measureText(p.textContent);return {baseline,width:box.width,canvasWidth:metric.width,ascent:metric.fontBoundingBoxAscent,descent:metric.fontBoundingBoxDescent};});
      const result=await observePaginationDocument({session,selector:'#doc',viewport});
      if(fonts.length!==1||fonts[0].familyName.toLowerCase()!=='malgun gothic'){unsupported(result);return;}
      assert.equal(result.status,'observed',JSON.stringify(result.diagnostics));
      assert.equal(result.sourceLedger[0].text,'한글');assert.equal(result.displayMapping.tokens.map(token=>token.text).join(''),'한글');
      assert.equal(validateDisplayMapping(result.displayMapping).valid,true);
      assert.equal(result.fontEvidence.length,1);assert.equal(result.fontEvidence[0].objectId,result.objects[0].id);
      assert.deepEqual(result.fontEvidence[0].fonts,fonts);assert.match(result.fingerprint,/^[a-f0-9]{64}$/);
      assert.ok(Number.isFinite(result.lineMeasurements[0].baseline));
      assert.ok(Math.abs(result.lineMeasurements[0].baseline-actual.baseline)<.1,JSON.stringify(actual));
      assert.ok(Math.abs(actual.width-actual.canvasWidth)<.5&&Number.isFinite(actual.ascent)&&Number.isFinite(actual.descent),JSON.stringify(actual));
    }finally{await session.context.close();}
  }finally{await runtime.close();}
});

test('unproven font fallback, mixed fonts and color emoji are not claimed as observed',async()=>{
  const variants=[plain('한글'),plain('A한🙂'),plain('🙂'),documentWith('<main id="doc" style="width:600px;height:100px"><p style="margin:10px;font:20px/30px \'DefinitelyMissingFont\',Arial">Hello</p></main>')];
  const viewport={width:800,height:600},runtime=await createBrowserRuntime({capture:false,checkLimits:true});
  try{for(const html of variants){const session=await runtime.render(await readInput({html}),{viewport});try{unsupported(await observePaginationDocument({session,selector:'#doc',viewport}));}finally{await session.context.close();}}}finally{await runtime.close();}
});

test('a document beyond 9000px is observed in tiles without changing viewport and retains full text coverage',async()=>{
  const viewport={width:800,height:600},text='Alpha Beta Gamma '.repeat(102).trim();
  const html=documentWith(`<main id="doc" style="width:600px;height:9400px"><p style="margin:10px;width:75px;font:20px/30px Arial">${text}</p></main>`);
  const runtime=await createBrowserRuntime({capture:false,checkLimits:true});
  try{
    const session=await runtime.render(await readInput({html}),{viewport});
    try{
      const screen=()=>session.page.evaluate(()=>({width:innerWidth,height:innerHeight,x:scrollX,y:scrollY,html:document.documentElement.outerHTML}));
      const before=await screen(),started=Date.now(),observed=await observePaginationDocument({session,selector:'#doc',viewport});
      assert.equal(observed.status,'observed',JSON.stringify({diagnostics:observed.diagnostics,elapsedMs:Date.now()-started}));
      assert.deepEqual(await screen(),before);assert.deepEqual(observed.viewport,viewport);assert.equal(observed.documentHeight,9400);
      assert.ok(observed.tileCoverage.count>=9,JSON.stringify(observed.tileCoverage));assert.equal(observed.sourceLedger[0].text,text);
      assert.ok(observed.lineMeasurements.length>=300);assert.ok(observed.lineMeasurements.at(-1).boxCssPx.y>9000);
      assert.equal(validateDisplayMapping(observed.displayMapping).valid,true);
      const expected=observed.sourceLedger.flatMap(item=>item.ranges.map(range=>({sourceId:item.sourceId,...range})));
      const output=adaptMappedPagination({...observed,mode:'paginate',pageHeight:600});assert.deepEqual(output.diagnostics,[],JSON.stringify(output.diagnostics));
      assert.ok(output.pages.length>=15);assert.equal(validateSourceCoverage(expected,output.pages.flatMap(page=>page.parts.flatMap(part=>part.sourceRanges))).valid,true);
    }finally{await session.context.close();}
  }finally{await runtime.close();}
});

test('unobserved whitespace collapse, BR, nested styles, hidden text and internal scrolling cannot silently succeed',async()=>{
  const variants=[
    plain('Hello  world'),plain('Hello <b>world</b>'),
    plain('Hidden','visibility:hidden'),plain('Clipped','height:10px;overflow:hidden'),
    plain('Scroll','height:10px;overflow:auto'),plain('Pretendard','font-family:Pretendard'),
    plain('Contained','contain:paint;height:10px'),plain('Strictly contained','contain:strict;height:10px'),
    plain('Fixed','position:fixed'),plain('Sticky','position:sticky'),plain('Auto visibility','content-visibility:auto'),
    plain('Transparent','color:transparent'),plain('Transparent fill','-webkit-text-fill-color:transparent'),plain('Transparent Color4','color:color(srgb 1 0 0 / 0)'),
    plain('Outlined','-webkit-text-stroke:1px red'),plain('Outer outline','outline:3px solid red'),plain('Hidden alpha','opacity:0'),
    plain('Emphasis','text-emphasis:filled red'),plain('Masked','-webkit-text-security:disc'),
    plain('Mirrored','scale:-1 1'),plain('Rotated','rotate:10deg'),plain('Moved','translate:5px 0'),
    documentWith('<style>#doc p::first-letter{font-size:30px}</style><main id="doc" style="width:600px;height:100px"><p style="margin:10px;font:20px/30px Arial">Hello</p></main>'),
    documentWith('<style>#doc p::first-letter{text-shadow:2px 2px red}</style><main id="doc" style="width:600px;height:100px"><p style="margin:10px;font:20px/30px Arial">Hello</p></main>'),
    documentWith('<style>#doc::first-letter{text-shadow:2px 2px red}</style><main id="doc" style="width:600px;height:100px"><p style="margin:10px;font:20px/30px Arial">Hello</p></main>'),
    documentWith('<style>#doc::first-line{color:red}</style><main id="doc" style="width:600px;height:100px"><p style="margin:10px;font:20px/30px Arial">Hello</p></main>'),
    documentWith('<style>#doc p::before{content:"";position:absolute;left:150px;width:10px;height:10px;background:red}</style><main id="doc" style="width:600px;height:100px"><p style="position:relative;margin:10px;font:20px/30px Arial">Hello</p></main>'),
    documentWith('<main id="doc" style="width:600px;height:100px;border:2px solid red"><p style="margin:10px;font:20px/30px Arial">Hello</p></main>'),
    documentWith('<main id="doc" style="width:600px;height:100px;outline:2px solid red"><p style="margin:10px;font:20px/30px Arial">Hello</p></main>'),
    documentWith('<main id="doc" style="width:600px;height:100px;text-decoration:underline red"><p style="margin:10px;font:20px/30px Arial">Hello</p></main>'),
    documentWith('<div style="text-decoration:underline red"><main id="doc" style="width:600px;height:100px"><p style="margin:10px;font:20px/30px Arial">Hello</p></main></div>'),
    documentWith('<main id="doc" style="width:600px;height:100px"><ul><li>Visible marker</li></ul></main>'),
    documentWith('<main id="doc" style="width:600px;height:100px"><p style="margin:10px;font:20px/30px Arial">Hello<img loading="lazy" src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+j1ioAAAAASUVORK5CYII="></p></main>'),
  ];
  const viewport={width:800,height:600},runtime=await createBrowserRuntime({capture:false,checkLimits:true});
  try{for(const html of variants){const session=await runtime.render(await readInput({html}),{viewport});try{unsupported(await observePaginationDocument({session,selector:'#doc',viewport}));}finally{await session.context.close();}}}finally{await runtime.close();}
});

test('changed viewport, blocked resources or source scripts reject observation before trusting a selection',async()=>{
  const viewport={width:800,height:600},runtime=await createBrowserRuntime({capture:false,checkLimits:true});
  try{
    const session=await runtime.render(await readInput({html:plain()}),{viewport});
    try{
      unsupported(await observePaginationDocument({session,selector:'#doc',viewport:{width:700,height:600}}));
      unsupported(await observePaginationDocument({session:{...session,sourceHadScripts:true},selector:'#doc',viewport}));
      unsupported(await observePaginationDocument({session:{...session,blocked:['https://blocked.invalid/a']},selector:'#doc',viewport}));
    }finally{await session.context.close();}
  }finally{await runtime.close();}
});

test('visible sibling content and a pointer-events:none overlay prevent section capture',async()=>{
  const html=documentWith('<main id="doc" style="width:600px;height:100px"><p style="margin:10px;font:20px/30px Arial">Hello</p></main><div id="cover" style="position:absolute;left:0;top:0;width:600px;height:100px;background:red;z-index:2;pointer-events:none"></div>');
  const viewport={width:800,height:600},runtime=await createBrowserRuntime({capture:false,checkLimits:true});
  try{
    const session=await runtime.render(await readInput({html}),{viewport});
    try{
      const covering=await session.page.evaluate(()=>{const p=document.querySelector('#doc p'),range=document.createRange();range.selectNodeContents(p.firstChild);const r=range.getBoundingClientRect();return document.elementFromPoint(r.x+r.width/2,r.y+r.height/2)?.id;});
      assert.notEqual(covering,'cover','hit testing cannot see a non-interactive overlay');
      const result=await observePaginationDocument({session,selector:'#doc',viewport});
      unsupported(result);
      assert.match(result.diagnostics[0].message,/바깥 콘텐츠/);
    }finally{await session.context.close();}
  }finally{await runtime.close();}
});

test('body pagination preserves verified inline colors and hard breaks in editable runs',async()=>{
  const html=documentWith('<p style="margin:10px;font:20px/30px Arial">Alpha <span style="color:red">Beta</span><br><span style="color:blue">Gamma</span></p>');
  const viewport={width:800,height:600},runtime=await createBrowserRuntime({capture:false,checkLimits:true});
  try{
    const session=await runtime.render(await readInput({html}),{viewport});
    try{
      const result=await observePaginationDocument({session,selector:'body',viewport});
      assert.equal(result.status,'observed',JSON.stringify(result.diagnostics));
      assert.equal(result.coverage?.outsideSelection,'verified-no-visible-content');
      assert.deepEqual(result.sourceLedger.map(source=>[source.text,source.style.color]),[
        ['Alpha ','rgb(0, 0, 0)'],['Beta','rgb(255, 0, 0)'],['\n','rgb(0, 0, 0)'],['Gamma','rgb(0, 0, 255)'],
      ]);
      const planned=adaptMappedPagination({...result,mode:'paginate',pageHeight:100});
      assert.deepEqual(planned.diagnostics,[],JSON.stringify(planned.diagnostics));
      const fragment=planned.pages[0].parts[0];
      assert.equal(fragment.content.text,'Alpha Beta\nGamma');
      assert.equal(fragment.content.runs.filter(run=>run.style.color==='rgb(255, 0, 0)').map(run=>run.text).join(''),'Beta');
      assert.equal(fragment.content.runs.filter(run=>run.style.color==='rgb(0, 0, 255)').map(run=>run.text).join(''),'Gamma');
    }finally{await session.context.close();}
  }finally{await runtime.close();}
});

test('a table row ledger and verified image retain independent source ownership',async()=>{
  const image='data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADElEQVR4nGOQTDkBAAHfAUZ76qvRAAAAAElFTkSuQmCC';
  const html=documentWith(`<style>tr{height:48px}td{padding:0;vertical-align:middle}</style>
    <main id="doc" style="width:600px;padding:12px;box-sizing:border-box">
      <h2 style="margin:0 0 8px;font:700 20px/30px Arial">Report</h2>
      <table style="width:240px;border-collapse:collapse;table-layout:fixed;font:16px/20px Arial">
        <tr><td>First</td><td>A</td></tr><tr><td rowspan="2">Merged</td><td>B</td></tr>
        <tr><td>C</td></tr><tr><td>Last</td><td>D</td></tr>
      </table>
      <img src="${image}" alt="" style="display:block;width:64px;height:64px;margin-top:8px">
    </main>`);
  const viewport={width:800,height:600},runtime=await createBrowserRuntime({capture:false,checkLimits:true});
  try{
    const input=await readInput({html}),session=await runtime.render(input,{viewport});
    try{
      const result=await observePaginationDocument({session,selector:'#doc',viewport,assets:input.assets,assetRefs:input.assetRefs});
      assert.equal(result.status,'observed',JSON.stringify(result.diagnostics));
      assert.deepEqual(result.objects.map(object=>object.kind),['text','table','image']);
      assert.equal(result.objects[0].content.runs[0].style.fontWeight,700);
      assert.equal(result.objects[2].content.assetId,input.assets.find(asset=>asset.mime==='image/png').id);
      assert.deepEqual(result.tableSourceLedger.map(cell=>cell.text),['First','A','Merged','B','C','Last','D']);
      assert.equal(new Set(result.tableSourceLedger.map(cell=>cell.sourceId)).size,result.tableSourceLedger.length);
      const plan=adaptMappedPagination({...result,mode:'paginate',pageHeight:130});
      assert.deepEqual(plan.diagnostics,[],JSON.stringify(plan.diagnostics));
      assert.equal(plan.pages[0].sourceEnd,result.objects[1].rowPaintBoxes[1].y);
      const tables=plan.pages.flatMap(page=>page.parts).filter(part=>part.kind==='table');
      assert.equal(tables.filter(part=>part.content.cells.some(cell=>cell.text==='Merged'&&cell.rowSpan===2)).length,1);
      assert.equal(plan.pages.flatMap(page=>page.parts).filter(part=>part.kind==='image').length,1);
    }finally{await session.context.close();}
  }finally{await runtime.close();}
});

test('cell backgrounds that are clipped or hidden do not become full PPT fills',async()=>{
  const viewport={width:800,height:600},runtime=await createBrowserRuntime({capture:false,checkLimits:true});
  try{
    for(const [tableStyle,cellStyle,text] of [
      ['border-collapse:separate;border-spacing:0;empty-cells:hide','background:red',''],
      ['border-collapse:collapse','background:red;background-clip:text','Cell text']
    ]){
      const html=documentWith(`<main id="doc" style="width:600px;padding:12px;box-sizing:border-box">
        <table style="width:240px;font:16px/20px Arial;${tableStyle}">
          <tr style="height:48px"><td style="padding:4px;${cellStyle}">${text}</td></tr>
        </table>
      </main>`);
      const input=await readInput({html}),session=await runtime.render(input,{viewport});
      try{
        const result=await observePaginationDocument({session,selector:'#doc',viewport});
        unsupported(result);
        assert.match(result.diagnostics[0].message,/셀 장식/);
      }finally{await session.context.close();}
    }
  }finally{await runtime.close();}
});

test('verified same-layer sibling overlap stays on one page in DOM paint order',async()=>{
  const image='data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADElEQVR4nGOQTDkBAAHfAUZ76qvRAAAAAElFTkSuQmCC';
  const html=documentWith(`<main id="doc" style="width:600px;padding:12px;box-sizing:border-box">
    <p style="height:60px;margin:0;position:relative;font:20px/30px Arial">Overlaid source</p>
    <img src="${image}" alt="" style="display:block;position:relative;top:-58px;width:64px;height:64px">
  </main>`);
  const viewport={width:800,height:600},runtime=await createBrowserRuntime({capture:false,checkLimits:true});
  try{
    const input=await readInput({html}),session=await runtime.render(input,{viewport});
    try{
      const covering=await session.page.evaluate(()=>{
        const p=document.querySelector('p'),range=document.createRange();
        range.selectNodeContents(p);
        const line=range.getBoundingClientRect();
        return document.elementFromPoint(line.x+3,line.y+3)?.tagName;
      });
      assert.equal(covering,'IMG');
      const result=await observePaginationDocument({session,selector:'#doc',viewport,assets:input.assets,assetRefs:input.assetRefs});
      assert.equal(result.status,'observed',JSON.stringify(result.diagnostics));
      const plan=adaptMappedPagination({...result,mode:'paginate',pageHeight:100});
      assert.deepEqual(plan.diagnostics,[],JSON.stringify(plan.diagnostics));
      assert.deepEqual(plan.pages[0].parts.map(part=>part.kind),['text','image']);
      assert.equal(plan.pages.slice(1).flatMap(page=>page.parts).length,0);
    }finally{await session.context.close();}
  }finally{await runtime.close();}
});

test('DOM changes during observation invalidate the fingerprint and return no partial output',async()=>{
  const viewport={width:800,height:600},runtime=await createBrowserRuntime({capture:false,checkLimits:true});
  try{
    const session=await runtime.render(await readInput({html:plain()}),{viewport});
    try{
      const screenshot=session.page.screenshot.bind(session.page);let count=0;
      session.page.screenshot=async options=>{if(++count===2)await session.page.evaluate(()=>{document.querySelector('#doc p').firstChild.data='Jello';});return screenshot(options);};
      unsupported(await observePaginationDocument({session,selector:'#doc',viewport}));
    }finally{await session.context.close();}
  }finally{await runtime.close();}
});
