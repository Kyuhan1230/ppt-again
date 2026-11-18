import test from 'node:test';
import assert from 'node:assert/strict';
import {adaptMappedPagination} from '../src/html-to-pptx/html-pagination-adapter.mjs';
import {validateSourceCoverage} from '../src/html-to-pptx/html-pagination.mjs';

function fixture(){
  const objectId='paragraph',normal={fontFamily:'Arial',fontSize:20,fontWeight:400},bold={...normal,fontWeight:700};
  const sourceLedger=[
    {sourceId:'a',objectId,kind:'text',text:'same  same ',ranges:[{start:0,end:11}],whiteSpace:'normal',style:normal},
    {sourceId:'br',objectId,kind:'br',text:'\n',ranges:[{start:0,end:1}],whiteSpace:'normal',style:normal},
    {sourceId:'b',objectId,kind:'text',text:'same',ranges:[{start:0,end:4}],whiteSpace:'normal',style:bold},
    {sourceId:'c',objectId,kind:'text',text:'ok',ranges:[{start:0,end:2}],whiteSpace:'normal',style:normal},
  ];
  const observations=[
    {sourceId:'a',objectId,start:0,end:4,lineId:'l1',kind:'visible',text:'same',style:normal},
    {sourceId:'a',objectId,start:4,end:5,lineId:'l1',kind:'visible',text:' ',style:normal},
    {sourceId:'a',objectId,start:5,end:6,lineId:'l1',kind:'collapsed',text:'',style:normal},
    {sourceId:'a',objectId,start:6,end:10,lineId:'l1',kind:'visible',text:'same',style:normal},
    {sourceId:'a',objectId,start:10,end:11,lineId:'l1',kind:'suppressed',text:'',style:normal},
    {sourceId:'br',objectId,start:0,end:1,lineId:'l2',kind:'hard-break',text:'\n',style:normal},
    {sourceId:'b',objectId,start:0,end:4,lineId:'l3',kind:'visible',text:'same',style:bold},
    {sourceId:'c',objectId,start:0,end:2,lineId:'l4',kind:'visible',text:'ok',style:normal},
  ];
  const token=(kind,lineId,text,style,sources=[],nextLineId)=>({kind,objectId,lineId,text,style,sources,...(nextLineId?{nextLineId}:{})});
  const tokens=[
    token('copy','l1','same',normal,[{sourceId:'a',start:0,end:4}]),
    token('collapsed-space','l1',' ',normal,[{sourceId:'a',start:4,end:5},{sourceId:'a',start:5,end:6}]),
    token('copy','l1','same',normal,[{sourceId:'a',start:6,end:10}]),
    token('suppressed-space','l1','',normal,[{sourceId:'a',start:10,end:11}]),
    token('layout-break','l1','\n',normal,[],'l2'),
    token('hard-break','l2','\n',normal,[{sourceId:'br',start:0,end:1}],'l3'),
    token('copy','l3','same',bold,[{sourceId:'b',start:0,end:4}]),
    token('layout-break','l3','\n',bold,[],'l4'),
    token('copy','l4','ok',normal,[{sourceId:'c',start:0,end:2}]),
  ];
  const lineIds=['l1','l2','l3','l4'],lineOwners=lineIds.map(lineId=>({lineId,objectId}));
  const lines=lineIds.map((lineId,i)=>({lineId,boxCssPx:{x:10,y:i*30,w:180,h:30},paintBoxCssPx:{x:10,y:i*30,w:180,h:30},baseline:i*30+22}));
  return {mode:'paginate',documentWidth:200,documentHeight:120,pageHeight:60,sourceObjectIds:[objectId],
    objects:[{id:objectId,kind:'text',parentId:null,splitSafe:true,style:{fill:'transparent',strokeWidth:0,backgroundImage:'none',opacity:1},content:{text:'same same\nsameok',runs:[{text:'same same\n',style:normal},{text:'same',style:bold},{text:'ok',style:normal}]},boxCssPx:{x:10,y:0,w:180,h:120},paintBoxCssPx:{x:10,y:0,w:180,h:120},lines}],
    displayMapping:{sourceObjectIds:[objectId],objectIds:[objectId],sourceBreakIds:['br'],lineIds,lineOwners,sourceLedger,observations,tokens},
    lineMeasurements:lines.map(line=>({...structuredClone(line),objectId}))};
}

test('display mapping paginates long text with source coverage, collapsed spaces, BR, styles and local baselines',()=>{
  const input=fixture(),before=structuredClone(input),result=adaptMappedPagination(input);
  assert.deepEqual(input,before);assert.deepEqual(result.diagnostics,[]);
  assert.equal(result.exportReady,false);
  assert.deepEqual(result.remainingChecks,['text-placement','model-validation','ppt-output']);
  assert.deepEqual(result.pages.map(p=>[p.index,p.sourceStart,p.sourceEnd,p.pageHeight]),[[0,0,60,60],[1,60,120,60]]);
  const parts=result.pages.flatMap(page=>page.parts);assert.equal(parts.length,2);
  assert.equal(new Set(parts.map(p=>p.fragmentId)).size,2);assert.ok(parts.every(p=>p.objectId==='paragraph'));
  assert.deepEqual(parts.map(p=>p.content.text),['same same\n','same\nok']);
  for(const part of parts){
    assert.equal(part.content.runs.map(run=>run.text).join(''),part.content.text);
    assert.deepEqual(part.lines.map(line=>line.boxCssPx.y),[0,30]);
    assert.deepEqual(part.lines.map(line=>line.baseline),[22,52]);
    assert.equal(part.baseline,22);
  }
  assert.ok(parts[1].content.runs.some(run=>run.text.includes('same')&&run.style.fontWeight===700));
  const expected=input.displayMapping.sourceLedger.flatMap(item=>item.ranges.map(r=>({sourceId:item.sourceId,...r})));
  assert.equal(validateSourceCoverage(expected,parts.flatMap(part=>part.sourceRanges)).valid,true);
  assert.deepEqual(parts.flatMap(part=>part.tokens).map(token=>token.kind),input.displayMapping.tokens.map(token=>token.kind));
  assert.equal(parts[0].tokens.find(token=>token.kind==='hard-break')?.handledBy,'page-boundary');
});

test('independent line measurement mismatch or missing mapping fails without partial pages',()=>{
  const moved=fixture();moved.lineMeasurements[1].baseline+=3;
  const missing=fixture();missing.lineMeasurements.pop();
  const lost=fixture();lost.displayMapping.tokens.splice(2,1);
  for(const input of [moved,missing,lost]){
    const result=adaptMappedPagination(input);assert.deepEqual(result.pages,[]);assert.ok(result.diagnostics.some(d=>d.severity==='error'),JSON.stringify(result));
  }
});

test('an automatic line break at the page boundary is recorded once without adding a blank line',()=>{
  const input=fixture();input.pageHeight=90;
  const result=adaptMappedPagination(input);assert.deepEqual(result.diagnostics,[]);
  const parts=result.pages.flatMap(page=>page.parts);assert.equal(parts.length,2);
  assert.deepEqual(parts.map(part=>part.content.text),['same same\n\nsame','ok']);
  assert.equal(parts.flatMap(part=>part.tokens).filter(token=>token.kind==='layout-break').length,2);
  assert.equal(parts[0].tokens.find(token=>token.kind==='layout-break'&&token.lineId==='l3')?.handledBy,'page-boundary');
  assert.equal(parts[1].content.text.startsWith('\n'),false);
  assert.deepEqual(parts[1].lines.map(line=>line.baseline),[22]);
});

test('continuation line offsets remain measurable and require later PPT placement verification',()=>{
  const input=fixture();input.documentHeight=135;
  input.objects[0].boxCssPx.h=135;input.objects[0].paintBoxCssPx.h=135;
  for(const [index,y] of [0,30,75,105].entries()){
    const line=input.objects[0].lines[index];line.boxCssPx.y=y;line.paintBoxCssPx.y=y;line.baseline=y+22;
    input.lineMeasurements[index]={...structuredClone(line),objectId:'paragraph'};
  }
  const result=adaptMappedPagination(input);assert.deepEqual(result.diagnostics,[]);assert.equal(result.exportReady,false);
  assert.ok(result.remainingChecks.includes('text-placement'));
  const continuation=result.pages[1].parts.find(part=>part.objectId==='paragraph');
  assert.equal(continuation.lines[0].boxCssPx.y,15);assert.equal(continuation.lines[0].baseline,37);
  assert.equal(continuation.placementEvidence.firstLineOffset,15);
  assert.equal(continuation.placementEvidence.requiresWriterVerification,true);
});

test('a BR inside one page remains in editable content rather than becoming a page break',()=>{
  const input=fixture();input.pageHeight=120;
  const result=adaptMappedPagination(input);assert.deepEqual(result.diagnostics,[]);
  const parts=result.pages.flatMap(page=>page.parts);assert.equal(parts.length,1);
  assert.equal(parts[0].content.text,'same same\n\nsame\nok');
  assert.notEqual(parts[0].tokens.find(token=>token.kind==='hard-break')?.handledBy,'page-boundary');
});

test('a preserved source newline is consumed once at a page boundary and remains a newline inside a page',()=>{
  for(const pageHeight of [60,120]){
    const input=fixture();input.pageHeight=pageHeight;
    const newline=input.displayMapping.sourceLedger.find(item=>item.sourceId==='br');newline.kind='text';newline.whiteSpace='pre-line';
    input.displayMapping.sourceBreakIds=[];
    input.displayMapping.observations.find(item=>item.sourceId==='br').kind='text-break';
    input.displayMapping.tokens.find(item=>item.kind==='hard-break').kind='text-break';
    const result=adaptMappedPagination(input);assert.deepEqual(result.diagnostics,[]);
    const parts=result.pages.flatMap(page=>page.parts),breakToken=parts.flatMap(part=>part.tokens).find(token=>token.kind==='text-break');
    assert.ok(breakToken);assert.equal(breakToken.handledBy==='page-boundary',pageHeight===60);
    assert.deepEqual(parts.map(part=>part.content.text),pageHeight===60?['same same\n','same\nok']:['same same\n\nsame\nok']);
    assert.ok(parts.flatMap(part=>part.sourceRanges).some(range=>range.sourceId==='br'&&range.start===0&&range.end===1));
  }
});

test('a page fragment containing only a source break cannot become an invalid empty text object',()=>{
  const input=fixture();input.pageHeight=30;
  const result=adaptMappedPagination(input);assert.deepEqual(result.pages,[]);
  assert.ok(result.diagnostics.some(d=>d.severity==='error'),JSON.stringify(result));
});

test('a page fragment with multiple BRs but no visible text is not an editable text object',()=>{
  const input=fixture(),objectId='paragraph',style=input.displayMapping.tokens[0].style;
  const lines=input.objects[0].lines.slice(0,3);input.documentHeight=90;input.objects[0].boxCssPx.h=90;input.objects[0].paintBoxCssPx.h=90;input.objects[0].lines=lines;
  input.objects[0].content={text:'\n\nA',runs:[{text:'\n\nA',style}]};input.lineMeasurements=lines.map(line=>({...structuredClone(line),objectId}));
  input.displayMapping={sourceObjectIds:[objectId],objectIds:[objectId],sourceBreakIds:['br1','br2'],lineIds:['l1','l2','l3'],lineOwners:['l1','l2','l3'].map(lineId=>({lineId,objectId})),
    sourceLedger:[{sourceId:'br1',objectId,kind:'br',text:'\n',ranges:[{start:0,end:1}],whiteSpace:'normal',style},{sourceId:'br2',objectId,kind:'br',text:'\n',ranges:[{start:0,end:1}],whiteSpace:'normal',style},{sourceId:'a',objectId,kind:'text',text:'A',ranges:[{start:0,end:1}],whiteSpace:'normal',style}],
    observations:[{sourceId:'br1',objectId,start:0,end:1,lineId:'l1',kind:'hard-break',text:'\n',style},{sourceId:'br2',objectId,start:0,end:1,lineId:'l2',kind:'hard-break',text:'\n',style},{sourceId:'a',objectId,start:0,end:1,lineId:'l3',kind:'visible',text:'A',style}],
    tokens:[{kind:'hard-break',objectId,lineId:'l1',nextLineId:'l2',text:'\n',style,sources:[{sourceId:'br1',start:0,end:1}]},{kind:'hard-break',objectId,lineId:'l2',nextLineId:'l3',text:'\n',style,sources:[{sourceId:'br2',start:0,end:1}]},{kind:'copy',objectId,lineId:'l3',text:'A',style,sources:[{sourceId:'a',start:0,end:1}]}]};
  const result=adaptMappedPagination(input);assert.deepEqual(result.pages,[]);assert.ok(result.diagnostics.some(d=>d.severity==='error'),JSON.stringify(result));
});

test('suppressed source space, object ownership and independent measurements cannot be forged',()=>{
  const omitted=fixture();omitted.displayMapping.tokens.splice(3,1);
  const forged=fixture();forged.displayMapping.observations[4].kind='collapsed';
  const owner=fixture();owner.lineMeasurements[0].objectId='other';
  const duplicate=fixture();duplicate.lineMeasurements.push(structuredClone(duplicate.lineMeasurements[0]));
  const wrongLine=fixture();wrongLine.displayMapping.tokens[6].lineId='l1';
  for(const input of [omitted,forged,owner,duplicate,wrongLine]){
    const result=adaptMappedPagination(input);assert.deepEqual(result.pages,[]);assert.ok(result.diagnostics.some(d=>d.severity==='error'),JSON.stringify(result));
  }
});

test('the existing editable text and run styles must match independently observed display tokens',()=>{
  const text=fixture();text.objects[0].content.text='tampered text';
  const style=fixture();style.objects[0].content.runs[1].style={...style.objects[0].content.runs[1].style,fontWeight:400};
  const runs=fixture();runs.objects[0].content.runs[0].text='different';
  for(const input of [text,style,runs]){const result=adaptMappedPagination(input);assert.deepEqual(result.pages,[]);assert.ok(result.diagnostics.some(d=>d.severity==='error'),JSON.stringify(result));}
});

test('oversized images and decorated long text are rejected without returning partial pages',()=>{
  const decorated=fixture();decorated.objects[0].style.fill='#ff0000';
  const image=fixture();image.documentHeight=200;image.sourceObjectIds.push('large');image.displayMapping.sourceObjectIds.push('large');image.displayMapping.objectIds.push('large');
  image.objects.push({id:'large',kind:'image',parentId:null,boxCssPx:{x:0,y:80,w:30,h:100},paintBoxCssPx:{x:0,y:80,w:30,h:100}});
  for(const input of [decorated,image]){const result=adaptMappedPagination(input);assert.deepEqual(result.pages,[]);assert.ok(result.diagnostics.some(d=>d.severity==='error'),JSON.stringify(result));}
});

test('adapter requires explicit pagination mode and a positive page height',()=>{
  for(const mutate of [input=>{delete input.mode;},input=>{input.mode='single';},input=>{input.pageHeight=0;},input=>{input.pageHeight=NaN;}]){
    const input=fixture();mutate(input);const result=adaptMappedPagination(input);assert.deepEqual(result.pages,[]);assert.ok(result.diagnostics.some(d=>d.severity==='error'));
  }
});
