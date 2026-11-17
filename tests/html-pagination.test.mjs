import test from 'node:test';
import assert from 'node:assert/strict';
import {planPagination,validateSourceCoverage,splitTableRows} from '../src/html-to-pptx/html-pagination.mjs';
import {validateTable} from '../src/html-to-pptx/tables.mjs';

const withPaint=input=>({mode:'paginate',...input,documentWidth:input.documentWidth??200,objects:input.objects.map(object=>({...object,...(object.kind==='text'&&!object.style?{style:{fill:'transparent',strokeWidth:0,backgroundImage:'none',opacity:1}}:{}),paintBoxCssPx:object.paintBoxCssPx??{...object.boxCssPx},...(object.lines?{lines:object.lines.map(line=>({...line,paintBoxCssPx:line.paintBoxCssPx??{...line.boxCssPx}}))}:{})}))});
const plan=input=>planPagination(withPaint(input));

test('long text splits at measured lines and keeps repeated UTF-16 source ranges and run styles',()=>{
  const fragments=['same🙂 ','same🙂 ','same🙂 ','same🙂'];
  let offset=0;
  const lines=fragments.map((text,i)=>{
    const start=offset;offset+=text.length;
    return {boxCssPx:{x:10,y:i*30,w:100,h:30},baseline:i*30+22,runs:[{sourceId:'text-node',start,end:offset,text,style:{bold:i%2===1}}]};
  });
  const input=withPaint({documentHeight:120,pageHeight:60,sourceObjectIds:['paragraph'],sourceLedger:[{sourceId:'text-node',text:fragments.join(''),ranges:[{start:0,end:offset}]}],objects:[{id:'paragraph',kind:'text',parentId:null,splitSafe:true,boxCssPx:{x:10,y:0,w:100,h:120},sourceRanges:[{sourceId:'text-node',start:0,end:offset}],lines}]});
  const before=structuredClone(input),result=planPagination(input);
  assert.deepEqual(input,before,'pagination must not mutate the original measurement');
  assert.deepEqual(result.diagnostics,[]);
  assert.deepEqual(result.pages.map(p=>[p.sourceStart,p.sourceEnd]),[[0,60],[60,120]]);
  assert.deepEqual(result.pages.map(p=>p.index),[0,1]);
  assert.ok(result.pages.every(p=>p.pageHeight===60),'every page retains the selected canvas height');
  const parts=result.pages.flatMap(page=>page.parts);
  assert.deepEqual(parts.map(p=>[p.objectId,p.lineStart,p.lineEnd]),[['paragraph',0,2],['paragraph',2,4]]);
  const runs=parts.flatMap(p=>p.runs);
  assert.deepEqual(runs,lines.flatMap(line=>line.runs),'repeated words and emoji must retain source ranges and run styles');
  assert.deepEqual(parts.flatMap(p=>p.lines),input.objects[0].lines,'line boxes and baselines remain in source coordinates until the output adapter translates them');
  assert.equal(validateSourceCoverage(input.objects[0].sourceRanges,runs).valid,true);
});

test('source coverage detects missing, duplicate and unexpected character ranges',()=>{
  const expected=[{sourceId:'a',start:0,end:4},{sourceId:'b',start:0,end:2}];
  assert.equal(validateSourceCoverage(expected,[{sourceId:'a',start:0,end:2},{sourceId:'a',start:2,end:4},{sourceId:'b',start:0,end:2}]).valid,true);
  for(const actual of [
    [{sourceId:'a',start:0,end:3},{sourceId:'b',start:0,end:2}],
    [{sourceId:'a',start:0,end:4},{sourceId:'a',start:2,end:4},{sourceId:'b',start:0,end:2}],
    [{sourceId:'a',start:0,end:5},{sourceId:'b',start:0,end:2}],
    [{sourceId:'a',start:0,end:4},{sourceId:'b',start:0,end:2},{sourceId:'extra',start:0,end:1}],
  ]){
    const result=validateSourceCoverage(expected,actual);assert.equal(result.valid,false,JSON.stringify(actual));assert.ok(result.errors.length);
  }
});

test('document pagination requires explicit paginate mode',()=>{
  const input=withPaint({documentHeight:100,pageHeight:100,sourceObjectIds:[],sourceLedger:[],objects:[]});
  for(const mode of [undefined,'single','auto',null]){
    const result=planPagination({...input,mode});assert.deepEqual(result.pages,[]);assert.ok(result.diagnostics.some(d=>d.severity==='error'),JSON.stringify({mode,result}));
  }
});

test('a long text box with a painted fill or border is not split without a decoration policy',()=>{
  const text='AB',lines=[{boxCssPx:{x:0,y:0,w:50,h:10},baseline:8,runs:[{sourceId:'node',start:0,end:1,text:'A',style:{}}]},{boxCssPx:{x:0,y:60,w:50,h:10},baseline:68,runs:[{sourceId:'node',start:1,end:2,text:'B',style:{}}]}];
  for(const style of [{fill:'#ff0000',strokeWidth:0,backgroundImage:'none',opacity:1},{fill:'transparent',strokeWidth:2,backgroundImage:'none',opacity:1}]){
    const result=plan({documentHeight:120,pageHeight:60,sourceObjectIds:['text'],sourceLedger:[{sourceId:'node',text,ranges:[{start:0,end:2}]}],objects:[{id:'text',kind:'text',parentId:null,splitSafe:true,style,boxCssPx:{x:0,y:0,w:50,h:120},sourceRanges:[{sourceId:'node',start:0,end:2}],lines}]});
    assert.deepEqual(result.pages,[]);assert.ok(result.diagnostics.some(d=>d.severity==='error'),JSON.stringify(result));
  }
});

const tableStyle={fontFamily:'Arial',fontSize:12,align:'left',valign:'top',padding:{top:1,right:1,bottom:1,left:1},borders:Object.fromEntries(['top','right','bottom','left'].map(side=>[side,{width:0,style:'none',color:'#000000'}]))};
const cell=(id,row,col,rowSpan=1,colSpan=1)=>({id,row,col,rowSpan,colSpan,text:id,runs:[{sourceId:`cell-${id}`,start:0,end:id.length,text:id,style:{bold:id==='merged'}}],style:structuredClone(tableStyle)});
const sampleTable=()=>({rows:4,columns:2,rowHeightsCssPx:[40,40,40,40],columnWidthsCssPx:[50,50],cells:[cell('merged',0,0,2),cell('r0',0,1),cell('r1',1,1),cell('r2',2,0,1,2),cell('r3',3,0,1,2)]});
const rowPaintBoxes=content=>{let y=0;return content.rowHeightsCssPx.map(h=>{const box={x:0,y,w:100,h};y+=h;return box;});};

test('table split respects a merged cell ending exactly at the boundary and rebases rows without losing cells',()=>{
  const table=sampleTable();assert.equal(validateTable(table).valid,true);
  const crossing=splitTableRows(table,0,1);assert.equal(crossing.content,null);assert.ok(crossing.diagnostics.some(d=>d.severity==='error'));
  const first=splitTableRows(table,0,2),second=splitTableRows(table,2,4);
  assert.deepEqual(first.diagnostics,[]);assert.deepEqual(second.diagnostics,[]);
  assert.equal(validateTable(first.content).valid,true,JSON.stringify(validateTable(first.content).errors));
  assert.equal(validateTable(second.content).valid,true,JSON.stringify(validateTable(second.content).errors));
  assert.deepEqual(first.content.rowHeightsCssPx,[40,40]);assert.deepEqual(second.content.rowHeightsCssPx,[40,40]);
  assert.deepEqual(second.content.cells.map(c=>[c.id,c.row,c.colSpan]),[['r2',0,2],['r3',1,2]]);
  assert.deepEqual([...first.content.cells,...second.content.cells].map(c=>c.id).sort(),table.cells.map(c=>c.id).sort());
  assert.deepEqual(first.content.cells.find(c=>c.id==='merged').style,table.cells[0].style);
});

test('table pagination chooses row boundaries without splitting merged cells or duplicating rows',()=>{
  const content=sampleTable(),objects=[{id:'table',kind:'table',parentId:null,splitSafe:true,boxCssPx:{x:0,y:0,w:100,h:160},rowPaintBoxes:rowPaintBoxes(content),content}];
  const sourceLedger=content.cells.map(c=>({sourceId:`cell-${c.id}`,text:c.text,ranges:[{start:0,end:c.text.length}]}));
  const result=plan({documentHeight:160,pageHeight:100,sourceObjectIds:['table'],sourceLedger,objects});
  assert.deepEqual(result.diagnostics,[]);assert.deepEqual(result.pages.map(p=>[p.sourceStart,p.sourceEnd]),[[0,80],[80,160]]);
  const parts=result.pages.flatMap(p=>p.parts);assert.deepEqual(parts.map(p=>[p.objectId,p.rowStart,p.rowEnd]),[['table',0,2],['table',2,4]]);
  const cells=parts.flatMap(p=>p.content.cells);assert.deepEqual(cells.map(c=>c.id).sort(),content.cells.map(c=>c.id).sort());
  for(const part of parts)assert.equal(validateTable(part.content).valid,true,JSON.stringify(validateTable(part.content).errors));
});

test('table split does not silently duplicate or drop a painted boundary with unknown ownership',()=>{
  const content=sampleTable();
  content.cells.find(c=>c.id==='r1').style.borders.bottom={width:1,style:'solid',color:'#000000'};
  content.cells.find(c=>c.id==='r2').style.borders.top={width:1,style:'solid',color:'#000000'};
  const objects=[{id:'table',kind:'table',parentId:null,splitSafe:true,boxCssPx:{x:0,y:0,w:100,h:160},rowPaintBoxes:rowPaintBoxes(content),content}];
  const sourceLedger=content.cells.map(c=>({sourceId:`cell-${c.id}`,text:c.text,ranges:[{start:0,end:c.text.length}]}));
  const result=plan({documentHeight:160,pageHeight:100,sourceObjectIds:['table'],sourceLedger,objects});
  assert.deepEqual(result.pages,[]);assert.ok(result.diagnostics.some(d=>d.severity==='error'),JSON.stringify(result));
});

test('a merged table cell taller than one page does not produce partial output',()=>{
  const content={rows:3,columns:2,rowHeightsCssPx:[40,40,40],columnWidthsCssPx:[50,50],cells:[cell('merged-three',0,0,3),cell('a',0,1),cell('b',1,1),cell('c',2,1)]};
  assert.equal(validateTable(content).valid,true);
  const sourceLedger=content.cells.map(c=>({sourceId:`cell-${c.id}`,text:c.text,ranges:[{start:0,end:c.text.length}]}));
  const objects=[{id:'table',kind:'table',parentId:null,splitSafe:true,boxCssPx:{x:0,y:0,w:100,h:120},rowPaintBoxes:rowPaintBoxes(content),content}];
  const result=plan({documentHeight:120,pageHeight:100,sourceObjectIds:['table'],sourceLedger,objects});
  assert.deepEqual(result.pages,[]);assert.ok(result.diagnostics.some(d=>d.severity==='error'),JSON.stringify(result));
});

test('table row paint must be measured and cannot be cut through text overflowing into the next row',()=>{
  const content=sampleTable(),base={id:'table',kind:'table',parentId:null,splitSafe:true,boxCssPx:{x:0,y:0,w:100,h:160},content};
  const rows=rowPaintBoxes(content);rows[1].h=45;
  const sourceLedger=content.cells.map(c=>({sourceId:`cell-${c.id}`,text:c.text,ranges:[{start:0,end:c.text.length}]}));
  for(const table of [base,{...base,rowPaintBoxes:rows}]){
    const result=plan({documentHeight:160,pageHeight:100,sourceObjectIds:['table'],sourceLedger,objects:[table]});
    assert.deepEqual(result.pages,[]);assert.ok(result.diagnostics.some(d=>d.severity==='error'),JSON.stringify(result));
  }
});

test('overlapping objects form an indivisible group and do not disappear at a page boundary',()=>{
  const objects=[
    {id:'picture',kind:'image',parentId:null,boxCssPx:{x:0,y:60,w:30,h:70}},
    {id:'label',kind:'shape',parentId:null,boxCssPx:{x:5,y:80,w:20,h:30}},
    {id:'overlay',kind:'shape',parentId:null,boxCssPx:{x:10,y:125,w:20,h:30}},
  ];
  const result=plan({documentHeight:180,pageHeight:100,sourceObjectIds:objects.map(o=>o.id),sourceLedger:[],objects});
  assert.deepEqual(result.diagnostics,[]);
  const assignment=result.pages.flatMap(p=>p.parts.map(part=>({page:p.index,id:part.objectId})));
  assert.deepEqual(assignment.map(a=>a.id).sort(),objects.map(o=>o.id).sort());
  assert.equal(new Set(assignment.map(a=>a.page)).size,1,JSON.stringify(assignment));
  assert.ok(result.pages[0].sourceEnd<=60,JSON.stringify(result.pages));
  assert.equal(result.pages.at(-1).sourceEnd,180);
  for(let i=1;i<result.pages.length;i++)assert.equal(result.pages[i-1].sourceEnd,result.pages[i].sourceStart);
});

test('an oversized indivisible image or overlapping group fails without partial pages',()=>{
  for(const objects of [
    [{id:'large',kind:'image',parentId:null,boxCssPx:{x:0,y:0,w:50,h:101}}],
    [{id:'one',kind:'shape',parentId:null,boxCssPx:{x:0,y:0,w:50,h:60}},{id:'two',kind:'shape',parentId:null,boxCssPx:{x:0,y:50,w:50,h:60}}],
  ]){
    const result=plan({documentHeight:150,pageHeight:100,sourceObjectIds:objects.map(o=>o.id),sourceLedger:[],objects});
    assert.deepEqual(result.pages,[]);assert.ok(result.diagnostics.some(d=>d.severity==='error'),JSON.stringify(result));
  }
});

test('zero-height lines use their painted stroke bounds when choosing a page boundary',()=>{
  const line={id:'rule',kind:'line',parentId:null,boxCssPx:{x:0,y:100,w:50,h:0},paintBoxCssPx:{x:0,y:99,w:50,h:2},style:{strokeWidth:2}};
  const result=plan({documentWidth:100,documentHeight:200,pageHeight:100,sourceObjectIds:['rule'],sourceLedger:[],objects:[line]});
  assert.deepEqual(result.diagnostics,[]);assert.ok(result.pages[0].sourceEnd<=99,JSON.stringify(result.pages));
  assert.deepEqual(result.pages.flatMap(page=>page.parts).map(part=>part.objectId),['rule']);
});

test('missing paint bounds and objects painted outside the document cannot be silently paginated',()=>{
  const object={id:'shape',kind:'shape',parentId:null,boxCssPx:{x:90,y:10,w:10,h:10}};
  const base={documentWidth:100,documentHeight:100,pageHeight:100,sourceObjectIds:['shape'],sourceLedger:[],objects:[object]};
  for(const input of [base,{...base,objects:[{...object,paintBoxCssPx:{x:89,y:9,w:12,h:12}}]}]){
    const result=planPagination(input);assert.deepEqual(result.pages,[]);assert.ok(result.diagnostics.some(d=>d.severity==='error'),JSON.stringify(result));
  }
});

test('many nonoverlapping columns with interleaved vertical bounds find a safe cut in bounded time',{timeout:5000},()=>{
  const objects=Array.from({length:600},(_,i)=>({id:`shape-${i}`,kind:'shape',parentId:null,boxCssPx:{x:i*2,y:1000-i,w:1,h:2}}));
  const result=plan({documentWidth:1200,documentHeight:1200,pageHeight:1000,sourceObjectIds:objects.map(o=>o.id),sourceLedger:[],objects});
  assert.deepEqual(result.diagnostics,[]);assert.ok(result.pages[0].sourceEnd<=401,JSON.stringify(result.pages));
  assert.equal(result.pages.flatMap(page=>page.parts).length,objects.length);
});

test('the final short source interval keeps the same output page height and is not stretched',()=>{
  const result=plan({documentHeight:230,pageHeight:100,sourceObjectIds:['last'],sourceLedger:[],objects:[{id:'last',kind:'shape',parentId:null,boxCssPx:{x:0,y:210,w:20,h:10}}]});
  assert.deepEqual(result.diagnostics,[]);assert.equal(result.pages.at(-1).sourceEnd,230);
  assert.ok(result.pages.every(page=>page.pageHeight===100),JSON.stringify(result.pages));
  assert.ok(result.pages.at(-1).sourceEnd-result.pages.at(-1).sourceStart<100);
  assert.deepEqual(result.pages.flatMap(p=>p.parts).map(p=>p.objectId),['last']);
});

test('invalid sizes, duplicate object ids and missing or cyclic parents fail without partial pages',()=>{
  const base={documentHeight:100,pageHeight:100,sourceObjectIds:[],sourceLedger:[],objects:[]};
  const shape=(id,parentId=null)=>({id,kind:'shape',parentId,boxCssPx:{x:0,y:0,w:10,h:10}});
  for(const input of [
    {...base,pageHeight:0},{...base,pageHeight:NaN},{...base,pageHeight:Infinity},
    {...base,sourceObjectIds:['same'],objects:[shape('same'),shape('same')]},
    {...base,sourceObjectIds:['one'],objects:[shape('one','missing')]},
    {...base,sourceObjectIds:['one','two'],objects:[shape('one','two'),shape('two','one')]},
    {...base,sourceObjectIds:['missing'],objects:[]},
    {...base,sourceObjectIds:[],objects:[shape('unexpected')]},
    {...base,sourceObjectIds:['outside'],objects:[{...shape('outside'),boxCssPx:{x:0,y:99,w:10,h:2}}]},
    {...base,documentHeight:20100},
  ]){
    const result=plan(input);assert.deepEqual(result.pages,[]);assert.ok(result.diagnostics.some(d=>d.severity==='error'),JSON.stringify(input));
  }
});

test('independent source ledger rejects missing, duplicated, fabricated and surrogate-splitting text',()=>{
  const source='A🙂B',style={bold:false};
  const run=(start,end,text=source.slice(start,end))=>({sourceId:'node',start,end,text,style});
  const line=(i,runs)=>({boxCssPx:{x:0,y:i*30,w:80,h:30},baseline:i*30+20,runs});
  const make=lines=>({documentHeight:60,pageHeight:30,sourceObjectIds:['text'],sourceLedger:[{sourceId:'node',text:source,ranges:[{start:0,end:source.length}]}],objects:[{id:'text',kind:'text',parentId:null,splitSafe:true,boxCssPx:{x:0,y:0,w:80,h:60},sourceRanges:[{sourceId:'node',start:0,end:source.length}],lines}]});
  for(const input of [
    make([line(0,[run(0,1)]),line(1,[run(3,4)])]),
    make([line(0,[run(0,3)]),line(1,[run(1,4)])]),
    make([line(0,[run(0,3,'fabricated')]),line(1,[run(3,4)])]),
    make([line(0,[run(0,2)]),line(1,[run(2,4)])]),
  ]){
    const result=plan(input);assert.deepEqual(result.pages,[]);assert.ok(result.diagnostics.some(d=>d.severity==='error'),JSON.stringify(result));
  }
});

test('a grapheme spanning two source nodes cannot be split between slides',()=>{
  const input={documentHeight:60,pageHeight:30,sourceObjectIds:['text'],sourceLedger:[{sourceId:'base',text:'e',ranges:[{start:0,end:1}]},{sourceId:'accent',text:'\u0301',ranges:[{start:0,end:1}]}],objects:[{id:'text',kind:'text',parentId:null,splitSafe:true,boxCssPx:{x:0,y:0,w:80,h:60},sourceRanges:[{sourceId:'base',start:0,end:1},{sourceId:'accent',start:0,end:1}],lines:[{boxCssPx:{x:0,y:0,w:80,h:30},baseline:20,runs:[{sourceId:'base',start:0,end:1,text:'e',style:{}}]},{boxCssPx:{x:0,y:30,w:80,h:30},baseline:50,runs:[{sourceId:'accent',start:0,end:1,text:'\u0301',style:{}}]}]}]};
  const result=plan(input);assert.deepEqual(result.pages,[]);assert.ok(result.diagnostics.some(d=>d.severity==='error'),JSON.stringify(result));
});
