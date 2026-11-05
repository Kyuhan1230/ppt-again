import test from 'node:test';
import assert from 'node:assert/strict';
import {projectDiagnostics,slideRanges,renderDiagnostics,validBounds} from '../web/html-to-pptx/diagnostic-view.js';
const make=(i,key='PPTX_TEXT_STYLE_UNSUPPORTED',severity='warning')=>({code:'PPTX_OBJECT_PLACEHOLDER',severity,reasonKey:key,reason:'Synthetic issue description',message:'Synthetic fallback',slideNumber:i%29+1,objectName:`Synthetic object ${i%10}`,objectType:'text'});
test('640 synthetic warnings are grouped by reason and generic duplicate suppression is presentation only',()=>{
 const raw=Array.from({length:320},(_,i)=>({...make(i),objectIndex:i+1})).flatMap(d=>[d,{...d,reasonKey:'PPTX_OBJECT_UNSUPPORTED',reason:'Generic detail'}]);
 const p=projectDiagnostics(raw);assert.equal(p.total,640);assert.equal(p.warningCount,640);assert.equal(p.suppressedGeneric,320);assert.equal(p.categories.length,1);assert.equal(p.categories[0].rows.length,320);assert.equal(p.affectedSlides.length,29);assert.equal(p.raw.length,640);assert.equal(raw[1].reason,'Generic detail');
});
test('master common issue groups once across 29 slides and compact ranges are stable',()=>{
 const p=projectDiagnostics(Array.from({length:29},(_,i)=>({code:'UNSUPPORTED_OBJECT',severity:'error',reasonKey:'PPTX_MASTER_ARTWORK_UNSUPPORTED',reason:'Master artwork',slideNumber:i+1,objectType:'background'})));
 assert.equal(p.categories.length,1);assert.equal(p.categories[0].rows.length,1);assert.equal(slideRanges(p.categories[0].slides),'1–29');assert.equal(slideRanges([4,2,2,1,7]),'1–2, 4, 7');
});
test('generic diagnostics with different slide/object/severity are not discarded',()=>{
 const d=make(0);const raw=[d,{...d,reasonKey:'PPTX_OBJECT_UNSUPPORTED',slideNumber:2},{...d,reasonKey:'PPTX_OBJECT_UNSUPPORTED',objectName:'Different'},{...d,reasonKey:'PPTX_OBJECT_UNSUPPORTED',severity:'error'}];const p=projectDiagnostics(raw);assert.equal(p.suppressedGeneric,0);assert.equal(p.errorCount,1);assert.equal(p.categories[0].severity,'error');
});
test('repeated warnings group by slide and object identity without merging same-named objects',()=>{
 const base={...make(0),slideNumber:1,objectName:'Same name',objectIndex:5};
 const projection=projectDiagnostics([base,{...base},{...base,objectIndex:6}]);
 assert.equal(projection.raw.length,3);assert.equal(projection.categories[0].count,2);assert.equal(projection.categories[0].rows.length,2);
 assert.deepEqual(projection.categories[0].rows.map(row=>row.objectIndex),[5,6]);
 const container=new Node('ul'),summary=new Node('p');
 renderDiagnostics({document:{createElement:t=>new Node(t)},container,summary,diagnostics:[base,{...base},{...base,objectIndex:6}]});
 assert.match(summary.textContent,/동일 객체 중복 1건 묶음/);
 const details=descendants(container).find(n=>n.tag==='details');details.open=true;details.events.toggle();
 assert.ok(descendants(container).some(n=>n.textContent==='분석 객체 5'));
 assert.ok(descendants(container).some(n=>n.textContent==='분석 객체 6'));
});
test('same-named warnings without object identity are not assumed to be the same object',()=>{
 const base={...make(0),slideNumber:1,objectName:'Same name'};
 const generic={...base,reasonKey:'PPTX_OBJECT_UNSUPPORTED',reason:'Generic cause'};
 const projection=projectDiagnostics([base,{...base},{...base,slideNumber:2},generic]);
 const specific=projection.categories.find(c=>c.reasonKey==='PPTX_TEXT_STYLE_UNSUPPORTED');
 assert.equal(specific.count,3);assert.equal(specific.rows.length,3);assert.equal(projection.suppressedGeneric,0);
 assert.ok(projection.categories.some(c=>c.reasonKey==='PPTX_OBJECT_UNSUPPORTED'));
});
test('objects with different exact causes in one category show each cause',()=>{
 const one={...make(0),reasonKey:'PPTX_CHART_UNSUPPORTED',objectName:'Chart A',objectType:'chart',objectIndex:1,reason:'First exact cause'};
 const two={...one,objectName:'Chart B',objectIndex:2,reason:'Second exact cause'};
 const container=new Node('ul'),summary=new Node('p');
 renderDiagnostics({document:{createElement:t=>new Node(t)},container,summary,diagnostics:[one,two]});
 const details=descendants(container).find(n=>n.tag==='details');details.open=true;details.events.toggle();
 assert.ok(descendants(container).some(n=>n.textContent==='First exact cause'));
 assert.ok(descendants(container).some(n=>n.textContent==='Second exact cause'));
});
class Node{constructor(tag){this.tag=tag;this.children=[];this.events={};this.open=false;this.textContent='';}append(...items){this.children.push(...items);}replaceChildren(...items){this.children=items;}addEventListener(type,fn){this.events[type]=fn;}}
const descendants=n=>[n,...n.children.flatMap(descendants)];
test('repeated warnings show one collapsed category with affected locations, without a repeated raw list',()=>{
 let count=0;const doc={createElement:t=>{count++;return new Node(t);}},container=new Node('ul'),summary=new Node('p');
 const raw=Array.from({length:640},(_,i)=>({...make(i),objectIndex:i+1}));
 const projection=renderDiagnostics({document:doc,container,summary,diagnostics:raw});
 assert.ok(count<25,`initial nodes ${count}`);const details=descendants(container).filter(n=>n.tag==='details');
 assert.equal(details.length,1);assert.ok(!details[0].open);assert.match(summary.textContent,/640/);
 assert.equal(projection.raw.length,640);
 details[0].open=true;details[0].events.toggle();
 assert.equal(descendants(container).filter(n=>n.textContent==='Synthetic issue description').length,1);
 assert.equal(descendants(container).filter(n=>n.className==='diagnostic-objects')[0].children.length,640);
});
test('preview actions exist only for available real slide previews; content stays text',()=>{
 const doc={createElement:t=>new Node(t)},container=new Node('ul'),summary=new Node('p');let selected=null;
 renderDiagnostics({document:doc,container,summary,diagnostics:[{...make(0),objectName:'<script>not executed</script>'},make(1)],slides:[{previewUrl:'/preview/1.png'},{}],onSelectSlide:i=>selected=i});
 const details=descendants(container).find(n=>n.tag==='details');details.open=true;details.events.toggle();const buttons=descendants(container).filter(n=>n.tag==='button');assert.equal(buttons.length,1);buttons[0].events.click();assert.equal(selected,0);assert.ok(descendants(container).some(n=>n.textContent==='<script>not executed</script>'));
});
test('unknown and locationless fatal errors remain visible in their own category',()=>{
 const p=projectDiagnostics([{code:'INPUT_INVALID',severity:'error',message:'Input invalid'}]);assert.equal(p.errorCount,1);assert.equal(p.categories.length,1);assert.equal(p.categories[0].reason,'Input invalid');assert.deepEqual(p.affectedSlides,[]);
});

test('only finite contained positive bounds can be shown; grouped rows retain per-slide bounds',()=>{
 for(const b of [null,{x:NaN,y:0,w:.2,h:.2},{x:-.1,y:0,w:.2,h:.2},{x:.9,y:0,w:.2,h:.2},{x:0,y:0,w:0,h:.2}])assert.ok(!validBounds(b));
 const one={x:.1,y:.2,w:.3,h:.4},two={x:.2,y:.1,w:.2,h:.3};assert.ok(validBounds(one));
 const p=projectDiagnostics([{...make(0),objectName:undefined,objectType:'background',bounds:one},{...make(1),objectName:undefined,objectType:'background',bounds:two}]);assert.equal(p.categories[0].rows.length,1);assert.deepEqual(p.categories[0].rows[0].locations,[{slideNumber:1,bounds:one},{slideNumber:2,bounds:two}]);
});
test('one blocking cause is first above thousands of representation diagnostics',()=>{
 const raw=Array.from({length:2051},(_,i)=>({...make(i),severity:'error'}));
 raw.push({severity:'error',blocking:true,code:'INPUT_INVALID',reasonKey:'PPTX_PARSE_FATAL',reason:'Document analysis stopped'});
 const p=projectDiagnostics(raw);
 assert.equal(p.categories[0].blocking,true);assert.equal(p.categories[0].reasonKey,'PPTX_PARSE_FATAL');
 assert.equal(p.fatalCount,1);assert.equal(p.representationCount,2051);assert.equal(p.total,2052);
 const container=new Node('ul'),summary=new Node('p');
 renderDiagnostics({document:{createElement:t=>new Node(t)},container,summary,diagnostics:raw});
 assert.match(summary.textContent,/변환 중단 원인 1건.*표현 제한 2051건/);
 const titles=descendants(container).filter(n=>n.className==='diagnostic-title');
 assert.match(titles[0].textContent,/변환 중단 원인.*문서 분석 실패/);
});
test('blocking and recoverable diagnostics never merge or hide a blocking generic cause',()=>{
 const base={...make(0),severity:'error'};
 const p=projectDiagnostics([base,{...base,blocking:true},{...base,blocking:true,reasonKey:'PPTX_OBJECT_UNSUPPORTED'}]);
 assert.equal(p.categories.length,3);assert.equal(p.fatalCount,2);assert.equal(p.representationCount,1);assert.equal(p.suppressedGeneric,0);
 assert.ok(p.categories.slice(0,2).every(c=>c.blocking));
 assert.equal(p.categories[2].blocking,false);
});
test('parser-prefixed inherited and theme reason keys have precise display labels',()=>{
 for(const [reasonKey,title] of [
  ['PPTX_INHERITED_TEXT_UNRESOLVED','상속 텍스트 해석 실패'],
  ['PPTX_THEME_COLOR_UNSUPPORTED','테마 색상 해석 실패'],
  ['PPTX_INHERITED_SHAPE_STYLE_UNSUPPORTED','상속 도형 서식']
 ])assert.equal(projectDiagnostics([{...make(0),reasonKey}]).categories[0].title,title);
});

test('round-trip warnings identify the PPT saving stage',()=>{
 for(const [reasonKey,title] of [
  ['HTML_PPTX_PARAGRAPH_LAYOUT_APPROXIMATED','PPT 재저장 시 문단 배치'],
  ['HTML_PPTX_SHADOW_OMITTED','PPT 재저장 시 그림자']
 ]) assert.equal(projectDiagnostics([{...make(0),reasonKey}]).categories[0].title,title);
});
