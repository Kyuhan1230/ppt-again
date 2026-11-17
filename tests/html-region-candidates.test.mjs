import test, {before, after} from 'node:test';
import assert from 'node:assert/strict';
import {readInput} from '../src/html-to-pptx/input.mjs';
import {createBrowserRuntime} from '../src/html-to-pptx/browser-runtime.mjs';
import {inspectInput,verifySelection} from '../src/html-to-pptx/html-regions.mjs';

let runtime;
before(async()=>{runtime=await createBrowserRuntime({capture:false,checkLimits:true});});
after(async()=>{await runtime?.close();});
const source=body=>`<!doctype html><html><body style="margin:0">${body}</body></html>`;
const input=html=>readInput({html});
const inspect=async(html,options={})=>inspectInput(await input(html),runtime,options);
const card=(text,width=280,height=180,style='')=>`<section style="box-sizing:border-box;display:inline-block;vertical-align:top;width:${width}px;height:${height}px;background:#ddd;${style}"><p>${text}</p></section>`;
const regionAt=(result,x,y,width,height)=>result.regions.find(r=>r.bounds&&Math.abs(r.bounds.x-x)<1&&Math.abs(r.bounds.y-y)<1&&Math.abs(r.bounds.width-width)<1&&Math.abs(r.bounds.height-height)<1);
const choiceFor=(result,ids)=>result.choices.find(c=>c.scope==='elements'&&ids.every(id=>c.regionIds.includes(id))&&c.regionIds.length===ids.length);

test('unmarked cards are proposed individually and as a group without being silently selected',async()=>{
  const html=source(`<main style="display:flex;width:900px">${card('one',280,180)}${card('two',300,180)}${card('three',320,190)}</main>`);
  const result=await inspect(html);
  assert.ok(result.snapshotId&&Array.isArray(result.regions)&&Array.isArray(result.choices));
  const a=regionAt(result,0,0,280,180),b=regionAt(result,280,0,300,180),c=regionAt(result,580,0,320,190);
  assert.ok(a&&b&&c,JSON.stringify(result.regions));
  assert.ok(result.choices.some(choice=>choice.scope==='document'));
  assert.ok(choiceFor(result,[a.id,b.id,c.id]),JSON.stringify(result.choices));
  assert.ok([a,b,c].every(r=>result.choices.some(choice=>choice.scope==='elements'&&choice.regionIds.length===1&&choice.regionIds[0]===r.id)));
  assert.ok(result.choices.every(choice=>Array.isArray(choice.evidence)&&choice.evidence.length));
  assert.equal(result.selectedChoiceId??null,null);
  assert.equal(choiceFor(result,[a.id,b.id,c.id]).convertible,true,'different page sizes can be fitted after an explicit choice');
});

test('background-only areas remain candidates and text outside them is counted',async()=>{
  const result=await inspect(source('<div style="width:500px;height:300px;background:#639"></div><p>external content</p>'));
  const region=regionAt(result,0,0,500,300);
  assert.ok(region,JSON.stringify(result.regions));
  const choice=choiceFor(result,[region.id]);
  assert.ok(choice,JSON.stringify(result.choices));
  assert.ok(choice.outsideContentCount>=1,JSON.stringify(choice));
});

test('visible offscreen and hidden elements are distinguished',async()=>{
  const result=await inspect(source('<div style="position:absolute;top:2500px;width:400px;height:300px;background:red">below</div><div style="display:none;width:400px;height:300px">hidden content</div>'));
  const below=regionAt(result,0,2500,400,300);
  assert.ok(below,JSON.stringify(result.regions));
  assert.equal(below.visibility,'offscreen');
  assert.ok(result.regions.some(r=>r.visibility==='hidden'&&r.bounds===null),JSON.stringify(result.regions));
  assert.ok(result.uncertainties.length>0);
});

test('parent-child containment differs from overlap between siblings',async()=>{
  const result=await inspect(source('<main style="position:relative;width:800px;height:500px;background:#eee"><div style="position:absolute;left:0;top:0;width:400px;height:300px;background:red">A</div><div style="position:absolute;left:250px;top:100px;width:400px;height:300px;background:blue">B</div></main>'));
  const parent=regionAt(result,0,0,800,500),a=regionAt(result,0,0,400,300),b=regionAt(result,250,100,400,300);
  assert.ok(parent&&a&&b,JSON.stringify(result.regions));
  assert.equal(a.parentId,parent.id);assert.equal(b.parentId,parent.id);
  assert.ok(a.overlaps.includes(b.id)&&b.overlaps.includes(a.id));
  assert.ok(!a.overlaps.includes(parent.id));
  assert.ok(!result.choices.some(choice=>choice.regionIds.includes(parent.id)&&choice.regionIds.includes(a.id)),'parent and child cannot form a duplicate-content group');
});

test('display contents retains a parent without fabricating its box',async()=>{
  const result=await inspect(source('<main style="display:contents"><div style="width:400px;height:300px;background:#ddd">child</div></main>'));
  const child=regionAt(result,0,0,400,300);
  assert.ok(child,JSON.stringify(result.regions));
  const parent=result.regions.find(r=>r.id===child.parentId);
  assert.ok(parent&&parent.bounds===null,JSON.stringify(result.regions));
});

test('script, internal clipping and generated content are reported as unobserved',async()=>{
  const result=await inspect(source('<style>aside::before{content:"generated"}</style><div style="width:400px;height:300px;background:#ddd">page</div><aside style="height:30px;overflow:auto"><p style="height:100px">scroll</p></aside><script>document.body.append("dynamic")</script>'));
  assert.ok(result.uncertainties.length>0,JSON.stringify(result));
  assert.ok(result.choices.length>0);
});

test('public candidate identifiers do not contain text, DOM paths or selectors',async()=>{
  const html=source('<section id="private-region-name" style="width:400px;height:300px;background:#ddd">CONFIDENTIAL-PHRASE-DO-NOT-EXPOSE</section>');
  const result=await inspect(html);
  const serialized=JSON.stringify(result);
  assert.doesNotMatch(serialized,/CONFIDENTIAL-PHRASE-DO-NOT-EXPOSE|private-region-name|data-source-path|section@\d+/);
  assert.ok(result.regions.every(r=>typeof r.id==='string'&&r.id.length>0));
  assert.equal(new Set(result.regions.map(r=>r.id)).size,result.regions.length);
});

test('a fresh inspection can verify its choice and rejects changed input, viewport, or mixed IDs',async()=>{
  const original=source('<section style="width:400px;height:300px;background:#ddd">ORIGINAL</section>');
  const result=await inspect(original);
  const choice=result.choices.find(c=>c.scope==='elements');
  assert.ok(choice,JSON.stringify(result));
  const selection={snapshotId:result.snapshotId,choiceId:choice.id};
  const verified=await verifySelection(await input(original),runtime,selection);
  assert.equal(verified.choice.id,choice.id);
  assert.deepEqual(verified.choice.regionIds,choice.regionIds);
  const forged=await verifySelection(await input(original),runtime,{...selection,regionIds:['forged'],selector:'body',bounds:{x:999,y:999,width:1,height:1}});
  assert.deepEqual(forged.choice.regionIds,choice.regionIds);
  const changedInput=await input(original.replace('ORIGINAL','CHANGED'));
  await assert.rejects(()=>verifySelection(changedInput,runtime,selection),{code:'STALE_INSPECTION'});
  const originalInput=await input(original);
  await assert.rejects(()=>verifySelection(originalInput,runtime,{...selection,viewport:{width:900,height:700}}),{code:'STALE_INSPECTION'});
  await assert.rejects(()=>verifySelection(originalInput,runtime,{...selection,choiceId:'not-an-existing-choice'}),{code:'STALE_INSPECTION'});
  const another=await inspect(original.replace('ORIGINAL','DIFFERENT'));
  await assert.rejects(()=>verifySelection(originalInput,runtime,{snapshotId:another.snapshotId,choiceId:choice.id}),{code:'STALE_INSPECTION'});
});

const svgBytes=color=>new TextEncoder().encode(`<svg xmlns="http://www.w3.org/2000/svg" width="40" height="40"><rect width="40" height="40" fill="${color}"/></svg>`);
test('an image asset change invalidates the old selection even when HTML is unchanged',async()=>{
  const html=source('<section style="width:400px;height:300px"><img src="secret-asset-name.svg" width="40" height="40"></section>');
  const first=await readInput({html,assets:[{path:'secret-asset-name.svg',bytes:svgBytes('red')}]});
  const changed=await readInput({html,assets:[{path:'secret-asset-name.svg',bytes:svgBytes('blue')}]});
  const result=await inspectInput(first,runtime);
  assert.doesNotMatch(JSON.stringify(result),/secret-asset-name\.svg/);
  const choice=result.choices.find(c=>c.scope==='elements');
  assert.ok(choice,JSON.stringify(result));
  await assert.rejects(()=>verifySelection(changed,runtime,{snapshotId:result.snapshotId,choiceId:choice.id}),{code:'STALE_INSPECTION'});
});

test('worker can inspect but rejects oversized inputs before starting a browser',async()=>{
  const {runConversion}=await import('../src/html-to-pptx/job-runner.mjs');
  const result=await runConversion({operation:'inspect-html',input:{html:source('<section style="width:400px;height:300px;background:#ddd">page</section>'),assets:[]}});
  assert.ok(result.inspection?.snapshotId);
  assert.ok(result.resourceUsage?.workerPid);
  await assert.rejects(()=>runConversion({operation:'inspect-html',input:{html:'x'.repeat(16*1024*1024+1),assets:[]}}),{code:'INPUT_LIMIT'});
  await assert.rejects(()=>runConversion({operation:'inspect-html',input:{html:'<p>ok</p>',assets:[{path:'large.png',bytes:new Uint8Array(16*1024*1024+1)}]}}),{code:'INPUT_LIMIT'});
});

test('explicit and declared-slide sources remain identifiable without guessing from size',async()=>{
  const html=source('<p>outside</p><section id="real-slide" data-slide style="width:400px;height:300px">page</section>');
  const declared=await inspect(html);
  assert.ok(declared.choices.some(c=>c.evidence.includes('declared-slide')));
  const explicit=await inspectInput(await readInput({html,slideSelector:'#real-slide'}),runtime);
  assert.ok(explicit.choices.some(c=>c.evidence.includes('explicit-selector')));
  assert.doesNotMatch(JSON.stringify(explicit),/#real-slide/);
});

test('candidate budget exhaustion fails explicitly rather than returning a partial successful list',async()=>{
  const many=source(Array.from({length:501},(_,i)=>`<section style="width:300px;height:200px;background:#ddd">item ${i}</section>`).join(''));
  await assert.rejects(()=>inspect(many),{code:'INPUT_LIMIT'});
});

for(const [label,fragment] of [
  ['internal scroll','<div style="width:400px;height:200px;overflow:auto"><section style="height:400px;background:#ddd">visible and clipped</section></div>'],
  ['overflow hidden','<div style="width:400px;height:200px;overflow:hidden"><section style="height:400px;background:#ddd">visible and clipped</section></div>'],
  ['generated content','<style>section::before{content:"generated";position:absolute;left:500px}</style><section style="width:400px;height:200px;background:#ddd">page</section>'],
  ['opacity composition','<section style="width:400px;height:200px;background:#ddd;opacity:.5">page</section>'],
  ['shadow composition','<section style="width:400px;height:200px;background:#ddd;box-shadow:10px 10px black">page</section>']
])test(`${label} is recorded as a limit to complete observation`,async()=>{
  const result=await inspect(source(fragment));
  assert.ok(result.uncertainties.length>0,`${label}: ${JSON.stringify(result)}`);
});

test('changed text wrapping invalidates the selection even when the enclosing rectangle is unchanged',async()=>{
  const html=source('<section style="width:400px;height:250px;background:#ddd;font:14px Arial">A long sentence that will wrap after the font size becomes much larger than its original size.</section>');
  const originalInput=await input(html);
  const result=await inspectInput(originalInput,runtime);
  const choice=result.choices.find(c=>c.scope==='elements');
  assert.ok(choice);
  const changedRuntime={...runtime,render:async(...args)=>{
    const session=await runtime.render(...args);
    await session.page.evaluate(()=>document.querySelector('section').style.setProperty('font-size','40px','important'));
    return session;
  }};
  await assert.rejects(()=>verifySelection(originalInput,changedRuntime,{snapshotId:result.snapshotId,choiceId:choice.id}),{code:'STALE_INSPECTION'});
});

test('clipped content is not counted as confirmed visible outside content',async()=>{
  const html=source('<section style="width:400px;height:300px;background:#ddd">page</section><aside style="position:relative;width:100px;height:20px;overflow:hidden"><span style="position:absolute;top:200px">not visible</span></aside>');
  const result=await inspect(html);
  const region=regionAt(result,0,0,400,300);
  const choice=choiceFor(result,[region.id]);
  assert.ok(choice);
  assert.equal(choice.outsideContentCount,0,JSON.stringify(choice));
  assert.ok(result.uncertainties.length>0);
});

test('unused supplied asset changes invalidate the input snapshot',async()=>{
  const html=source('<section style="width:400px;height:300px;background:#ddd">page</section>');
  const first=await readInput({html,assets:[{path:'unused.svg',bytes:svgBytes('red')}]});
  const changed=await readInput({html,assets:[{path:'unused.svg',bytes:svgBytes('blue')}]});
  const result=await inspectInput(first,runtime);
  const choice=result.choices.find(c=>c.scope==='elements');
  assert.ok(choice);
  await assert.rejects(()=>verifySelection(changed,runtime,{snapshotId:result.snapshotId,choiceId:choice.id}),{code:'STALE_INSPECTION'});
});

test('the URL to asset bytes mapping is part of the snapshot',async()=>{
  const html=source('<section style="width:400px;height:300px"><img src="a.svg"><img src="b.svg"></section>');
  const first=await readInput({html,assets:[{path:'a.svg',bytes:svgBytes('red')},{path:'b.svg',bytes:svgBytes('red')}]});
  const changed=await readInput({html,assets:[{path:'a.svg',bytes:svgBytes('red')},{path:'b.svg',bytes:svgBytes('blue')}]});
  const result=await inspectInput(first,runtime);
  const choice=result.choices.find(c=>c.scope==='elements');
  assert.ok(choice);
  await assert.rejects(()=>verifySelection(changed,runtime,{snapshotId:result.snapshotId,choiceId:choice.id}),{code:'STALE_INSPECTION'});
});

test('valid SVG roots are reported as verified SVG candidates',async()=>{
  const html='<svg viewBox="0 0 800 600" width="800" height="600" xmlns="http://www.w3.org/2000/svg"><rect width="100" height="100" fill="red"/></svg>';
  const parsed=await input(html);
  assert.equal(parsed.profile,'svg-slides-v1');
  const result=await inspectInput(parsed,runtime);
  assert.ok(result.choices.some(c=>c.evidence.includes('validated-svg')),JSON.stringify(result));
});

test('explicit html selector is offered without revealing the selector',async()=>{
  const html=source('<section style="width:400px;height:300px">page</section>');
  const parsed=await readInput({html,slideSelector:'html'});
  const result=await inspectInput(parsed,runtime);
  assert.ok(result.choices.some(c=>c.evidence.includes('explicit-selector')),JSON.stringify(result));
});

test('a descendant can restore visibility hidden by its parent',async()=>{
  const html=source('<main style="visibility:hidden"><section style="visibility:visible;width:400px;height:300px;background:red">visible</section></main>');
  const result=await inspect(html);
  const child=regionAt(result,0,0,400,300);
  assert.ok(child,JSON.stringify(result.regions));
  assert.equal(child.visibility,'visible');
  assert.ok(result.uncertainties.length>0);
});

test('unstable animation requires a new inspection instead of accepting an old snapshot',async()=>{
  const html=source('<style>@keyframes change{from{opacity:1}to{opacity:.5}}</style><section style="width:400px;height:300px;background:red;animation:change 1s infinite">animated</section>');
  const parsed=await input(html);
  const result=await inspectInput(parsed,runtime);
  const choice=result.choices.find(c=>c.scope==='elements');
  assert.ok(choice);
  assert.ok(result.uncertainties.length>0);
  await assert.rejects(()=>verifySelection(parsed,runtime,{snapshotId:result.snapshotId,choiceId:choice.id}),{code:'STALE_INSPECTION'});
});

test('document candidate bounds cover an image placed at negative coordinates',async()=>{
  const img='data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+j1ioAAAAASUVORK5CYII=';
  const result=await inspect(source(`<img src="${img}" style="position:absolute;left:-200px;top:-100px;width:100px;height:80px"><section style="width:400px;height:300px">page</section>`));
  const documentChoice=result.choices.find(c=>c.scope==='document');
  assert.ok(documentChoice?.bounds,JSON.stringify(result));
  assert.ok(documentChoice.bounds.x<=-200&&documentChoice.bounds.y<=-100,JSON.stringify(documentChoice));
});

test('similar siblings with substantially different internal layout are not proposed as a repeated group',async()=>{
  const html=source('<main style="display:flex"><section style="position:relative;width:400px;height:300px;background:#ddd"><p style="position:absolute;top:10px;left:20px;margin:0">A</p></section><section style="position:relative;width:400px;height:300px;background:#ddd"><p style="position:absolute;top:220px;left:20px;margin:0">B</p></section></main>');
  const result=await inspect(html);
  const a=regionAt(result,0,0,400,300),b=regionAt(result,400,0,400,300);
  assert.ok(a&&b,JSON.stringify(result.regions));
  assert.equal(choiceFor(result,[a.id,b.id]),undefined,JSON.stringify(result.choices));
});

test('visible absolutely positioned descendants protrude outside their candidate but clipped descendants do not',async()=>{
  for(const [overflow,expectedOutside] of [['visible',true],['hidden',false]]){
    const html=source(`<section style="position:relative;width:400px;height:300px;background:#ddd;overflow:${overflow}"><span style="position:absolute;left:500px;top:0">OUTSIDE</span></section>`);
    const result=await inspect(html);
    const region=regionAt(result,0,0,400,300);
    assert.ok(region,JSON.stringify(result.regions));
    const choice=choiceFor(result,[region.id]);
    assert.ok(choice,JSON.stringify(result.choices));
    assert.equal(choice.outsideContentCount>0,expectedOutside,JSON.stringify(choice));
    if(!expectedOutside)assert.ok(result.uncertainties.length>0);
  }
});

test('validated embedded pages are separate indexed candidates and offer a page group',async()=>{
  const data={pages:[{html:'<section style="width:400px;height:300px;background:#ddd">first</section>',source:'a'},{html:'<section style="width:600px;height:350px;background:#ddd">second</section>',source:'a'}],sources:{a:{styles:''}}};
  const parsed=await input(`<script type="application/json">${JSON.stringify(data)}</script>`);
  assert.equal(parsed.embeddedSlides?.length,2,JSON.stringify(parsed.diagnostics));
  const result=await inspectInput(parsed,runtime);
  assert.equal(result.regions.length,2,JSON.stringify(result.regions));
  assert.deepEqual(result.regions.map(r=>r.pageIndex),[0,1]);
  assert.notEqual(result.regions[0].id,result.regions[1].id);
  assert.ok(result.regions.every(r=>r.evidence.includes('validated-embedded-slide')));
  const group=choiceFor(result,result.regions.map(r=>r.id));
  assert.ok(group,JSON.stringify(result.choices));
  assert.equal(group.convertible,false);
  const verified=await verifySelection(parsed,runtime,{snapshotId:result.snapshotId,choiceId:group.id});
  assert.deepEqual(verified.choice.regionIds,group.regionIds);
});

test('large top-level p and h1 elements are offered as candidates',async()=>{
  for(const tag of ['p','h1']){
    const result=await inspect(source(`<${tag} style="margin:0;width:600px;height:400px">Large standalone text</${tag}>`));
    const region=regionAt(result,0,0,600,400);
    assert.ok(region,`${tag}: ${JSON.stringify(result.regions)}`);
    assert.ok(choiceFor(result,[region.id]),`${tag}: ${JSON.stringify(result.choices)}`);
  }
});

test('hundreds of small paragraphs inside a slide do not become hundreds of slide candidates',async()=>{
  const paragraphs=Array.from({length:300},(_,i)=>`<p style="margin:0;font:1px/1px Arial">line ${i}</p>`).join('');
  const result=await inspect(source(`<section style="width:600px;height:400px;overflow:hidden">${paragraphs}</section>`));
  const section=regionAt(result,0,0,600,400);
  assert.ok(section,JSON.stringify(result.regions));
  assert.ok(result.regions.length<20,`small paragraphs must not be individual candidates: ${result.regions.length}`);
});

test('explicit or declared image selection survives merging an identical transparent wrapper',async()=>{
  const img='data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+j1ioAAAAASUVORK5CYII=';
  const html=source(`<section style="width:400px;height:300px"><img id="only" data-slide src="${img}" style="display:block;width:400px;height:300px"></section>`);
  for(const [selector,evidence] of [['#only','explicit-selector'],[null,'declared-slide']]){
    const parsed=await readInput({html,slideSelector:selector});
    const result=await inspectInput(parsed,runtime);
    const chosen=result.choices.find(c=>c.evidence.includes(evidence));
    assert.ok(chosen,`${evidence}: ${JSON.stringify(result)}`);
    assert.equal(result.choices[0].id,chosen.id);
    assert.equal(chosen.regionIds.length,1);
    const region=result.regions.find(r=>r.id===chosen.regionIds[0]);
    assert.deepEqual(region?.bounds,{x:0,y:0,width:400,height:300});
  }
});

test('equivalent transparent wrappers are one choice while their parent relationships remain available',async()=>{
  const img='data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+j1ioAAAAASUVORK5CYII=';
  const html=source(`<main style="width:400px;height:300px"><section style="width:400px;height:300px"><img src="${img}" style="display:block;width:400px;height:300px"></section></main>`);
  const result=await inspect(html);
  const matches=result.regions.filter(r=>r.bounds?.x===0&&r.bounds?.y===0&&r.bounds?.width===400&&r.bounds?.height===300);
  assert.equal(matches.length,3,JSON.stringify(result.regions));
  const ids=matches.map(r=>r.id),choices=result.choices.filter(c=>c.scope==='elements'&&c.regionIds.length===1&&ids.includes(c.regionIds[0]));
  assert.equal(choices.length,1,JSON.stringify(result.choices));
  assert.deepEqual(new Set([choices[0].regionIds[0],...(choices[0].equivalentRegionIds??[])]),new Set(ids));
  assert.equal(matches.filter(r=>ids.includes(r.parentId)).length,2,JSON.stringify(matches));
});

test('hundreds of small text divs in a declared slide do not exhaust the candidate budget silently',async()=>{
  const objects=Array.from({length:501},(_,i)=>`<div style="position:absolute;left:0;top:${i}px;width:10px;height:1px;font:1px/1px Arial">${i}</div>`).join('');
  const result=await inspect(source(`<section data-slide style="position:relative;width:800px;height:600px">${objects}</section>`));
  const slide=regionAt(result,0,0,800,600);
  assert.ok(slide,JSON.stringify(result.regions));
  assert.ok(result.choices.some(c=>c.evidence.includes('declared-slide')));
  assert.ok(result.regions.length<30,`small objects must not exhaust the candidate budget: ${result.regions.length}`);
  assert.ok(result.uncertainties.includes('objects-inside-declared-slide-not-separate-candidates'),JSON.stringify(result.uncertainties));
});

test('small background, image, offscreen and repeated regions remain candidate options within unmarked containers',async()=>{
  const img='data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+j1ioAAAAASUVORK5CYII=';
  const html=source(`<section style="position:relative;width:800px;height:600px"><aside style="position:absolute;left:10px;top:10px;width:20px;height:20px;background:red"></aside><img src="${img}" style="position:absolute;left:40px;top:10px;width:20px;height:20px"><div style="position:absolute;left:10px;top:3000px;width:30px;height:20px">offscreen</div><div style="position:absolute;left:10px;top:50px;width:40px;height:20px">A</div><div style="position:absolute;left:60px;top:50px;width:40px;height:20px">B</div></section>`);
  const result=await inspect(html);
  const bg=regionAt(result,10,10,20,20),image=regionAt(result,40,10,20,20),offscreen=regionAt(result,10,3000,30,20);
  const a=regionAt(result,10,50,40,20),b=regionAt(result,60,50,40,20);
  assert.ok(bg&&image&&offscreen&&a&&b,JSON.stringify(result.regions));
  assert.equal(offscreen.visibility,'offscreen');
  assert.ok(result.choices.some(choice=>choice.scope==='elements'&&choice.regionIds.includes(a.id)&&choice.regionIds.includes(b.id)),JSON.stringify(result.choices));
});

test('hundreds of colored shapes and images inside a declared slide do not exhaust the candidate budget',async()=>{
  const img='data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+j1ioAAAAASUVORK5CYII=';
  const objects=Array.from({length:501},(_,i)=>{
    const x=(i%40)*6,y=Math.floor(i/40)*6;
    return `<div style="position:absolute;left:${x}px;top:${y}px;width:4px;height:4px;background:red"></div><img src="${img}" style="position:absolute;left:${x}px;top:${y+100}px;width:4px;height:4px">`;
  }).join('');
  const result=await inspect(source(`<section data-slide style="position:relative;width:800px;height:600px">${objects}</section>`));
  assert.ok(regionAt(result,0,0,800,600));
  assert.ok(result.choices.some(c=>c.evidence.includes('declared-slide')));
  assert.ok(result.regions.length<30,`ordinary slide objects must not exhaust the candidate budget: ${result.regions.length}`);
  assert.ok(result.uncertainties.includes('objects-inside-declared-slide-not-separate-candidates'),JSON.stringify(result.uncertainties));
});

test('small colored shapes and images remain candidates inside unmarked undecided containers',async()=>{
  const img='data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+j1ioAAAAASUVORK5CYII=';
  const html=source(`<main style="position:relative;width:800px;height:600px"><div style="position:absolute;left:10px;top:10px;width:20px;height:20px;background:red"></div><img src="${img}" style="position:absolute;left:40px;top:10px;width:20px;height:20px"></main>`);
  const result=await inspect(html);
  const shape=regionAt(result,10,10,20,20),image=regionAt(result,40,10,20,20);
  assert.ok(shape&&image,JSON.stringify(result.regions));
  assert.ok(choiceFor(result,[shape.id])&&choiceFor(result,[image.id]),JSON.stringify(result.choices));
});

test('a new worker re-inspects and verifies the submitted snapshot and choice',async()=>{
  const {runConversion}=await import('../src/html-to-pptx/job-runner.mjs');
  const html=source('<section style="width:400px;height:300px;background:#ddd">PRIVATE TEXT</section>');
  const inspected=await runConversion({operation:'inspect-html',input:{html,assets:[]}});
  const choice=inspected.inspection.choices.find(c=>c.scope==='elements');
  assert.ok(choice,JSON.stringify(inspected.inspection));
  const selection={html,assets:[],snapshotId:inspected.inspection.snapshotId,choiceId:choice.id};
  const verified=await runConversion({operation:'verify-html-selection',input:selection});
  assert.ok(verified.resourceUsage?.workerPid);
  assert.notEqual(verified.resourceUsage.workerPid,inspected.resourceUsage.workerPid,'verification must run in a new worker');
  assert.equal(verified.verification.snapshotId,selection.snapshotId);
  assert.equal(verified.verification.choice.id,choice.id);
  assert.deepEqual(verified.verification.choice.regionIds,choice.regionIds);
  assert.doesNotMatch(JSON.stringify(verified.verification),/PRIVATE TEXT|data-source-path/);
  await assert.rejects(()=>runConversion({operation:'verify-html-selection',input:{...selection,choiceId:'c-not-a-real-choice'}}),{code:'STALE_INSPECTION',statusCode:409});
  await assert.rejects(()=>runConversion({operation:'verify-html-selection',input:{...selection,html:html.replace('PRIVATE TEXT','changed content')}}),{code:'STALE_INSPECTION',statusCode:409});
});

test('worker selection reinspection enforces input limits and honors cancellation before spawning',async()=>{
  const {runConversion}=await import('../src/html-to-pptx/job-runner.mjs');
  const selection={html:source('<section>page</section>'),assets:[],snapshotId:'s-bad',choiceId:'c-bad'};
  const checkPreflight=code=>error=>error.code===code&&error.resourceUsage?.workerPid===null;
  await assert.rejects(()=>runConversion({operation:'verify-html-selection',input:{...selection,html:'x'.repeat(16*1024*1024+1)}}),checkPreflight('INPUT_LIMIT'));
  await assert.rejects(()=>runConversion({operation:'verify-html-selection',input:{...selection,assets:[{path:'oversized.png',bytes:new Uint8Array(16*1024*1024+1)}]}}),checkPreflight('INPUT_LIMIT'));
  await assert.rejects(()=>runConversion({operation:'verify-html-selection',input:{...selection,regionIds:['forged']}}),checkPreflight('INPUT_INVALID'));
  await assert.rejects(()=>runConversion({operation:'verify-html-selection',input:{...selection,bounds:{x:0,y:0,width:999,height:999}}}),checkPreflight('INPUT_INVALID'));
  await assert.rejects(()=>runConversion({operation:'verify-html-selection',input:{...selection,assets:[{path:'x.svg',bytes:svgBytes('red'),extra:'forged'}]}}),checkPreflight('INPUT_INVALID'));
  const controller=new AbortController();controller.abort();
  await assert.rejects(()=>runConversion({operation:'verify-html-selection',input:selection,signal:controller.signal}),checkPreflight('JOB_CANCELLED'));
});
