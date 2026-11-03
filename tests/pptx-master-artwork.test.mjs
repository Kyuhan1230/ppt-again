import test from 'node:test';
import assert from 'node:assert/strict';
import {readPptx} from '../src/html-to-pptx/read-pptx.mjs';
import {masterFixture} from './master-artwork-fixture.mjs';

test('visible master and layout shapes and layout-owned pictures precede slide artwork',async()=>{
 const r=await readPptx(await masterFixture(),{allowPartial:true,fontPolicy:'malgun-gothic'});
 assert.equal(r.status,'parsed',JSON.stringify(r.diagnostics));
 const nodes=r.model.slides[0].nodes,byName=name=>nodes.find(n=>n.id===name);
 const m=byName('master:master-shape'),l=byName('layout:layout-shape'),pic=byName('layout:layout-picture'),s=byName('slide-shape');
 assert.ok(m&&l&&pic&&s,JSON.stringify(nodes.map(n=>n.id)));
 assert.equal(m.style.fill.toLowerCase(),'#ff0000');assert.equal(l.style.fill.toLowerCase(),'#00ff00');
 assert.equal(pic.kind,'image');assert.ok(r.model.assets.some(a=>a.id===pic.content.assetId));
 assert.ok(m.zOrder<l.zOrder&&l.zOrder<pic.zOrder&&pic.zOrder<s.zOrder);
 assert.ok(!r.diagnostics.some(d=>d.reasonKey==='PPTX_MASTER_ARTWORK_UNSUPPORTED'));
});

test('showMasterSp on the slide suppresses both inherited layers',async()=>{
 const r=await readPptx(await masterFixture({slideHidden:true}),{allowPartial:true,fontPolicy:'malgun-gothic'});
 const names=r.model.slides[0].nodes.map(n=>n.id);
 assert.ok(names.includes('slide-shape'));assert.ok(!names.includes('master:master-shape'));assert.ok(!names.includes('layout:layout-shape'));assert.ok(!names.includes('layout:layout-picture'));
});

test('showMasterSp on the layout suppresses master artwork, but retains layout artwork',async()=>{
 const r=await readPptx(await masterFixture({layoutHidden:true}),{allowPartial:true,fontPolicy:'malgun-gothic'});
 const names=r.model.slides[0].nodes.map(n=>n.id);
 assert.ok(!names.includes('master:master-shape'));assert.ok(names.includes('layout:layout-shape'));assert.ok(names.includes('layout:layout-picture'));assert.ok(names.includes('slide-shape'));
});

test('unsupported layout shapes remain diagnosed without hiding supported slide content',async()=>{
 const r=await readPptx(await masterFixture({unsupported:true}),{allowPartial:true,fontPolicy:'malgun-gothic'});
 assert.equal(r.status,'parsed',JSON.stringify(r.diagnostics));
 assert.ok(r.model.slides[0].nodes.some(n=>n.id==='slide-shape'));
 assert.ok(r.diagnostics.some(d=>d.reasonKey==='PPTX_SHAPE_GRADIENT_UNSUPPORTED'&&d.objectName==='layout-shape'));
});

test('layout picture placeholders and explicitly hidden artwork are not rendered',async()=>{
 const r=await readPptx(await masterFixture({placeholderPic:true,hiddenShape:true}),{allowPartial:true,fontPolicy:'malgun-gothic'});
 assert.equal(r.status,'parsed',JSON.stringify(r.diagnostics));
 const names=r.model.slides[0].nodes.map(n=>n.id);
 assert.ok(!names.includes('layout:layout-picture'));assert.ok(!names.includes('layout:hidden-artwork'));
 assert.ok(names.includes('slide-picture'));assert.ok(names.includes('layout:layout-shape'));
});
