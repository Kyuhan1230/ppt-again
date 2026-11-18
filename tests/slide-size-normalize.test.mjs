import test from 'node:test';
import assert from 'node:assert/strict';
import {normalizeSlideSizes} from '../src/html-to-pptx/slide-size-policy.mjs';
import {canvasDiagnostics} from '../src/html-to-pptx/canvas.mjs';

test('contain scales editable geometry and absolute PPT text formatting without mutating the source',()=>{
  const capture={width:800,height:600,renderScale:1,content:'computed-background-only',sourceChildrenCopied:0};
  const textStyle={fontSize:20,lineHeight:'30px',letterSpacing:'2px',padding:{top:2,right:3,bottom:4,left:5},outline:{width:2},shadow:{x:2,y:3,blur:4}};
  const para={beforePx:2,afterPx:3,leftPx:4,rightPx:5,indentPx:-1,lineHeight:'30px',runs:[{text:'A',style:textStyle}],sourceSpacing:{spcBef:{kind:'spcPts',value:200},lnSpc:{kind:'spcPct',value:100000}},bullet:{sourceSize:{kind:'buSzPts',value:800}},sourceEnd:{mode:'explicit',attrs:{sz:'1200',spc:'100',kern:'50',b:'1'}}};
  const nodes=[
    {id:'text',kind:'text',boxCssPx:{x:100,y:100,w:200,h:60},style:textStyle,content:{runs:para.runs,paragraphs:[para]}},
    {id:'line',kind:'line',boxCssPx:{x:10,y:20,w:100,h:50},style:{strokeWidth:2,endpoints:{start:{x:10,y:20},end:{x:110,y:70}}},content:{}},
    {id:'image',kind:'image',boxCssPx:{x:20,y:30,w:200,h:100},style:{capture},content:{crop:{left:10000,top:0,right:0,bottom:0},editableVector:{width:100,height:40,strokeWidth:2}}},
    {id:'table',kind:'table',boxCssPx:{x:0,y:0,w:200,h:100},style:{},content:{columnWidthsCssPx:[100,100],rowHeightsCssPx:[50,50],cells:[{style:{fontSize:10,padding:{top:1,right:1,bottom:1,left:1},borders:{top:{width:2}}},runs:[],paragraphs:[]}]}}
  ];
  const model={slides:[{id:'one',widthCssPx:800,heightCssPx:600,background:{color:'#ff0000',capture},nodes}],assets:[],pptxSizeEmu:{width:800*9525,height:600*9525}};
  const measurements=[{slideId:'one',objects:[{nodeId:'text',boxCssPx:{x:100,y:100,w:200,h:60},lines:[{x:100,y:100,w:100,h:30}],baseline:125,lineAdvance:30}]}];
  const before=structuredClone({model,measurements});
  const result=normalizeSlideSizes(model,measurements,{mode:'contain',targetWidthEmu:1600*9525,targetHeightEmu:900*9525});
  assert.deepEqual({model,measurements},before);
  const slide=result.model.slides[0],text=slide.nodes[0];
  assert.deepEqual(text.boxCssPx,{x:350,y:150,w:300,h:90});
  assert.equal(text.style.fontSize,30);assert.equal(text.style.letterSpacing,'3px');
  assert.deepEqual(text.content.paragraphs[0].sourceSpacing,{spcBef:{kind:'spcPts',value:300},lnSpc:{kind:'spcPct',value:100000}});
  assert.deepEqual(text.content.paragraphs[0].sourceEnd.attrs,{sz:'1800',spc:'150',kern:'75',b:'1'});
  assert.equal(text.content.paragraphs[0].bullet.sourceSize.value,1200);
  assert.deepEqual(slide.nodes[1].style.endpoints,{start:{x:215,y:30},end:{x:365,y:105}});
  assert.deepEqual(slide.nodes[2].content.crop,nodes[2].content.crop);
  assert.deepEqual(slide.nodes[2].content.editableVector,{width:150,height:60,strokeWidth:3});
  assert.deepEqual(slide.nodes[3].content.columnWidthsCssPx,[150,150]);
  assert.equal(slide.nodes[3].content.cells[0].style.borders.top.width,3);
  assert.equal(slide.fitBackground.background.capture.renderScale,2/3);
  assert.equal(slide.nodes[2].style.capture.renderScale,2/3);
  assert.equal(result.measurements[0].objects[0].baseline,187.5);
});

test('inconsistent original PPT metadata requires an explicit output size and is reported',()=>{
  const original={width:1600*9525,height:900*9525};
  const model={slides:[{id:'one',widthCssPx:800,heightCssPx:600,background:{color:'#f00'},nodes:[]}],assets:[],pptxSizeEmu:original};
  assert.throws(()=>normalizeSlideSizes(model,[],{mode:'contain'}),{code:'INPUT_INVALID'});
  const fitted=normalizeSlideSizes(model,[],{mode:'contain',targetWidthEmu:1600*9525,targetHeightEmu:900*9525});
  assert.equal(fitted.sizePolicy.originalAspectMismatch,true);
  assert.deepEqual(model.pptxSizeEmu,original);
});

test('PPT canvas dimensions below one inch are rejected for default, explicit and original sizes',()=>{
  const model={slides:[{id:'one',widthCssPx:8192,heightCssPx:1,background:{color:'#fff'},nodes:[]}],assets:[]};
  assert.ok(canvasDiagnostics(model).length);
  assert.throws(()=>normalizeSlideSizes(model,[]),{code:'INPUT_INVALID'});
  assert.throws(()=>normalizeSlideSizes({...model,pptxSizeEmu:{width:914399,height:914400}},[]),{code:'INPUT_INVALID'});
  assert.throws(()=>normalizeSlideSizes(model,[],{mode:'contain',targetWidthEmu:914399,targetHeightEmu:914400}),{code:'INPUT_INVALID'});
  assert.deepEqual([model.slides[0].widthCssPx,model.slides[0].heightCssPx],[8192,1]);
});

test('valid original EMU tolerates browser CSS rounding while real aspect changes remain invalid',()=>{
  const size={width:9144001,height:6858000};
  const model={slides:[{id:'one',widthCssPx:960,heightCssPx:720,background:{color:'#fff'},nodes:[]}],assets:[],pptxSizeEmu:size};
  assert.deepEqual(canvasDiagnostics(model),[]);
  const normalized=normalizeSlideSizes(model,[]);
  assert.deepEqual(normalized.model.pptxSizeEmu,size);
  assert.deepEqual(canvasDiagnostics(normalized.model),[]);
  const changed=structuredClone(model);changed.slides[0].heightCssPx=700;
  assert.ok(canvasDiagnostics(changed).length);
  assert.throws(()=>normalizeSlideSizes(changed,[]),{code:'INPUT_INVALID'});
});

test('original size recognition does not depend on mixed slide order',()=>{
  const size={width:1280*9525,height:720*9525};
  const make=(id,w,h)=>({id,widthCssPx:w,heightCssPx:h,background:{color:'#fff'},nodes:[{id:'text',kind:'text',boxCssPx:{x:100,y:100,w:200,h:40},style:{fontSize:20},content:{runs:[]}}]});
  const slides=[make('wide',1280,720),make('tall',800,600)],model={slides,assets:[],pptxSizeEmu:size};
  const policy={mode:'contain',targetWidthEmu:1600*9525,targetHeightEmu:900*9525};
  const a=normalizeSlideSizes(model,[],policy),b=normalizeSlideSizes({...model,slides:[...slides].reverse()},[],policy);
  for(const id of ['wide','tall'])assert.deepEqual(a.model.slides.find(s=>s.id===id).nodes[0].boxCssPx,b.model.slides.find(s=>s.id===id).nodes[0].boxCssPx);
  const defaultA=normalizeSlideSizes(model,[]),defaultB=normalizeSlideSizes({...model,slides:[...slides].reverse()},[]);
  for(const id of ['wide','tall']){
    const first=defaultA.sizePolicy.slides.find(s=>s.slideId===id),second=defaultB.sizePolicy.slides.find(s=>s.slideId===id);
    assert.ok(Math.abs(first.physicalScale-second.physicalScale)<1e-9,JSON.stringify({first,second}));
  }
  const bad={...model,slides:[make('tall',800,600)]};assert.throws(()=>normalizeSlideSizes(bad,[]),{code:'INPUT_INVALID'});
  assert.equal(normalizeSlideSizes(bad,[],policy).sizePolicy.originalAspectMismatch,true);
});

test('source PPT absolute paragraph points track physical text scaling after CSS canvas resize',()=>{
  const para={runs:[],sourceSpacing:{spcBef:{kind:'spcPts',value:200}},bullet:{sourceSize:{kind:'buSzPts',value:800}}};
  const node={id:'text',kind:'text',boxCssPx:{x:0,y:0,w:300,h:40},style:{fontSize:20},content:{paragraphs:[para]}};
  const model={slides:[{id:'one',widthCssPx:640,heightCssPx:360,background:{color:'#fff'},nodes:[node]}],assets:[],pptxSizeEmu:{width:1280*9525,height:720*9525}};
  const result=normalizeSlideSizes(model,[]),p=result.model.slides[0].nodes[0].content.paragraphs[0];
  assert.equal(p.sourceSpacing.spcBef.value,400);
  assert.equal(p.bullet.sourceSize.value,1600);
});
