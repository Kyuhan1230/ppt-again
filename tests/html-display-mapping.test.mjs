import test from 'node:test';
import assert from 'node:assert/strict';
import {validateDisplayMapping} from '../src/html-to-pptx/html-display-mapping.mjs';

function fixture(){
  const normal={fontFamily:'Arial',fontSize:20,fontWeight:400};
  const bold={fontFamily:'Arial',fontSize:20,fontWeight:700};
  const input={
    sourceObjectIds:['paragraph'],objectIds:['paragraph'],sourceBreakIds:['br-1'],lineIds:['line-1','line-2','line-3'],
    sourceLedger:[
      {sourceId:'a',kind:'text',text:'same  same ',ranges:[{start:0,end:11}],whiteSpace:'normal',style:normal},
      {sourceId:'br-1',kind:'br',text:'\n',ranges:[{start:0,end:1}],whiteSpace:'normal',style:normal},
      {sourceId:'b',kind:'text',text:'same',ranges:[{start:0,end:4}],whiteSpace:'normal',style:bold},
    ],
    observations:[
      {sourceId:'a',start:0,end:4,lineId:'line-1',kind:'visible',text:'same',style:normal},
      {sourceId:'a',start:4,end:5,lineId:'line-1',kind:'visible',text:' ',style:normal},
      {sourceId:'a',start:5,end:6,lineId:'line-1',kind:'collapsed',text:'',style:normal},
      {sourceId:'a',start:6,end:10,lineId:'line-1',kind:'visible',text:'same',style:normal},
      {sourceId:'a',start:10,end:11,lineId:'line-1',kind:'suppressed',text:'',style:normal},
      {sourceId:'br-1',start:0,end:1,lineId:'line-2',kind:'hard-break',text:'\n',style:normal},
      {sourceId:'b',start:0,end:4,lineId:'line-3',kind:'visible',text:'same',style:bold},
    ],
    tokens:[
      {kind:'copy',sources:[{sourceId:'a',start:0,end:4}],text:'same',style:normal,lineId:'line-1'},
      {kind:'collapsed-space',sources:[{sourceId:'a',start:4,end:5},{sourceId:'a',start:5,end:6}],text:' ',style:normal,lineId:'line-1'},
      {kind:'copy',sources:[{sourceId:'a',start:6,end:10}],text:'same',style:normal,lineId:'line-1'},
      {kind:'suppressed-space',sources:[{sourceId:'a',start:10,end:11}],text:'',style:normal,lineId:'line-1'},
      {kind:'layout-break',sources:[],text:'\n',style:normal,lineId:'line-1',nextLineId:'line-2'},
      {kind:'hard-break',sources:[{sourceId:'br-1',start:0,end:1}],text:'\n',style:normal,lineId:'line-2',nextLineId:'line-3'},
      {kind:'copy',sources:[{sourceId:'b',start:0,end:4}],text:'same',style:bold,lineId:'line-3'},
    ],
  };
  input.lineOwners=input.lineIds.map(lineId=>({lineId,objectId:'paragraph'}));
  for(const key of ['sourceLedger','observations','tokens'])input[key]=input[key].map(item=>({...item,objectId:'paragraph'}));
  return input;
}

test('display mapping accepts measured copies, collapsed and suppressed spaces, BR and automatic line breaks',()=>{
  const input=fixture(),before=structuredClone(input),result=validateDisplayMapping(input);
  assert.equal(result.valid,true,JSON.stringify(result.errors));assert.deepEqual(result.errors,[]);
  assert.deepEqual(input,before,'validation must not mutate independent source or display ledgers');
  assert.equal(input.tokens.map(token=>token.text).join(''),'same same\n\nsame');
});

test('display mapping rejects a missing or duplicated source range even when repeated words look identical',()=>{
  const missing=fixture();missing.tokens.splice(2,1);
  const duplicate=fixture();duplicate.tokens.splice(2,0,structuredClone(duplicate.tokens[0]));
  for(const input of [missing,duplicate]){
    const result=validateDisplayMapping(input);assert.equal(result.valid,false);assert.ok(result.errors.length);
  }
});

test('independent observations cannot be missing, duplicated, fabricated or reordered',()=>{
  const missing=fixture();missing.observations.splice(3,1);
  const duplicate=fixture();duplicate.observations.splice(2,0,structuredClone(duplicate.observations[0]));
  const fabricated=fixture();fabricated.observations[0].text='fake';
  const moved=fixture();const first=moved.tokens[0].sources;moved.tokens[0].sources=moved.tokens[2].sources;moved.tokens[2].sources=first;
  for(const input of [missing,duplicate,fabricated,moved]){
    const result=validateDisplayMapping(input);assert.equal(result.valid,false);assert.ok(result.errors.length);
  }
});

test('a single collapsed space cannot merge different observed text styles',()=>{
  const input=fixture();input.observations[2].style={...input.observations[2].style,fontWeight:700};
  const result=validateDisplayMapping(input);assert.equal(result.valid,false);assert.ok(result.errors.length);
});

test('display mapping rejects fabricated copy text, omitted objects and incorrectly styled text',()=>{
  const fabricated=fixture();fabricated.tokens[0].text='fake';
  const omitted=fixture();omitted.objectIds=[];
  const wrongStyle=fixture();wrongStyle.tokens[6].style=wrongStyle.tokens[0].style;
  for(const input of [fabricated,omitted,wrongStyle]){
    const result=validateDisplayMapping(input);assert.equal(result.valid,false);assert.ok(result.errors.length);
  }
});

test('only CSS-collapsible whitespace may collapse and visible mid-line spaces cannot be suppressed',()=>{
  const nonSpace=fixture();nonSpace.tokens[1].sources[0].start=3;
  const preserve=fixture();preserve.sourceLedger[0].whiteSpace='pre';
  const nbsp=fixture();nbsp.sourceLedger[0].text='same\u00a0 same ';
  const middle=fixture();middle.tokens[1].kind='suppressed-space';middle.tokens[1].text='';
  for(const input of [nonSpace,preserve,nbsp,middle]){
    const result=validateDisplayMapping(input);assert.equal(result.valid,false);assert.ok(result.errors.length);
  }
});

test('observed visible whitespace must be consistent with CSS white-space rules',()=>{
  const style={fontFamily:'Arial',fontSize:20,fontWeight:400},objectId='text';
  const make=(text,whiteSpace)=>({sourceObjectIds:[objectId],objectIds:[objectId],sourceBreakIds:[],lineIds:['one'],lineOwners:[{lineId:'one',objectId}],
    sourceLedger:[{sourceId:'node',objectId,kind:'text',text,ranges:[{start:0,end:text.length}],whiteSpace,style}],
    observations:[{sourceId:'node',objectId,start:0,end:text.length,lineId:'one',kind:'visible',text,style}],
    tokens:[{kind:'copy',objectId,sources:[{sourceId:'node',start:0,end:text.length}],text,style,lineId:'one'}]});
  const kept=validateDisplayMapping(make('a  b','pre'));assert.equal(kept.valid,true,JSON.stringify(kept.errors));
  const edges=validateDisplayMapping(make(' a ','pre'));assert.equal(edges.valid,true,JSON.stringify(edges.errors));
  for(const [text,whiteSpace] of [['a  b','normal'],['a  b','nowrap'],['a  b','pre-line'],[' a','normal'],['a ','normal'],['a\tb','normal'],['a\nb','normal'],['a\nb','pre-line']]){
    const result=validateDisplayMapping(make(text,whiteSpace));assert.equal(result.valid,false,JSON.stringify({text,whiteSpace}));assert.ok(result.errors.length);
  }
  const adjacent=make('a  b','normal');
  adjacent.observations=[{sourceId:'node',objectId,start:0,end:2,lineId:'one',kind:'visible',text:'a ',style},{sourceId:'node',objectId,start:2,end:4,lineId:'one',kind:'visible',text:' b',style}];
  adjacent.tokens=[{kind:'copy',objectId,sources:[{sourceId:'node',start:0,end:2}],text:'a ',style,lineId:'one'},{kind:'copy',objectId,sources:[{sourceId:'node',start:2,end:4}],text:' b',style,lineId:'one'}];
  const result=validateDisplayMapping(adjacent);assert.equal(result.valid,false);assert.ok(result.errors.length);
});

test('hard breaks require a real unique BR node and automatic breaks connect adjacent measured lines',()=>{
  const absent=fixture();absent.sourceBreakIds=[];
  const duplicate=fixture();duplicate.tokens.splice(6,0,structuredClone(duplicate.tokens[5]));
  const fake=fixture();fake.tokens[5].sources=[{sourceId:'a',start:10,end:11}];
  const skip=fixture();skip.tokens[4].nextLineId='line-3';
  const invented=fixture();invented.tokens[4].sources=[{sourceId:'a',start:10,end:11}];
  for(const input of [absent,duplicate,fake,skip,invented]){
    const result=validateDisplayMapping(input);assert.equal(result.valid,false);assert.ok(result.errors.length);
  }
});

function preservedNewline(whitespace,newline){
  const style={fontFamily:'Arial',fontSize:20,fontWeight:400},text=`a${newline}b`,end=1+newline.length;
  return {sourceObjectIds:['text'],objectIds:['text'],sourceBreakIds:[],lineIds:['one','two'],lineOwners:[{lineId:'one',objectId:'text'},{lineId:'two',objectId:'text'}],
    sourceLedger:[{sourceId:'node',objectId:'text',kind:'text',text,ranges:[{start:0,end:text.length}],whiteSpace:whitespace,style}],
    observations:[{sourceId:'node',objectId:'text',start:0,end:1,lineId:'one',kind:'visible',text:'a',style},{sourceId:'node',objectId:'text',start:1,end,lineId:'one',kind:'text-break',text:'\n',style},{sourceId:'node',objectId:'text',start:end,end:end+1,lineId:'two',kind:'visible',text:'b',style}],
    tokens:[{kind:'copy',objectId:'text',sources:[{sourceId:'node',start:0,end:1}],text:'a',style,lineId:'one'},{kind:'text-break',objectId:'text',sources:[{sourceId:'node',start:1,end}],text:'\n',style,lineId:'one',nextLineId:'two'},{kind:'copy',objectId:'text',sources:[{sourceId:'node',start:end,end:end+1}],text:'b',style,lineId:'two'}]};
}

test('preserved LF and CRLF text-node newlines become one text break, not a duplicated automatic break',()=>{
  for(const whitespace of ['pre','pre-wrap','pre-line','break-spaces'])for(const newline of ['\n','\r\n']){
    const input=preservedNewline(whitespace,newline),result=validateDisplayMapping(input);
    assert.equal(result.valid,true,JSON.stringify({whitespace,newline,errors:result.errors}));
    assert.equal(input.tokens.map(token=>token.text).join(''),'a\nb');
  }
});

test('preserved newlines cannot be copied and followed by a second layout break or invented under normal whitespace',()=>{
  const duplicated=preservedNewline('pre-line','\n');duplicated.tokens[1].kind='copy';delete duplicated.tokens[1].nextLineId;
  duplicated.tokens.splice(2,0,{kind:'layout-break',objectId:'text',sources:[],text:'\n',style:duplicated.tokens[0].style,lineId:'one',nextLineId:'two'});
  const invented=preservedNewline('normal','\n');
  const divided=preservedNewline('pre','\r\n');divided.tokens[1].sources[0].end=2;
  for(const input of [duplicated,invented,divided]){const result=validateDisplayMapping(input);assert.equal(result.valid,false);assert.ok(result.errors.length);}
});

test('consecutive automatic breaks cannot fabricate an unobserved empty line',()=>{
  const style={fontFamily:'Arial',fontSize:20,fontWeight:400},objectId='text';
  const input={sourceObjectIds:[objectId],objectIds:[objectId],sourceBreakIds:[],lineIds:['one','two','three'],lineOwners:['one','two','three'].map(lineId=>({lineId,objectId})),
    sourceLedger:[{sourceId:'node',objectId,kind:'text',text:'ab',ranges:[{start:0,end:2}],whiteSpace:'normal',style}],
    observations:[{sourceId:'node',objectId,start:0,end:1,lineId:'one',kind:'visible',text:'a',style},{sourceId:'node',objectId,start:1,end:2,lineId:'three',kind:'visible',text:'b',style}],
    tokens:[{kind:'copy',objectId,sources:[{sourceId:'node',start:0,end:1}],text:'a',style,lineId:'one'},{kind:'layout-break',objectId,sources:[],text:'\n',style,lineId:'one',nextLineId:'two'},{kind:'layout-break',objectId,sources:[],text:'\n',style,lineId:'two',nextLineId:'three'},{kind:'copy',objectId,sources:[{sourceId:'node',start:1,end:2}],text:'b',style,lineId:'three'}]};
  const result=validateDisplayMapping(input);assert.equal(result.valid,false);assert.ok(result.errors.length);
});

test('an empty line cannot be claimed without a source break or observed text',()=>{
  const input={sourceObjectIds:['text'],objectIds:['text'],sourceBreakIds:[],lineIds:['empty'],lineOwners:[{lineId:'empty',objectId:'text'}],sourceLedger:[],observations:[],tokens:[]};
  const result=validateDisplayMapping(input);assert.equal(result.valid,false);assert.ok(result.errors.length);
});

test('normal whitespace cannot reappear at a soft-wrapped line boundary',()=>{
  const style={fontFamily:'Arial',fontSize:20,fontWeight:400},objectId='text';
  const input={sourceObjectIds:[objectId],objectIds:[objectId],sourceBreakIds:[],lineIds:['one','two'],lineOwners:[{lineId:'one',objectId},{lineId:'two',objectId}],
    sourceLedger:[{sourceId:'left',objectId,kind:'text',text:'a',ranges:[{start:0,end:1}],whiteSpace:'normal',style},{sourceId:'right',objectId,kind:'text',text:' b',ranges:[{start:0,end:2}],whiteSpace:'normal',style}],
    observations:[{sourceId:'left',objectId,start:0,end:1,lineId:'one',kind:'visible',text:'a',style},{sourceId:'right',objectId,start:0,end:2,lineId:'two',kind:'visible',text:' b',style}],
    tokens:[{kind:'copy',objectId,sources:[{sourceId:'left',start:0,end:1}],text:'a',style,lineId:'one'},{kind:'layout-break',objectId,sources:[],text:'\n',style,lineId:'one',nextLineId:'two'},{kind:'copy',objectId,sources:[{sourceId:'right',start:0,end:2}],text:' b',style,lineId:'two'}]};
  const result=validateDisplayMapping(input);assert.equal(result.valid,false);assert.ok(result.errors.length);
});

test('a combining grapheme crossing source nodes cannot be separated by a measured line break',()=>{
  const style={fontFamily:'Arial',fontSize:20,fontWeight:400};
  const input={sourceObjectIds:['text'],objectIds:['text'],sourceBreakIds:[],lineIds:['one','two'],lineOwners:[{lineId:'one',objectId:'text'},{lineId:'two',objectId:'text'}],
    sourceLedger:[{sourceId:'base',kind:'text',text:'e',ranges:[{start:0,end:1}],whiteSpace:'normal',style},{sourceId:'mark',kind:'text',text:'\u0301',ranges:[{start:0,end:1}],whiteSpace:'normal',style}],
    observations:[{sourceId:'base',start:0,end:1,lineId:'one',kind:'visible',text:'e',style},{sourceId:'mark',start:0,end:1,lineId:'two',kind:'visible',text:'\u0301',style}],
    tokens:[{kind:'copy',sources:[{sourceId:'base',start:0,end:1}],text:'e',style,lineId:'one'},{kind:'layout-break',sources:[],text:'\n',style,lineId:'one',nextLineId:'two'},{kind:'copy',sources:[{sourceId:'mark',start:0,end:1}],text:'\u0301',style,lineId:'two'}]};
  for(const key of ['sourceLedger','observations','tokens'])input[key]=input[key].map(item=>({...item,objectId:'text'}));
  const result=validateDisplayMapping(input);assert.equal(result.valid,false);assert.ok(result.errors.length);
});

test('a combining grapheme across differently styled source nodes stays intact on one line',()=>{
  const normal={fontFamily:'Arial',fontSize:20,fontWeight:400},bold={...normal,fontWeight:700};
  const input={sourceObjectIds:['text'],objectIds:['text'],sourceBreakIds:[],lineIds:['one'],lineOwners:[{lineId:'one',objectId:'text'}],
    sourceLedger:[{sourceId:'base',objectId:'text',kind:'text',text:'e',ranges:[{start:0,end:1}],whiteSpace:'normal',style:normal},{sourceId:'mark',objectId:'text',kind:'text',text:'\u0301',ranges:[{start:0,end:1}],whiteSpace:'normal',style:bold}],
    observations:[{sourceId:'base',objectId:'text',start:0,end:1,lineId:'one',kind:'visible',text:'e',style:normal},{sourceId:'mark',objectId:'text',start:0,end:1,lineId:'one',kind:'visible',text:'\u0301',style:bold}],
    tokens:[{kind:'copy',objectId:'text',sources:[{sourceId:'base',start:0,end:1}],text:'e',style:normal,lineId:'one'},{kind:'copy',objectId:'text',sources:[{sourceId:'mark',start:0,end:1}],text:'\u0301',style:bold,lineId:'one'}]};
  const result=validateDisplayMapping(input);assert.equal(result.valid,true,JSON.stringify(result.errors));
});

test('separate paragraph ownership is preserved and a break cannot bridge unrelated objects',()=>{
  const style={fontFamily:'Arial',fontSize:20,fontWeight:400},ids=['first','second'];
  const input={sourceObjectIds:ids,objectIds:ids,sourceBreakIds:[],lineIds:['line-a','line-b'],lineOwners:[{lineId:'line-a',objectId:'first'},{lineId:'line-b',objectId:'second'}],
    sourceLedger:[{sourceId:'a',objectId:'first',kind:'text',text:'A',ranges:[{start:0,end:1}],whiteSpace:'normal',style},{sourceId:'b',objectId:'second',kind:'text',text:'B',ranges:[{start:0,end:1}],whiteSpace:'normal',style}],
    observations:[{sourceId:'a',objectId:'first',start:0,end:1,lineId:'line-a',kind:'visible',text:'A',style},{sourceId:'b',objectId:'second',start:0,end:1,lineId:'line-b',kind:'visible',text:'B',style}],
    tokens:[{kind:'copy',objectId:'first',sources:[{sourceId:'a',start:0,end:1}],text:'A',style,lineId:'line-a'},{kind:'copy',objectId:'second',sources:[{sourceId:'b',start:0,end:1}],text:'B',style,lineId:'line-b'}]};
  assert.equal(validateDisplayMapping(input).valid,true,JSON.stringify(validateDisplayMapping(input).errors));
  const wrongLine=structuredClone(input);wrongLine.tokens[1].lineId='line-a';
  const wrongOwner=structuredClone(input);wrongOwner.observations[1].objectId='first';
  const bridged=structuredClone(input);bridged.tokens.splice(1,0,{kind:'layout-break',objectId:'first',sources:[],text:'\n',style,lineId:'line-a',nextLineId:'line-b'});
  for(const invalid of [wrongLine,wrongOwner,bridged]){const result=validateDisplayMapping(invalid);assert.equal(result.valid,false);assert.ok(result.errors.length);}
});

test('a combining grapheme cannot be separated into different objects without proven block separation',()=>{
  const style={fontFamily:'Arial',fontSize:20,fontWeight:400};
  const input={sourceObjectIds:['first','second'],objectIds:['first','second'],sourceBreakIds:[],lineIds:['line-a','line-b'],lineOwners:[{lineId:'line-a',objectId:'first'},{lineId:'line-b',objectId:'second'}],
    sourceLedger:[{sourceId:'a',objectId:'first',kind:'text',text:'e',ranges:[{start:0,end:1}],whiteSpace:'normal',style},{sourceId:'b',objectId:'second',kind:'text',text:'\u0301',ranges:[{start:0,end:1}],whiteSpace:'normal',style}],
    observations:[{sourceId:'a',objectId:'first',start:0,end:1,lineId:'line-a',kind:'visible',text:'e',style},{sourceId:'b',objectId:'second',start:0,end:1,lineId:'line-b',kind:'visible',text:'\u0301',style}],
    tokens:[{kind:'copy',objectId:'first',sources:[{sourceId:'a',start:0,end:1}],text:'e',style,lineId:'line-a'},{kind:'copy',objectId:'second',sources:[{sourceId:'b',start:0,end:1}],text:'\u0301',style,lineId:'line-b'}]};
  const result=validateDisplayMapping(input);assert.equal(result.valid,false);assert.ok(result.errors.length);
});
