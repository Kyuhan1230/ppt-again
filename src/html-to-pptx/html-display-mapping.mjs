import {isDeepStrictEqual} from 'node:util';
import {validateSourceCoverage} from './html-pagination.mjs';

const MAX_SOURCES=10000,MAX_SPANS=200000,MAX_TEXT=2000000;
const record=value=>value!==null&&typeof value==='object'&&!Array.isArray(value);
const id=value=>typeof value==='string'&&value.length>0&&value.length<=1024;
const range=value=>record(value)&&id(value.sourceId)&&Number.isSafeInteger(value.start)&&Number.isSafeInteger(value.end)&&value.start>=0&&value.end>value.start;
const key=value=>JSON.stringify([value.sourceId,value.start,value.end]);
const whitespace=/^[\t\n\f\r ]+$/;
const collapsible=(source,raw)=>['normal','nowrap'].includes(source.whiteSpace)||source.whiteSpace==='pre-line'&&!/[\r\n]/.test(raw);
const validStyle=value=>{
  if(!record(value))return false;
  const seen=new WeakSet(),pending=[[value,0]];let nodes=0;
  while(pending.length){
    const [item,depth]=pending.pop();if(++nodes>512||depth>12)return false;
    if(item===null||typeof item==='boolean')continue;
    if(typeof item==='number'){if(!Number.isFinite(item))return false;continue;}
    if(typeof item==='string'){if(item.length>8192)return false;continue;}
    if(typeof item!=='object'||seen.has(item))return false;seen.add(item);
    if(!Array.isArray(item)&&Object.getPrototypeOf(item)!==Object.prototype&&Object.getPrototypeOf(item)!==null)return false;
    const keys=Object.keys(item);if(keys.length>128)return false;
    for(const key of keys){if(key.length>256)return false;pending.push([item[key],depth+1]);}
  }
  return true;
};

// This validates independently observed data. The reader must still prove that
// its source and rendered-observation ledgers describe the selected DOM.
export function validateDisplayMapping(input){
  try{return validateDisplayMappingChecked(input);}catch{return {valid:false,errors:['표시 대응 자료를 확인할 수 없습니다.']};}
}
function validateDisplayMappingChecked(input){
  const errors=[],fail=message=>{errors.push(message);return {valid:false,errors};};
  if(!record(input)||!['sourceLedger','observations','tokens','sourceObjectIds','objectIds','sourceBreakIds','lineIds','lineOwners'].every(name=>Array.isArray(input[name])))return fail('원본·표시·객체 장부가 필요합니다.');
  if(input.sourceLedger.length>MAX_SOURCES||input.observations.length>MAX_SPANS||input.tokens.length>MAX_SPANS||input.lineIds.length>MAX_SPANS||input.sourceBreakIds.length>MAX_SOURCES||input.sourceObjectIds.length>MAX_SOURCES||input.objectIds.length>MAX_SOURCES)return fail('표시 대응 자료가 지원 한도를 초과합니다.');
  const unique=items=>items.every(id)&&new Set(items).size===items.length;
  const objectIds=new Set(input.objectIds);
  if(!unique(input.sourceObjectIds)||!unique(input.objectIds)||input.sourceObjectIds.length!==input.objectIds.length||input.sourceObjectIds.some(value=>!objectIds.has(value)))return fail('원본 객체가 누락되었거나 중복되었습니다.');
  if(!unique(input.sourceBreakIds)||!unique(input.lineIds))return fail('줄 또는 줄바꿈 ID가 올바르지 않습니다.');
  if(input.lineOwners.length!==input.lineIds.length||input.lineOwners.some((entry,index)=>!record(entry)||entry.lineId!==input.lineIds[index]||!objectIds.has(entry.objectId)))return fail('측정 줄의 원본 객체 관계가 올바르지 않습니다.');
  const lineOwners=new Map(input.lineOwners.map(entry=>[entry.lineId,entry.objectId]));
  const sources=new Map(),expected=[],breaks=new Set();let length=0,ledgerText='';
  for(const source of input.sourceLedger){
    if(!record(source)||!id(source.sourceId)||sources.has(source.sourceId)||!objectIds.has(source.objectId)||!['text','br'].includes(source.kind)||typeof source.text!=='string'||!Array.isArray(source.ranges)||!['normal','nowrap','pre','pre-wrap','pre-line','break-spaces'].includes(source.whiteSpace)||!validStyle(source.style)||(length+=source.text.length)>MAX_TEXT)return fail('원본 텍스트 장부가 올바르지 않습니다.');
    if(source.kind==='br'&&(source.text!=='\n'||source.ranges.length!==1||source.ranges[0]?.start!==0||source.ranges[0]?.end!==1))return fail('원본 줄바꿈 노드가 올바르지 않습니다.');
    if(source.kind==='br')breaks.add(source.sourceId);
    const boundaries=new Set([0,source.text.length]);for(const entry of new Intl.Segmenter('und',{granularity:'grapheme'}).segment(source.text))boundaries.add(entry.index);
    let cursor=0;const base=ledgerText.length;
    for(const span of source.ranges){const item={sourceId:source.sourceId,start:span?.start,end:span?.end};if(expected.length>=MAX_SPANS||!range(item)||item.start!==cursor||item.end>source.text.length||!boundaries.has(item.start)||!boundaries.has(item.end))return fail('원본 글자 범위가 누락되었거나 올바르지 않습니다.');expected.push(item);cursor=item.end;}
    if(cursor!==source.text.length)return fail('원본 글자 범위가 누락되었습니다.');
    ledgerText+=source.text;sources.set(source.sourceId,{...source,boundaries,base});
  }
  if(input.sourceBreakIds.length!==breaks.size||input.sourceBreakIds.some(value=>!breaks.has(value)))return fail('원본 BR 장부가 누락되었거나 중복되었습니다.');
  const lineIndex=new Map(input.lineIds.map((value,index)=>[value,index])),observed=new Map();
  for(const item of input.observations){
    if(!range(item)||!lineIndex.has(item.lineId)||!['visible','collapsed','suppressed','hard-break','text-break'].includes(item.kind)||typeof item.text!=='string'||!validStyle(item.style))return fail('표시 관측 범위가 올바르지 않습니다.');
    const source=sources.get(item.sourceId),k=key(item);
    if(!source||item.objectId!==source.objectId||lineOwners.get(item.lineId)!==source.objectId||item.end>source.text.length||!source.boundaries.has(item.start)||!source.boundaries.has(item.end)||observed.has(k)||!isDeepStrictEqual(item.style,source.style))return fail('표시 관측과 원본 객체·범위 또는 서식이 다릅니다.');
    const raw=source.text.slice(item.start,item.end);
    if(item.kind==='hard-break'){
      if(source.kind!=='br'||item.start!==0||item.end!==1||item.text!=='\n')return fail('BR 관측이 원본 줄바꿈과 다릅니다.');
    }else if(source.kind!=='text')return fail('BR을 일반 글자로 처리할 수 없습니다.');
    else if(item.kind==='text-break'){
      if(!['pre','pre-wrap','pre-line','break-spaces'].includes(source.whiteSpace)||!/^\r\n$|^[\r\n]$/.test(raw)||item.text!=='\n')return fail('원문 개행의 표시를 확인할 수 없습니다.');
    }
    else if(item.kind==='visible'){
      if(/[\r\n\t]/.test(item.text)||(['normal','nowrap','pre-line'].includes(source.whiteSpace)&&/ {2}/.test(item.text))||item.text!==raw&&!(whitespace.test(raw)&&collapsible(source,raw)&&item.text===' '))return fail('표시 글자가 원본 관측과 다릅니다.');
    }else if(!whitespace.test(raw)||!collapsible(source,raw)||item.text!=='')return fail('비표시 또는 축약 공백의 근거가 없습니다.');
    observed.set(k,item);
  }
  if(!validateSourceCoverage(expected,input.observations).valid)return fail('표시 관측에서 원본 글자가 누락되었거나 중복되었습니다.');
  if(input.tokens.length&&!input.lineIds.length)return fail('측정된 줄이 없습니다.');
  const visibleCounts=new Map(),lineEvidence=new Set();for(const token of input.tokens){if(['copy','collapsed-space'].includes(token?.kind))visibleCounts.set(token.lineId,(visibleCounts.get(token.lineId)??0)+1);if(['copy','collapsed-space','hard-break','text-break'].includes(token?.kind))lineEvidence.add(token.lineId);}
  if(input.lineIds.some(lineId=>!lineEvidence.has(lineId)))return fail('원본 또는 실제 표시 근거가 없는 빈 줄입니다.');
  const seenVisible=new Map(),seen=new Set(),consumed=[],graphemes=new Set([0,ledgerText.length]);
  for(const entry of new Intl.Segmenter('und',{granularity:'grapheme'}).segment(ledgerText))graphemes.add(entry.index);
  let currentLine=0,cursor=0,generatedBreaks=0,lastObservedStyle=null,previousVisibleSpace=null;
  for(const token of input.tokens){
    if(record(token)&&token.lineId===input.lineIds[currentLine+1]&&lineOwners.get(input.lineIds[currentLine])!==lineOwners.get(token.lineId)){
      if(!graphemes.has(cursor))return fail('하나의 글자 조합을 서로 다른 객체로 나눌 수 없습니다.');
      currentLine++;lastObservedStyle=null;previousVisibleSpace=null;
    }
    if(!record(token)||!['copy','collapsed-space','suppressed-space','hard-break','text-break','layout-break'].includes(token.kind)||!Array.isArray(token.sources)||token.sources.length>MAX_SPANS-consumed.length||!validStyle(token.style)||typeof token.text!=='string'||token.lineId!==input.lineIds[currentLine]||token.objectId!==lineOwners.get(token.lineId))return fail('표시 토큰의 객체 또는 줄 순서가 올바르지 않습니다.');
    const matches=[];
    for(const span of token.sources){
      if(!range(span))return fail('표시 토큰의 원본 범위가 올바르지 않습니다.');
      const k=key(span),observation=observed.get(k),source=sources.get(span.sourceId);
      if(!observation||seen.has(k)||observation.lineId!==token.lineId||source.objectId!==token.objectId||source.base+span.start!==cursor)return fail('표시 토큰의 관측 범위가 누락·중복되었거나 순서가 다릅니다.');
      seen.add(k);consumed.push(span);matches.push({observation,source,raw:source.text.slice(span.start,span.end)});cursor+=span.end-span.start;lastObservedStyle=observation.style;
    }
    if(token.kind==='copy'){
      const match=matches[0];if(matches.length!==1||match.observation.kind!=='visible'||token.text!==match.raw||token.text!==match.observation.text||!isDeepStrictEqual(token.style,match.observation.style))return fail('일반 글자의 내용 또는 서식이 실제 표시와 다릅니다.');
    }else if(token.kind==='collapsed-space'){
      const visible=matches.filter(m=>m.observation.kind==='visible');
      if(!matches.length||visible.length!==1||visible[0].observation.text!==' '||token.text!==' '||matches.some(m=>!whitespace.test(m.raw)||!collapsible(m.source,m.raw)||!['visible','collapsed'].includes(m.observation.kind)||!isDeepStrictEqual(m.observation.style,visible[0].observation.style)||m.source.whiteSpace!==visible[0].source.whiteSpace)||!isDeepStrictEqual(token.style,visible[0].observation.style))return fail('축약 공백의 표시 또는 서식을 확정할 수 없습니다.');
    }else if(token.kind==='suppressed-space'){
      const before=seenVisible.get(token.lineId)??0,after=(visibleCounts.get(token.lineId)??0)-before;
      if(!matches.length||token.text!==''||matches.some(m=>m.observation.kind!=='suppressed'||!whitespace.test(m.raw)||!collapsible(m.source,m.raw)||!isDeepStrictEqual(token.style,m.observation.style))||before>0&&after>0)return fail('표시되지 않은 공백을 줄 경계에서 확인할 수 없습니다.');
    }else if(token.kind==='hard-break'){
      const match=matches[0];if(matches.length!==1||match.observation.kind!=='hard-break'||token.text!=='\n'||!isDeepStrictEqual(token.style,match.observation.style))return fail('BR 토큰이 원본 줄바꿈 관측과 다릅니다.');
    }else if(token.kind==='text-break'){
      const match=matches[0];if(matches.length!==1||match.observation.kind!=='text-break'||token.text!=='\n'||!isDeepStrictEqual(token.style,match.observation.style))return fail('원문 개행 토큰이 실제 표시와 다릅니다.');
    }else if(token.sources.length||token.text!=='\n'||!lastObservedStyle||!isDeepStrictEqual(token.style,lastObservedStyle)||!lineEvidence.has(token.lineId)||!lineEvidence.has(token.nextLineId))return fail('자동 줄바꿈은 실제 표시된 인접 줄 사이의 별도 표식이어야 합니다.');
    if(['copy','collapsed-space'].includes(token.kind)){
      const collapses=['normal','nowrap','pre-line'].includes(matches[0].source.whiteSpace);
      const before=seenVisible.get(token.lineId)??0,after=(visibleCounts.get(token.lineId)??0)-before-1;
      if(collapses&&(token.text.startsWith(' ')&&before===0||token.text.endsWith(' ')&&after===0))return fail('줄 시작 또는 끝의 축약 공백을 표시 글자로 확정할 수 없습니다.');
      if(previousVisibleSpace!==null&&token.text.startsWith(' ')&&(previousVisibleSpace||collapses))return fail('인접 공백의 실제 축약을 확인할 수 없습니다.');
      previousVisibleSpace=token.text.endsWith(' ')?collapses:null;
      seenVisible.set(token.lineId,(seenVisible.get(token.lineId)??0)+1);
    }
    if(['hard-break','text-break','layout-break'].includes(token.kind)){
      if(token.nextLineId!==input.lineIds[currentLine+1]||lineOwners.get(token.nextLineId)!==token.objectId||!graphemes.has(cursor))return fail('줄 경계가 같은 객체의 다음 줄 또는 글자 조합과 다릅니다.');
      if(token.kind==='layout-break')generatedBreaks++;currentLine++;previousVisibleSpace=null;
    }
  }
  if(currentLine!==Math.max(0,input.lineIds.length-1)||seen.size!==observed.size||cursor!==ledgerText.length||!validateSourceCoverage(expected,consumed).valid)return fail('출력에서 원본 글자·줄·BR이 누락되었거나 중복되었습니다.');
  return {valid:true,errors:[],generatedBreaks};
}
