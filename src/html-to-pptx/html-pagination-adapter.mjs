import {isDeepStrictEqual} from 'node:util';
import {planPagination,validateSourceCoverage} from './html-pagination.mjs';
import {validateDisplayMapping} from './html-display-mapping.mjs';

const fail=message=>({pages:[],diagnostics:[{severity:'error',code:'PAGINATION_INVALID',message}]});
const unsupported=message=>({pages:[],diagnostics:[{severity:'error',code:'PAGINATION_UNSUPPORTED',message}]});
const key=value=>JSON.stringify([value.sourceId,value.start,value.end]);
const sameBox=(a,b)=>a&&b&&['x','y','w','h'].every(field=>Number.isFinite(a[field])&&Number.isFinite(b[field])&&Math.abs(a[field]-b[field])<.000001);
const idsMatch=(a,b)=>{if(!Array.isArray(a)||!Array.isArray(b)||a.length!==b.length)return false;const left=new Set(a),right=new Set(b);return left.size===a.length&&right.size===b.length&&a.every(id=>right.has(id));};
const sameRuns=(a,b)=>{
  if(!Array.isArray(a)||!Array.isArray(b)||a.some(run=>typeof run?.text!=='string'||!run.style)||b.some(run=>typeof run?.text!=='string'||!run.style))return false;
  a=a.filter(run=>run.text);b=b.filter(run=>run.text);let i=0,j=0,x=0,y=0;
  while(i<a.length&&j<b.length){
    if(!isDeepStrictEqual(a[i].style,b[j].style))return false;
    const length=Math.min(a[i].text.length-x,b[j].text.length-y);
    if(a[i].text.slice(x,x+length)!==b[j].text.slice(y,y+length))return false;
    x+=length;y+=length;if(x===a[i].text.length){i++;x=0;}if(y===b[j].text.length){j++;y=0;}
  }
  return i===a.length&&j===b.length;
};

// Pure bridge between independently observed source/display data and the
// page planner. It does not create a validated deck or claim PPT export.
export function adaptMappedPagination(input){
  try{return adaptMappedPaginationChecked(input);}catch{return fail('분할 자료를 확인할 수 없습니다.');}
}
function adaptMappedPaginationChecked(input){
  if(!input||typeof input!=='object'||!Array.isArray(input.objects)||!Array.isArray(input.lineMeasurements)||!input.displayMapping)return fail('분할 대상과 독립 측정 장부가 필요합니다.');
  const mapping=input.displayMapping,validation=validateDisplayMapping(mapping);
  if(!validation.valid)return fail('원본과 표시 텍스트의 대응을 검증할 수 없습니다.');
  if(!idsMatch(input.sourceObjectIds,mapping.sourceObjectIds)||!idsMatch(input.objects.map(object=>object?.id),mapping.objectIds))return fail('원본 객체 장부와 분할 대상이 다릅니다.');
  if(input.objects.some(object=>object.kind==='text'&&object.content?.paragraphs))return unsupported('문단별 번호·간격의 분할 근거를 확인할 수 없습니다.');
  const objectMap=new Map(input.objects.map(object=>[object.id,object]));
  const sourceMap=new Map(mapping.sourceLedger.map(source=>[source.sourceId,source]));
  const observationMap=new Map(mapping.observations.map(observation=>[key(observation),observation]));
  const tokensByLine=new Map(),tokensByObject=new Map(),linesByObject=new Map();
  for(const [index,token] of mapping.tokens.entries()){
    if(!tokensByLine.has(token.lineId))tokensByLine.set(token.lineId,[]);tokensByLine.get(token.lineId).push({index,token});
    if(!tokensByObject.has(token.objectId))tokensByObject.set(token.objectId,[]);tokensByObject.get(token.objectId).push(token);
  }
  const semanticRunsByObject=new Map();
  for(const [objectId,items] of tokensByObject){
    const next=[],runs=[];let following=null,previous=null,suppressedAdded=false;
    for(let i=items.length-1;i>=0;i--){next[i]=following;if(!['layout-break','suppressed-space'].includes(items[i].kind))following=items[i];}
    for(const [i,token] of items.entries()){
      if(token.kind==='layout-break')continue;
      if(token.kind==='suppressed-space'){
        const after=next[i],word=t=>t&&['copy','collapsed-space'].includes(t.kind)&&t.text;
        if(!suppressedAdded&&word(previous)&&word(after)&&!previous.text.endsWith(' ')&&!after.text.startsWith(' ')){runs.push({text:' ',style:token.style});suppressedAdded=true;}
      }else{if(token.text)runs.push({text:token.text,style:token.style});previous=token;suppressedAdded=false;}
    }
    semanticRunsByObject.set(objectId,runs);
  }
  for(const owner of mapping.lineOwners){if(!linesByObject.has(owner.objectId))linesByObject.set(owner.objectId,[]);linesByObject.get(owner.objectId).push(owner.lineId);}
  const measured=new Map();
  for(const item of input.lineMeasurements){
    if(!item||typeof item.lineId!=='string'||measured.has(item.lineId))return fail('독립 줄 측정이 누락되었거나 중복되었습니다.');
    measured.set(item.lineId,item);
  }
  if(measured.size!==mapping.lineIds.length)return fail('독립 줄 측정 수가 일치하지 않습니다.');
  const seenLines=new Set(),expectedByObject=new Map();
  for(const source of mapping.sourceLedger){if(!expectedByObject.has(source.objectId))expectedByObject.set(source.objectId,[]);for(const span of source.ranges)expectedByObject.get(source.objectId).push({sourceId:source.sourceId,start:span.start,end:span.end});}
  const sourcesByObject=new Map();
  for(const [kind,sources] of [['text',mapping.sourceLedger],['table',input.tableSourceLedger??[]]])for(const source of sources){
    if(objectMap.get(source.objectId)?.kind!==kind)return fail('원본 글자의 소유 객체를 확인할 수 없습니다.');
    if(!sourcesByObject.has(source.objectId))sourcesByObject.set(source.objectId,[]);
    sourcesByObject.get(source.objectId).push(source);
  }
  const plannerLedger=input.objects.flatMap(object=>sourcesByObject.get(object.id)??[]);
  const prepared=[];
  for(const object of input.objects){
    if(!object||typeof object!=='object')return fail('분할 객체가 올바르지 않습니다.');
    if(object.kind!=='text'){
      if(object.kind==='table'){
        const expected=sourcesByObject.get(object.id)??[];
        if(!Array.isArray(object.sourceRanges)||!Array.isArray(object.content?.cells)||!validateSourceCoverage(expected.flatMap(source=>source.ranges.map(span=>({...span,sourceId:source.sourceId}))),object.sourceRanges).valid)return unsupported('표 셀의 독립 원문 기록을 확인할 수 없습니다.');
        const cells=new Map(expected.map(source=>[source.sourceId,source]));
        if(object.content.cells.some(cell=>cell.runs.some(run=>{
          if(!run.text)return false;
          const source=cells.get(run.sourceId);
          return !source||source.row!==cell.row||source.col!==cell.col;
        })))return fail('표 글자가 원본 셀을 벗어났습니다.');
      }
      prepared.push(structuredClone(object));continue;
    }
    const semanticRuns=semanticRunsByObject.get(object.id)??[];
    if(typeof object.content?.text!=='string'||!sameRuns(object.content.runs,semanticRuns)||object.content.text!==semanticRuns.map(run=>run.text).join(''))return fail('문단의 현재 내용·서식과 독립 표시 관측이 다릅니다.');
    if(!Array.isArray(object.lines)||!object.lines.length)return fail('문단의 실제 줄 측정이 필요합니다.');
    const expectedOrder=linesByObject.get(object.id)??[];
    if(!isDeepStrictEqual(object.lines.map(line=>line?.lineId),expectedOrder))return fail('문단의 실제 줄 순서가 다릅니다.');
    const lines=[];
    for(const line of object.lines){
      const reference=measured.get(line.lineId);
      if(seenLines.has(line.lineId)||!reference||reference.objectId!==object.id||!sameBox(line.boxCssPx,reference.boxCssPx)||!sameBox(line.paintBoxCssPx,reference.paintBoxCssPx)||!Number.isFinite(line.baseline)||!Number.isFinite(reference.baseline)||Math.abs(line.baseline-reference.baseline)>.000001)return fail('문단의 줄 위치 또는 기준선이 독립 측정과 다릅니다.');
      seenLines.add(line.lineId);const runs=[];
      for(const {token} of tokensByLine.get(line.lineId)??[])for(const span of token.sources){
        const source=sourceMap.get(span.sourceId),observed=observationMap.get(key(span));
        if(!source||!observed||source.objectId!==object.id)return fail('줄의 원본 범위를 확인할 수 없습니다.');
        runs.push({...span,text:source.text.slice(span.start,span.end),style:structuredClone(observed.style)});
      }
      lines.push({...structuredClone(line),runs});
    }
    prepared.push({...structuredClone(object),sourceRanges:expectedByObject.get(object.id)??[],lines});
  }
  if(seenLines.size!==mapping.lineIds.length)return fail('원본 객체와 연결되지 않은 측정 줄이 있습니다.');
  const planned=planPagination({mode:input.mode,documentWidth:input.documentWidth,documentHeight:input.documentHeight,pageHeight:input.pageHeight,sourceObjectIds:input.sourceObjectIds,sourceLedger:plannerLedger,objects:prepared});
  if(planned.diagnostics.length)return planned;
  const linePage=new Map();
  for(const page of planned.pages)for(const part of page.parts)if(part.kind==='text')for(const line of part.lines){if(linePage.has(line.lineId))return fail('측정 줄이 여러 장표에 중복되었습니다.');linePage.set(line.lineId,page.index);}
  if(linePage.size!==mapping.lineIds.length)return fail('측정 줄이 장표에서 누락되었습니다.');
  const fragmentCounts=new Map();for(const page of planned.pages)for(const part of page.parts)fragmentCounts.set(part.objectId,(fragmentCounts.get(part.objectId)??0)+1);
  const assignedTokens=new Set(),usedIds=new Set(input.objects.map(object=>object.id)),pages=[];
  for(const page of planned.pages){
    const output={index:page.index,sourceStart:page.sourceStart,sourceEnd:page.sourceEnd,pageHeight:page.pageHeight,parts:[]};
    for(const part of page.parts){
      const object=objectMap.get(part.objectId);if(!object)return fail('분할 객체의 원본을 찾을 수 없습니다.');
      const base=`${object.id}::page-${page.index+1}`;let fragmentId=base,suffix=1;while(usedIds.has(fragmentId))fragmentId=`${base}-${suffix++}`;usedIds.add(fragmentId);
      const top=Math.max(object.boxCssPx.y,page.sourceStart),bottom=Math.min(object.boxCssPx.y+object.boxCssPx.h,page.sourceEnd);
      const boxCssPx={...object.boxCssPx,y:top-page.sourceStart,h:Math.max(0,bottom-top)};
      const fragment={objectId:object.id,fragmentId,kind:object.kind,boxCssPx,sourceRanges:[],tokens:[],parentObjectId:object.parentId,sourceFragmentCount:fragmentCounts.get(object.id),originalBoxCssPx:structuredClone(object.boxCssPx),originalPaintBoxCssPx:structuredClone(object.paintBoxCssPx),...(object.sourceRef?{originalSourceRef:structuredClone(object.sourceRef)}:{}),...(object.style?{originalStyle:structuredClone(object.style)}:{}),...(Number.isInteger(object.zOrder)?{originalZOrder:object.zOrder}:{}),...(object.pptxOrigin?{originalPptxOrigin:structuredClone(object.pptxOrigin)}:{}),...(object.kind!=='text'&&object.content?{originalContent:structuredClone(object.content)}:{})};
      if(part.kind==='text'){
        const runs=[],lines=[];
        for(const line of part.lines){
          const local={lineId:line.lineId,boxCssPx:{...line.boxCssPx,y:line.boxCssPx.y-page.sourceStart},paintBoxCssPx:{...line.paintBoxCssPx,y:line.paintBoxCssPx.y-page.sourceStart},baseline:line.baseline-page.sourceStart};lines.push(local);
          for(const {index,token} of tokensByLine.get(line.lineId)??[]){
            if(assignedTokens.has(index))return fail('표시 토큰이 여러 장표에 중복되었습니다.');assignedTokens.add(index);
            const atPageBoundary=['layout-break','hard-break','text-break'].includes(token.kind)&&linePage.get(token.nextLineId)!==page.index;
            const recorded={...structuredClone(token),tokenIndex:index,...(atPageBoundary?{handledBy:'page-boundary'}:{})};fragment.tokens.push(recorded);
            fragment.sourceRanges.push(...structuredClone(token.sources));
            if(token.text&&!atPageBoundary)runs.push({text:token.text,style:structuredClone(token.style)});
          }
        }
        if(!runs.some(run=>/[^\r\n]/.test(run.text)))return unsupported('글자 없이 줄바꿈만 있는 장표 조각은 안전하게 만들 수 없습니다.');
        fragment.lineStart=part.lineStart;fragment.lineEnd=part.lineEnd;fragment.lines=lines;fragment.baseline=lines[0]?.baseline;
        fragment.placementEvidence={boxTop:boxCssPx.y,firstLineTop:lines[0].boxCssPx.y,firstBaseline:lines[0].baseline,firstLineOffset:lines[0].boxCssPx.y-boxCssPx.y,requiresWriterVerification:true};
        fragment.content={text:runs.map(run=>run.text).join(''),runs};
      }else if(part.kind==='table'){
        const rowOffset=object.content.rowHeightsCssPx.slice(0,part.rowStart).reduce((sum,height)=>sum+height,0);
        const height=part.content.rowHeightsCssPx.reduce((sum,value)=>sum+value,0);
        fragment.boxCssPx={...object.boxCssPx,y:object.boxCssPx.y+rowOffset-page.sourceStart,h:height};
        fragment.content=part.content;
        fragment.sourceCells=part.sourceCells;
        fragment.rowStart=part.rowStart;
        fragment.rowEnd=part.rowEnd;
        fragment.sourceRanges=part.content.cells.flatMap(cell=>cell.runs.filter(run=>run.text).map(({sourceId,start,end})=>({sourceId,start,end})));
        if(part.tableBoundaryPolicy)fragment.tableBoundaryPolicy=part.tableBoundaryPolicy;
      }
      output.parts.push(fragment);
    }
    pages.push(output);
  }
  if(assignedTokens.size!==mapping.tokens.length)return fail('표시 토큰이 장표에서 누락되었습니다.');
  const expected=plannerLedger.flatMap(source=>source.ranges.map(span=>({sourceId:source.sourceId,start:span.start,end:span.end})));
  if(!validateSourceCoverage(expected,pages.flatMap(page=>page.parts.flatMap(part=>part.sourceRanges))).valid)return fail('분할 후 원본 글자 범위가 누락되었거나 중복되었습니다.');
  // The current PPT writer does not position text using measured baselines.
  // The consumer must construct/validate a deck and verify box, padding and
  // vertical alignment against the emitted lines before claiming an export.
  return {pages,diagnostics:[],exportReady:false,remainingChecks:['text-placement','model-validation','ppt-output']};
}
