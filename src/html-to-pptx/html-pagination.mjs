import {validateTable} from './tables.mjs';

const LIMIT_OBJECTS=10000,LIMIT_LINES=50000,LIMIT_TEXT=2000000,LIMIT_PAGES=200;
const record=value=>value!==null&&typeof value==='object'&&!Array.isArray(value);
const identifier=value=>typeof value==='string'&&value.length>0&&value.length<=1024;
const failure=(code,message,reasonKey)=>({pages:[],diagnostics:[{severity:'error',code,message,...(reasonKey?{reasonKey}:{})}]});
const rangeValid=r=>record(r)&&identifier(r.sourceId)&&Number.isSafeInteger(r.start)&&Number.isSafeInteger(r.end)&&r.start>=0&&r.end>r.start;

// Intervals use UTF-16 offsets, not the text's value: the same words may occur
// repeatedly. Adjacent intervals may be combined, overlapping ones may not.
export function validateSourceCoverage(expected,actual){
  const errors=[];
  const collect=(items,label)=>{
    const groups=new Map();
    if(!Array.isArray(items)||items.length>LIMIT_LINES*10){errors.push(`${label}: invalid range list`);return groups;}
    for(const r of items){
      if(!rangeValid(r)){errors.push(`${label}: invalid range`);continue;}
      if(!groups.has(r.sourceId))groups.set(r.sourceId,[]);
      groups.get(r.sourceId).push([r.start,r.end]);
    }
    for(const [id,ranges] of groups){
      ranges.sort((a,b)=>a[0]-b[0]||a[1]-b[1]);const merged=[];
      for(const current of ranges){const last=merged.at(-1);if(last&&current[0]<last[1])errors.push(`${label}: overlapping ranges`);else if(last&&current[0]===last[1])last[1]=current[1];else merged.push([...current]);}
      groups.set(id,merged);
    }
    return groups;
  };
  const a=collect(expected,'expected'),b=collect(actual,'actual');
  for(const id of new Set([...a.keys(),...b.keys()]))if(JSON.stringify(a.get(id)??[])!==JSON.stringify(b.get(id)??[]))errors.push('source coverage differs');
  return {valid:errors.length===0,errors};
}

// A table boundary may coincide with a merged cell's end, but may never pass
// through a merged cell. Native cells and their formatting are copied intact.
export function splitTableRows(content,start,end){
  try{return splitTableRowsChecked(content,start,end);}catch{return {content:null,diagnostics:[{severity:'error',code:'PAGINATION_INVALID',message:'표 분할 자료를 확인할 수 없습니다.'}]};}
}
function splitTableRowsChecked(content,start,end){
  const diagnostics=[],bad=message=>diagnostics.push({severity:'error',code:'PAGINATION_UNSUPPORTED',message});
  if(!validateTable(content).valid||!Number.isInteger(start)||!Number.isInteger(end)||start<0||end>content?.rows||start>=end){bad('표 행 범위 또는 표 격자가 올바르지 않습니다.');return {content:null,diagnostics};}
  if(content.cells.some(c=>c.row<start&&c.row+c.rowSpan>start||c.row<end&&c.row+c.rowSpan>end)){bad('병합된 셀을 가로지르는 위치에서는 표를 나눌 수 없습니다.');return {content:null,diagnostics};}
  const result=structuredClone(content);
  result.rows=end-start;result.rowHeightsCssPx=content.rowHeightsCssPx.slice(start,end);
  result.cells=result.cells.filter(c=>c.row>=start&&c.row<end).map(c=>({...c,row:c.row-start}));
  if(!validateTable(result).valid){bad('나눈 표의 격자를 검증할 수 없습니다.');return {content:null,diagnostics};}
  return {content:result,diagnostics};
}

const validBox=(b,zeroAxis=false)=>record(b)&&['x','y','w','h'].every(k=>Number.isFinite(b[k]))&&b.x>=0&&b.y>=0&&(zeroAxis?b.w>=0&&b.h>=0&&(b.w>0||b.h>0):b.w>0&&b.h>0);
const inside=(a,b)=>a.x>=b.x-.05&&a.y>=b.y-.05&&a.x+a.w<=b.x+b.w+.05&&a.y+a.h<=b.y+b.h+.05;
const union=(a,b)=>{const x=Math.min(a.x,b.x),y=Math.min(a.y,b.y);return {x,y,w:Math.max(a.x+a.w,b.x+b.w)-x,h:Math.max(a.y+a.h,b.y+b.h)-y};};
const transparent=value=>value==='transparent'||typeof value==='string'&&/^rgba\([^)]*[, /]0(?:\.0+)?\)$/.test(value);
const plainTextPaint=style=>record(style)&&transparent(style.fill)&&style.strokeWidth===0&&style.backgroundImage==='none'&&style.opacity===1&&!style.assetId&&!style.shadow&&!style.capture&&(!style.boxShadow||style.boxShadow==='none');

// This is an internal planning contract. The caller must supply an independent
// source ledger and reader-verified splitSafe data; no HTML is trusted here.
export function planPagination(input){
  try{return planPaginationChecked(input);}catch{return failure('PAGINATION_INVALID','분할 자료를 확인할 수 없습니다.');}
}
function planPaginationChecked(input){
  const invalid=message=>failure('PAGINATION_INVALID',message),unsupported=message=>failure('PAGINATION_UNSUPPORTED',message);
  if(input?.mode!=='paginate')return invalid('여러 장으로 나누려면 분할 방식을 명시적으로 선택해 주세요.');
  if(!record(input)||!Number.isFinite(input.documentWidth)||!Number.isFinite(input.documentHeight)||!Number.isFinite(input.pageHeight)||input.documentWidth<=0||input.documentHeight<=0||input.pageHeight<=0||!Array.isArray(input.objects)||input.objects.length>LIMIT_OBJECTS||!Array.isArray(input.sourceObjectIds)||!Array.isArray(input.sourceLedger))return invalid('문서 크기 또는 원본 장부가 올바르지 않습니다.');
  if(input.sourceObjectIds.length!==input.objects.length||input.sourceLedger.length>LIMIT_OBJECTS)return invalid('원본 객체 장부가 올바르지 않습니다.');
  const ids=new Map();
  for(const object of input.objects){
    if(!record(object)||!identifier(object.id)||ids.has(object.id)||!['text','table','image','shape','line','chart','group'].includes(object.kind)||!validBox(object.boxCssPx,object.kind==='line')||!validBox(object.paintBoxCssPx)||!inside(object.boxCssPx,object.paintBoxCssPx)||object.boxCssPx.x+object.boxCssPx.w>input.documentWidth+.000001||object.boxCssPx.y+object.boxCssPx.h>input.documentHeight+.000001||object.paintBoxCssPx.x+object.paintBoxCssPx.w>input.documentWidth+.000001||object.paintBoxCssPx.y+object.paintBoxCssPx.h>input.documentHeight+.000001||!(object.parentId===null||identifier(object.parentId)))return invalid('객체 ID·종류·위치 또는 실제 표시 범위가 올바르지 않습니다.');
    if(object.kind==='line'&&(object.boxCssPx.w===0||object.boxCssPx.h===0)){
      const stroke=object.style?.strokeWidth,box=object.boxCssPx,paint=object.paintBoxCssPx;
      if(!Number.isFinite(stroke)||stroke<=0||box.h===0&&(paint.y>box.y-stroke/2+.000001||paint.y+paint.h<box.y+stroke/2-.000001)||box.w===0&&(paint.x>box.x-stroke/2+.000001||paint.x+paint.w<box.x+stroke/2-.000001))return invalid('선의 실제 두께를 확인할 수 없습니다.');
    }
    ids.set(object.id,object);
  }
  const originalIds=new Set();
  for(const id of input.sourceObjectIds){if(!identifier(id)||originalIds.has(id)||!ids.has(id))return invalid('원본 객체가 누락되었거나 중복되었습니다.');originalIds.add(id);}
  for(const object of input.objects){
    let id=object.parentId,depth=0;const seen=new Set([object.id]);
    while(id!==null){const parent=ids.get(id);if(++depth>64||!parent||parent.kind!=='group'||seen.has(id)||!inside(object.boxCssPx,parent.boxCssPx)||!inside(object.paintBoxCssPx,parent.paintBoxCssPx))return invalid('객체 그룹 관계가 올바르지 않습니다.');seen.add(id);id=parent.parentId;}
  }
  const sources=new Map(),expected=[];let textLength=0,ledgerText='';
  for(const item of input.sourceLedger){
    if(!record(item)||!identifier(item.sourceId)||sources.has(item.sourceId)||typeof item.text!=='string'||(textLength+=item.text.length)>LIMIT_TEXT||!Array.isArray(item.ranges))return invalid('원본 텍스트 장부가 올바르지 않습니다.');
    const boundaries=new Set([0,item.text.length]);for(const entry of new Intl.Segmenter('und',{granularity:'grapheme'}).segment(item.text))boundaries.add(entry.index);
    const spans=[];let previous=0;
    for(const r of item.ranges){
      const range={sourceId:item.sourceId,start:r?.start,end:r?.end};if(!rangeValid(range)||range.end>item.text.length||!boundaries.has(range.start)||!boundaries.has(range.end)||range.start!==previous)return invalid('원본 글자 범위가 올바르지 않습니다.');
      expected.push(range);const last=spans.at(-1);
      if(last&&last.end===range.start)last.end=range.end;else spans.push({start:range.start,end:range.end,base:ledgerText.length});
      ledgerText+=item.text.slice(range.start,range.end);previous=range.end;
    }
    if(previous!==item.text.length)return invalid('원본 텍스트 장부에 누락된 글자 범위가 있습니다.');
    sources.set(item.sourceId,{text:item.text,boundaries,spans});
  }
  if(!validateSourceCoverage(expected,expected).valid)return invalid('원본 텍스트 장부가 중복됩니다.');
  const actual=[],atoms=[],objectAtoms=new Map();let lineCount=0;
  const addAtom=atom=>{atoms.push(atom);if(!objectAtoms.has(atom.object))objectAtoms.set(atom.object,[]);objectAtoms.get(atom.object).push(atom);};
  const checkRun=run=>{
    if(!rangeValid(run)||typeof run.text!=='string'||!record(run.style))return false;
    const source=sources.get(run.sourceId);
    if(!source||run.end>source.text.length||!source.boundaries.has(run.start)||!source.boundaries.has(run.end)||run.text!==source.text.slice(run.start,run.end))return false;
    actual.push(run);return true;
  };
  for(const object of input.objects){
    const box=object.boxCssPx,paint=object.paintBoxCssPx,area=union(box,paint),large=area.h>input.pageHeight+.000001,grouped=object.parentId!==null;
    if(object.kind==='text'){
      if(!Array.isArray(object.sourceRanges)||!Array.isArray(object.lines)||!object.lines.length||(lineCount+=object.lines.length)>LIMIT_LINES)return invalid('문단의 원본 범위와 측정된 줄이 필요합니다.');
      const runs=[];let previous=-Infinity;
      for(const line of object.lines){
        if(!record(line)||!validBox(line.boxCssPx)||!validBox(line.paintBoxCssPx)||!inside(line.boxCssPx,box)||!inside(line.boxCssPx,line.paintBoxCssPx)||!inside(line.paintBoxCssPx,paint)||!Number.isFinite(line.baseline)||line.baseline<line.boxCssPx.y-.05||line.baseline>line.boxCssPx.y+line.boxCssPx.h+.05||line.boxCssPx.y<previous||!Array.isArray(line.runs)||!line.runs.length)return invalid('문단 줄의 위치·실제 표시 범위·기준선 또는 텍스트가 올바르지 않습니다.');
        previous=line.boxCssPx.y;
        for(const run of line.runs){if(!checkRun(run))return invalid('측정된 글자와 원본 텍스트 범위가 다릅니다.');runs.push(run);}
      }
      if(!validateSourceCoverage(object.sourceRanges,runs).valid)return invalid('문단의 원본 글자가 누락되었거나 중복되었습니다.');
      if(large&&object.splitSafe===true&&!grouped){
        if(!plainTextPaint(object.style))return unsupported('문단 배경·테두리 또는 효과의 분할 방법을 확인할 수 없습니다.');
        for(const [i,line] of object.lines.entries())addAtom({object,index:i,box:union(line.boxCssPx,line.paintBoxCssPx),kind:'line'});
      }
      else addAtom({object,box:area,kind:'whole'});
    }else if(object.kind==='table'){
      const content=object.content;
      if(!validateTable(content).valid||Math.abs(content.rowHeightsCssPx.reduce((a,b)=>a+b,0)-box.h)>.05||Math.abs(content.columnWidthsCssPx.reduce((a,b)=>a+b,0)-box.w)>.05)return invalid('표의 격자와 위치가 일치하지 않습니다.');
      const runs=[];for(const cell of content.cells)for(const run of cell.runs){if(run.text==='')continue;if(!checkRun(run))return invalid('표 셀 글자와 원본 텍스트 범위가 다릅니다.');runs.push(run);}
      if(object.sourceRanges!==undefined&&(!Array.isArray(object.sourceRanges)||!validateSourceCoverage(object.sourceRanges,runs).valid))return invalid('표의 원본 글자가 누락되었거나 중복되었습니다.');
      if(large&&object.splitSafe===true&&!grouped){
        if(!Array.isArray(object.rowPaintBoxes)||object.rowPaintBoxes.length!==content.rows)return unsupported('표 행별 실제 표시 범위를 확인할 수 없습니다.');
        let y=box.y;
        for(const [i,height] of content.rowHeightsCssPx.entries()){
          const row={x:box.x,y,w:box.w,h:height},rowPaint=object.rowPaintBoxes[i];
          if(!validBox(rowPaint)||!inside(row,rowPaint)||!inside(rowPaint,paint))return unsupported('표 행의 실제 표시 범위가 올바르지 않습니다.');
          addAtom({object,index:i,box:union(row,rowPaint),kind:'row'});y+=height;
        }
        for(const cell of content.cells)if(cell.rowSpan>1){
          const painted=object.rowPaintBoxes.slice(cell.row,cell.row+cell.rowSpan),left=Math.min(...painted.map(r=>r.x)),top=Math.min(...painted.map(r=>r.y)),right=Math.max(...painted.map(r=>r.x+r.w)),bottom=Math.max(...painted.map(r=>r.y+r.h));
          addAtom({object,box:{x:left,y:top,w:right-left,h:bottom-top},kind:'merge'});
        }
      }else addAtom({object,box:area,kind:'whole'});
    }else addAtom({object,box:area,kind:'whole'});
  }
  if(!validateSourceCoverage(expected,actual).valid)return invalid('원본 텍스트가 누락되었거나 중복되었습니다.');
  // A horizontal cut is forbidden inside every painted vertical interval.
  // Merging intersecting intervals also keeps transitively overlapping paint
  // together, including overlaps between distinct source objects.
  if(atoms.length>LIMIT_LINES)return invalid('분할 대상이 지원 한도를 초과했습니다.');
  const ordered=atoms.map(a=>({top:a.box.y,bottom:a.box.y+a.box.h})).sort((a,b)=>a.top-b.top||a.bottom-b.bottom),intervals=[];
  for(const interval of ordered){const last=intervals.at(-1);if(last&&interval.top<last.bottom-.000001)last.bottom=Math.max(last.bottom,interval.bottom);else intervals.push({...interval});}
  if(intervals.some(g=>g.bottom-g.top>input.pageHeight+.000001))return failure('PAGINATION_UNSUPPORTED','한 장보다 큰 객체 또는 겹친 객체 묶음을 나눌 수 없습니다.','PAGINATION_OVERSIZE');
  const pages=[];let top=0;
  while(top<input.documentHeight){
    if(pages.length>=LIMIT_PAGES)return unsupported('분할 결과가 200장을 초과합니다.');
    let end=Math.min(top+input.pageHeight,input.documentHeight),low=0,high=intervals.length;
    while(low<high){const mid=(low+high)>>1;if(intervals[mid].top<end-.000001)low=mid+1;else high=mid;}
    const interval=intervals[low-1];if(interval&&interval.bottom>end+.000001)end=interval.top;
    if(end<=top+.000001)return unsupported('안전한 장표 경계를 찾을 수 없습니다.');
    pages.push({index:pages.length,sourceStart:top,sourceEnd:end,pageHeight:input.pageHeight,parts:[]});top=end;
  }
  const pageFor=box=>pages.find(page=>box.y>=page.sourceStart-.000001&&box.y+box.h<=page.sourceEnd+.000001);
  for(const object of input.objects){
    const relevant=(objectAtoms.get(object)??[]).filter(a=>a.kind!=='merge');
    const pageGroups=new Map();for(const atom of relevant){const page=pageFor(atom.box);if(!page)return unsupported('객체를 한 장에 배치할 수 없습니다.');if(!pageGroups.has(page))pageGroups.set(page,[]);pageGroups.get(page).push(atom);}
    for(const [page,items] of pageGroups){
      const part={objectId:object.id,kind:object.kind};
      if(items[0].kind==='line'){
        const start=items[0].index,end=items.at(-1).index+1;if(items.length!==end-start)return unsupported('문단의 줄 순서를 보존할 수 없습니다.');
        part.lineStart=start;part.lineEnd=end;part.lines=structuredClone(object.lines.slice(start,end));part.runs=part.lines.flatMap(line=>line.runs);
      }else if(items[0].kind==='row'){
        const start=items[0].index,end=items.at(-1).index+1;if(items.length!==end-start)return unsupported('표 행 순서를 보존할 수 없습니다.');
        const paintedBoundary=cut=>cut>0&&cut<object.content.rows&&object.content.cells.some(cell=>{
          const edge=cell.row===cut?cell.style.borders.top:cell.row+cell.rowSpan===cut?cell.style.borders.bottom:null;
          return edge&&edge.width>0&&edge.style!=='none';
        });
        if((paintedBoundary(start)||paintedBoundary(end))&&object.tableBoundaryPolicy!=='repeat')return unsupported('표 분할 경계의 테두리 반복 정책이 필요합니다.');
        const split=splitTableRows(object.content,start,end);if(split.diagnostics.length)return {pages:[],diagnostics:split.diagnostics};
        part.rowStart=start;part.rowEnd=end;part.content=split.content;if(paintedBoundary(start)||paintedBoundary(end))part.tableBoundaryPolicy='repeat';
        part.sourceCells=object.content.cells.filter(c=>c.row>=start&&c.row<end).map(c=>({row:c.row,col:c.col,rowSpan:c.rowSpan,colSpan:c.colSpan}));
      }else if(object.kind==='text'){
        part.lineStart=0;part.lineEnd=object.lines.length;part.lines=structuredClone(object.lines);part.runs=part.lines.flatMap(line=>line.runs);
      }else if(object.kind==='table'){
        part.rowStart=0;part.rowEnd=object.content.rows;part.content=structuredClone(object.content);part.sourceCells=object.content.cells.map(c=>({row:c.row,col:c.col,rowSpan:c.rowSpan,colSpan:c.colSpan}));
      }
      page.parts.push(part);
    }
  }
  const emittedWithPages=pages.flatMap(page=>page.parts.flatMap(part=>(part.runs??(part.kind==='table'?part.content.cells.flatMap(c=>c.runs.filter(r=>r.text!=='')):[])).map(run=>({run,page:page.index}))));
  const emitted=emittedWithPages.map(item=>item.run);
  if(!validateSourceCoverage(expected,emitted).valid)return invalid('나눈 장표에서 원본 텍스트가 누락되었거나 중복되었습니다.');
  // A Unicode grapheme may cross DOM text nodes (for example e + combining
  // accent in separate spans). A page boundary must not split that grapheme.
  const graphemes=new Set([0,ledgerText.length]);for(const part of new Intl.Segmenter('und',{granularity:'grapheme'}).segment(ledgerText))graphemes.add(part.index);
  let cursor=0,previousPage=null;
  for(const {run,page} of emittedWithPages){
    const spans=sources.get(run.sourceId).spans;let low=0,high=spans.length;
    while(low<high){const mid=(low+high)>>1;if(spans[mid].end<=run.start)low=mid+1;else high=mid;}
    const span=spans[low];if(!span||run.start<span.start||run.end>span.end||span.base+run.start-span.start!==cursor)return unsupported('원본 텍스트의 읽기 순서를 보존할 수 없습니다.');
    if(previousPage!==null&&page!==previousPage&&!graphemes.has(cursor))return unsupported('하나의 글자 조합을 서로 다른 장표로 나눌 수 없습니다.');
    cursor+=run.end-run.start;previousPage=page;
  }
  return {pages,diagnostics:[]};
}
