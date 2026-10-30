import {diagnostic} from './model.mjs';
import {EMU_PER_CSS_PX} from './canvas.mjs';
const elements=n=>[...(n?.childNodes??[])].filter(c=>c.nodeType===1);
const child=(n,name)=>elements(n).find(c=>c.localName===name);
const attr=(n,k,fallback=null)=>n?.hasAttribute(k)?n.getAttribute(k):fallback;
const px=(n,k,fallback=0)=>Number(attr(n,k,fallback))/EMU_PER_CSS_PX;
const A='http://schemas.openxmlformats.org/drawingml/2006/main';

export function readNativeTable(frame,{slideId,sourceId,boxCssPx,readText,readColor}){
  const diagnostics=[],issue=message=>diagnostics.push(diagnostic('UNSUPPORTED_OBJECT',message,{slideId,nodeId:sourceId}));
  const graphic=child(frame,'graphic'),data=child(graphic,'graphicData'),table=child(data,'tbl');
  if(graphic?.namespaceURI!==A||attr(data,'uri')!=='http://schemas.openxmlformats.org/drawingml/2006/table'||table?.namespaceURI!==A||elements(data).length!==1){issue('차트·SmartArt는 아직 표로 읽을 수 없습니다.');return {node:null,diagnostics};}
  const columns=elements(child(table,'tblGrid')),rows=elements(table).filter(n=>n.localName==='tr');
  if(!rows.length||rows.length>100||!columns.length||columns.length>50||rows.length*columns.length>2000){issue('표의 행·열·셀 수 제한을 초과했습니다.');return {node:null,diagnostics};}
  if(columns.some(n=>n.localName!=='gridCol'||n.namespaceURI!==A)){issue('표 열 정의가 잘못되었습니다.');return {node:null,diagnostics};}
  const columnWidthsCssPx=columns.map(n=>px(n,'w')),rowHeightsCssPx=rows.map(n=>px(n,'h')),cells=[];
  const themed=!!child(child(table,'tblPr'),'tableStyleId'),covered=new Map();
  for(const [row,tr] of rows.entries()){
    const rowCells=elements(tr).filter(n=>n.localName==='tc');
    if(rowCells.length!==columns.length){issue('표 셀 수와 열 수가 일치하지 않습니다.');continue;}
    for(const [col,tc] of rowCells.entries()){
      const rowSpan=Number(attr(tc,'rowSpan',1)),colSpan=Number(attr(tc,'gridSpan',1)),hm=attr(tc,'hMerge','0'),vm=attr(tc,'vMerge','0');
      if(![rowSpan,colSpan].every(n=>Number.isInteger(n)&&n>0)||![hm,vm].every(v=>['0','1','false','true'].includes(v))){issue('병합 셀의 범위·표시가 잘못되었습니다.');continue;}
      const hMerge=['1','true'].includes(hm),vMerge=['1','true'].includes(vm),anchor=covered.get(`${row}:${col}`);
      if(hMerge||vMerge){
        if(!anchor||hMerge!==(col>anchor.col)||vMerge!==(row>anchor.row)||rowSpan>anchor.rowSpan||colSpan>anchor.colSpan)issue('병합 셀과 시작 셀의 관계가 일치하지 않습니다.');
        if([...tc.getElementsByTagNameNS('*','t')].some(n=>n.textContent.length))issue('병합으로 가려진 셀에 문구가 있어 보존할 수 없습니다.');
        continue;
      }
      if(anchor||row+rowSpan>rows.length||col+colSpan>columns.length){issue('병합 셀이 겹치거나 격자 범위를 벗어났습니다.');continue;}
      const start={row,col,rowSpan,colSpan};
      for(let r=row;r<row+rowSpan;r++)for(let c=col;c<col+colSpan;c++){const key=`${r}:${c}`;if(covered.has(key))issue('병합 셀이 겹칩니다.');covered.set(key,start);}
      const props=child(tc,'tcPr'),tx=child(tc,'txBody');
      if(!props||!tx){issue('표 셀의 내용·서식이 없습니다.');continue;}
      if(attr(props,'vert','horz')!=='horz'||!['t','ctr','b'].includes(attr(props,'anchor','t'))||elements(props).some(n=>!['lnL','lnR','lnT','lnB','solidFill','noFill','cell3D','extLst'].includes(n.localName))||child(props,'cell3D'))issue('세로쓰기·복잡한 표 셀 효과는 지원하지 않습니다.');
      const text=readText(tx),borders={};
      for(const [side,name]of [['left','lnL'],['right','lnR'],['top','lnT'],['bottom','lnB']]){
        const line=child(props,name),none=child(line,'noFill');
        if(!line&&themed)issue('테마에서 상속된 표 테두리는 지원하지 않습니다.');
        if(line&&!none&&(!child(line,'solidFill')||attr(child(line,'prstDash'),'val','solid')!=='solid'||['headEnd','tailEnd'].some(n=>attr(child(line,n),'type','none')!=='none')))issue('단색 실선 외의 표 테두리는 지원하지 않습니다.');
        borders[side]={color:none||!line?'transparent':readColor(child(line,'solidFill')),width:none||!line?0:px(line,'w'),style:none||!line?'none':'solid'};
      }
      if(themed&&!child(props,'solidFill')&&!child(props,'noFill'))issue('테마에서 상속된 표 배경은 지원하지 않습니다.');
      const style={...text.style,fill:readColor(child(props,'solidFill')),borders,padding:{left:px(props,'marL',91440),right:px(props,'marR',91440),top:px(props,'marT',45720),bottom:px(props,'marB',45720)},valign:({t:'top',ctr:'mid',b:'bottom'})[attr(props,'anchor','t')]};
      cells.push({row,col,rowSpan,colSpan,text:text.content.text,runs:text.content.runs,style});
    }
  }
  // PowerPoint can retain the old graphicFrame extent after adding rows. Its
  // visible table size (and native Shape.Width/Height) follows the current grid.
  const bounds={...boxCssPx,w:columnWidthsCssPx.reduce((a,v)=>a+v,0),h:rowHeightsCssPx.reduce((a,v)=>a+v,0)};
  return {node:diagnostics.length?null:{id:sourceId,kind:'table',boxCssPx:bounds,content:{rows:rows.length,columns:columns.length,columnWidthsCssPx,rowHeightsCssPx,cells},style:{opacity:1}},diagnostics};
}
