import {pptCharacterSpacing} from './text-spacing.mjs';
import {paragraphTextRuns} from './write-paragraphs.mjs';

// PptxGenJS owns the native table XML; measured cell sizes remain explicit.
export function addNativeTable(slide,node,{pointScale,resolveFont,paint}){
  const {content:c,boxCssPx:b}=node,inch=v=>v*pointScale/72;
  const textStyle=s=>{
    const color=paint(s.color),options={...resolveFont(s),fontSize:s.fontSize*pointScale,italic:s.fontStyle==='italic',color:color.color,transparency:color.transparency,charSpacing:pptCharacterSpacing(s,pointScale),underline:s.textDecoration?.includes('underline')?{style:'sng'}:undefined,...(Number.isInteger(s.baseline)&&s.baseline!==0&&Math.abs(s.baseline)<=100000?{baseline:s.baseline/50}:{}),breakLine:false};
    if(s.outline){const edge=paint(s.outline.color);if(edge.transparency!==0||!Number.isFinite(s.outline.width)||s.outline.width<=0)throw new Error('불투명 단색 외의 표 글자 외곽선은 지원하지 않습니다.');options.outline={color:edge.color,size:s.outline.width*pointScale};}
    if(s.highlight){const highlight=paint(s.highlight);if(highlight.transparency!==0)throw new Error('불투명 단색 외의 표 글자 형광펜은 지원하지 않습니다.');options.highlight=highlight.color;}
    return options;
  };
  const rows=Array.from({length:c.rows},()=>[]);
  for(const cell of [...c.cells].sort((a,b)=>a.row-b.row||a.col-b.col)){
    const s=cell.style,p=s.padding??{top:0,right:0,bottom:0,left:0};
    const padding=[p.top,p.right,p.bottom,p.left];
    // PptxGenJS 4 interprets top margins >= 1 as points, smaller values as inches.
    const margin=padding.map(v=>p.top*pointScale>=72?v*pointScale:inch(v));
    const border=['top','right','bottom','left'].map(side=>{
      const edge=s.borders[side],color=paint(edge.color);
      if(edge.style==='none'||edge.width===0||color.transparency===100)return {type:'none',pt:0,color:'FFFFFF'};
      if(color.transparency!==0)throw new Error('반투명 표 테두리는 아직 지원하지 않습니다.');
      return {type:'solid',pt:edge.width*pointScale,color:color.color};
    });
    const runs=cell.paragraphs?cell.paragraphs.flatMap((paragraph,index)=>{
      const items=paragraphTextRuns(paragraph,r=>({text:r.text,options:textStyle(r.style)}),s);
      if(index<cell.paragraphs.length-1)items.at(-1).options.breakLine=true;
      return items;
    }):(cell.runs.length?cell.runs:[{text:cell.text,style:s}]).map(r=>({text:r.text,options:textStyle(r.style)}));
    rows[cell.row].push({text:runs,options:{...textStyle(s),rowspan:cell.rowSpan,colspan:cell.colSpan,margin,border,fill:paint(s.fill??'transparent'),align:s.align??'left',valign:s.valign==='mid'?'middle':s.valign??'top',lineSpacing:s.lineHeight!=='normal'&&parseFloat(s.lineHeight)>0?parseFloat(s.lineHeight)*pointScale:undefined,paraSpaceBefore:0,paraSpaceAfter:0}});
  }
  slide.addTable(rows,{x:inch(b.x),y:inch(b.y),w:inch(b.w),h:inch(b.h),colW:c.columnWidthsCssPx.map(inch),rowH:c.rowHeightsCssPx.map(inch),autoPage:false,objectName:node.id,margin:0,fontFace:resolveFont(c.cells[0].style).fontFace,fontSize:c.cells[0].style.fontSize*pointScale,lang:'ko-KR',breakLine:false});
  return {sourceNodeId:node.id,sourceRef:node.sourceRef,pptObjectName:node.id,support:'native-table',parentId:node.parentId};
}
