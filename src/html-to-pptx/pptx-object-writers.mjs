import {assetDataUrl} from './assets.mjs';
import {diagnostic} from './model.mjs';
import {addNativeTable} from './write-table.mjs';
import {addNativeChart} from './write-chart.mjs';
import {pptCharacterSpacing} from './text-spacing.mjs';
import {paragraphTextRuns} from './write-paragraphs.mjs';

// Writers share the caller's slide, font and diagnostic state. The caller
// alone controls drawing order, generated background IDs and objectMap entries.
export function writePptxTable(node,{slide,pointScale,resolveTextFont,source,paint}){
  return addNativeTable(slide,node,{pointScale,resolveFont:s=>resolveTextFont(s,source),paint});
}
export function writePptxChart(node,{slide,pointScale,resolveFont,source}){
  return addNativeChart(slide,node,{pointScale,resolveFont:s=>resolveFont(s,source)});
}

// An ordinary transparent picture may reveal the shape's original fill.
// Captured background artwork already contains that fill and must not use this.
export function writePptxBackgroundFill(node,box,{slide,pptx,paint},objectName){
  const st=node.style;
  slide.addShape(pptx.ShapeType.rect,{...box,objectName,fill:paint(st.fill,st.fillOpacity??1),line:paint('transparent')});
}

export function writePptxText(node,box,{slide,pptx,source,measured,resolveTextFont,paint,pointScale,inchScale}){
  const st=node.style;
  const runOptions=s=>({...resolveTextFont(s,source),fontSize:s.fontSize*pointScale,italic:s.fontStyle==='italic',color:paint(s.color,st.opacity).color,transparency:paint(s.color,st.opacity).transparency,charSpacing:pptCharacterSpacing(s,pointScale),underline:s.textDecoration?.includes('underline')?{style:'sng'}:undefined,...(Number.isInteger(s.baseline)&&s.baseline!==0&&Math.abs(s.baseline)<=100000?{baseline:s.baseline/50}:{}),breakLine:false});
  const makeRun=r=>{
    const options=runOptions(r.style),outline=r.style.outline;
    if(outline){
      const color=paint(outline.color,st.opacity??1);
      if(color.transparency!==0||!Number.isFinite(outline.width)||outline.width<=0)throw new Error('불투명 단색 외의 글자 외곽선은 지원하지 않습니다.');
      options.outline={color:color.color,size:outline.width*pointScale};
    }
    if(r.style.highlight){const highlight=paint(r.style.highlight);if(highlight.transparency!==0)throw new Error('불투명 단색 외의 글자 형광펜은 지원하지 않습니다.');options.highlight=highlight.color;}
    return {text:r.text,options};
  };
  let runs=node.content.paragraphs?node.content.paragraphs.flatMap((para,index)=>{
    const items=paragraphTextRuns(para,makeRun,st),spacing=parseFloat(para.lineHeight)*(para.lineHeight.endsWith('px')?1:(para.runs.find(r=>!r.marker)?.style.fontSize??st.fontSize))*pointScale;
    for(const item of items)Object.assign(item.options,{align:para.align,lineSpacing:spacing,paraSpaceBefore:para.beforePx*pointScale,paraSpaceAfter:para.afterPx*pointScale});
    if(index<node.content.paragraphs.length-1)items.at(-1).options.breakLine=true;
    return items;
  }):(node.content.runs??[{text:node.content.text,style:st}]).map(makeRun);
  if(node.content.tableCellGroup){
    const lines=[];for(const run of runs){const segments=run.text.split('\n');for(let i=0;i<segments.length;i++){if(segments[i])lines.push({text:segments[i],options:{...run.options}});if(i<segments.length-1){if(!lines.length)lines.push({text:'',options:{...run.options}});lines.at(-1).options.breakLine=true;}}}runs=lines;
  }
  const padding=st.padding??{left:0,right:0,top:0,bottom:0};
  // Browser and Office advances can differ slightly in a tight auto-sized
  // label. Preserve a measured single line without enlarging its box or
  // shrinking type. Longer edits may overflow and need a fresh HTML export.
  const layout=measured.measurements.find(m=>m.slideId===source.id).objects.find(o=>o.nodeId===node.id);
  const {bold:defaultBold,italic:defaultItalic,...textDefaults}=runOptions(st);
  const shapeOptions=node.pptxOrigin?{
    shape:st.radius>0?pptx.ShapeType.roundRect:pptx.ShapeType.rect,
    rectRadius:st.radius>0?st.radius*inchScale:undefined,
    fill:paint(st.fill,(st.opacity??1)*(st.fillOpacity??1)),
    line:{...paint(st.stroke,(st.strokeWidth??0)>0?(st.opacity??1)*(st.strokeOpacity??1):0),width:(st.strokeWidth??0)*pointScale,dashType:st.dash||st.borderStyle==='dashed'?'dash':st.borderStyle==='dotted'?'sysDot':'solid'},
    isTextBox:node.pptxOrigin.textBox
  }:{isTextBox:true};
  slide.addText(runs,{...box,...textDefaults,...shapeOptions,margin:[padding.left,padding.right,padding.bottom,padding.top].map(v=>v*pointScale),lineSpacing:st.lineHeight!=='normal'&&parseFloat(st.lineHeight)>0?parseFloat(st.lineHeight)*pointScale:undefined,paraSpaceBefore:0,paraSpaceAfter:0,align:st.align??'left',valign:st.valign??'top',fit:'none',wrap:!['pre','nowrap'].includes(st.whiteSpace)&&layout.lines.length!==1,lang:'ko-KR'});
}

export function writePptxShape(node,box,{slide,pptx,inches,paint,pointScale,inchScale}){
  const b=node.boxCssPx,st=node.style;
  const stroke=paint(st.stroke,(st.strokeWidth??0)>0?(st.opacity??1)*(st.strokeOpacity??1):0),half=st.svgGeometry||stroke.transparency===100?0:(st.strokeWidth??0)/2;
  if(b.w<=half*2||b.h<=half*2)throw new Error('Border is wider than the shape');
  slide.addShape(st.shape==='ellipse'?pptx.ShapeType.ellipse:st.radius>0?pptx.ShapeType.roundRect:pptx.ShapeType.rect,{...box,x:inches(b.x+half),y:inches(b.y+half),w:inches(b.w-2*half),h:inches(b.h-2*half),rectRadius:st.shape==='ellipse'?undefined:Math.max(0,(st.radius??0)-half)*inchScale,fill:paint(st.assetId?'transparent':st.fill,(st.opacity??1)*(st.fillOpacity??1)),line:{...stroke,width:(st.strokeWidth??0)*pointScale,dashType:st.dash||st.borderStyle==='dashed'?'dash':st.borderStyle==='dotted'?'sysDot':'solid'}});
}

export function writePptxLine(node,box,{slide,pptx,inches,paint,pointScale,profile}){
  const b=node.boxCssPx,st=node.style;
  const horizontal=b.w>=b.h,t=st.strokeWidth||Math.min(b.w,b.h)||1;
  if(st.svgGeometry&&st.endpoints){
    const {start,end}=st.endpoints;
    slide.addShape(pptx.ShapeType.line,{...box,x:inches(Math.min(start.x,end.x)),y:inches(Math.min(start.y,end.y)),w:inches(Math.abs(end.x-start.x)),h:inches(Math.abs(end.y-start.y)),flipH:end.x<start.x,flipV:end.y<start.y,line:{...paint(st.stroke,st.opacity),width:st.strokeWidth*pointScale}});
    profile.lineRepresentations.strokes++;
  }else if(st.linePaint?.kind==='bar'){
    const {thickness,horizontal:barHorizontal}=st.linePaint;
    slide.addShape(pptx.ShapeType.rect,{...box,w:inches(barHorizontal?b.w:thickness),h:inches(barHorizontal?thickness:b.h),fill:paint(st.linePaint.color,st.opacity),line:paint('transparent')});
    profile.lineRepresentations.filledBars++;
  }else{
    slide.addShape(pptx.ShapeType.line,{...box,x:inches(b.x+(horizontal?0:t/2)),y:inches(b.y+(horizontal?t/2:0)),w:horizontal?inches(b.w):0,h:horizontal?0:inches(b.h),line:{...paint(st.stroke??st.fill,st.opacity),width:t*pointScale,dashType:st.dash?'dash':'solid'}});
    profile.lineRepresentations.strokes++;
  }
}

export function writePptxImage(node,box,{slide,pptx,source,measured,inches,paint,pointScale,diagnostics}){
  const b=node.boxCssPx,st=node.style;
  const asset=measured.model.assets.find(a=>a.id===node.content.assetId);if(!asset)throw new Error('Picture asset missing');
  if(node.content.editableVector){
    const v=node.content.editableVector;
    slide.addShape(pptx.ShapeType[v.shape],{...box,x:inches(b.x+(b.w-v.width)/2),y:inches(b.y+(b.h-v.height)/2),w:inches(v.width),h:inches(v.height),rotate:v.angle,fill:paint(v.fill==='gradient'?v.gradient.from:v.fill==='none'?'transparent':v.fill,st.opacity),line:{...paint(v.stroke==='none'?'transparent':v.stroke,st.opacity),width:v.strokeWidth*pointScale}});
  }else{
    const crop=node.content.crop;
    const fullWidth=crop?box.w/(1-(crop.left+crop.right)/100000):box.w,fullHeight=crop?box.h/(1-(crop.top+crop.bottom)/100000):box.h;
    const sizing=crop?{type:'crop',x:fullWidth*crop.left/100000,y:fullHeight*crop.top/100000,w:box.w,h:box.h}:undefined;
    slide.addImage({...box,w:fullWidth,h:fullHeight,sizing,data:assetDataUrl(asset),rounding:node.content.rounding===true,transparency:(1-(st.opacity??1))*100});
    if(asset.mime==='image/svg+xml'&&!node.content.artworkOnly)diagnostics.push(diagnostic('UNSUPPORTED_OBJECT','SVG는 그림 객체입니다. 내부 글자와 경로를 개별 편집할 수 없습니다.',{slideId:source.id,nodeId:node.id,sourceRef:node.sourceRef},'warning'));
  }
}
