import {validateChart,chartAxis} from './charts.mjs';

// PptxGenJS writes both the native chart and its internal editable workbook.
export function addNativeChart(slide,node,{pointScale,resolveFont}){
  const {content:c,boxCssPx:b}=node,validation=validateChart(c),font=resolveFont(node.style);
  if(!validation.valid)throw new Error(`Invalid chart: ${validation.errors.map(e=>e.path).join(', ')}`);
  if(typeof font?.fontFace!=='string'||!font.fontFace.trim()||/[\u0000-\u001f\ufffe\uffff]/.test(font.fontFace)||!Number.isFinite(pointScale)||pointScale<=0||!Number.isFinite(node.style.fontSize)||node.style.fontSize<=0)throw new Error('Chart font or coordinate scale is unavailable');
  // PptxGenJS 4 inserts chart typefaces directly into XML attributes.
  const face=font.fontFace.replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&apos;'}[c]));
  const size=node.style.fontSize*pointScale,strokeScale=size/12,color='223845',inch=v=>Math.round(v*pointScale*12700)/914400;
  const options={x:inch(b.x),y:inch(b.y),w:inch(b.w),h:inch(b.h),objectName:node.id,
    fontFace:face,lang:'ko-KR',showTitle:!!c.title,title:c.title,titleFontFace:face,titleFontSize:size,titleBold:false,titleColor:color,titleAlign:'center',
    showLegend:c.legend,legendPos:'b',legendFontFace:face,legendFontSize:size,legendColor:color,
    chartArea:{fill:{color:'FFFFFF',transparency:100},roundedCorners:false},plotArea:{fill:{color:'FFFFFF',transparency:100}},
    chartColors:c.type==='pie'?c.categories.map(()=>c.series[0].color.slice(1)):c.series.map(s=>s.color.slice(1)),
    showLabel:c.type==='pie',showValue:false,showPercent:false,showSerName:false,showLeaderLines:false,showDataTable:false,
    dataLabelPosition:'ctr',dataLabelFontFace:face,dataLabelFontSize:size,dataLabelFontBold:false,dataLabelFontItalic:false,dataLabelColor:c.type==='pie'?'FFFFFF':color,
    valLabelFormatCode:'General',displayBlanksAs:'gap',
  };
  const axis=chartAxis(c);
  if(axis)Object.assign(options,{
    catAxisHidden:false,catAxisLabelFontFace:face,catAxisLabelFontSize:size,catAxisLabelFontBold:false,catAxisLabelFontItalic:false,catAxisLabelColor:color,catAxisLabelPos:'low',
    catAxisLineColor:'475569',catAxisLineSize:strokeScale,catAxisLineShow:true,catAxisMajorTickMark:'none',catAxisMinorTickMark:'none',catAxisOrientation:'minMax',catGridLine:{style:'none'},
    valAxisHidden:false,valAxisLabelFontFace:face,valAxisLabelFontSize:size,valAxisLabelFontBold:false,valAxisLabelFontItalic:false,valAxisLabelColor:color,valAxisLabelPos:'low',
    valAxisLineColor:'475569',valAxisLineSize:strokeScale,valAxisLineShow:true,valAxisMajorTickMark:'none',valAxisMinorTickMark:'none',valAxisOrientation:'minMax',valAxisLabelFormatCode:'General',
    valAxisMinVal:axis.min,valAxisMaxVal:axis.max,valAxisMajorUnit:axis.majorUnit,valGridLine:{color:'CBD5E1',size:strokeScale,style:'solid'},
    catAxisTitle:c.axes.xTitle??'',showCatAxisTitle:!!c.axes.xTitle,catAxisTitleFontFace:face,catAxisTitleFontSize:size,catAxisTitleColor:color,
    valAxisTitle:c.axes.yTitle??'',showValAxisTitle:!!c.axes.yTitle,valAxisTitleFontFace:face,valAxisTitleFontSize:size,valAxisTitleColor:color,
  });
  if(c.type==='bar')Object.assign(options,{barDir:'col',barGrouping:'clustered',barGapWidthPct:33,barOverlapPct:0});
  if(c.type==='line'){
    const marker=Math.round(6*strokeScale);if(marker<2||marker>72)throw new Error('Chart marker size is outside the native format range');
    Object.assign(options,{lineSize:2*strokeScale,lineSmooth:false,lineDataSymbol:'circle',lineDataSymbolSize:marker,lineDataSymbolLineSize:strokeScale});
  }
  if(c.type==='pie')Object.assign(options,{firstSliceAng:0,dataNoEffects:true,dataBorder:{color:'FFFFFF',pt:strokeScale}});
  // PptxGenJS 4.0.1 uses `value || ''` in numeric workbook cells. Serialize
  // only zero in this copy so it writes <v>0</v>; the source stays numeric.
  const series=c.series.map(s=>({name:s.name,labels:[...c.categories],values:s.values.map(value=>value===0?'0':value)}));
  slide.addChart(c.type,series,options);
  return {sourceNodeId:node.id,sourceRef:node.sourceRef,pptObjectName:node.id,support:'native-chart',parentId:node.parentId};
}
