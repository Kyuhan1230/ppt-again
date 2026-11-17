import {assetDataUrl} from './assets.mjs';
import {verifiedBackgroundDataUrl,validateBackgroundCapture} from './background-assets.mjs';
import {renderChartSvg} from './charts.mjs';
import {chartHashes} from './chart-input.mjs';
import {presetVectorMarkup} from './preset-vector.mjs';

export const esc=v=>String(v??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
export const number=v=>String(Number(Number(v).toFixed(10)));
const backgroundCache=new WeakMap();
function backgroundUrl(assets,id){
  let cache=backgroundCache.get(assets);
  if(!cache){cache=new Map();backgroundCache.set(assets,cache);}
  if(cache.has(id))return cache.get(id);
  const url=verifiedBackgroundDataUrl(assets.get(id));cache.set(id,url);return url;
}
// PPT uses the element's alpha independently from the fill/stroke alpha.
// Keep that distinction when emitting CSS so transparent artwork reveals the
// same underlying colour in both output formats.
function colorWithOpacity(value,opacity=1){
  const color=value??'transparent';
  if(!Number.isFinite(opacity)||opacity<0||opacity>1)throw new Error('색상 불투명도가 잘못되었습니다.');
  if(opacity===1||color==='transparent')return color;
  let rgb,alpha=1;
  if(/^#[\da-f]{3,8}$/i.test(color)){
    let hex=color.slice(1);if(hex.length===3||hex.length===4)hex=[...hex].map(c=>c+c).join('');
    if(hex.length===8){alpha=parseInt(hex.slice(6),16)/255;hex=hex.slice(0,6);}
    if(hex.length===6)rgb=[0,2,4].map(i=>parseInt(hex.slice(i,i+2),16));
  }else{
    const match=/^rgba?\(([^)]+)\)$/i.exec(color);
    if(match){const values=match[1].split(/[, /]+/).map(Number);rgb=values.slice(0,3);alpha=values[3]??1;}
  }
  if(!rgb||rgb.length!==3||rgb.some(n=>!Number.isFinite(n)||n<0||n>255)||!Number.isFinite(alpha)||alpha<0||alpha>1)throw new Error('색상 불투명도를 보존할 수 없습니다.');
  return `rgba(${rgb.join(',')},${number(alpha*opacity)})`;
}
function backgroundCss(background,assets,imageKey,bounds){
  if(background?.assetId){
    const capture=background.capture;
    validateBackgroundCapture(capture,assets.get(background.assetId),bounds);
    const color=capture?'transparent':colorWithOpacity(background?.fill??background?.color??'transparent',background?.fillOpacity??1);
    return `background:${color} url("${backgroundUrl(assets,background.assetId)}") 0 0 / 100% 100% no-repeat;`;
  }
  if(background?.[imageKey]&&background[imageKey]!=='none')throw new Error('배경 효과를 보존할 그림 자산이 없습니다.');
  return `background:${colorWithOpacity(background?.fill??background?.color??'transparent',background?.fillOpacity??1)};`;
}
export function slideBackgroundCss(background,assets,bounds){
  return backgroundCss({...background,color:background?.color??'#fff'},assets,'image',bounds);
}
const fontCss=(s,shift=true)=>{
  const baseline=Number.isInteger(s.baseline)&&Math.abs(s.baseline)<=100000?s.baseline:0;
  return `font-family:${JSON.stringify(s.fontFamily)};font-size:${number(s.fontSize*(baseline?2/3:1))}px;font-weight:${s.fontWeight??400};font-style:${s.fontStyle??'normal'};color:${s.color??'#000'};letter-spacing:${s.letterSpacing==='normal'||s.letterSpacing==null?'normal':number(parseFloat(s.letterSpacing))+'px'};${s.disableKerning===true?'font-kerning:none;':''}text-decoration:${s.textDecoration??'none'};${baseline&&shift?`vertical-align:${number(s.fontSize*baseline/100000)}px;`:''}${s.outline?`-webkit-text-stroke:${number(s.outline.width)}px ${s.outline.color};`:''}`;
};
const runMarkup=(runs,shift=true,markerWidth=0,hangingPeriod=false,glyphOffset=0)=>runs.map((r,index)=>{
  const marker=r.marker===true,fixedMarker=marker&&markerWidth>0;
  const markerCss=fixedMarker?`display:inline-block;width:${number(markerWidth)}px;text-indent:0;white-space:pre;line-height:0;`:marker?'line-height:0;':'';
  const highlight=/^#[a-f\d]{6}$/i.test(r.style.highlight??'')?r.style.highlight:null;
  // Chromium does not implement hanging-punctuation. At the end of a
  // left-aligned PPT table paragraph, let a final period extend into the
  // inset without consuming extra line width, as Office does.
  const hang=hangingPeriod&&!marker&&index===runs.length-1&&r.text.endsWith('.');
  const content=hang?`${esc(r.text.slice(0,-1))}<span data-pptx-hanging-punct style="display:inline-block;width:0;line-height:0;white-space:nowrap;text-indent:0">.</span>`:esc(r.text);
  return `<span${marker?' data-pptx-list-marker=""':''}${highlight?' data-pptx-highlight=""':''}${Number.isInteger(r.style.baseline)&&r.style.baseline!==0&&Math.abs(r.style.baseline)<=100000?` data-pptx-baseline="${r.style.baseline}"`:''} style="${esc(markerCss+(glyphOffset?`position:relative;top:${number(glyphOffset)}px;`:'')+fontCss(r.style,shift)+(highlight?`background-color:${highlight};`:''))}">${content}</span>`;
}).join('');
const paragraphMarkup=(paragraphs,baseStyle,explicitSeparators=true)=>paragraphs.map((p,index)=>{
  const paragraphStyle=p.runs.find(run=>!run.marker&&run.text.trim())?.style??p.runs[0]?.style??baseStyle;
  const css=`display:block;${fontCss({...paragraphStyle,textDecoration:'none'},false)}line-height:${p.lineHeight};text-align:${p.align};padding-top:${number(p.beforePx)}px;padding-bottom:${number(p.afterPx)}px;padding-left:${number(p.leftPx)}px;padding-right:${number(p.rightPx)}px;text-indent:${number(p.indentPx)}px;`;
  const empty=!p.runs.some(run=>run.text.length),last=index===paragraphs.length-1;
  const visible=p.runs.filter(run=>run.text.length),first=visible[0]?.style;
  const uniformBaseline=Boolean(first?.baseline&&visible.every(run=>run.style.baseline===first.baseline&&run.style.fontSize===first.fontSize));
  // PowerPoint places glyphs lower within an expanded percentage line box
  // than Chromium does. Preserve the line box and move its glyphs only for
  // top-aligned text, where native and HTML block heights agree.
  const ratio=Number(p.lineHeight),glyphOffset=explicitSeparators&&(!baseStyle.valign||baseStyle.valign==='top')&&Number.isFinite(ratio)&&ratio>1.2?(ratio-1.2)*paragraphStyle.fontSize/4:0;
  const bullet=p.bullet?` data-pptx-bullet-type="${esc(p.bullet.type)}" data-pptx-bullet-level="${p.bullet.level??0}"${p.bullet.type==='char'?` data-pptx-bullet-char="${esc(p.bullet.char)}"`:` data-pptx-bullet-format="${esc(p.bullet.format)}" data-pptx-bullet-start="${p.bullet.start??1}" data-pptx-bullet-restart="${p.bullet.restart===false?0:1}"`}${p.bullet.followTextColor?' data-pptx-bullet-follow-text-color="1"':''}${p.bullet.followTextFont?' data-pptx-bullet-follow-text-font="1"':''}`:'';
  const marker=p.runs.find(r=>r.marker),bulletSize=p.bullet?.sourceSize?` data-pptx-source-bullet-size="${esc(JSON.stringify({size:p.bullet.sourceSize,markerSize:marker?.style.fontSize??paragraphStyle.fontSize,textSize:paragraphStyle.fontSize}))}"`:'';
  const fontSize=Math.max(...p.runs.filter(r=>!r.marker).map(r=>r.style.fontSize),paragraphStyle.fontSize);
  const sourceSpacing=p.sourceSpacing?` data-pptx-source-spacing="${esc(JSON.stringify({values:p.sourceSpacing,fontSize,beforePx:p.beforePx,afterPx:p.afterPx,linePx:parseFloat(p.lineHeight)*(p.lineHeight.endsWith('px')?1:paragraphStyle.fontSize)}))}"`:'';
  const sourceEnd=p.sourceEnd?` data-pptx-source-end="${esc(JSON.stringify(p.sourceEnd))}"`:'';
  return `<span data-pptx-paragraph${bullet}${bulletSize}${sourceSpacing}${sourceEnd}${p.hangingPunct===true?' data-pptx-hanging-punctuation="1"':''} style="${esc(css)}">${runMarkup(p.runs,!uniformBaseline,Math.max(0,-p.indentPx),!explicitSeparators&&p.hangingPunct===true&&p.align==='left',glyphOffset)}${empty&&(last||!explicitSeparators)?'<br data-pptx-empty>':''}${!last&&explicitSeparators?'\n':''}</span>`;
}).join('');

function writeText(ctx,output){
  const {node,s,b,slide,assets,usedIds,originDocument}=ctx;
  let {style,body,tag,extra}=output;
          const p=s.padding??{};
          if(node.pptxOrigin&&node.pptxOrigin.documentId===originDocument)extra+=` data-pptx-origin="${esc(JSON.stringify(node.pptxOrigin))}"`;
          if(node.content.placeholderDescription)extra=` data-placeholder-marker="" role="img" tabindex="0" aria-label="${esc(node.content.placeholderDescription)}" title="${esc(node.content.placeholderDescription)}"`;
          style+=fontCss({...s,textDecoration:'none'},false)+`line-height:${s.lineHeight??'normal'};white-space:${s.whiteSpace==='pre'?'pre':'pre-wrap'};overflow:visible;text-align:${s.align??'left'};padding:${p.top??0}px ${p.right??0}px ${p.bottom??0}px ${p.left??0}px;align-content:${({mid:'center',bottom:'end'})[s.valign]??'start'};`;
          if(s.tabSize!==undefined)style+=`tab-size:${number(s.tabSize)}px;`;
          if(s.fill&&s.fill!=='transparent')style+=`background:${colorWithOpacity(s.fill,s.fillOpacity??1)};`;
          if(s.strokeWidth>0&&s.stroke&&s.stroke!=='transparent'){
            // PowerPoint centers shape strokes on the boundary; CSS borders
            // otherwise shrink the text area and change its line wrapping.
            extra+=' data-pptx-shape-outline=""';
            style+=`outline:${number(s.strokeWidth)}px solid ${colorWithOpacity(s.stroke,s.strokeOpacity??1)};outline-offset:${number(-s.strokeWidth/2)}px;`;
          }
          if(s.radius>0)style+=`border-radius:${number(s.radius)}px;`;
          body=node.content.paragraphs?paragraphMarkup(node.content.paragraphs,s):runMarkup(node.content.runs??[{text:node.content.text,style:s}]);
          const allowance=s.align==='left'&&(node.content.paragraphs??[]).every(p=>p.align==='left')?s.wrapAllowance??0:0;
          if(['mid','bottom'].includes(s.valign)||allowance>0){
            // Block align-content falls back to the top when text overflows.
            // Office keeps the complete text block centered or bottom-aligned.
            if(['mid','bottom'].includes(s.valign))style+=`display:flex;flex-direction:column;justify-content:${s.valign==='mid'?'center':'flex-end'};`;
            body=`<span data-pptx-text-content style="display:block;flex:0 0 auto;min-width:0;${allowance?`margin-right:${number(-allowance)}px;`:''}">${body}</span>`;
          }

  Object.assign(output,{style,body,tag,extra});
}
function writeTable(ctx,output){
  const {node,s,b,slide,assets,usedIds,originDocument}=ctx;
  let {style,body,tag,extra}=output;
          tag='table';style+='table-layout:fixed;border-collapse:collapse;border-spacing:0;';
          const c=node.content;
          const gridWidth=c.columnWidthsCssPx.reduce((a,v)=>a+v,0),gridHeight=c.rowHeightsCssPx.reduce((a,v)=>a+v,0);
          body='<colgroup>'+c.columnWidthsCssPx.map(w=>`<col style="width:${number(w/gridWidth*100)}%">`).join('')+'</colgroup><tbody>';
          for(let row=0;row<c.rows;row++){
            body+=`<tr style="height:${number(c.rowHeightsCssPx[row]/gridHeight*100)}%">`;
            for(const cell of c.cells.filter(c=>c.row===row).sort((a,b)=>a.col-b.col)){
              const cs=cell.style,p=cs.padding;
              // Collapsed HTML borders occupy layout space while PowerPoint's
              // centered strokes do not. Avoid growing each content-sized row.
              const borderInset=(side,padding)=>{const b=cs.borders[side];return b.style==='solid'&&b.width>0?Math.min(padding,Math.max(1,Math.floor(b.width))/2):0;};
              // For horizontal layout, only compensate fractional borders that
              // Chromium inflates to a whole pixel; full-pixel lines already
              // match the native table's text inset more closely unchanged.
              const horizontalInset=(side,padding)=>cs.borders[side].width<1?borderInset(side,padding):0;
              const topInset=borderInset('top',p.top),rightInset=horizontalInset('right',p.right),bottomInset=borderInset('bottom',p.bottom),leftInset=horizontalInset('left',p.left);
              let css=fontCss({...cs,textDecoration:'none'},false)+`box-sizing:border-box;background:${cs.fill??'transparent'};line-height:${cs.lineHeight??'normal'};white-space:pre-wrap;text-align:${cs.align??'left'};vertical-align:${cs.valign==='mid'?'middle':cs.valign??'top'};padding:${number(p.top-topInset)}px ${number(p.right-rightInset)}px ${number(p.bottom-bottomInset)}px ${number(p.left-leftInset)}px;`;
              let nativeBorders='';
              for(const side of ['top','right','bottom','left']){
                const border=cs.borders[side],thin=border.style==='solid'&&border.width>=.25&&border.width<=.5&&/^#[a-f\d]{6}$/i.test(border.color??'');
                let paint=border.color;
                if(thin){
                  const rgb=[1,3,5].map(i=>parseInt(border.color.slice(i,i+2),16));paint=`rgba(${rgb.join(',')},0.5)`;
                  nativeBorders+=` data-pptx-native-border-${side}="${number(border.width)} ${border.color.toLowerCase()}"`;
                }
                css+=`border-${side}:${number(border.width)}px ${border.style} ${paint};`;
              }
              const content=cell.paragraphs?paragraphMarkup(cell.paragraphs,cs,false):runMarkup(cell.runs.length?cell.runs:[{text:cell.text,style:cs}]);
              body+=`<td${cell.rowSpan>1?` rowspan="${cell.rowSpan}"`:''}${cell.colSpan>1?` colspan="${cell.colSpan}"`:''}${topInset?` data-pptx-inset-top="${number(topInset)}"`:''}${rightInset?` data-pptx-inset-right="${number(rightInset)}"`:''}${bottomInset?` data-pptx-inset-bottom="${number(bottomInset)}"`:''}${leftInset?` data-pptx-inset-left="${number(leftInset)}"`:''}${nativeBorders} style="${esc(css)}">${content}</td>`;
            }
            body+='</tr>';
          }
          body+='</tbody>';

  Object.assign(output,{style,body,tag,extra});
}
function writeChart(ctx,output){
  const {node,s,b,slide,assets,usedIds,originDocument}=ctx;
  let {style,body,tag,extra}=output;
          tag='figure';style+=fontCss(s)+'margin:0;';
          const id=JSON.stringify([slide.id,node.id]),svg=renderChartSvg(node.content,{width:b.w,height:b.h,fontFamily:s.fontFamily,fontSize:s.fontSize}),hashes=chartHashes(node.content,svg);
          extra=` data-chart-id="${esc(id)}" data-chart-json-sha256="${hashes.json}" data-chart-svg-sha256="${hashes.svg}"`;
          const json=JSON.stringify(node.content).replaceAll('<','\\u003c');
          body=svg+`<script type="application/json" data-chart-for="${esc(id)}">${json}</script>`;

  Object.assign(output,{style,body,tag,extra});
}
function writeShape(ctx,output){
  const {node,s,b,slide,assets,usedIds,originDocument}=ctx;
  let {style,body,tag,extra}=output;
          style+=backgroundCss(s,assets,'backgroundImage',b)+`border:${s.strokeWidth??0}px solid ${colorWithOpacity(s.stroke??'transparent',s.strokeOpacity??1)};border-radius:${s.shape==='ellipse'?'50%':(s.radius??0)+'px'};`;
          extra=` data-geometry="${s.svgGeometry?'centerline':'css-border'}" data-shape="${s.shape==='ellipse'?'ellipse':'rect'}"`;

  Object.assign(output,{style,body,tag,extra});
}
function writeLine(ctx,output){
  const {node,s,b,slide,assets,usedIds,originDocument}=ctx;
  let {style,body,tag,extra}=output;
          if(s.connectorPoints){
            if(!Array.isArray(s.connectorPoints)||s.connectorPoints.length!==4||!s.connectorPoints.every(p=>Number.isFinite(p?.x)&&Number.isFinite(p?.y)))throw new Error('연결선 좌표가 잘못되었습니다.');
            return s.connectorPoints.slice(0,-1).map((start,index)=>{
              const end=s.connectorPoints[index+1],x=Math.min(start.x,end.x),y=Math.min(start.y,end.y),w=Math.abs(end.x-start.x),h=Math.abs(end.y-start.y);
              let id=node.id;if(index){id=`${node.id}::segment-${index+1}`;while(usedIds.has(id))id+='-';usedIds.add(id);}
              const css=`position:absolute;box-sizing:border-box;left:${number(x)}px;top:${number(y)}px;width:${number(w)}px;height:${number(h)}px;z-index:${node.zOrder};opacity:${s.opacity??1};`;
              return `<div data-source-id="${esc(id)}" data-kind="line"${node.parentId?` data-parent-id="${esc(node.parentId)}"`:''} style="${esc(css)}"><svg xmlns="http://www.w3.org/2000/svg" width="${number(Math.max(1,w))}" height="${number(Math.max(1,h))}" style="overflow:visible;display:block"><line x1="${number(start.x-x)}" y1="${number(start.y-y)}" x2="${number(end.x-x)}" y2="${number(end.y-y)}" style="stroke:${esc(s.stroke??'#000')};stroke-width:${number(s.strokeWidth??1)}px"/></svg></div>`;
            }).join('\n');
          }
          const e=s.endpoints??{start:{x:b.x,y:b.y},end:{x:b.x+b.w,y:b.y+b.h}};
          const nativeWidth=s.strokeWidth??1,visibleWidth=s.renderStrokeWidth??nativeWidth,adjusted=visibleWidth!==nativeWidth;
          body=`<svg xmlns="http://www.w3.org/2000/svg" width="${number(Math.max(1,b.w))}" height="${number(Math.max(1,b.h))}" style="overflow:visible;display:block"><line${adjusted?` data-pptx-native-width="${number(nativeWidth)}"`:''} x1="${number(e.start.x-b.x)}" y1="${number(e.start.y-b.y)}" x2="${number(e.end.x-b.x)}" y2="${number(e.end.y-b.y)}" style="stroke:${esc(s.stroke??'#000')};stroke-width:${number(visibleWidth)}px;${adjusted?'shape-rendering:crispEdges;':''}"/></svg>`;

  Object.assign(output,{style,body,tag,extra});
}
function writeImage(ctx,output){
  const {node,s,b,slide,assets,usedIds,originDocument}=ctx;
  let {style,body,tag,extra}=output;
          const asset=assets.get(node.content.assetId);if(!asset.mime.startsWith('image/'))throw new Error('그림에는 그림 자산만 사용할 수 있습니다.');
          if(node.content.editableVector){
            const vector=presetVectorMarkup(node.content.editableVector);
            tag='svg';extra=` xmlns="http://www.w3.org/2000/svg" data-preset-vector="" data-pptx-vector="${esc(JSON.stringify(node.content.editableVector))}" viewBox="0 0 ${number(vector.width)} ${number(vector.height)}"`;body=vector.body;style+='display:block;';
          }else{
            tag='img';extra=` alt="" src="${esc(assetDataUrl(asset))}"${node.content.duotoneBaked?' data-duotone-baked=""':''}${node.content.snippedTextBackground?' data-snip-background=""':''}`;style+='display:block;';
            if(node.content.rounding){extra+=' data-pptx-rounded-image=""';style+='border-radius:50%;';}
            const crop=node.content.crop;if(crop)style+=`object-fit:fill;object-view-box:inset(${number(crop.top/1000)}% ${number(crop.right/1000)}% ${number(crop.bottom/1000)}% ${number(crop.left/1000)}%);`;
          }

  Object.assign(output,{style,body,tag,extra});
}
function writeGroup(ctx,output){output.style+='pointer-events:none;';}
const writers={text:writeText,table:writeTable,chart:writeChart,shape:writeShape,line:writeLine,image:writeImage,group:writeGroup};

// Ordering, ID allocation and shared context belong to the slide traversal.
export function writeHtmlObject(node,{slide,assets,usedIds,originDocument}){
  const b=node.boxCssPx,s=node.style;
  const output={style:`position:absolute;box-sizing:border-box;left:${number(b.x)}px;top:${number(b.y)}px;width:${number(b.w)}px;height:${number(b.h)}px;z-index:${node.zOrder};opacity:${s.opacity??1};`,body:'',tag:'div',extra:''};
  const writer=writers[node.kind];
  if(!writer)throw new Error('지원하지 않는 HTML 출력 객체입니다.');
  const complete=writer({node,s,b,slide,assets,usedIds,originDocument},output);
  if(complete!==undefined)return complete;
  let {style,body,tag,extra}=output;
  if(s.shadow){const sh=s.shadow;style+=`box-shadow:${number(sh.x)}px ${number(sh.y)}px ${number(sh.blur)}px rgba(${sh.r},${sh.g},${sh.b},${number(sh.alpha)});`;extra+=' data-pptx-outer-shadow=""';}
  return `<${tag} data-source-id="${esc(node.id)}" data-kind="${node.kind}"${node.parentId?` data-parent-id="${esc(node.parentId)}"`:''}${extra} style="${esc(style)}">${tag==='img'?'':body+`</${tag}>`}`;
}
