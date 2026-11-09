import {validateModel,diagnostic} from './model.mjs';
import {assetDataUrl} from './assets.mjs';
import {canvasSize} from './canvas.mjs';
import {renderChartSvg} from './charts.mjs';
import {chartHashes} from './chart-input.mjs';
import {presetVectorMarkup} from './preset-vector.mjs';
const esc=v=>String(v??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const number=v=>String(Number(Number(v).toFixed(10)));
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

// Current content lives once, in the DOM. Metadata contains only document structure.
export function exportHtml(model){
  const diagnostics=validateModel(model).errors.map(e=>diagnostic('INPUT_INVALID',`${e.path}: ${e.message}`));
  if(diagnostics.length)return {status:'failed',html:null,diagnostics};
  try{
    // PPT embedding permission does not grant web redistribution permission.
    for(const family of new Set((model.fonts??[]).map(f=>f.family)))diagnostics.push(diagnostic('FONT_NOT_PORTABLE',`${family}: 글꼴 파일은 HTML에 포함하지 않았습니다.`,{},'warning'));
    const size=canvasSize(model),assets=new Map(model.assets.map(a=>[a.id,a]));
    // PowerPoint can name a full installed face as a family (e.g. "Arial Bold").
    // Browsers may only match its typographic family unless local() names the face.
    const localFaces=new Set(),addFace=style=>{const family=style?.fontFamily;if(localFaces.size<32&&typeof family==='string'&&/^[\p{L}\p{N} ._-]{1,150} (?:Bold|ExtraBold|SemiBold|DemiBold|UltraBold|Black|Heavy)$/iu.test(family))localFaces.add(family);};
    for(const slide of model.slides)for(const node of slide.nodes){
      addFace(node.style);for(const run of node.content?.runs??[])addFace(run.style);
      for(const cell of node.content?.cells??[]){addFace(cell.style);for(const run of cell.runs??[])addFace(run.style);}
    }
    const localFaceCss=[...localFaces].map(family=>`@font-face{font-family:${JSON.stringify(family)};src:local(${JSON.stringify(family)});font-weight:400;font-style:normal}`).join('');
    const slides=model.slides.map(slide=>{
      const usedIds=new Set(slide.nodes.map(node=>node.id));
      const markup=[...slide.nodes].sort((a,b)=>a.zOrder-b.zOrder).map(node=>{
        const b=node.boxCssPx,s=node.style;
        let style=`position:absolute;box-sizing:border-box;left:${number(b.x)}px;top:${number(b.y)}px;width:${number(b.w)}px;height:${number(b.h)}px;z-index:${node.zOrder};opacity:${s.opacity??1};`,body='',tag='div',extra='';
        if(node.kind==='text'){
          const p=s.padding??{};
          if(node.content.placeholderDescription)extra=` data-placeholder-marker="" role="img" tabindex="0" aria-label="${esc(node.content.placeholderDescription)}" title="${esc(node.content.placeholderDescription)}"`;
          style+=fontCss({...s,textDecoration:'none'},false)+`line-height:${s.lineHeight??'normal'};white-space:${s.whiteSpace==='pre'?'pre':'pre-wrap'};overflow:visible;text-align:${s.align??'left'};padding:${p.top??0}px ${p.right??0}px ${p.bottom??0}px ${p.left??0}px;align-content:${({mid:'center',bottom:'end'})[s.valign]??'start'};`;
          if(s.tabSize!==undefined)style+=`tab-size:${number(s.tabSize)}px;`;
          if(s.fill&&s.fill!=='transparent')style+=`background:${s.fill};`;
          if(s.strokeWidth>0&&s.stroke&&s.stroke!=='transparent'){
            // PowerPoint centers shape strokes on the boundary; CSS borders
            // otherwise shrink the text area and change its line wrapping.
            extra+=' data-pptx-shape-outline=""';
            style+=`outline:${number(s.strokeWidth)}px solid ${s.stroke};outline-offset:${number(-s.strokeWidth/2)}px;`;
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
        }else if(node.kind==='table'){
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
        }else if(node.kind==='chart'){
          tag='figure';style+=fontCss(s)+'margin:0;';
          const id=JSON.stringify([slide.id,node.id]),svg=renderChartSvg(node.content,{width:b.w,height:b.h,fontFamily:s.fontFamily,fontSize:s.fontSize}),hashes=chartHashes(node.content,svg);
          extra=` data-chart-id="${esc(id)}" data-chart-json-sha256="${hashes.json}" data-chart-svg-sha256="${hashes.svg}"`;
          const json=JSON.stringify(node.content).replaceAll('<','\\u003c');
          body=svg+`<script type="application/json" data-chart-for="${esc(id)}">${json}</script>`;
        }else if(node.kind==='shape'){
          style+=`background:${s.fill??'transparent'};border:${s.strokeWidth??0}px solid ${s.stroke??'transparent'};border-radius:${s.shape==='ellipse'?'50%':(s.radius??0)+'px'};`;
          extra=` data-geometry="${s.svgGeometry?'centerline':'css-border'}" data-shape="${s.shape==='ellipse'?'ellipse':'rect'}"`;
        }else if(node.kind==='line'){
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
        }else if(node.kind==='image'){
          const asset=assets.get(node.content.assetId);if(!asset.mime.startsWith('image/'))throw new Error('그림에는 그림 자산만 사용할 수 있습니다.');
          if(node.content.editableVector){
            const vector=presetVectorMarkup(node.content.editableVector);
            tag='svg';extra=` xmlns="http://www.w3.org/2000/svg" data-preset-vector="" data-pptx-vector="${esc(JSON.stringify(node.content.editableVector))}" viewBox="0 0 ${number(vector.width)} ${number(vector.height)}"`;body=vector.body;style+='display:block;';
          }else{
            tag='img';extra=` alt="" src="${esc(assetDataUrl(asset))}"${node.content.duotoneBaked?' data-duotone-baked=""':''}${node.content.snippedTextBackground?' data-snip-background=""':''}`;style+='display:block;';
            if(node.content.rounding){extra+=' data-pptx-rounded-image=""';style+='border-radius:50%;';}
            const crop=node.content.crop;if(crop)style+=`object-fit:fill;object-view-box:inset(${number(crop.top/1000)}% ${number(crop.right/1000)}% ${number(crop.bottom/1000)}% ${number(crop.left/1000)}%);`;
          }
        }else if(node.kind==='group')style+='pointer-events:none;';
        else throw new Error('지원하지 않는 HTML 출력 객체입니다.');
        if(s.shadow){const sh=s.shadow;style+=`box-shadow:${number(sh.x)}px ${number(sh.y)}px ${number(sh.blur)}px rgba(${sh.r},${sh.g},${sh.b},${number(sh.alpha)});`;extra+=' data-pptx-outer-shadow=""';}
        return `<${tag} data-source-id="${esc(node.id)}" data-kind="${node.kind}"${node.parentId?` data-parent-id="${esc(node.parentId)}"`:''}${extra} style="${esc(style)}">${tag==='img'?'':body+`</${tag}>`}`;
      }).join('\n');
      const w=number(slide.widthCssPx),h=number(slide.heightCssPx);
      return `<div class="slide-view" style="max-width:${w}px;aspect-ratio:${w}/${h}"><section data-slide data-slide-id="${esc(slide.id)}" data-ppt-width-emu="${size.width}" data-ppt-height-emu="${size.height}"${slide.animationOmitted?' data-animation-omitted=""':''}${slide.transition?` data-pptx-transition="${esc(JSON.stringify(slide.transition))}"`:''} style="position:relative;width:${w}px;height:${h}px;background:${esc(slide.background?.color??'#fff')};transform:scale(min(1,calc((100vw - 32px)/${w}px)));transform-origin:0 0;overflow:hidden">${markup}</section></div>`;
    }).join('\n');
    const missing=model.slides.flatMap((slide,index)=>slide.nodes.filter(node=>node.content.placeholderDescription).map(node=>({slide:index+1,description:node.content.placeholderDescription})));
    const missingSummary=missing.length?`<details data-placeholder-summary style="max-width:60rem;margin:1rem auto;padding:1rem;background:#fff;color:#7f1d1d"><summary>변환하지 못한 항목 ${missing.length}개 확인</summary><ol>${missing.map(item=>`<li>${item.slide}번 장표: ${esc(item.description)}</li>`).join('')}</ol></details>`:'';
    // In Edge, compositing each slide avoids coloured subpixel text fringes
    // that differ from PowerPoint's grayscale slide rendering. This rendering
    // hint leaves the editable text, geometry and opacity unchanged.
    const html=`<!doctype html><html lang="ko" data-slidehandoff="${model.version}"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><meta http-equiv="Content-Security-Policy" content="default-src 'none'; style-src 'unsafe-inline'; img-src data:; font-src data:; script-src 'none'; base-uri 'none'"><title>편집 가능한 장표</title><style>html,body{margin:0;padding:0}body{padding:16px;background:#e5e7eb}.slide-view{width:100%;margin:0 auto 24px;overflow:hidden}section[data-slide]{will-change:transform}*{font-synthesis:weight style}${localFaceCss}</style></head><body>${slides}${missingSummary}</body></html>`;
    if(Buffer.byteLength(html)>16*1024*1024)throw new Error('출력 HTML이 재입력 한도 16MiB를 초과합니다.');
    return {status:'generated',html,diagnostics};
  }catch(e){return {status:'failed',html:null,diagnostics:[diagnostic('EXPORT_FAILED',e.message)]};}
}
