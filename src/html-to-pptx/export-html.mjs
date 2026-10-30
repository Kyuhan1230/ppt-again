import {validateModel,diagnostic} from './model.mjs';
import {assetDataUrl} from './assets.mjs';
import {canvasSize} from './canvas.mjs';
import {renderChartSvg} from './charts.mjs';
import {chartHashes} from './chart-input.mjs';
const esc=v=>String(v??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const number=v=>String(Number(Number(v).toFixed(10)));
const fontCss=s=>`font-family:${JSON.stringify(s.fontFamily)};font-size:${number(s.fontSize)}px;font-weight:${s.fontWeight??400};font-style:${s.fontStyle??'normal'};color:${s.color??'#000'};letter-spacing:${s.letterSpacing==='normal'||s.letterSpacing==null?'normal':number(parseFloat(s.letterSpacing))+'px'};text-decoration:${s.textDecoration??'none'};${s.outline?`-webkit-text-stroke:${number(s.outline.width)}px ${s.outline.color};`:''}`;

// Current content lives once, in the DOM. Metadata contains only document structure.
export function exportHtml(model){
  const diagnostics=validateModel(model).errors.map(e=>diagnostic('INPUT_INVALID',`${e.path}: ${e.message}`));
  if(diagnostics.length)return {status:'failed',html:null,diagnostics};
  try{
    // PPT embedding permission does not grant web redistribution permission.
    for(const family of new Set((model.fonts??[]).map(f=>f.family)))diagnostics.push(diagnostic('FONT_NOT_PORTABLE',`${family}: 글꼴 파일은 HTML에 포함하지 않았습니다.`,{},'warning'));
    const size=canvasSize(model),assets=new Map(model.assets.map(a=>[a.id,a]));
    const slides=model.slides.map(slide=>{
      const markup=[...slide.nodes].sort((a,b)=>a.zOrder-b.zOrder).map(node=>{
        const b=node.boxCssPx,s=node.style;
        let style=`position:absolute;box-sizing:border-box;left:${number(b.x)}px;top:${number(b.y)}px;width:${number(b.w)}px;height:${number(b.h)}px;z-index:${node.zOrder};opacity:${s.opacity??1};`,body='',tag='div',extra='';
        if(node.kind==='text'){
          const p=s.padding??{};
          style+=fontCss(s)+`line-height:${s.lineHeight??'normal'};white-space:pre-wrap;overflow:visible;text-align:${s.align??'left'};padding:${p.top??0}px ${p.right??0}px ${p.bottom??0}px ${p.left??0}px;align-content:${({mid:'center',bottom:'end'})[s.valign]??'start'};`;
          body=(node.content.runs??[{text:node.content.text,style:s}]).map(r=>`<span style="${esc(fontCss(r.style))}">${esc(r.text)}</span>`).join('');
        }else if(node.kind==='table'){
          tag='table';style+='table-layout:fixed;border-collapse:collapse;border-spacing:0;';
          const c=node.content;
          const gridWidth=c.columnWidthsCssPx.reduce((a,v)=>a+v,0),gridHeight=c.rowHeightsCssPx.reduce((a,v)=>a+v,0);
          body='<colgroup>'+c.columnWidthsCssPx.map(w=>`<col style="width:${number(w/gridWidth*100)}%">`).join('')+'</colgroup><tbody>';
          for(let row=0;row<c.rows;row++){
            body+=`<tr style="height:${number(c.rowHeightsCssPx[row]/gridHeight*100)}%">`;
            for(const cell of c.cells.filter(c=>c.row===row).sort((a,b)=>a.col-b.col)){
              const cs=cell.style,p=cs.padding;
              let css=fontCss(cs)+`box-sizing:border-box;background:${cs.fill??'transparent'};line-height:${cs.lineHeight??'normal'};white-space:pre-wrap;text-align:${cs.align??'left'};vertical-align:${cs.valign==='mid'?'middle':cs.valign??'top'};padding:${p.top}px ${p.right}px ${p.bottom}px ${p.left}px;`;
              for(const side of ['top','right','bottom','left']){const border=cs.borders[side];css+=`border-${side}:${number(border.width)}px ${border.style} ${border.color};`;}
              const content=(cell.runs.length?cell.runs:[{text:cell.text,style:cs}]).map(r=>`<span style="${esc(fontCss(r.style))}">${esc(r.text)}</span>`).join('');
              body+=`<td${cell.rowSpan>1?` rowspan="${cell.rowSpan}"`:''}${cell.colSpan>1?` colspan="${cell.colSpan}"`:''} style="${esc(css)}">${content}</td>`;
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
          const e=s.endpoints??{start:{x:b.x,y:b.y},end:{x:b.x+b.w,y:b.y+b.h}};
          body=`<svg xmlns="http://www.w3.org/2000/svg" width="${number(Math.max(1,b.w))}" height="${number(Math.max(1,b.h))}" style="overflow:visible;display:block"><line x1="${number(e.start.x-b.x)}" y1="${number(e.start.y-b.y)}" x2="${number(e.end.x-b.x)}" y2="${number(e.end.y-b.y)}" style="stroke:${esc(s.stroke??'#000')};stroke-width:${number(s.strokeWidth??1)}px"/></svg>`;
        }else if(node.kind==='image'){
          const asset=assets.get(node.content.assetId);if(!asset.mime.startsWith('image/'))throw new Error('그림에는 그림 자산만 사용할 수 있습니다.');
          tag='img';extra=` alt="" src="${esc(assetDataUrl(asset))}"`;style+='display:block;';
        }else if(node.kind==='group')style+='pointer-events:none;';
        else throw new Error('지원하지 않는 HTML 출력 객체입니다.');
        return `<${tag} data-source-id="${esc(node.id)}" data-kind="${node.kind}"${node.parentId?` data-parent-id="${esc(node.parentId)}"`:''}${extra} style="${esc(style)}">${tag==='img'?'':body+`</${tag}>`}`;
      }).join('\n');
      const w=number(slide.widthCssPx),h=number(slide.heightCssPx);
      return `<div class="slide-view" style="max-width:${w}px;aspect-ratio:${w}/${h}"><section data-slide data-slide-id="${esc(slide.id)}" data-ppt-width-emu="${size.width}" data-ppt-height-emu="${size.height}" style="position:relative;width:${w}px;height:${h}px;background:${esc(slide.background?.color??'#fff')};transform:scale(min(1,calc((100vw - 32px)/${w}px)));transform-origin:0 0;overflow:hidden">${markup}</section></div>`;
    }).join('\n');
    const html=`<!doctype html><html lang="ko" data-slidehandoff="${model.version}"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><meta http-equiv="Content-Security-Policy" content="default-src 'none'; style-src 'unsafe-inline'; img-src data:; font-src data:; script-src 'none'; base-uri 'none'"><title>편집 가능한 장표</title><style>html,body{margin:0;padding:0}body{padding:16px;background:#e5e7eb}.slide-view{width:100%;margin:0 auto 24px;overflow:hidden}*{font-synthesis:none}</style></head><body>${slides}</body></html>`;
    if(Buffer.byteLength(html)>16*1024*1024)throw new Error('출력 HTML이 재입력 한도 16MiB를 초과합니다.');
    return {status:'generated',html,diagnostics};
  }catch(e){return {status:'failed',html:null,diagnostics:[diagnostic('EXPORT_FAILED',e.message)]};}
}
