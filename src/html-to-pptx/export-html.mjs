import {validateModel,diagnostic} from './model.mjs';
import {assetDataUrl} from './assets.mjs';
import {canvasSize} from './canvas.mjs';
const esc=v=>String(v??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const number=v=>String(Number(Number(v).toFixed(10)));
const fontCss=s=>`font-family:${JSON.stringify(s.fontFamily)};font-size:${number(s.fontSize)}px;font-weight:${s.fontWeight??400};font-style:${s.fontStyle??'normal'};color:${s.color??'#000'};letter-spacing:${s.letterSpacing==='normal'||s.letterSpacing==null?'normal':number(parseFloat(s.letterSpacing))+'px'};text-decoration:${s.textDecoration??'none'};${s.outline?`-webkit-text-stroke:${number(s.outline.width)}px ${s.outline.color};`:''}`;

// Current content lives once, in the DOM. Metadata contains only document structure.
export function exportHtml(model){
  const diagnostics=validateModel(model).errors.map(e=>diagnostic('INPUT_INVALID',`${e.path}: ${e.message}`));
  if(diagnostics.length)return {status:'failed',html:null,diagnostics};
  try{
    const size=canvasSize(model),assets=new Map(model.assets.map(a=>[a.id,a]));
    const slides=model.slides.map(slide=>{
      const markup=[...slide.nodes].sort((a,b)=>a.zOrder-b.zOrder).map(node=>{
        const b=node.boxCssPx,s=node.style;
        let style=`position:absolute;box-sizing:border-box;left:${number(b.x)}px;top:${number(b.y)}px;width:${number(b.w)}px;height:${number(b.h)}px;z-index:${node.zOrder};opacity:${s.opacity??1};`,body='',tag='div',extra='';
        if(node.kind==='text'){
          const p=s.padding??{};
          style+=fontCss(s)+`line-height:${s.lineHeight??'normal'};white-space:pre-wrap;overflow:visible;text-align:${s.align??'left'};padding:${p.top??0}px ${p.right??0}px ${p.bottom??0}px ${p.left??0}px;align-content:${({mid:'center',bottom:'end'})[s.valign]??'start'};`;
          body=(node.content.runs??[{text:node.content.text,style:s}]).map(r=>`<span style="${esc(fontCss(r.style))}">${esc(r.text)}</span>`).join('');
        }else if(node.kind==='shape'){
          style+=`background:${s.fill??'transparent'};border:${s.strokeWidth??0}px solid ${s.stroke??'transparent'};border-radius:${s.shape==='ellipse'?'50%':(s.radius??0)+'px'};`;
          extra=` data-geometry="${s.svgGeometry?'centerline':'css-border'}" data-shape="${s.shape==='ellipse'?'ellipse':'rect'}"`;
        }else if(node.kind==='line'){
          const e=s.endpoints??{start:{x:b.x,y:b.y},end:{x:b.x+b.w,y:b.y+b.h}};
          body=`<svg xmlns="http://www.w3.org/2000/svg" width="${number(Math.max(1,b.w))}" height="${number(Math.max(1,b.h))}" style="overflow:visible;display:block"><line x1="${number(e.start.x-b.x)}" y1="${number(e.start.y-b.y)}" x2="${number(e.end.x-b.x)}" y2="${number(e.end.y-b.y)}" style="stroke:${esc(s.stroke??'#000')};stroke-width:${number(s.strokeWidth??1)}px"/></svg>`;
        }else if(node.kind==='image'){
          tag='img';extra=` alt="" src="${esc(assetDataUrl(assets.get(node.content.assetId)))}"`;style+='display:block;';
        }else if(node.kind==='group')style+='pointer-events:none;';
        else throw new Error('지원하지 않는 HTML 출력 객체입니다.');
        return `<${tag} data-source-id="${esc(node.id)}" data-kind="${node.kind}"${node.parentId?` data-parent-id="${esc(node.parentId)}"`:''}${extra} style="${esc(style)}">${tag==='img'?'':body+`</${tag}>`}`;
      }).join('\n');
      const w=number(slide.widthCssPx),h=number(slide.heightCssPx);
      return `<div class="slide-view" style="max-width:${w}px;aspect-ratio:${w}/${h}"><section data-slide data-slide-id="${esc(slide.id)}" data-ppt-width-emu="${size.width}" data-ppt-height-emu="${size.height}" style="position:relative;width:${w}px;height:${h}px;background:${esc(slide.background?.color??'#fff')};transform:scale(min(1,calc((100vw - 32px)/${w}px)));transform-origin:0 0;overflow:hidden">${markup}</section></div>`;
    }).join('\n');
    const html=`<!doctype html><html lang="ko" data-slidehandoff="1"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><meta http-equiv="Content-Security-Policy" content="default-src 'none'; style-src 'unsafe-inline'; img-src data:; font-src data:; script-src 'none'; base-uri 'none'"><title>편집 가능한 장표</title><style>html,body{margin:0;padding:0}body{padding:16px;background:#e5e7eb}.slide-view{width:100%;margin:0 auto 24px;overflow:hidden}*{font-synthesis:none}</style></head><body>${slides}</body></html>`;
    if(Buffer.byteLength(html)>16*1024*1024)throw new Error('출력 HTML이 재입력 한도 16MiB를 초과합니다.');
    return {status:'generated',html,diagnostics};
  }catch(e){return {status:'failed',html:null,diagnostics:[diagnostic('EXPORT_FAILED',e.message)]};}
}
