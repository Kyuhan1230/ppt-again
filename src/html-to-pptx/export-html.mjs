import {validateModel,diagnostic} from './model.mjs';
import {canvasSize} from './canvas.mjs';
import {esc,number,writeHtmlObject,slideBackgroundCss} from './html-object-writers.mjs';
// Current content lives once, in the DOM. Metadata contains only document structure.
export function exportHtml(model){
  const diagnostics=validateModel(model).errors.map(e=>diagnostic('INPUT_INVALID',`${e.path}: ${e.message}`));
  if(diagnostics.length)return {status:'failed',html:null,diagnostics};
  try{
    // PPT embedding permission does not grant web redistribution permission.
    for(const family of new Set((model.fonts??[]).map(f=>f.family)))diagnostics.push(diagnostic('FONT_NOT_PORTABLE',`${family}: 글꼴 파일은 HTML에 포함하지 않았습니다.`,{},'warning'));
    const size=canvasSize(model),assets=new Map(model.assets.map(a=>[a.id,a]));
    // JPEG marker validation cannot prove complete pixel decoding in this
    // synchronous writer. Keep the limitation visible at the affected object.
    for(const slide of model.slides){
      if(assets.get(slide.background?.assetId)?.mime==='image/jpeg')diagnostics.push(diagnostic('BACKGROUND_JPEG_UNVERIFIED','JPEG 배경의 그림 전체를 읽을 수 있는지는 확인하지 못했습니다.',{slideId:slide.id},'warning'));
      for(const node of slide.nodes)if(node.kind==='shape'&&assets.get(node.style?.assetId)?.mime==='image/jpeg')diagnostics.push(diagnostic('BACKGROUND_JPEG_UNVERIFIED','JPEG 배경의 그림 전체를 읽을 수 있는지는 확인하지 못했습니다.',{slideId:slide.id,nodeId:node.id},'warning'));
    }
    // This is relation metadata, not a source of text or an authenticity
    // claim. The visible DOM remains the only source of editable content.
    const originDocuments=new Set(model.slides.flatMap(slide=>slide.nodes.map(node=>node.pptxOrigin?.documentId).filter(Boolean)));
    const originDocument=originDocuments.size===1?[...originDocuments][0]:null;
    if(originDocuments.size>1)diagnostics.push(diagnostic('PPTX_OBJECT_SPLIT','서로 다른 문서의 객체 관계를 함께 확인할 수 없어 현재 객체를 분리해 유지합니다.',{},'warning'));
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
      const markup=[...slide.nodes].sort((a,b)=>a.zOrder-b.zOrder).map(node=>writeHtmlObject(node,{slide,assets,usedIds,originDocument})).join('\n');
      const w=number(slide.widthCssPx),h=number(slide.heightCssPx);
      return `<div class="slide-view" style="max-width:${w}px;aspect-ratio:${w}/${h}"><section data-slide data-slide-id="${esc(slide.id)}" data-ppt-width-emu="${size.width}" data-ppt-height-emu="${size.height}"${slide.animationOmitted?' data-animation-omitted=""':''}${slide.transition?` data-pptx-transition="${esc(JSON.stringify(slide.transition))}"`:''} style="position:relative;width:${w}px;height:${h}px;${esc(slideBackgroundCss(slide.background,assets,{w:slide.widthCssPx,h:slide.heightCssPx}))}transform:scale(min(1,calc((100vw - 32px)/${w}px)));transform-origin:0 0;overflow:hidden">${markup}</section></div>`;
    }).join('\n');
    const missing=model.slides.flatMap((slide,index)=>slide.nodes.filter(node=>node.content.placeholderDescription).map(node=>({slide:index+1,description:node.content.placeholderDescription})));
    const missingSummary=missing.length?`<details data-placeholder-summary style="max-width:60rem;margin:1rem auto;padding:1rem;background:#fff;color:#7f1d1d"><summary>변환하지 못한 항목 ${missing.length}개 확인</summary><ol>${missing.map(item=>`<li>${item.slide}번 장표: ${esc(item.description)}</li>`).join('')}</ol></details>`:'';
    // In Edge, compositing each slide avoids coloured subpixel text fringes
    // that differ from PowerPoint's grayscale slide rendering. This rendering
    // hint leaves the editable text, geometry and opacity unchanged.
    const html=`<!doctype html><html lang="ko" data-slidehandoff="${model.version}"${originDocument?` data-pptx-origin-document="${esc(originDocument)}"`:''}><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><meta http-equiv="Content-Security-Policy" content="default-src 'none'; style-src 'unsafe-inline'; img-src data:; font-src data:; script-src 'none'; base-uri 'none'"><title>편집 가능한 장표</title><style>html,body{margin:0;padding:0}body{padding:16px;background:#e5e7eb}.slide-view{width:100%;margin:0 auto 24px;overflow:hidden}section[data-slide]{will-change:transform}*{font-synthesis:weight style}${localFaceCss}</style></head><body>${slides}${missingSummary}</body></html>`;
    if(Buffer.byteLength(html)>16*1024*1024)throw new Error('출력 HTML이 재입력 한도 16MiB를 초과합니다.');
    return {status:'generated',html,diagnostics};
  }catch(e){return {status:'failed',html:null,diagnostics:[diagnostic('EXPORT_FAILED',e.message)]};}
}
