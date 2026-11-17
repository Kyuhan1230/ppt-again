// CSS pixels describe the document, not the size of a browser window.
export const EMU_PER_POINT=12700, EMU_PER_CSS_PX=9525, MIN_SLIDE_EMU=914400, MAX_SLIDE_EMU=51206400;
// Chromium can quantize each measured CSS edge by approximately 1/64px.
// Compare aspect ratios in CSS pixels, allowing that uncertainty on both axes.
export function sameCanvasAspect(size,slide){
  const width=slide?.widthCssPx,height=slide?.heightCssPx;
  if(!Number.isFinite(width)||!Number.isFinite(height)||width<=0||height<=0)return false;
  const epsilon=Math.min(.02,Math.min(width,height)*.001);
  return Math.abs(width*size.height-height*size.width)<=epsilon*(size.width+size.height);
}
export function canvasSize(deck){
  const first=deck.slides[0];
  if(deck.pptxSizeEmu)return deck.pptxSizeEmu;
  const scale=960*EMU_PER_POINT/Math.max(first.widthCssPx,first.heightCssPx);
  return {width:Math.round(first.widthCssPx*scale),height:Math.round(first.heightCssPx*scale)};
}
export function canvasDiagnostics(deck){
  if(!deck.slides?.length)return [];
  const size=deck.pptxSizeEmu===undefined?canvasSize(deck):deck.pptxSizeEmu;
  if(!size||!['width','height'].every(k=>Number.isSafeInteger(size[k])&&size[k]>=MIN_SLIDE_EMU&&size[k]<=MAX_SLIDE_EMU))return ['PPT dimensions must be integer EMU between 1 and 56 inches'];
  if(deck.pptxSizeEmu!==undefined&&deck.slides.some(s=>!sameCanvasAspect(size,s)))return ['Original PPT and logical canvas aspect ratios differ'];
  return [];
}
