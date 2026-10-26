// CSS pixels describe the document, not the size of a browser window.
export const EMU_PER_POINT=12700, EMU_PER_CSS_PX=9525;
export function canvasSize(deck){
  const first=deck.slides[0];
  if(deck.pptxSizeEmu)return deck.pptxSizeEmu;
  const scale=960*EMU_PER_POINT/Math.max(first.widthCssPx,first.heightCssPx);
  return {width:Math.round(first.widthCssPx*scale),height:Math.round(first.heightCssPx*scale)};
}
export function canvasDiagnostics(deck){
  const size=deck.pptxSizeEmu;if(size===undefined)return [];
  if(!size||!['width','height'].every(k=>Number.isSafeInteger(size[k])&&size[k]>0&&size[k]<=51206400))return ['Original PPT dimensions must be positive integer EMU within 56 inches'];
  if(deck.slides.some(s=>Math.abs(size.width/s.widthCssPx-size.height/s.heightCssPx)>0.001))return ['Original PPT and logical canvas aspect ratios differ'];
  return [];
}
