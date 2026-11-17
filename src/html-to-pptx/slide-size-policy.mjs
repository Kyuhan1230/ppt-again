import {canvasSize,sameCanvasAspect,EMU_PER_CSS_PX,EMU_PER_POINT,MIN_SLIDE_EMU,MAX_SLIDE_EMU} from './canvas.mjs';

// Work on a copy: measurements and captures still describe the source rendering.
export function normalizeSlideSizes(model,measurements,policy={mode:'contain'}){
  const deck=structuredClone(model);
  // Aliases within one slide must stay shared; aliases across slides cannot,
  // because each slide may have a different scale and offset.
  deck.slides=model.slides.map(slide=>structuredClone(slide));
  const measured=measurements.map(measurement=>structuredClone(measurement)),first=deck.slides[0];
  if(!first)throw new Error('장표 크기를 결정할 장표가 없습니다.');
  const original=deck.pptxSizeEmu,custom=policy.targetWidthEmu!==undefined;
  if(original!==undefined&&(!original||!['width','height'].every(key=>Number.isSafeInteger(original[key])&&original[key]>=MIN_SLIDE_EMU&&original[key]<=MAX_SLIDE_EMU)))throw Object.assign(new Error('원본 PPT 장표 크기는 각 변이 1~56인치여야 합니다.'),{code:'INPUT_INVALID'});
  const originalAspectMismatch=Boolean(original&&!deck.slides.some(slide=>sameCanvasAspect(original,slide)));
  if(originalAspectMismatch&&!custom)throw Object.assign(new Error('원본 PPT 크기 정보와 모든 장표의 비율이 다릅니다. 원본 PPT에서 HTML을 다시 만들거나 출력 크기를 직접 지정해 주세요.'),{code:'INPUT_INVALID',reasonKey:'SIZE_ORIGINAL_ASPECT_MISMATCH'});
  const size=custom?{width:policy.targetWidthEmu,height:policy.targetHeightEmu}:canvasSize(deck);
  if(!['width','height'].every(key=>Number.isSafeInteger(size[key])&&size[key]>=MIN_SLIDE_EMU&&size[key]<=MAX_SLIDE_EMU))throw Object.assign(new Error('PPT 장표 크기는 각 변이 1~56인치여야 합니다. 다른 출력 크기를 선택해 주세요.'),{code:'INPUT_INVALID',reasonKey:'SIZE_TARGET_OUT_OF_RANGE'});
  const width=custom?size.width/EMU_PER_CSS_PX:first.widthCssPx,height=custom?size.height/EMU_PER_CSS_PX:original?width*original.height/original.width:first.heightCssPx;
  const oldPointScale=original?EMU_PER_CSS_PX/EMU_PER_POINT:canvasSize(model).width/model.slides[0].widthCssPx/EMU_PER_POINT,newPointScale=size.width/width/EMU_PER_POINT;
  const report={mode:'contain',targetWidthEmu:size.width,targetHeightEmu:size.height,marginColor:policy.marginColor??'#ffffff',source:custom?'custom':original?'original':'first-slide',...(originalAspectMismatch?{originalAspectMismatch:true}:{}),slides:[]};
  const boxes=new WeakSet(),styles=new WeakSet(),paragraphs=new WeakSet();
  const number=(obj,key,s)=>{if(typeof obj?.[key]==='number'&&Number.isFinite(obj[key]))obj[key]*=s;};
  const length=(obj,key,s)=>{const v=obj?.[key];if(typeof v==='number')obj[key]=v*s;else if(typeof v==='string'&&/^-?\d+(?:\.\d+)?px$/.test(v))obj[key]=`${parseFloat(v)*s}px`;};
  const box=(b,s,x,y)=>{if(!b||boxes.has(b))return;boxes.add(b);number(b,'x',s);number(b,'y',s);if(Number.isFinite(b.x))b.x+=x;if(Number.isFinite(b.y))b.y+=y;number(b,'w',s);number(b,'h',s);};
  const style=(st,s,x,y)=>{if(!st||styles.has(st))return;styles.add(st);
    for(const key of ['fontSize','strokeWidth','renderStrokeWidth','radius','wrapAllowance','tabSize'])number(st,key,s);
    for(const key of ['lineHeight','letterSpacing'])length(st,key,s);
    for(const key of ['top','right','bottom','left'])number(st.padding,key,s);
    for(const border of Object.values(st.borders??{}))number(border,'width',s);
    number(st.outline,'width',s);for(const k of ['x','y','blur'])number(st.shadow,k,s);number(st.linePaint,'thickness',s);
    for(const p of Object.values(st.endpoints??{})){if(p&&Number.isFinite(p.x)&&Number.isFinite(p.y)){p.x=p.x*s+x;p.y=p.y*s+y;}}
    if(st.capture)st.capture={...st.capture,renderScale:st.capture.renderScale/s};
  };
  const para=(p,s,physical)=>{if(!p||paragraphs.has(p))return;paragraphs.add(p);
    for(const key of ['beforePx','afterPx','leftPx','rightPx','indentPx'])number(p,key,s);length(p,'lineHeight',s);
    for(const run of p.runs??[])style(run.style,s,0,0);
    for(const value of Object.values(p.sourceSpacing??{}))if(value.kind==='spcPts')value.value=Math.round(value.value*physical);
    if(p.bullet?.sourceSize?.kind==='buSzPts')p.bullet.sourceSize.value=Math.round(p.bullet.sourceSize.value*physical);
    if(p.sourceEnd?.mode==='explicit')for(const key of ['spc','sz','kern'])if(p.sourceEnd.attrs[key]!==undefined)p.sourceEnd.attrs[key]=String(Math.round(Number(p.sourceEnd.attrs[key])*physical));
  };
  for(const slide of deck.slides){
    const sw=slide.widthCssPx,sh=slide.heightCssPx,s=Math.min(width/sw,height/sh),x=Math.max(0,(width-sw*s)/2),y=Math.max(0,(height-sh*s)/2),physical=s*newPointScale/oldPointScale;
    report.slides.push({slideId:slide.id,sourceWidth:sw,sourceHeight:sh,scale:s,physicalScale:s*newPointScale/(EMU_PER_CSS_PX/EMU_PER_POINT),offsetX:x,offsetY:y,offsetXEmu:Math.round(x*newPointScale*EMU_PER_POINT),offsetYEmu:Math.round(y*newPointScale*EMU_PER_POINT)});
    if(slide.background?.capture)slide.background.capture={...slide.background.capture,renderScale:slide.background.capture.renderScale/s};
    if(x>1e-8||y>1e-8){slide.fitBackground={boxCssPx:{x,y,w:sw*s,h:sh*s},background:slide.background};slide.background={color:report.marginColor};}
    for(const node of slide.nodes){box(node.boxCssPx,s,x,y);style(node.style,s,x,y);
      for(const run of node.content.runs??[])style(run.style,s,0,0);
      for(const p of node.content.paragraphs??[])para(p,s,physical);
      for(const key of ['columnWidthsCssPx','rowHeightsCssPx'])if(node.content[key])node.content[key]=node.content[key].map(v=>v*s);
      for(const cell of node.content.cells??[]){style(cell.style,s,0,0);for(const run of cell.runs??[])style(run.style,s,0,0);for(const p of cell.paragraphs??[])para(p,s,physical);}
      if(node.content.editableVector)for(const key of ['width','height','strokeWidth'])number(node.content.editableVector,key,s);
    }
    const m=measured.find(v=>v.slideId===slide.id);
    if(m)for(const object of m.objects??[]){box(object.boxCssPx,s,x,y);for(const line of object.lines??[])box(line,s,x,y);if(Number.isFinite(object.baseline))object.baseline=object.baseline*s+y;number(object,'lineAdvance',s);}
    slide.widthCssPx=width;slide.heightCssPx=height;
  }
  if(custom||original)deck.pptxSizeEmu=size;
  return {model:deck,measurements:measured,sizePolicy:report};
}
