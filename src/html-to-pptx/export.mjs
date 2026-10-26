import PptxGenJS from 'pptxgenjs';
import {diagnostic,validateModel} from './model.mjs';
import {assetDataUrl} from './assets.mjs';
import {groupDiagnostics,applyFlatGroups} from './groups.mjs';
import {canvasSize,EMU_PER_POINT} from './canvas.mjs';

function paint(value,opacity=1){
  if(!value||value==='transparent')return {color:'FFFFFF',transparency:100};
  let hex,alpha=1;
  if(/^#[\da-f]{3,8}$/i.test(value)){
    hex=value.slice(1);if(hex.length===3||hex.length===4)hex=[...hex].map(c=>c+c).join('');
    if(hex.length===8){alpha=parseInt(hex.slice(6),16)/255;hex=hex.slice(0,6);}
  }else{
    const rgb=value.match(/^rgba?\(([^)]+)\)$/);
    if(rgb){const [r,g,b,a]=rgb[1].split(/[, /]+/).map(Number);hex=[r,g,b].map(n=>Math.round(n).toString(16).padStart(2,'0')).join('');alpha=a??1;}
  }
  if(!hex||!/^([\da-f]{6})$/i.test(hex)||!Number.isFinite(alpha))throw new Error(`Unsupported color: ${value}`);
  return {color:hex.toUpperCase(),transparency:Math.max(0,Math.min(100,(1-alpha*opacity)*100))};
}

export async function exportPptx(measured){
  const diagnostics=[...(measured?.diagnostics??[])],objectMap=[];
  const profile={engine:'PptxGenJS',version:'4.0.1',sourceDigest:measured?.model?.sourceDigest??null,environment:measured?.environment??null,visualQuality:'unmeasured',fontMappings:[],lineRepresentations:{filledBars:0,strokes:0}};
  const failed=()=>({status:'failed',pptxBytes:null,objectMap:[],profile,diagnostics});
  const issue=(code,message,slide=null,node=null)=>diagnostics.push(diagnostic(code,message,{slideId:slide?.id,nodeId:node?.id,sourceRef:node?.sourceRef}));
  const resolveFont=style=>{
    const mapping=measured.environment.fontMappings?.find(m=>m.sourceFamily===style.fontFamily&&m.sourceWeight===style.fontWeight);
    if(mapping){if(!profile.fontMappings.some(m=>m.sourceFamily===mapping.sourceFamily&&m.sourceWeight===mapping.sourceWeight)){
      profile.fontMappings.push(mapping);
      if(mapping.approximate)diagnostics.push(diagnostic('FONT_UNAVAILABLE',`글꼴 굵기 ${mapping.sourceWeight}는 설치된 ${mapping.nativeWeight} 굵기로 출력합니다: ${mapping.sourceFamily}. 정확한 화면 일치는 미확인입니다.`,{},'warning'));
    }return {fontFace:mapping.nativeFamily,bold:mapping.nativeBold};}
    const font=measured.environment.fonts?.find(f=>f.family===style.fontFamily);
    if(!font?.available||font.custom||['serif','sans-serif','monospace','system-ui','cursive','fantasy'].includes(style.fontFamily))return null;
    return {fontFace:style.fontFamily,bold:style.fontWeight>=600};
  };
  if(diagnostics.some(d=>d.severity==='error'))return failed();
  const validation=validateModel(measured);
  if(!validation.valid){issue('EXPORT_FAILED',`입력 모델 검증 실패: ${validation.errors.map(e=>e.path).join(', ')}`);return failed();}
  for(const slide of measured.model.slides){
    const m=measured.measurements.find(m=>m.slideId===slide.id);
    if(!m||m.assetsReady!==true||m.fontsReady!==true||m.layoutStable!==true)issue('RENDER_INCOMPLETE','자산·글꼴·배치의 측정 완료가 필요합니다.',slide);
    if(slide.background?.image&&slide.background.image!=='none'&&!slide.background.assetId)issue('RENDER_INCOMPLETE','장식 배경의 분리된 자산이 없습니다.',slide);
    for(const {group,message} of groupDiagnostics(slide))issue('UNSUPPORTED_OBJECT',message,slide,group);
    for(const n of slide.nodes){
      if(n.kind==='shape'&&n.style.backgroundImage&&n.style.backgroundImage!=='none'&&!n.style.assetId)issue('RENDER_INCOMPLETE','객체의 분리된 장식 배경 자산이 없습니다.',slide,n);
      if(n.kind==='text'&&!m?.objects.some(o=>o.nodeId===n.id&&o.lines?.length&&Number.isFinite(o.baseline)))issue('RENDER_INCOMPLETE','텍스트의 실제 줄·기준선 측정이 없습니다.',slide,n);
      if(n.kind==='text')for(const style of [n.style,...(n.content.runs??[]).map(r=>r.style)])if(!resolveFont(style))issue('FONT_UNAVAILABLE',`PPT에서 사용할 글꼴 대응을 확인할 수 없습니다: ${style.fontFamily} ${style.fontWeight}`,slide,n);
    }
  }
  if(diagnostics.some(d=>d.severity==='error'))return failed();
  try{
    const pptx=new PptxGenJS(),first=measured.model.slides[0],sizeEmu=canvasSize(measured.model),pointScale=sizeEmu.width/first.widthCssPx/EMU_PER_POINT,inchScale=pointScale/72;
    const emu=value=>Math.round(value*pointScale*12700),inches=value=>emu(value)/914400;
    profile.targetSizePt={width:first.widthCssPx*pointScale,height:first.heightCssPx*pointScale};profile.pointScale=pointScale;
    pptx.defineLayout({name:'HTML_SLIDES',width:sizeEmu.width/914400,height:sizeEmu.height/914400});pptx.layout='HTML_SLIDES';
    pptx.title='HTML 장표';pptx.author='HTML to editable PPTX';pptx.subject='Native text and shapes; visual validation is separate';pptx.lang='ko-KR';
    for(const [index,source] of measured.model.slides.entries()){
      if(source.widthCssPx!==first.widthCssPx||source.heightCssPx!==first.heightCssPx)throw new Error('Mixed slide sizes are unsupported');
      const slide=pptx.addSlide();slide.background=paint(source.background?.color??'#ffffff');
      let objectId=2;
      const objectNames=new Set(source.nodes.map(n=>n.id));
      const addBackground=(assetId,box,derivedFrom=null,sourceRef=null,opacity=1)=>{
        const asset=measured.model.assets.find(a=>a.id===assetId);if(!asset)throw new Error('Background asset missing');
        const baseName=derivedFrom?`${derivedFrom}::art`:`${source.id}::background`;
        let name=baseName,suffix=1;
        while(objectNames.has(name))name=`${baseName}::derived-${suffix++}`;
        objectNames.add(name);
        slide.addImage({...box,data:assetDataUrl(asset),objectName:name,transparency:(1-opacity)*100});
        objectMap.push({sourceSlideId:source.id,sourceNodeId:name,derivedFrom,sourceRef:sourceRef??{slideId:source.id,path:'background'},pptSlideNumber:index+1,pptObjectId:objectId++,pptObjectName:name,support:'background-art',parentId:null,assetSha256:asset.sha256});
      };
      if(source.background?.assetId)addBackground(source.background.assetId,{x:0,y:0,w:inches(source.widthCssPx),h:inches(source.heightCssPx)});
      for(const node of [...source.nodes].sort((a,b)=>a.zOrder-b.zOrder)){
        if(node.kind==='group')continue;
        const b=node.boxCssPx,st=node.style,box={x:inches(b.x),y:inches(b.y),w:inches(b.w),h:inches(b.h),objectName:node.id};
        if(node.kind==='shape'&&st.assetId)addBackground(st.assetId,box,node.id,node.sourceRef,st.opacity??1);
        if(node.kind==='text'){
          const runOptions=s=>({...resolveFont(s),fontSize:s.fontSize*pointScale,italic:s.fontStyle==='italic',color:paint(s.color,st.opacity).color,transparency:paint(s.color,st.opacity).transparency,charSpacing:s.letterSpacing&&s.letterSpacing!=='normal'?parseFloat(s.letterSpacing)*pointScale:0,underline:s.textDecoration?.includes('underline')?{style:'sng'}:undefined,breakLine:false});
          const runs=(node.content.runs??[{text:node.content.text,style:st}]).map(r=>{
            const options=runOptions(r.style),outline=r.style.outline;
            if(outline){
              const color=paint(outline.color,st.opacity??1);
              if(color.transparency!==0||!Number.isFinite(outline.width)||outline.width<=0)throw new Error('불투명 단색 외의 글자 외곽선은 지원하지 않습니다.');
              options.outline={color:color.color,size:outline.width*pointScale};
            }
            return {text:r.text,options};
          });
          const padding=st.padding??{left:0,right:0,top:0,bottom:0};
          // Browser and Office advances can differ slightly in a tight auto-sized
          // label. Preserve a measured single line without enlarging its box or
          // shrinking type. Longer edits may overflow and need a fresh HTML export.
          const layout=measured.measurements.find(m=>m.slideId===source.id).objects.find(o=>o.nodeId===node.id);
          slide.addText(runs,{...box,...runOptions(st),margin:[padding.top,padding.right,padding.bottom,padding.left].map(v=>v*pointScale),lineSpacing:st.lineHeight!=='normal'&&parseFloat(st.lineHeight)>0?parseFloat(st.lineHeight)*pointScale:undefined,paraSpaceBefore:0,paraSpaceAfter:0,align:st.align??'left',valign:st.valign??'top',fit:'none',wrap:layout.lines.length!==1,isTextBox:true,lang:'ko-KR'});
        }else if(node.kind==='shape'){
          const stroke=paint(st.stroke,(st.strokeWidth??0)>0?(st.opacity??1)*(st.strokeOpacity??1):0),half=st.svgGeometry||stroke.transparency===100?0:(st.strokeWidth??0)/2;
          if(b.w<=half*2||b.h<=half*2)throw new Error('Border is wider than the shape');
          slide.addShape(st.shape==='ellipse'?pptx.ShapeType.ellipse:st.radius>0?pptx.ShapeType.roundRect:pptx.ShapeType.rect,{...box,x:inches(b.x+half),y:inches(b.y+half),w:inches(b.w-2*half),h:inches(b.h-2*half),rectRadius:st.shape==='ellipse'?undefined:Math.max(0,(st.radius??0)-half)*inchScale,fill:paint(st.assetId?'transparent':st.fill,(st.opacity??1)*(st.fillOpacity??1)),line:{...stroke,width:(st.strokeWidth??0)*pointScale,dashType:st.dash||st.borderStyle==='dashed'?'dash':st.borderStyle==='dotted'?'sysDot':'solid'}});
        }else if(node.kind==='line'){
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
        }else if(node.kind==='image'){
          const asset=measured.model.assets.find(a=>a.id===node.content.assetId);if(!asset)throw new Error('Picture asset missing');
          slide.addImage({...box,data:assetDataUrl(asset),transparency:(1-(st.opacity??1))*100});
          if(asset.mime==='image/svg+xml'&&!node.content.artworkOnly)diagnostics.push(diagnostic('UNSUPPORTED_OBJECT','SVG는 그림 객체입니다. 내부 글자와 경로를 개별 편집할 수 없습니다.',{slideId:source.id,nodeId:node.id,sourceRef:node.sourceRef},'warning'));
        }
        objectMap.push({sourceSlideId:source.id,sourceNodeId:node.id,sourceRef:node.sourceRef,pptSlideNumber:index+1,pptObjectId:objectId++,pptObjectName:node.id,support:node.kind==='text'?'native-text':node.kind==='image'?(node.content.artworkOnly?'background-art':'picture'):'native-shape',parentId:node.parentId});
      }
      slide.addNotes('텍스트·도형은 편집 가능한 원생 객체입니다. 화면 1px 품질과 PowerPoint 편집 검증은 별도 관문입니다.');
    }
    const raw=new Uint8Array(await pptx.write({outputType:'uint8array'})),pptxBytes=applyFlatGroups(raw,measured.model,objectMap,pointScale);
    return {status:'generated',pptxBytes,objectMap,profile,diagnostics};
  }catch(e){issue('EXPORT_FAILED',`PPT 생성 실패: ${e.message}`);return failed();}
}
