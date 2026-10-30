import PptxGenJS from 'pptxgenjs';
import {diagnostic,validateModel} from './model.mjs';
import {assetDataUrl,sha256} from './assets.mjs';
import {groupDiagnostics,applyFlatGroups} from './groups.mjs';
import {canvasSize,EMU_PER_POINT} from './canvas.mjs';
import {addNativeTable} from './write-table.mjs';
import {addNativeChart} from './write-chart.mjs';
import {chartText} from './charts.mjs';
import {planFontEmbedding,embedPptxFonts} from './font-embedding.mjs';

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

// inspectFont returns sorted, disjoint ranges; binary search bounds work for fragmented cmaps.
function hasGlyph(ranges,cp){
  let lo=0,hi=ranges.length;
  while(lo<hi){const mid=(lo+hi)>>>1,[start,end]=ranges[mid];if(cp<start)hi=mid;else if(cp>end)lo=mid+1;else return true;}
  return false;
}

export async function exportPptx(measured,{fontEmbedding='off'}={}){
  const diagnostics=[...(measured?.diagnostics??[])],objectMap=[];
  const profile={engine:'PptxGenJS',version:'4.0.1',sourceDigest:measured?.model?.sourceDigest??null,environment:measured?.environment??null,visualQuality:'unmeasured',fontMappings:[],lineRepresentations:{filledBars:0,strokes:0}};
  const failed=()=>({status:'failed',pptxBytes:null,objectMap:[],profile,diagnostics});
  const issue=(code,message,slide=null,node=null)=>diagnostics.push(diagnostic(code,message,{slideId:slide?.id,nodeId:node?.id,sourceRef:node?.sourceRef}));
  let fontPlan={faces:[],report:[]};const embeddedStyles=new Map(),verifiedSourceHashes=new Map();
  const styleKey=style=>JSON.stringify([String(style?.fontFamily??'').trim().toLowerCase(),style?.fontWeight==='normal'?400:style?.fontWeight==='bold'?700:Number(style?.fontWeight??400),style?.fontStyle??'normal']);
  const embeddedKey=(style,slide)=>JSON.stringify([slide?.id??null,styleKey(style)]);
  const resolveFont=(style,slide)=>{
    const embedded=embeddedStyles.get(embeddedKey(style,slide));
    if(embedded)return {fontFace:embedded.family,bold:embedded.weight===700};
    const scoped=(measured.environment.fontMappings??[]).filter(m=>m.sourceFamily===style.fontFamily&&(m.slideId?m.slideId===slide?.id:measured.environment.inputProfile!=='embedded-static-documents-v1'));
    let mapping=scoped.find(m=>m.sourceWeight===style.fontWeight);
    if(!mapping&&(style.fontStyle??'normal')==='normal'&&Number.isFinite(style.fontWeight)){
      const hashes=new Set(scoped.map(m=>m.sourceSha256));
      const hash=[...hashes][0];
      if(hashes.size===1&&/^[a-f0-9]{64}$/i.test(hash??'')&&!verifiedSourceHashes.has(hash))verifiedSourceHashes.set(hash,(measured.model.assets??[]).some(a=>a.mime?.startsWith('font/')&&a.sha256===hash&&sha256(a.bytes)===hash));
      if(hashes.size===1&&verifiedSourceHashes.get(hash)){
        const candidates=scoped.filter(m=>Number.isFinite(m.nativeWeight)&&Math.abs(m.nativeWeight-style.fontWeight)<=50).sort((a,b)=>Math.abs(a.nativeWeight-style.fontWeight)-Math.abs(b.nativeWeight-style.fontWeight)||b.nativeWeight-a.nativeWeight);
        if(candidates.length)mapping={...candidates[0],matchedSourceWeight:candidates[0].sourceWeight,sourceWeight:style.fontWeight,approximate:true,layoutCalibration:'unmeasured'};
      }
    }
    if(mapping){if(!profile.fontMappings.some(m=>m.sourceFamily===mapping.sourceFamily&&m.sourceWeight===mapping.sourceWeight&&m.slideId===mapping.slideId&&m.sourceSha256===mapping.sourceSha256)){
      profile.fontMappings.push(mapping);
      if(mapping.approximate)diagnostics.push(diagnostic('FONT_UNAVAILABLE',`글꼴 굵기 ${mapping.sourceWeight}는 설치된 ${mapping.nativeWeight} 굵기로 출력합니다: ${mapping.sourceFamily}. 정확한 화면 일치는 미확인입니다.`,{},'warning'));
    }return {fontFace:mapping.nativeFamily,bold:mapping.nativeBold};}
    const font=measured.environment.fonts?.find(f=>f.family===style.fontFamily);
    if(!font?.available||font.custom||['serif','sans-serif','monospace','system-ui','cursive','fantasy'].includes(style.fontFamily))return null;
    return {fontFace:style.fontFamily,bold:style.fontWeight>=600};
  };
  // PptxGenJS inserts text/table typefaces directly into XML attributes.
  // Chart output escapes its own typeface separately.
  const resolveTextFont=(style,slide)=>{
    const font=resolveFont(style,slide);
    return font&&{...font,fontFace:font.fontFace.replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&apos;'}[c]))};
  };
  if(!['off','auto','required'].includes(fontEmbedding)){issue('INPUT_INVALID','글꼴 포함 옵션이 올바르지 않습니다.');return failed();}
  if(diagnostics.some(d=>d.severity==='error'))return failed();
  const validation=validateModel(measured);
  if(!validation.valid){
    const missingFont=fontEmbedding!=='off'&&validation.errors.some(e=>/\.fonts\[\d+\]\.assetId$/.test(e.path));
    issue(missingFont?'FONT_SOURCE_UNSUPPORTED':'EXPORT_FAILED',missingFont?'포함할 글꼴 파일이 없습니다.':`입력 모델 검증 실패: ${validation.errors.map(e=>e.path).join(', ')}`);return failed();
  }
  if(fontEmbedding!=='off'){
    const uses=[];
    for(const slide of measured.model.slides)for(const node of slide.nodes){
      const add=(style,text='')=>uses.push({style,text,slide,node});
      if(node.kind==='text'){
        add(node.style);
        const runs=node.content.runs??[{text:node.content.text,style:node.style}];
        if(!Array.isArray(runs)||runs.some(r=>typeof r?.text!=='string'||!r.style)||runs.map(r=>r.text).join('')!==node.content.text){issue('INPUT_INVALID','현재 글자와 서식 구간이 일치하지 않습니다.',slide,node);continue;}
        for(const run of runs)add(run.style,run.text);
      }else if(node.kind==='table'){
        for(const cell of node.content.cells){add(cell.style);for(const run of cell.runs.length?cell.runs:[{text:cell.text,style:cell.style}])add(run.style,run.text);}
      }else if(node.kind==='chart'){add(node.style);for(const text of chartText(node.content))add(node.style,text);}
    }
    const declared=measured.model.fonts??[],selected=new Map(),resolvedUses=new Map(),unembedded=new Set();
    for(const use of uses){
      const [family,weight,style]=JSON.parse(styleKey(use.style));
      const scopedIds=measured.environment.inputProfile==='embedded-static-documents-v1'
        ?new Set((measured.environment.fontUsages??[]).filter(u=>u.slideId===use.slide.id&&u.nodeId===use.node.id).map(u=>u.fontId)):null;
      const candidates=declared.filter(face=>face.cssFamily.trim().toLowerCase()===family&&(!scopedIds||scopedIds.has(face.id)));
      if(!candidates.length){
        const custom=measured.environment.fonts?.some(f=>f?.custom&&typeof f.family==='string'&&f.family.trim().toLowerCase()===family);
        if(fontEmbedding==='required'||custom){const format=diagnostics.some(d=>d.code==='FONT_FORMAT_UNSUPPORTED');issue(format?'FONT_FORMAT_UNSUPPORTED':'FONT_SOURCE_UNSUPPORTED',format?'이 웹폰트 형식은 아직 포함할 수 없습니다.':'사용한 글꼴 파일을 포함할 수 없습니다.',use.slide,use.node);}
        else if(!unembedded.has(family)){unembedded.add(family);diagnostics.push(diagnostic('FONT_NOT_PORTABLE','일부 글꼴 파일은 포함하지 않습니다.',{},'warning'));}
        continue;
      }
      if(style!=='normal'||![400,700].includes(weight)){issue('FONT_STYLE_UNSUPPORTED','일반·굵게의 정적 글꼴만 포함할 수 있습니다.',use.slide,use.node);continue;}
      const matches=candidates.filter(face=>face.weight===weight&&face.style===style);
      if(matches.length!==1){issue('FONT_STYLE_MISMATCH','사용한 스타일의 글꼴 파일을 하나로 확정할 수 없습니다.',use.slide,use.node);continue;}
      const face=matches[0],asset=measured.model.assets.find(a=>a.id===face.assetId);
      if(!asset){issue('FONT_SOURCE_UNSUPPORTED','사용한 글꼴 파일이 없습니다.',use.slide,use.node);continue;}
      if(asset.mime!=='font/ttf'){issue('FONT_FORMAT_UNSUPPORTED','정적 TTF 글꼴만 포함할 수 있습니다.',use.slide,use.node);continue;}
      selected.set(face.id,face);
      resolvedUses.set(use,face);
    }
    if(diagnostics.some(d=>d.severity==='error'))return failed();
    fontPlan=planFontEmbedding({...measured.model,fonts:[...selected.values()]});
    diagnostics.push(...fontPlan.diagnostics);profile.fontEmbedding=fontPlan.report;
    if(fontEmbedding==='required'&&!fontPlan.faces.length)issue('FONT_SOURCE_UNSUPPORTED','포함할 글꼴 파일이 없습니다.');
    if(diagnostics.some(d=>d.severity==='error'))return failed();
    for(const [use,face] of resolvedUses){
      const prepared=fontPlan.faces.find(f=>f.family.toLowerCase()===face.family.toLowerCase()&&f.weight===face.weight&&f.style===face.style);
      embeddedStyles.set(embeddedKey(use.style,use.slide),prepared);
    }
    for(const use of uses){
      const face=embeddedStyles.get(embeddedKey(use.style,use.slide));if(!face)continue;
      if([...use.text].some(c=>{const cp=c.codePointAt(0);return ![9,10,13].includes(cp)&&!hasGlyph(face.meta.glyphCoverage,cp);})){issue('FONT_GLYPH_MISSING','제공 글꼴이 현재 문구의 일부 문자를 포함하지 않습니다.',use.slide,use.node);}
    }
    if(diagnostics.some(d=>d.severity==='error'))return failed();
  }
  for(const slide of measured.model.slides){
    const m=measured.measurements.find(m=>m.slideId===slide.id);
    if(!m||m.assetsReady!==true||m.fontsReady!==true||m.layoutStable!==true)issue('RENDER_INCOMPLETE','자산·글꼴·배치의 측정 완료가 필요합니다.',slide);
    if(slide.background?.image&&slide.background.image!=='none'&&!slide.background.assetId)issue('RENDER_INCOMPLETE','장식 배경의 분리된 자산이 없습니다.',slide);
    for(const {group,message} of groupDiagnostics(slide))issue('UNSUPPORTED_OBJECT',message,slide,group);
    for(const n of slide.nodes){
      if(n.kind==='shape'&&n.style.backgroundImage&&n.style.backgroundImage!=='none'&&!n.style.assetId)issue('RENDER_INCOMPLETE','객체의 분리된 장식 배경 자산이 없습니다.',slide,n);
      if(n.kind==='text'&&!m?.objects.some(o=>o.nodeId===n.id&&o.lines?.length&&Number.isFinite(o.baseline)))issue('RENDER_INCOMPLETE','텍스트의 실제 줄·기준선 측정이 없습니다.',slide,n);
      if(n.kind==='text')for(const style of [n.style,...(n.content.runs??[]).map(r=>r.style)])if(!resolveFont(style,slide))issue('FONT_UNAVAILABLE',`PPT에서 사용할 글꼴 대응을 확인할 수 없습니다: ${style.fontFamily} ${style.fontWeight}`,slide,n);
      if(n.kind==='table')for(const cell of n.content.cells)for(const style of [cell.style,...cell.runs.map(r=>r.style)])if(!resolveFont(style,slide))issue('FONT_UNAVAILABLE',`PPT 표에서 사용할 글꼴 대응을 확인할 수 없습니다: ${style.fontFamily} ${style.fontWeight}`,slide,n);
      if(n.kind==='chart'&&!resolveFont(n.style,slide))issue('FONT_UNAVAILABLE',`PPT 차트에서 사용할 글꼴 대응을 확인할 수 없습니다: ${n.style.fontFamily} ${n.style.fontWeight}`,slide,n);
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
        if(node.kind==='table'){
          const entry=addNativeTable(slide,node,{pointScale,resolveFont:s=>resolveTextFont(s,source),paint});
          objectMap.push({...entry,sourceSlideId:source.id,pptSlideNumber:index+1,pptObjectId:objectId++});continue;
        }
        if(node.kind==='chart'){
          const entry=addNativeChart(slide,node,{pointScale,resolveFont:s=>resolveFont(s,source)});
          objectMap.push({...entry,sourceSlideId:source.id,pptSlideNumber:index+1,pptObjectId:objectId++});continue;
        }
        const b=node.boxCssPx,st=node.style,box={x:inches(b.x),y:inches(b.y),w:inches(b.w),h:inches(b.h),objectName:node.id};
        if(node.kind==='shape'&&st.assetId)addBackground(st.assetId,box,node.id,node.sourceRef,st.opacity??1);
        if(node.kind==='text'){
          const runOptions=s=>({...resolveTextFont(s,source),fontSize:s.fontSize*pointScale,italic:s.fontStyle==='italic',color:paint(s.color,st.opacity).color,transparency:paint(s.color,st.opacity).transparency,charSpacing:s.letterSpacing&&s.letterSpacing!=='normal'?parseFloat(s.letterSpacing)*pointScale:0,underline:s.textDecoration?.includes('underline')?{style:'sng'}:undefined,breakLine:false});
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
    const raw=new Uint8Array(await pptx.write({outputType:'uint8array'}));let pptxBytes=applyFlatGroups(raw,measured.model,objectMap,pointScale);
    if(fontPlan.faces.length){
      try{pptxBytes=embedPptxFonts(pptxBytes,fontPlan.faces,{container:'eot'});}catch(e){issue(e.code??'EXPORT_FAILED','글꼴을 PPT에 포함하지 못했습니다.');return failed();}
      profile.fontEmbedding=fontPlan.report.map(entry=>({...entry,status:'included'}));
    }
    return {status:'generated',pptxBytes,objectMap,profile,diagnostics};
  }catch(e){issue('EXPORT_FAILED',`PPT 생성 실패: ${e.message}`);return failed();}
}
