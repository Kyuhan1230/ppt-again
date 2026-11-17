import path from 'node:path';
import {diagnostic} from './model.mjs';
import {resolveAsset,sha256} from './assets.mjs';
import {validPresetVector,presetVectorMarkup} from './preset-vector.mjs';
import {EMU_PER_CSS_PX} from './canvas.mjs';
import {applyPngDuotone,duotoneColors} from './png-duotone.mjs';
import {snippedTextSvg} from './snip-text.mjs';
import {readNativeTable} from './read-table.mjs';
import {readChartPackage} from './read-chart.mjs';

const elements=n=>[...(n?.childNodes??[])].filter(c=>c.nodeType===1);
const child=(n,name)=>elements(n).find(c=>c.localName===name);
const all=(n,name)=>n?[...n.getElementsByTagNameNS('*',name)]:[];
const attr=(n,k,fallback=null)=>n?.hasAttribute(k)?n.getAttribute(k):fallback;
const num=(n,k,fallback=0)=>Number(attr(n,k,fallback));
const px=n=>n/EMU_PER_CSS_PX;
const C='http://schemas.openxmlformats.org/drawingml/2006/chart';
const R='http://schemas.openxmlformats.org/officeDocument/2006/relationships',A='http://schemas.openxmlformats.org/drawingml/2006/main';
const mimeByExt={'.png':'image/png','.jpg':'image/jpeg','.jpeg':'image/jpeg','.svg':'image/svg+xml'};

// Read only the picture's content/appearance; the visitor owns object order.
export function readPictureObject({object,explicitLine,lineRef,color,inheritedStroke,roundedPicture,pr,shapeGeometry,allowPartial,diagnostics,slide,id,objectContexts,referencedWidth,node,ownerRels,parts,issue,relationRecords,ownerFile,contentTypeFor,model,colorMap,colorScheme,pictureShadow,pictureApprox,index}){
  let pictureBorder=null;
            const line=explicitLine,lineFill=child(line,'solidFill');
            if(lineFill&&child(line,'noFill'))throw new Error('그림 테두리 채움이 중복되었습니다.');
            const linePaint=lineFill?color(lineFill):inheritedStroke;
            const geometrySupported=roundedPicture||!child(pr,'custGeom')&&[null,'rect'].includes(shapeGeometry);
            const lineSupported=(!line||attr(line,'cmpd','sng')==='sng'&&attr(line,'algn','ctr')==='ctr'&&(!child(line,'prstDash')||attr(child(line,'prstDash'),'val')==='solid')&&elements(line).every(n=>n.namespaceURI===A&&['solidFill','noFill','prstDash','round','bevel','miter'].includes(n.localName)))&&geometrySupported;
            if((line||lineRef)&&linePaint&&linePaint!=='transparent'&&!lineSupported){
              if(allowPartial)diagnostics.push({...diagnostic('PPTX_IMAGE_APPEARANCE_APPROXIMATED','그림의 복잡한 테두리는 표시하지 않았습니다.',{slideId:slide.id,nodeId:id},'warning'),slideNumber:index+1,...objectContexts.get(`${slide.id}:${id}`),reasonKey:'PPTX_IMAGE_APPEARANCE_APPROXIMATED'});
              else issue('그림의 복잡한 테두리는 아직 지원하지 않습니다.',id);
            }else if((line||lineRef)&&!child(line,'noFill')&&linePaint&&linePaint!=='transparent'&&lineSupported){
              const rawWidth=line?.hasAttribute('w')?num(line,'w',NaN):null,width=rawWidth===null?(referencedWidth||1):px(rawWidth);
              if(!Number.isFinite(width)||width<0||width>1024)throw new Error('그림 테두리 너비가 잘못되었습니다.');
              if(width>0)pictureBorder={stroke:linePaint,strokeWidth:width,shape:roundedPicture?'ellipse':'rect'};
            }
            node.kind='image';const blip=all(object,'blip')[0],target=ownerRels.get(attr(blip,'r:embed'));
            const effectsOnBlip=elements(blip),duotones=effectsOnBlip.filter(n=>n.namespaceURI===A&&n.localName==='duotone');
            const blipFill=child(object,'blipFill'),rect=child(blipFill,'srcRect');
            let crop;
            if(rect){
              crop=Object.fromEntries([['left','l'],['top','t'],['right','r'],['bottom','b']].map(([key,name])=>[key,num(rect,name)]));
              const fillRect=child(child(blipFill,'stretch'),'fillRect');
              const valid=Object.values(crop).every(v=>Number.isInteger(v)&&v>=0&&v<100000)&&crop.left+crop.right<100000&&crop.top+crop.bottom<100000;
              if(!valid)throw new Error('그림 자르기 좌표가 잘못되었습니다.');
              if(!child(blipFill,'stretch')||fillRect&&['l','t','r','b'].some(side=>num(fillRect,side)!==0))issue('이 그림의 자르기 또는 채우기는 지원하지 않습니다.',id);
              if(!Object.values(crop).some(Boolean))crop=undefined;
            }
            if(!target||!parts[target]){issue('내장 그림 자산이 없습니다.',id);return {skip:true};}
            const extension=path.posix.extname(target).toLowerCase(),declaredMime=mimeByExt[extension];if(!declaredMime){issue('이 그림 형식은 지원하지 않습니다.',id);return {skip:true};}
            // Some files swap JPEG/PNG labels. PptxGenJS can also write SVG
            // bytes into the raster fallback while linking the same bytes as
            // an SVG. Only recover that exact, explicitly linked case; the
            // regular asset size, MIME and SVG safety checks still apply.
            const fallbackBytes=parts[target],raster=['image/png','image/jpeg'].includes(declaredMime);
            const jpeg=raster&&fallbackBytes.length>=3&&fallbackBytes[0]===0xff&&fallbackBytes[1]===0xd8&&fallbackBytes[2]===0xff;
            const png=raster&&fallbackBytes.length>=8&&[137,80,78,71,13,10,26,10].every((v,i)=>fallbackBytes[i]===v);
            const vectors=all(blip,'svgBlip'),vector=vectors.length===1?vectors[0]:null,ext=vector?.parentNode,extList=ext?.parentNode;
            const vectorRelation=vector&&relationRecords(ownerFile).get(attr(vector,'r:embed')),vectorTarget=vectorRelation?.target;
            const svgFallback=raster&&!jpeg&&!png&&vector?.namespaceURI==='http://schemas.microsoft.com/office/drawing/2016/SVG/main'&&
              ext?.namespaceURI===A&&ext.localName==='ext'&&attr(ext,'uri')==='{96DAC541-7B7A-43D3-8B79-37D633B846F1}'&&
              extList?.namespaceURI===A&&extList.localName==='extLst'&&extList.parentNode===blip&&vectorRelation?.type===`${R}/image`&&
              vectorTarget&&path.posix.extname(vectorTarget).toLowerCase()==='.svg'&&contentTypeFor(vectorTarget)==='image/svg+xml'&&parts[vectorTarget]&&
              Buffer.from(fallbackBytes).equals(Buffer.from(parts[vectorTarget]));
            const bytes=svgFallback?parts[vectorTarget]:fallbackBytes;
            const mime=svgFallback?'image/svg+xml':jpeg?'image/jpeg':png?'image/png':declaredMime;
            const assetName='image'+(svgFallback?'.svg':jpeg?'.jpg':png?'.png':extension);
            let asset=resolveAsset(assetName,[{path:assetName,mime,bytes}]),duotoneBaked=false;
            if(allowPartial&&mime==='image/png'&&duotones.length===1&&effectsOnBlip.every(n=>n.namespaceURI===A&&['duotone','alphaModFix','extLst'].includes(n.localName))){
              const endpoints=duotoneColors(duotones[0],key=>{
                const mapped=attr(colorMap,key,key),slot=child(colorScheme,mapped),values=elements(slot);
                return values.length===1?values[0]:null;
              });
              if(endpoints){
                const transformed=applyPngDuotone(asset.bytes,endpoints);
                if(transformed){asset=resolveAsset('duotone.png',[{path:'duotone.png',mime:'image/png',bytes:transformed}]);duotoneBaked=true;}
              }
            }
            const imageEffects=effectsOnBlip.some(n=>!['alphaModFix','extLst'].includes(n.localName)&&!(duotoneBaked&&n===duotones[0]))||all(blip,'imgEffect').length>0;
            if(duotoneBaked)diagnostics.push({...diagnostic('PPTX_IMAGE_DUOTONE_BAKED','듀오톤 색상 효과를 그림에 적용했습니다. 효과를 별도 설정으로 편집할 수 없습니다.',{slideId:slide.id,nodeId:id},'warning'),slideNumber:index+1,...objectContexts.get(`${slide.id}:${id}`),reasonKey:'PPTX_IMAGE_DUOTONE_BAKED'});
            if(imageEffects){
              if(allowPartial&&!pictureApprox)diagnostics.push({...diagnostic('PPTX_IMAGE_APPEARANCE_APPROXIMATED','사진의 색상 및 그림 효과를 생략해 원본과 다를 수 있습니다.',{slideId:slide.id,nodeId:id},'warning'),slideNumber:index+1,...objectContexts.get(`${slide.id}:${id}`),reasonKey:'PPTX_IMAGE_APPEARANCE_APPROXIMATED'});
              else if(!allowPartial)issue('그림의 색상 또는 효과는 아직 지원하지 않습니다.',id,'PPTX_IMAGE_UNSUPPORTED');
            }
            if(!model.assets.some(a=>a.id===asset.id))model.assets.push(asset);
            node.content={assetId:asset.id,...(crop?{crop}:{}),...(duotoneBaked?{duotoneBaked:true}:{}),...(roundedPicture?{rounding:true}:{})};node.style.opacity=num(child(blip,'alphaModFix'),'amt',100000)/100000;
            if(pictureShadow){node.style.shadow=pictureShadow;diagnostics.push({...diagnostic('PPTX_SHADOW_APPROXIMATED','사진의 바깥 그림자를 CSS로 근사했으며 PPT로 다시 저장할 때 빠질 수 있습니다.',{slideId:slide.id,nodeId:id},'warning'),slideNumber:index+1,...objectContexts.get(`${slide.id}:${id}`),reasonKey:'PPTX_SHADOW_APPROXIMATED'});}
  return {pictureBorder};
}

// Read shape, connector or text content after common geometry was checked.
export function readShapeTextObject({object,pr,box,inheritedFill,inheritedStroke,inheritedStrokeWidth,color,referencedWidth,shadow,issue,id,allowPartial,isLine,diagnostics,slide,index,objectContexts,node,xfrm,readText,snipped,model,seen,parentId}){
  let lineArrowSpecs=[];
            const geometry=child(pr,'prstGeom'),shape=attr(geometry,'prst'),line=child(pr,'ln'),tx=child(object,'txBody');
            let radius=0;
            if(shape==='roundRect'){
              const adjustments=elements(child(geometry,'avLst'));let value=16667;
              if(adjustments.length){
                const guide=adjustments[0],match=/^val\s+(-?\d+)$/.exec(attr(guide,'fmla',''));
                if(adjustments.length!==1||guide.namespaceURI!==A||guide.localName!=='gd'||attr(guide,'name')!=='adj'||!match||!Number.isSafeInteger(Number(match[1])))issue('둥근 사각형의 이 모서리 조정식은 지원하지 않습니다.',id);
                else value=Math.max(0,Math.min(50000,Number(match[1])));
              }
              radius=Math.min(box.w,box.h)*value/100000;
            }
            Object.assign(node.style,{svgGeometry:true,shape:shape==='ellipse'?'ellipse':'rect',fill:inheritedFill??color(child(pr,'solidFill')),stroke:inheritedStroke??color(child(line,'solidFill')),strokeWidth:inheritedStroke!==null?inheritedStrokeWidth:child(line,'noFill')?0:line?.hasAttribute('w')?px(num(line,'w')):referencedWidth,radius});
            if(shadow)node.style.shadow=shadow;
            if(child(line,'prstDash')&&attr(child(line,'prstDash'),'val')!=='solid')issue('점선·화살표 선은 아직 지원하지 않습니다.',id);
            const arrowEnds=elements(line).filter(n=>['headEnd','tailEnd'].includes(n.localName)&&attr(n,'type')!=='none');
            if(arrowEnds.length){
              const arrowApprox=allowPartial&&isLine&&arrowEnds.every(n=>n.namespaceURI===A&&attr(n,'type')==='triangle'&&['sm','med','lg'].includes(attr(n,'w','med'))&&['sm','med','lg'].includes(attr(n,'len','med')));
              if(arrowApprox&&['line','straightConnector1'].includes(shape))lineArrowSpecs=arrowEnds.map(end=>({end:end.localName==='headEnd'?'head':'tail',width:attr(end,'w','med'),length:attr(end,'len','med')}));
              else if(arrowApprox)diagnostics.push({...diagnostic('PPTX_ARROWHEAD_OMITTED','삼각형 화살표 끝은 표시하지 않았습니다.',{slideId:slide.id,nodeId:id},'warning'),slideNumber:index+1,...objectContexts.get(`${slide.id}:${id}`),reasonKey:'PPTX_ARROWHEAD_OMITTED'});
              else issue('화살표 선은 아직 지원하지 않습니다.',id);
            }
            if(shape==='line'||object.localName==='cxnSp'){
              node.kind='line';const flipH=num(xfrm,'flipH'),flipV=num(xfrm,'flipV');
              node.style.endpoints={start:{x:box.x+(flipH?box.w:0),y:box.y+(flipV?box.h:0)},end:{x:box.x+(flipH?0:box.w),y:box.y+(flipV?0:box.h)}};
              if(shape==='bentConnector3'){
                const guides=elements(child(geometry,'avLst'));
                const guide=guides[0],match=guide&&/^val\s+(-?\d+)$/.exec(attr(guide,'fmla',''));
                if(guides.length>1||guide&&(guide.namespaceURI!==A||guide.localName!=='gd'||attr(guide,'name')!=='adj1'||!match||!Number.isSafeInteger(Number(match[1]))))issue('이 연결선의 조정식은 아직 지원하지 않습니다.',id);
                else{
                  const ratio=guide?Math.max(0,Math.min(100000,Number(match[1])))/100000:.5;
                  const {start,end}=node.style.endpoints,midX=start.x+(end.x-start.x)*ratio;
                  node.style.connectorPoints=[start,{x:midX,y:start.y},{x:midX,y:end.y},end];
                }
              }else if(shape!=='line'&&shape!=='straightConnector1')issue('이 연결선 모양은 아직 지원하지 않습니다.',id);
            }else if(tx&&all(tx,'t').length){
              node.kind='text';const text=readText(tx,id);node.content=text.content;Object.assign(node.style,text.style);
              if(snipped){
                const svg=node.style.strokeWidth===0?snippedTextSvg(box,snipped,node.style.fill):null;
                if(!svg)issue('이 잘린 모서리 텍스트 도형의 채움 또는 선은 아직 지원하지 않습니다.',id);
                else{
                  const asset=resolveAsset('snip.svg',[{path:'snip.svg',mime:'image/svg+xml',bytes:Buffer.from(svg)}]);
                  if(!model.assets.some(a=>a.id===asset.id))model.assets.push(asset);
                  let backgroundId=`${id}::background`;while(seen.has(backgroundId))backgroundId+='-';seen.add(backgroundId);
                  slide.nodes.push({...node,id:backgroundId,kind:'image',zOrder:slide.nodes.length,content:{assetId:asset.id,snippedTextBackground:true},style:{opacity:1},sourceRef:{slideId:slide.id,elementId:backgroundId}});
                  node.zOrder=slide.nodes.length;Object.assign(node.style,{fill:'transparent',stroke:'transparent',strokeWidth:0,radius:0});
                  // PowerPoint's preset text rectangle excludes half of the
                  // snipped corner on each affected side.
                  node.style.padding[snipped.flipH?'left':'right']+=snipped.cutX/2;
                  node.style.padding[snipped.flipV?'bottom':'top']+=snipped.cutY/2;
                  diagnostics.push({...diagnostic('PPTX_SHAPE_IMAGE_APPROXIMATED','잘린 모서리 배경은 SVG 그림으로 보존하고 글자는 별도로 편집할 수 있습니다.',{slideId:slide.id,nodeId:id},'warning'),slideNumber:index+1,...objectContexts.get(`${slide.id}:${id}`),reasonKey:'PPTX_SHAPE_IMAGE_APPROXIMATED'});
                }
              }else if(shape&&!['rect','roundRect'].includes(shape))issue('이 텍스트 도형의 모양은 아직 지원하지 않습니다.',id);
              else if(object.localName==='sp'){
                const cNvSpPr=child(child(object,'nvSpPr'),'cNvSpPr');
                node.pptxOrigin={version:1,documentId:model.sourceDigest,slideId:slide.id,objectId:id,parentId,role:'text-shape',textBox:['1','true'].includes(attr(cNvSpPr,'txBox','0'))};
              }
            }else if(!['rect','roundRect','ellipse'].includes(shape))issue('지원 도형은 사각형·둥근 사각형·타원입니다.',id);
  return lineArrowSpecs;
}

// Table and chart reading delegates to their existing validators.
export function readGraphicFrame({object,node,ownerFile,relationRecords,parts,chartPackages,budget,contentTypeFor,workbooks,renderChartImages,usedCharts,usedWorkbooks,blockedChartWorkbooks,unverifiedChartTargets,layer,nv,doc,slide,index,id,base,diagnostics,pendingChartImages,chartPointCounter,model,box,readText,color,allowPartial,map}){
            const graphic=child(object,'graphic'),data=child(graphic,'graphicData');
            if(attr(data,'uri')===C){
              const refs=elements(data),chart=refs[0],rid=chart?.getAttributeNS(R,'id'),relation=relationRecords(ownerFile).get(rid);
              if(graphic.namespaceURI!==A||data.namespaceURI!==A||refs.length!==1||chart?.namespaceURI!==C||chart.localName!=='chart'||relation?.type!==`${R}/chart`||!relation.target.startsWith('ppt/charts/')||!parts[relation.target])throw new Error('차트 관계 또는 내장 부분이 잘못되었습니다.');
              const chartPath=relation.target;
              if(!chartPackages.has(chartPath))chartPackages.set(chartPath,readChartPackage(parts,chartPath,{budget,contentTypeFor,workbooks,allowImageFallback:typeof renderChartImages==='function'}));
              const result=chartPackages.get(chartPath);
              // Workbook ownership is a validated package fact, independent of
              // whether this chart's visual presentation can be represented.
              if(result.verifiedWorkbookPath){usedCharts.add(chartPath);usedWorkbooks.add(result.verifiedWorkbookPath);}
              // A structurally linked but invalid workbook remains a blocking
              // chart error; distinguish it from an actually unrelated package.
              if(!result.verifiedWorkbookPath&&result.referencedWorkbookPath&&result.diagnostics.some(d=>d.severity==='error'&&d.recoverable!==true))blockedChartWorkbooks.set(chartPath,result.referencedWorkbookPath);
              // A malformed chart can prevent the externalData id from being
              // checked. A package relation is evidence to investigate, not
              // proof that the file is either connected or safe.
              if(!result.referencedWorkbookPath&&result.diagnostics.some(d=>d.severity==='error'&&d.recoverable!==true))for(const r of relationRecords(chartPath).values())if(r.type===`${R}/package`&&r.target.startsWith('ppt/embeddings/')&&parts[r.target])unverifiedChartTargets.add(r.target);
              if(result.imageFallbackEligible&&typeof renderChartImages==='function'){
                const shapeId=Number(attr(nv,'id')),matches=all(doc,'cNvPr').filter(item=>Number(attr(item,'id'))===shapeId);
                if(layer!=='slide'||!Number.isSafeInteger(shapeId)||shapeId<=0||matches.length!==1){
                  const ambiguous=layer==='slide'&&Number.isSafeInteger(shapeId)&&shapeId>0&&matches.length>1;
                  diagnostics.push({...diagnostic('UNSUPPORTED_OBJECT','차트의 원본 위치를 안전하게 확인할 수 없어 그림으로 대체하지 못했습니다.',{slideId:slide.id,nodeId:id}),slideNumber:index+1,objectName:base,objectType:'chart',reasonKey:ambiguous?'PPTX_CHART_IMAGE_ID_AMBIGUOUS':'PPTX_CHART_IMAGE_FAILED',blocking:true});
                }else{
                  slide.nodes.push(node);pendingChartImages.push({slideNumber:index+1,slideId:slide.id,nodeId:id,shapeId,node,objectName:base});
                }
                return;
              }
              diagnostics.push(...result.diagnostics.map(d=>{
                const blocking=d.severity==='error'&&d.recoverable!==true;
                const reasonKey=d.reasonKey??(blocking?({CHART_DATA_MISMATCH:'PPTX_CHART_DATA_MISMATCH',INPUT_INVALID:'PPTX_CHART_STRUCTURE_INVALID',LIMIT_EXCEEDED:'PPTX_CHART_LIMIT_EXCEEDED',UNSUPPORTED_OBJECT:'PPTX_CHART_UNSUPPORTED'})[d.code]??'PPTX_CHART_VALIDATION_FAILED':undefined);
                return {...d,slideId:slide.id,nodeId:id,slideNumber:index+1,objectName:base,objectType:'chart',...(blocking?{blocking:true}:{}),...(reasonKey?{reasonKey}:{})};
              }));
              if(result.content){
                if((chartPointCounter.value+=result.content.categories.length*result.content.series.length)>10000)throw Object.assign(new Error('문서의 차트 데이터가 10,000점을 초과했습니다.'),{code:'LIMIT_EXCEEDED'});
                usedCharts.add(chartPath);usedWorkbooks.add(result.workbookPath);
                slide.nodes.push({...node,kind:'chart',content:structuredClone(result.content),style:{...node.style,...result.style}});model.version=2;
              }
              return;
            }
            const table=readNativeTable(object,{slideId:slide.id,sourceId:id,boxCssPx:box,readText:tx=>readText(tx,id),readColor:color,allowPartial});
            diagnostics.push(...table.diagnostics.map(d=>({...d,...(d.severity==='error'&&d.recoverable!==true?{blocking:true}:{}),slideNumber:index+1,objectName:base,objectType:'table'})));
            if(table.node){
              // Table dimensions come from the native grid, not its frame.
              // Apply the composed group scale to that grid just as for other
              // object geometry. Cell text, padding and borders remain an
              // approximation covered by the enclosing group warning.
              table.node.content.columnWidthsCssPx=table.node.content.columnWidthsCssPx.map(width=>width*map.sx);
              table.node.content.rowHeightsCssPx=table.node.content.rowHeightsCssPx.map(height=>height*map.sy);
              table.node.boxCssPx.w=table.node.content.columnWidthsCssPx.reduce((sum,width)=>sum+width,0);
              table.node.boxCssPx.h=table.node.content.rowHeightsCssPx.reduce((sum,height)=>sum+height,0);
              slide.nodes.push({...node,...table.node});model.version=2;
            }return;

}

// Only bounded, solid-color primitive geometry is converted to inline SVG.
export function rotatedPresetArtwork(object,pr,xfrm,box,map,resolveColor,colorScheme,colorMap,themeDoc,uniqueKey){
  const tx=child(object,'txBody');
  if(object.localName!=='sp'||['r','fld','br','t'].some(name=>all(tx,name).length)||num(xfrm,'flipH')||Math.abs(map.sx-map.sy)>1e-7)return null;
  const geometry=child(pr,'prstGeom'),shape=attr(geometry,'prst');
  if(!['ellipse','downArrow','rightArrow','bentArrow'].includes(shape)||num(xfrm,'flipV')&&shape!=='bentArrow'||elements(geometry).some(n=>n.namespaceURI!==A||n.localName!=='avLst')||elements(geometry).length>1)return null;
  const adjustments={adj1:shape==='bentArrow'?25000:50000,adj2:shape==='bentArrow'?25000:50000},seen=new Set();
  for(const guide of elements(child(geometry,'avLst'))){
    const name=attr(guide,'name'),formula=attr(guide,'fmla',''),match=/^val (-?\d+)$/.exec(formula);
    if(!['downArrow','rightArrow'].includes(shape)||guide.namespaceURI!==A||guide.localName!=='gd'||elements(guide).length||!['adj1','adj2'].includes(name)||seen.has(name)||!match)return null;
    const value=Number(match[1]);if(!Number.isSafeInteger(value)||value<0||value>2147483647)return null;
    seen.add(name);adjustments[name]=value;
  }
  if(elements(pr).some(n=>!['xfrm','prstGeom','noFill','solidFill','gradFill','ln'].includes(n.localName)))return null;
  const style=child(object,'style'),effect=child(style,'effectRef');
  if(effect&&num(effect,'idx')!==0)return null;
  const rgb=fill=>{
    const colors=elements(fill),c=colors[0];
    if(colors.length!==1||c.namespaceURI!==A||!['srgbClr','schemeClr'].includes(c.localName))return null;
    const safeEffects=source=>elements(source).every(effect=>{
      const raw=attr(effect,'val','');return effect.namespaceURI===A&&['lumMod','lumOff','shade','alpha'].includes(effect.localName)&&elements(effect).length===0&&/^\d+$/.test(raw)&&Number(raw)<=100000&&(effect.localName!=='alpha'||Number(raw)===100000);
    });
    if(!safeEffects(c))return null;
    if(c.localName==='srgbClr'&&!/^[a-f\d]{6}$/i.test(attr(c,'val','')))return null;
    let source=c;
    if(c.localName==='schemeClr'){
      const key=attr(c,'val','');if(!/^[a-zA-Z0-9]+$/.test(key))return null;
      const sources=elements(child(colorScheme,attr(colorMap,key,key)));
      if(sources.length!==1||sources[0].namespaceURI!==A||!['srgbClr','sysClr','prstClr'].includes(sources[0].localName)||!safeEffects(sources[0]))return null;
      source=sources[0];
    }
    // OOXML shade acts in linear RGB; multiplying the encoded sRGB channels
    // makes dark theme outlines noticeably too dark.
    const effects=[...elements(source),...(source===c?[]:elements(c))],shades=effects.filter(e=>e.localName==='shade');
    if(shades.length){
      if(shades.length!==1||effects.some(e=>!['shade','alpha'].includes(e.localName)))return null;
      const raw=source.localName==='sysClr'?attr(source,'lastClr',''):source.localName==='srgbClr'?attr(source,'val',''):({black:'000000',white:'FFFFFF'})[attr(source,'val')];
      if(!/^[a-f\d]{6}$/i.test(raw??''))return null;
      const factor=Number(attr(shades[0],'val'))/100000;
      const channels=[0,2,4].map(i=>{
        const s=parseInt(raw.slice(i,i+2),16)/255,linear=s<=.04045?s/12.92:((s+.055)/1.055)**2.4,shaded=linear*factor;
        return Math.round(255*(shaded<=.0031308?12.92*shaded:1.055*shaded**(1/2.4)-.055)).toString(16).padStart(2,'0');
      });
      return `#${channels.join('')}`;
    }
    const value=resolveColor(fill,'');
    return /^#[a-f\d]{6}$/i.test(value)?value:null;
  };
  const fillNode=child(pr,'solidFill'),noFill=child(pr,'noFill'),gradFill=child(pr,'gradFill');
  if([fillNode,noFill,gradFill].filter(Boolean).length>1)return null;
  const themeEntry=(list,ref)=>{
    const raw=attr(ref,'idx','');if(!/^\d+$/.test(raw)||Number(raw)<1||!Number.isSafeInteger(Number(raw)))return null;
    return elements(all(themeDoc,list)[0])[Number(raw)-1];
  };
  const inheritedPaint=(entry,ref)=>{
    const c=elements(entry);return entry?.namespaceURI===A&&entry.localName==='solidFill'&&c.length===1&&c[0].namespaceURI===A&&c[0].localName==='schemeClr'&&attr(c[0],'val')==='phClr'&&!elements(c[0]).length?rgb(ref):null;
  };
  let fill=noFill?'none':gradFill?'gradient':fillNode?rgb(fillNode):inheritedPaint(themeEntry('fillStyleLst',child(style,'fillRef')),child(style,'fillRef')),gradient;
  if(gradFill){
    const onlyAttrs=(node,names)=>[...(node?.attributes??[])].every(a=>!a.namespaceURI&&names.includes(a.name));
    if(gradFill.namespaceURI!==A||!onlyAttrs(gradFill,['flip','rotWithShape'])||attr(gradFill,'flip','none')!=='none'||attr(gradFill,'rotWithShape','1')!=='1')return null;
    const parts=elements(gradFill),list=child(gradFill,'gsLst'),lin=child(gradFill,'lin'),tile=child(gradFill,'tileRect');
    if(parts.length!==3||!list||!lin||!tile||parts.some(p=>p.namespaceURI!==A)||!onlyAttrs(list,[])||!onlyAttrs(lin,['ang','scaled'])||attr(lin,'ang')!=='0'||attr(lin,'scaled','1')!=='1'||elements(lin).length||!onlyAttrs(tile,[])||elements(tile).length)return null;
    const stops=elements(list);if(stops.length!==2)return null;
    const colors=stops.map((stop,i)=>{
      const values=elements(stop),c=values[0];
      if(stop.namespaceURI!==A||stop.localName!=='gs'||!onlyAttrs(stop,['pos'])||attr(stop,'pos')!==String(i*100000)||values.length!==1||c.namespaceURI!==A||c.localName!=='srgbClr'||!onlyAttrs(c,['val'])||elements(c).length||!/^[a-f\d]{6}$/i.test(attr(c,'val','')))return null;
      return `#${attr(c,'val').toLowerCase()}`;
    });
    if(colors.some(c=>!c))return null;
    gradient={id:`g-${sha256(uniqueKey).slice(0,20)}`,from:colors[0],to:colors[1]};
  }
  if(!fill)return null;
  const line=child(pr,'ln')??themeEntry('lnStyleLst',child(style,'lnRef'));if(!line||line.namespaceURI!==A||line.localName!=='ln')return null;
  if(attr(line,'cap','flat')!=='flat'||attr(line,'cmpd','sng')!=='sng'||attr(line,'algn','ctr')!=='ctr')return null;
  if(elements(line).some(n=>!['solidFill','noFill','prstDash','headEnd','tailEnd','miter'].includes(n.localName)))return null;
  if(child(line,'prstDash')&&attr(child(line,'prstDash'),'val')!=='solid')return null;
  if(['headEnd','tailEnd'].some(n=>child(line,n)&&attr(child(line,n),'type','none')!=='none'))return null;
  const lineFill=child(line,'solidFill'),lineNone=child(line,'noFill');
  if(Boolean(lineFill)===Boolean(lineNone))return null;
  const stroke=lineNone?'none':child(pr,'ln')?rgb(lineFill):inheritedPaint(lineFill,child(style,'lnRef'));if(!stroke)return null;
  const rawWidth=attr(line,'w','12700'),rawAngle=attr(xfrm,'rot','0');
  if(!/^\d+$/.test(rawWidth)||!/^[-]?\d+$/.test(rawAngle)||!Number.isSafeInteger(Number(rawWidth))||!Number.isSafeInteger(Number(rawAngle))||Math.abs(Number(rawAngle))>2147483647)return null;
  const width=lineNone?0:px(Number(rawWidth))*map.sx,angle=Number(rawAngle)/60000%360;
  const vector={shape,width:box.w,height:box.h,angle,fill,stroke,strokeWidth:width,...adjustments,...(gradient?{gradient}:{}),...(num(xfrm,'flipV')?{flipV:true}:{})};
  if(!validPresetVector(vector))return null;
  const artwork=presetVectorMarkup(vector),w=artwork.width,h=artwork.height;
  return {svg:artwork.svg,vector,box:{x:box.x+box.w/2-w/2,y:box.y+box.h/2-h/2,w,h}};
}
