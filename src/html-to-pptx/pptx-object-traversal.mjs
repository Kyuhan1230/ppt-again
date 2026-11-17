import path from 'node:path';
import {diagnostic} from './model.mjs';
import {resolveAsset} from './assets.mjs';
import {EMU_PER_CSS_PX} from './canvas.mjs';
import {readMathTableFallback,stretchImageFragment,readMathTransform} from './math-table-fallback.mjs';
import {validPresetVector,presetVectorMarkup} from './preset-vector.mjs';
import {snippedTextGeometry} from './snip-text.mjs';
import {simpleOuterShadow} from './shadow.mjs';
import {nearlyEllipticalGeometry} from './ellipse-geometry.mjs';
import {readPictureObject,readShapeTextObject,readGraphicFrame,rotatedPresetArtwork} from './pptx-object-readers.mjs';

const elements=n=>[...(n?.childNodes??[])].filter(c=>c.nodeType===1);
const child=(n,name)=>elements(n).find(c=>c.localName===name);
const all=(n,name)=>n?[...n.getElementsByTagNameNS('*',name)]:[];
const attr=(n,k,fallback=null)=>n?.hasAttribute(k)?n.getAttribute(k):fallback;
const num=(n,k,fallback=0)=>Number(attr(n,k,fallback));
const px=n=>n/EMU_PER_CSS_PX;
const P='http://schemas.openxmlformats.org/presentationml/2006/main',R='http://schemas.openxmlformats.org/officeDocument/2006/relationships',A='http://schemas.openxmlformats.org/drawingml/2006/main';
const MC='http://schemas.openxmlformats.org/markup-compatibility/2006',MATH='http://schemas.openxmlformats.org/officeDocument/2006/math';

// One visitor per slide owns IDs, z-order, group coordinates and diagnostics
// across the master, layout and slide layers; readers never traverse separately.
export function createPptxVisitor({file,doc,rels,relationships,allowPartial,objectContexts,slide,index,model,diagnostics,relationRecords,parts,validateMathPng,renderChartImages,chartPackages,budget,contentTypeFor,workbooks,usedCharts,usedWorkbooks,blockedChartWorkbooks,unverifiedChartTargets,pendingChartImages,chartPointCounter,colorMap,colorScheme,themeDoc,issue,placeholderContext,readText,color,setActiveNodeId,setCurrentContext}){
  const seen=new Set();let objectCount=0;
  const visit=(parent,parentId=null,map={sx:1,sy:1,tx:0,ty:0},depth=0,ownerFile=file,layer='slide')=>{
    const ownerRels=ownerFile===file?rels:relationships(ownerFile);
    for(const object of elements(parent)){
      if(['nvGrpSpPr','grpSpPr','extLst'].includes(object.localName))continue;
      // Placeholder objects define inherited properties, not independently
      // visible text/artwork. Slide objects matching them are parsed below.
      const nonVisualName={sp:'nvSpPr',pic:'nvPicPr',graphicFrame:'nvGraphicFramePr',cxnSp:'nvCxnSpPr',grpSp:'nvGrpSpPr'}[object.localName];
      const nonVisual=child(object,nonVisualName);
      if(layer!=='slide'&&(child(child(nonVisual,'nvPr'),'ph')||['1','true'].includes(attr(child(nonVisual,'cNvPr'),'hidden','0'))))continue;
      if(++objectCount>10000)throw new Error('장표 객체 한도를 초과했습니다.');
      const nv=all(object,'cNvPr')[0],base=attr(nv,'name')||`object-${attr(nv,'id',String(objectCount))}`;
      const key=layer==='slide'?base:`${layer}:${base}`;
      let id=key,suffix=1;while(seen.has(id))id=`${key}-${suffix++}`;seen.add(id);
      setActiveNodeId(id);
      const currentObjectContext={slideId:slide.id,slideNumber:index+1,nodeId:id};setCurrentContext(currentObjectContext);
      objectContexts.set(`${slide.id}:${id}`,{objectName:base,objectIndex:objectCount,objectType:object.localName==='pic'?'image':object.localName==='grpSp'?'group':object.localName==='graphicFrame'?(all(object,'tbl').length?'table':'chart'):all(object,'t').length?'text':'shape'});
      Object.assign(currentObjectContext,objectContexts.get(`${slide.id}:${id}`));
      if(allowPartial&&object.namespaceURI===MC&&object.localName==='AlternateContent'){
        const alternatives=elements(object),choices=alternatives.filter(n=>n.namespaceURI===MC&&n.localName==='Choice'),fallbacks=alternatives.filter(n=>n.namespaceURI===MC&&n.localName==='Fallback');
        const fallback=fallbacks.length===1?elements(fallbacks[0]):[],shape=fallback.length===1?fallback[0]:null;
        const math=choices.some(c=>[...c.getElementsByTagNameNS(MATH,'oMath')].length>0);
        const tableFallback=math&&alternatives.every(n=>n.namespaceURI===MC&&['Choice','Fallback'].includes(n.localName))&&alternatives.at(-1)===fallbacks[0]?readMathTableFallback(shape):null;
        if(tableFallback){
          const relation=relationRecords(ownerFile).get(tableFallback.rid);
          if(!relation||relation.type!==`${R}/image`||!relation.target.startsWith('ppt/media/')||!parts[relation.target]||path.posix.extname(relation.target).toLowerCase()!=='.png')throw new Error('수식 그림 관계 또는 이미지가 잘못되었습니다.');
          validateMathPng(relation.target);
          const asset=resolveAsset('math-fallback.png',[{path:'math-fallback.png',mime:'image/png',bytes:parts[relation.target]}]);
          if(!model.assets.some(a=>a.id===asset.id))model.assets.push(asset);
          objectContexts.set(`${slide.id}:${id}`,{...objectContexts.get(`${slide.id}:${id}`),objectType:'image'});
          for(const [fragmentIndex,fragment] of tableFallback.fragments.entries()){
            let fragmentId=`${id}::math-cell-${fragmentIndex+1}`;while(seen.has(fragmentId))fragmentId+='-';seen.add(fragmentId);
            const box={x:px(fragment.box.x*map.sx+map.tx),y:px(fragment.box.y*map.sy+map.ty),w:px(fragment.box.w*map.sx),h:px(fragment.box.h*map.sy)};
            const crop=Object.values(fragment.crop).some(Boolean)?fragment.crop:undefined;
            slide.nodes.push({id:fragmentId,kind:'image',boxCssPx:box,zOrder:slide.nodes.length,content:{assetId:asset.id,...(crop?{crop}:{})},style:{opacity:1},sourceRef:{slideId:slide.id,elementId:fragmentId},parentId});
          }
          diagnostics.push({...diagnostic('PPTX_MATH_IMAGE_FALLBACK','수식은 그림으로 보존되며 수식을 개별 편집할 수 없습니다. 그림 배치는 원본과 다를 수 있습니다.',{slideId:slide.id,nodeId:id},'warning'),slideNumber:index+1,objectName:base,objectType:'image',reasonKey:'PPTX_MATH_IMAGE_FALLBACK'});
          if(tableFallback.omittedBorder)diagnostics.push({...diagnostic('PPTX_MATH_TABLE_BORDER_OMITTED','수식 대체 표의 테두리는 별도 선으로 옮기지 않아 이미지에 포함되지 않은 부분이 빠질 수 있습니다.',{slideId:slide.id,nodeId:id},'warning'),slideNumber:index+1,objectName:base,objectType:'image',reasonKey:'PPTX_MATH_TABLE_BORDER_OMITTED'});
          continue;
        }
        const pr=shape?.namespaceURI===P&&shape.localName==='sp'?child(shape,'spPr'):null,xfrm=child(pr,'xfrm'),off=child(xfrm,'off'),ext=child(xfrm,'ext');
        const geom=child(pr,'prstGeom'),fill=child(pr,'blipFill'),blip=child(fill,'blip'),stretch=child(fill,'stretch'),fillRect=child(stretch,'fillRect');
        const mathFallback=math&&alternatives.every(n=>n.namespaceURI===MC&&['Choice','Fallback'].includes(n.localName))&&alternatives.at(-1)===fallbacks[0]&&shape&&pr?.namespaceURI===P&&xfrm?.namespaceURI===A&&geom?.namespaceURI===A&&attr(geom,'prst')==='rect'&&fill?.namespaceURI===A&&blip?.namespaceURI===A&&stretch?.namespaceURI===A&&
          elements(pr).length===3&&elements(pr).every(n=>n.namespaceURI===A&&['xfrm','prstGeom','blipFill'].includes(n.localName))&&elements(xfrm).length===2&&elements(xfrm).every(n=>n.namespaceURI===A&&['off','ext'].includes(n.localName))&&
          elements(fill).length===2&&elements(fill).every(n=>n.namespaceURI===A&&['blip','stretch'].includes(n.localName))&&elements(blip).length===0&&elements(stretch).length<=1&&elements(stretch).every(n=>n.namespaceURI===A&&n.localName==='fillRect')&&
          !all(shape,'t').some(t=>t.textContent.trim());
        if(mathFallback){
          const transform=readMathTransform(xfrm);
          if(transform.rotation||transform.flipH||transform.flipV){issue('회전·뒤집힌 수식 그림은 아직 지원하지 않습니다.',id);continue;}
          const rid=blip.getAttributeNS(R,'embed'),relation=relationRecords(ownerFile).get(rid);
          if(blip.hasAttributeNS(R,'link')||!relation||relation.type!==`${R}/image`||!relation.target.startsWith('ppt/media/')||!parts[relation.target]||path.posix.extname(relation.target).toLowerCase()!=='.png')throw new Error('수식 그림 관계 또는 이미지가 잘못되었습니다.');
          validateMathPng(relation.target);
          const asset=resolveAsset('math-fallback.png',[{path:'math-fallback.png',mime:'image/png',bytes:parts[relation.target]}]);
          if(!model.assets.some(a=>a.id===asset.id))model.assets.push(asset);
          const fragment=stretchImageFragment(transform.box,fillRect);
          if(!fragment){issue('수식 그림의 채우기가 도형 밖에 있어 표시할 수 없습니다.',id);continue;}
          const box={x:px(fragment.box.x*map.sx+map.tx),y:px(fragment.box.y*map.sy+map.ty),w:px(fragment.box.w*map.sx),h:px(fragment.box.h*map.sy)};
          const crop=Object.values(fragment.crop).some(Boolean)?fragment.crop:undefined;
          objectContexts.set(`${slide.id}:${id}`,{...objectContexts.get(`${slide.id}:${id}`),objectType:'image'});
          slide.nodes.push({id,kind:'image',boxCssPx:box,zOrder:slide.nodes.length,content:{assetId:asset.id,...(crop?{crop}:{})},style:{opacity:1},sourceRef:{slideId:slide.id,elementId:id},parentId});
          diagnostics.push({...diagnostic('PPTX_MATH_IMAGE_FALLBACK','수식은 그림으로 보존되며 수식을 개별 편집할 수 없습니다. 그림 배치는 원본과 다를 수 있습니다.',{slideId:slide.id,nodeId:id},'warning'),slideNumber:index+1,objectName:base,objectType:'image',reasonKey:'PPTX_MATH_IMAGE_FALLBACK'});
          continue;
        }
      }
      if(!['sp','cxnSp','pic','grpSp','graphicFrame'].includes(object.localName)){issue('이 객체는 아직 지원하지 않습니다.',id);continue;}
      const group=object.localName==='grpSp',pr=child(object,group?'grpSpPr':'spPr');
      let xfrm=object.localName==='graphicFrame'?child(object,'xfrm'):child(pr,'xfrm');
      if(!xfrm&&object.localName==='sp'&&layer==='slide'&&!parentId){
        const inheritedPlaceholder=placeholderContext(object);
        if(inheritedPlaceholder.ph&&!inheritedPlaceholder.ambiguousLayout)xfrm=child(child(inheritedPlaceholder.matchedLayout,'spPr'),'xfrm')??child(child(inheritedPlaceholder.masterShape,'spPr'),'xfrm');
      }
      if(!xfrm){issue('마스터에 상속된 객체 위치는 아직 지원하지 않습니다.',id);continue;}
      const position=child(xfrm,'off'),extent=child(xfrm,'ext'),isLine=object.localName==='cxnSp'||attr(child(pr,'prstGeom'),'prst')==='line';
      if(!position||!extent||!['x','y'].every(k=>Number.isFinite(num(position,k,NaN)))||!['cx','cy'].every(k=>Number.isFinite(num(extent,k,NaN))&&num(extent,k)>=0)||!Number.isFinite(num(xfrm,'rot')))throw new Error('객체 좌표·크기·회전값이 잘못되었습니다.');
      if(!isLine&&(num(extent,'cx')===0||num(extent,'cy')===0)){issue('너비 또는 높이가 0인 객체는 화면 재현을 확정할 수 없습니다.',id,'PPTX_ZERO_SIZE_OBJECT_UNSUPPORTED');continue;}
      const off=child(xfrm,'off'),ext=child(xfrm,'ext');
      const raw={x:num(off,'x'),y:num(off,'y'),w:num(ext,'cx'),h:num(ext,'cy')};
      const box={x:px(raw.x*map.sx+map.tx),y:px(raw.y*map.sy+map.ty),w:px(raw.w*map.sx),h:px(raw.h*map.sy)};
      // Bounds describe known unrotated PPT coordinates, not a rendered original image.
      const bounds={x:box.x/slide.widthCssPx,y:box.y/slide.heightCssPx,w:box.w/slide.widthCssPx,h:box.h/slide.heightCssPx};
      if(Object.values(bounds).every(Number.isFinite)&&bounds.x>=0&&bounds.y>=0&&bounds.w>0&&bounds.h>0&&bounds.x+bounds.w<=1&&bounds.y+bounds.h<=1)objectContexts.get(`${slide.id}:${id}`).bounds=bounds;
      const node={id,kind:'shape',boxCssPx:box,zOrder:slide.nodes.length,content:{},style:{opacity:1},sourceRef:{slideId:slide.id,elementId:id},parentId};
      let pictureBorder=null;
      let lineArrowSpecs=[];
      const transformed=num(xfrm,'rot')!==0||(!['cxnSp'].includes(object.localName)&&!['line'].includes(attr(child(pr,'prstGeom'),'prst'))&&(num(xfrm,'flipH')||num(xfrm,'flipV')));
      const textShape=child(object,'txBody');
      const snipped=allowPartial&&object.localName==='sp'&&all(textShape,'t').some(t=>t.textContent.length)?snippedTextGeometry(child(pr,'prstGeom'),xfrm,raw,map):null;
      if(transformed&&!snipped||allowPartial&&object.localName==='sp'&&['downArrow','rightArrow'].includes(attr(child(pr,'prstGeom'),'prst'))){
        const artwork=allowPartial&&rotatedPresetArtwork(object,pr,xfrm,box,map,color,colorScheme,colorMap,themeDoc,`${slide.id}:${id}`);
        if(!artwork){if(transformed){issue('회전·뒤집힌 객체는 아직 지원하지 않습니다.',id);continue;}}
        else{
          const asset=resolveAsset('shape.svg',[{path:'shape.svg',mime:'image/svg+xml',bytes:Buffer.from(artwork.svg)}]);
          if(!model.assets.some(a=>a.id===asset.id))model.assets.push(asset);
          slide.nodes.push({...node,kind:'image',boxCssPx:artwork.box,content:{assetId:asset.id,editableVector:artwork.vector}});
          if(!['ellipse','downArrow','rightArrow'].includes(artwork.vector.shape))diagnostics.push({...diagnostic('PPTX_SHAPE_IMAGE_APPROXIMATED','HTML에서는 SVG 도형을 편집할 수 있지만, 다시 만든 PPT에서는 그림으로 저장되어 개별 편집할 수 없습니다.',{slideId:slide.id,nodeId:id},'warning'),slideNumber:index+1,...objectContexts.get(`${slide.id}:${id}`),reasonKey:'PPTX_SHAPE_IMAGE_APPROXIMATED'});
          continue;
        }
      }
      if(object.localName==='graphicFrame'){
        readGraphicFrame({object,node,ownerFile,relationRecords,parts,chartPackages,budget,contentTypeFor,workbooks,renderChartImages,usedCharts,usedWorkbooks,blockedChartWorkbooks,unverifiedChartTargets,layer,nv,doc,slide,index,id,base,diagnostics,pendingChartImages,chartPointCounter,model,box,readText,color,allowPartial,map});
        continue;
      }
      if(group){
        if(depth>=32){issue('그룹 중첩 깊이가 지원 한도를 초과합니다.',id);continue;}
        node.kind='group';slide.nodes.push(node);
        const co=child(xfrm,'chOff'),ce=child(xfrm,'chExt');
        if(!co||!ce||!['x','y'].every(k=>Number.isFinite(num(co,k,NaN)))||!['cx','cy'].every(k=>Number.isFinite(num(ce,k,NaN))&&num(ce,k)>=0))throw new Error('그룹 좌표·배율이 잘못되었습니다.');
        if(num(ce,'cx')===0||num(ce,'cy')===0){issue('크기가 0인 그룹 내부 좌표는 화면 재현을 확정할 수 없습니다.',id,'PPTX_ZERO_SIZE_OBJECT_UNSUPPORTED');continue;}
        const sx=raw.w/num(ce,'cx'),sy=raw.h/num(ce,'cy');
        if(!Number.isFinite(sx)||!Number.isFinite(sy)||sx<=0||sy<=0)throw new Error('그룹 좌표·배율이 잘못되었습니다.');
        if(Math.abs(sx-1)>1e-7||Math.abs(sy-1)>1e-7){
          if(!allowPartial){issue(Math.abs(sx-sy)>1e-7?'비균일 그룹 배율은 지원하지 않습니다.':'배율이 적용된 그룹의 글자·선 굵기는 아직 보존할 수 없습니다.',id);continue;}
          diagnostics.push({...diagnostic('PPTX_GROUP_SCALE_APPROXIMATED','배율이 적용된 그룹의 위치와 크기를 옮겼지만 글자 크기와 선 굵기는 원본과 다를 수 있습니다.',{slideId:slide.id,nodeId:id},'warning'),slideNumber:index+1,...objectContexts.get(`${slide.id}:${id}`),reasonKey:'PPTX_GROUP_SCALE_APPROXIMATED'});
        }
        visit(object,id,{sx:map.sx*sx,sy:map.sy*sy,tx:map.tx+map.sx*(raw.x-num(co,'x')*sx),ty:map.ty+map.sy*(raw.y-num(co,'y')*sy)},depth+1,ownerFile,layer);continue;
      }
      const unsupported=['custGeom','gradFill','pattFill','blipFill','effectDag','scene3d','sp3d'];
      const unsupportedPaint=elements(pr).some(n=>unsupported.includes(n.localName)&&!['gradFill','custGeom'].includes(n.localName));
      const effects=elements(child(pr,'effectLst'));
      const complexPaint=unsupportedPaint||effects.length;
      const pictureApprox=object.localName==='pic'&&allowPartial&&(child(pr,'gradFill')||child(pr,'custGeom')||complexPaint);
      const shapeGeometry=attr(child(pr,'prstGeom'),'prst');
      const roundedPicture=object.localName==='pic'&&(shapeGeometry==='ellipse'||allowPartial&&nearlyEllipticalGeometry(child(pr,'custGeom')));
      const pictureShadow=roundedPicture&&allowPartial&&effects.length===1&&Math.abs(map.sx-map.sy)<1e-7?simpleOuterShadow(effects[0],map.sx):null;
      const shadowApprox=allowPartial&&object.localName==='sp'&&['rect','roundRect'].includes(shapeGeometry)&&Boolean(child(pr,'solidFill'))&&!child(pr,'gradFill')&&!child(pr,'custGeom')&&!unsupportedPaint&&effects.length===1&&effects[0].namespaceURI===A&&effects[0].localName==='outerShdw';
      const shadow=shadowApprox&&Math.abs(map.sx-map.sy)<1e-7?simpleOuterShadow(effects[0],map.sx):null;
      if(pictureApprox){
        diagnostics.push({...diagnostic('PPTX_IMAGE_APPEARANCE_APPROXIMATED',roundedPicture?'사진의 자유형 마스크를 타원으로 근사했으며 일부 효과가 원본과 다를 수 있습니다.':'사진은 보존하지만 자유형 마스크와 그림자 등 일부 효과가 생략되어 원본과 다를 수 있습니다.',{slideId:slide.id,nodeId:id},'warning'),slideNumber:index+1,...objectContexts.get(`${slide.id}:${id}`),reasonKey:'PPTX_IMAGE_APPEARANCE_APPROXIMATED'});
      }else if(shadowApprox){
        const code=shadow?'PPTX_SHADOW_APPROXIMATED':'PPTX_SHADOW_OMITTED';
        diagnostics.push({...diagnostic(code,shadow?'바깥 그림자는 CSS로 근사했으며 PPT로 다시 저장할 때 빠질 수 있습니다.':'바깥 그림자는 표시하지 않았습니다.',{slideId:slide.id,nodeId:id},'warning'),slideNumber:index+1,...objectContexts.get(`${slide.id}:${id}`),reasonKey:code});
      }else{
        if(child(pr,'gradFill'))issue('도형 그라데이션 채움은 아직 지원하지 않습니다.',id,'PPTX_SHAPE_GRADIENT_UNSUPPORTED');
        if(child(pr,'custGeom'))issue('자유형 벡터 경로 도형은 아직 지원하지 않습니다.',id,'PPTX_CUSTOM_GEOMETRY_UNSUPPORTED');
        if(complexPaint)issue('복잡한 도형·채움·효과는 아직 지원하지 않습니다.',id);
      }
      const objectStyle=child(object,'style');
      const styleEntry=(list,ref)=>{const idx=num(ref,'idx');return Number.isInteger(idx)&&idx>0?elements(all(themeDoc,list)[0])[idx-1]:null;};
      const inheritedPaint=(fill,ref)=>{
        if(fill?.localName!=='solidFill')return null;
        const c=elements(fill);if(c.length!==1)return null;
        if(c[0].localName==='schemeClr'&&attr(c[0],'val')==='phClr')return elements(c[0]).length?null:color(ref);
        return color(fill);
      };
      let inheritedFill=null,inheritedStroke=null,inheritedStrokeWidth=0,styleUnsupported=false;
      const fillRef=child(objectStyle,'fillRef'),lineRef=child(objectStyle,'lnRef'),effectRef=child(objectStyle,'effectRef');
      if(fillRef&&!elements(pr).some(n=>['solidFill','noFill','blipFill','gradFill','pattFill','grpFill'].includes(n.localName))){
        inheritedFill=num(fillRef,'idx')===0?'transparent':inheritedPaint(styleEntry('fillStyleLst',fillRef),fillRef);
        if(inheritedFill===null)styleUnsupported=true;
      }
      const explicitLine=child(pr,'ln');
      // A shape may override only the line color while retaining the
      // theme's line width. Color and width inherit independently.
      const referencedLine=styleEntry('lnStyleLst',lineRef),referencedWidth=referencedLine?.localName==='ln'&&referencedLine.namespaceURI===A&&Number.isFinite(num(referencedLine,'w',NaN))&&num(referencedLine,'w')>=0?px(num(referencedLine,'w')):0;
      if(lineRef&&!elements(explicitLine).some(n=>['solidFill','noFill','gradFill','pattFill'].includes(n.localName))){
        const entry=styleEntry('lnStyleLst',lineRef);
        if(num(lineRef,'idx')===0)inheritedStroke='transparent';
        else if(entry?.localName==='ln'&&entry.namespaceURI===A&&attr(entry,'cap','flat')==='flat'&&attr(entry,'cmpd','sng')==='sng'&&attr(entry,'algn','ctr')==='ctr'&&Number.isFinite(num(entry,'w',NaN))&&num(entry,'w')>=0&&(!child(entry,'prstDash')||attr(child(entry,'prstDash'),'val')==='solid')&&elements(entry).every(n=>n.namespaceURI===A&&['solidFill','noFill','prstDash','miter'].includes(n.localName))){
          inheritedStroke=child(entry,'noFill')?'transparent':inheritedPaint(child(entry,'solidFill'),lineRef);
          inheritedStrokeWidth=inheritedStroke==='transparent'?0:px(num(explicitLine,'w',num(entry,'w')));
        }
        if(inheritedStroke===null)styleUnsupported=true;
      }
      if(effectRef&&num(effectRef,'idx')!==0){
        const effect=styleEntry('effectStyleLst',effectRef);
        if(!effect||elements(child(effect,'effectLst')).length||elements(effect).some(n=>n.localName!=='effectLst'))styleUnsupported=true;
      }
      if(styleUnsupported)issue('상속된 도형 채움·선·효과는 아직 지원하지 않습니다.',id);
      if(object.localName==='pic'){
        const picture=readPictureObject({object,explicitLine,lineRef,color,inheritedStroke,roundedPicture,pr,shapeGeometry,allowPartial,diagnostics,slide,id,objectContexts,referencedWidth,node,ownerRels,parts,issue,relationRecords,ownerFile,contentTypeFor,model,colorMap,colorScheme,pictureShadow,pictureApprox,index});
        if(picture.skip)continue;
        pictureBorder=picture.pictureBorder;
      }else{
        lineArrowSpecs=readShapeTextObject({object,pr,box,inheritedFill,inheritedStroke,inheritedStrokeWidth,color,referencedWidth,shadow,issue,id,allowPartial,isLine,diagnostics,slide,index,objectContexts,node,xfrm,readText,snipped,model,seen,parentId});
      }
      slide.nodes.push(node);
      if(pictureBorder&&!diagnostics.some(d=>d.nodeId===id&&d.severity==='error'&&d.recoverable===true)){
        let borderId=`${id}::outline`;while(seen.has(borderId))borderId+='-';seen.add(borderId);
        slide.nodes.push({id:borderId,kind:'shape',boxCssPx:{...box},zOrder:slide.nodes.length,content:{},style:{opacity:1,svgGeometry:true,fill:'transparent',radius:0,...pictureBorder},sourceRef:{slideId:slide.id,elementId:borderId},parentId});
      }
      if(lineArrowSpecs.length&&!diagnostics.some(d=>d.nodeId===id&&d.severity==='error'&&d.recoverable===true)){
        const {start,end}=node.style.endpoints??{},stroke=node.style.stroke,strokeWidth=node.style.strokeWidth;
        let made=false;
        for(const spec of lineArrowSpecs){
          const tip=spec.end==='head'?start:end,base=spec.end==='head'?end:start,dx=tip?.x-base?.x,dy=tip?.y-base?.y,distance=Math.hypot(dx,dy);
          if(!Number.isFinite(distance)||distance<=0||!Number.isFinite(strokeWidth)||strokeWidth<=0||!/^#[a-f\d]{6}$/i.test(stroke??'')){
            diagnostics.push({...diagnostic('PPTX_ARROWHEAD_OMITTED','이 화살표 끝은 표시하지 않았습니다.',{slideId:slide.id,nodeId:id},'warning'),slideNumber:index+1,...objectContexts.get(`${slide.id}:${id}`),reasonKey:'PPTX_ARROWHEAD_OMITTED'});continue;
          }
          const unit=Math.max(1,strokeWidth),length=({sm:5,med:7,lg:10})[spec.length]*unit,width=({sm:4,med:6,lg:8})[spec.width]*unit,angle=Math.atan2(dy,dx)*180/Math.PI;
          const vector={shape:'triangle',width:length,height:width,angle,fill:stroke,stroke:'none',strokeWidth:0,adj1:0,adj2:0};
          if(!validPresetVector(vector)){diagnostics.push({...diagnostic('PPTX_ARROWHEAD_OMITTED','이 화살표 끝은 표시하지 않았습니다.',{slideId:slide.id,nodeId:id},'warning'),slideNumber:index+1,...objectContexts.get(`${slide.id}:${id}`),reasonKey:'PPTX_ARROWHEAD_OMITTED'});continue;}
          const artwork=presetVectorMarkup(vector),center={x:tip.x-dx/distance*length/2,y:tip.y-dy/distance*length/2},asset=resolveAsset('arrow.svg',[{path:'arrow.svg',mime:'image/svg+xml',bytes:Buffer.from(artwork.svg)}]);
          if(!model.assets.some(a=>a.id===asset.id))model.assets.push(asset);
          let arrowId=`${id}::arrow-${spec.end}`;while(seen.has(arrowId))arrowId+='-';seen.add(arrowId);
          slide.nodes.push({id:arrowId,kind:'image',boxCssPx:{x:center.x-artwork.width/2,y:center.y-artwork.height/2,w:artwork.width,h:artwork.height},zOrder:slide.nodes.length,content:{assetId:asset.id,editableVector:vector},style:{opacity:1},sourceRef:{slideId:slide.id,elementId:arrowId},parentId});made=true;
        }
        if(made){
          if(strokeWidth<1)node.style.renderStrokeWidth=1;
          diagnostics.push({...diagnostic('PPTX_SHAPE_IMAGE_APPROXIMATED','화살표 끝을 SVG로 근사했습니다. 다시 만든 PPT에서는 선과 그림으로 분리됩니다.',{slideId:slide.id,nodeId:id},'warning'),slideNumber:index+1,...objectContexts.get(`${slide.id}:${id}`),reasonKey:'PPTX_SHAPE_IMAGE_APPROXIMATED'});
        }
      }
    }
  };
  return visit;
}
