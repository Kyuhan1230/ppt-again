import path from 'node:path';
import {diagnostic,validateModel} from './model.mjs';
import {resolveAsset,sha256} from './assets.mjs';
import {unwrapPptxFont} from './font-embedding.mjs';
import {decideFontEmbedding} from './font-policy.mjs';
import {openPptxPackage} from './pptx-package-reader.mjs';
import {readSlideText} from './pptx-text-reader.mjs';
import {createPptxVisitor} from './pptx-object-traversal.mjs';
import {preparePptxSlide,preparePlaceholderContext,visitPptxSlideLayers,themeBackground} from './pptx-slide-reader.mjs';
import {validatePngStructure} from './png-duotone.mjs';
import {readTransition} from './transitions.mjs';
import {safeError} from './error-log.mjs';

const elements=n=>[...(n?.childNodes??[])].filter(c=>c.nodeType===1);
const child=(n,name)=>elements(n).find(c=>c.localName===name);
const all=(n,name)=>n?[...n.getElementsByTagNameNS('*',name)]:[];
const attr=(n,k,fallback=null)=>n?.hasAttribute(k)?n.getAttribute(k):fallback;
const num=(n,k,fallback=0)=>Number(attr(n,k,fallback));
const P='http://schemas.openxmlformats.org/presentationml/2006/main',R='http://schemas.openxmlformats.org/officeDocument/2006/relationships';

// The optional renderer is called only after package and chart-data validation.
export async function readPptx(bytes,{fontPolicy='original',allowPartial=false}={}, {renderChartImages}={}){
  const selectedFont=({'malgun-gothic':'Malgun Gothic',pretendard:'Pretendard'})[fontPolicy];
  const objectContexts=new Map();
  let currentContext={};
  const diagnostics=[],pendingChartImages=[],model={version:1,sourceDigest:sha256(bytes??''),slides:[],assets:[]};
  const fail=(code,message,slideId=null,nodeId=null)=>diagnostics.push(diagnostic(code,message,{slideId,nodeId}));
  const done=()=>{const failed=diagnostics.some(d=>d.severity==='error');return {status:failed?'failed':'parsed',model:failed?null:model,diagnostics};};
  const fatalReasons=new Map([
    ['객체 좌표·크기·회전값이 잘못되었습니다.','PPTX_INVALID_COORDINATES'],
    ['그룹 좌표·배율이 잘못되었습니다.','PPTX_INVALID_COORDINATES'],
    ['그림 자르기 좌표가 잘못되었습니다.','PPTX_INVALID_CROP'],
    ['현재 차트에 연결되지 않은 내장 문서는 지원하지 않습니다.','PPTX_UNUSED_EMBEDDED_PACKAGE'],
    ['차트 오류로 내장 문서의 연결을 확인하지 못했습니다.','PPTX_EMBEDDED_LINK_UNVERIFIED'],
    ['검증한 차트 workbook 외의 패키지 관계는 지원하지 않습니다.','PPTX_UNEXPECTED_PACKAGE_RELATIONSHIP'],
    ['잘못된 PPTX XML입니다.','PPTX_XML_INVALID'],
    ['PPTX 필수 부분이 없습니다.','PPTX_REQUIRED_PART_MISSING'],
    ['DTD·엔터티는 지원하지 않습니다.','PPTX_XML_UNSAFE'],
    ['외부 참조가 있는 PPTX는 지원하지 않습니다.','PPTX_EXTERNAL_REFERENCE'],
    ['외부 PPTX 경로입니다.','PPTX_EXTERNAL_REFERENCE']
  ]);
  try{
    if(!['original','malgun-gothic','pretendard'].includes(fontPolicy)||typeof allowPartial!=='boolean')throw new Error('PPTX 가져오기 옵션이 올바르지 않습니다.');
    if(!(bytes instanceof Uint8Array)||bytes.length<22||bytes.length>64*1024*1024)throw new Error('PPTX 바이트는 64MiB 이내여야 합니다.');
    const {budget,parts,xml,relationCache,relationRecords,relationships,contentTypeFor,validateMathPng}=openPptxPackage(bytes);
    const workbooks=new Map(),chartPackages=new Map(),usedWorkbooks=new Set(),usedCharts=new Set(),blockedChartWorkbooks=new Map(),unverifiedChartTargets=new Set(),chartPointCounter={value:0};
    const presentation=xml('ppt/presentation.xml'),size=all(presentation,'sldSz')[0];
    model.pptxSizeEmu={width:num(size,'cx'),height:num(size,'cy')};
    const slideRefs=all(presentation,'sldId');if(!slideRefs.length||slideRefs.length>200)throw new Error('장표는 1~200장이어야 합니다.');
    const presRels=relationships('ppt/presentation.xml');
    const tableStylesPart=[...relationRecords('ppt/presentation.xml').values()].find(r=>r.type===`${R}/tableStyles`)?.target;
    const tableStylesDoc=tableStylesPart&&parts[tableStylesPart]?xml(tableStylesPart):null;
    const fontLists=all(presentation,'embeddedFontLst');
    if(fontLists.length){
      const reject=(code,message)=>{throw Object.assign(new Error(message),{code});};
      const list=fontLists[0],root=presentation.documentElement;
      if(fontLists.length!==1||list.namespaceURI!==P||list.parentNode!==root||root.namespaceURI!==P)reject('FONT_INVALID','내장 글꼴 목록의 구조가 잘못되었습니다.');
      const relationNodes=all(xml('ppt/_rels/presentation.xml.rels'),'Relationship'),fontRels=new Map(relationNodes.filter(n=>attr(n,'Type')===`${R}/font`).map(n=>[attr(n,'Id'),n]));
      const embedded=elements(list),faceCount=embedded.reduce((n,font)=>n+elements(font).filter(c=>c.localName!=='font').length,0);
      if(embedded.length>32||faceCount>32)reject('FONT_ASSET_LIMIT','내장 글꼴은 최대 32개 face까지 지원합니다.');
      const variants={regular:[400,'normal'],bold:[700,'normal'],italic:[400,'italic'],boldItalic:[700,'italic']};
      const used=new Set(),slots=new Set(),cached=new Map();let total=0;model.fonts=[];
      for(const item of embedded){
        const nodes=elements(item),names=nodes.filter(n=>n.localName==='font'),family=attr(names[0],'typeface');
        if(item.namespaceURI!==P||item.localName!=='embeddedFont'||names.length!==1||names[0].namespaceURI!==P||!family||!family.trim()||nodes.length<2)reject('FONT_INVALID','내장 글꼴 family 또는 구조가 잘못되었습니다.');
        for(const node of nodes.filter(n=>n.localName!=='font')){
          if(node.namespaceURI!==P||!Object.hasOwn(variants,node.localName))reject('FONT_STYLE_UNSUPPORTED','지원하지 않는 내장 글꼴 스타일입니다.');
          const [weight,style]=variants[node.localName],key=JSON.stringify([family.toLowerCase(),weight,style]);
          if(slots.has(key))reject('FONT_STYLE_MISMATCH','동일 family·스타일의 내장 글꼴이 중복됩니다.');slots.add(key);
          const id=node.getAttributeNS(R,'id'),target=presRels.get(id);
          if(!id||!fontRels.has(id)||!target?.startsWith('ppt/fonts/')||!parts[target])reject('FONT_INVALID','내장 글꼴 관계 또는 파일이 없습니다.');
          const contentType=contentTypeFor(target);
          if(!used.has(target)){total+=parts[target].length;used.add(target);}if(total>32*1024*1024)reject('FONT_ASSET_LIMIT','내장 글꼴 합계가 32 MiB를 초과했습니다.');
          // Explicit replacement never decodes or redistributes the embedded font.
          // Relationship, structure and package/face byte limits still apply.
          if(selectedFont){diagnostics.push({...diagnostic('PPTX_FONT_REPLACED','선택한 글꼴을 사용하며 원본 내장 글꼴 파일은 포함하지 않습니다.',{},'warning'),fontFamily:family,objectType:'font',reasonKey:'PPTX_EMBEDDED_FONT_IGNORED'});continue;}
          if(contentType!=='application/x-fontdata')reject('FONT_FORMAT_UNSUPPORTED','내장 글꼴 ContentType을 지원하지 않습니다.');
          if(!cached.has(target)){
            try{cached.set(target,unwrapPptxFont(parts[target]));}
            catch(error){
              // Preserve the authored font family without decoding or exporting
              // an unsupported embedded container. The browser must use a
              // locally installed font; malformed fonts remain fatal.
              if(!allowPartial||error.code!=='FONT_FORMAT_UNSUPPORTED')throw error;
              diagnostics.push({...diagnostic('FONT_FORMAT_UNSUPPORTED','원본 글꼴 이름은 유지하지만 내장 글꼴 파일은 포함하지 않습니다. 설치된 글꼴과 줄바꿈을 확인해 주세요.',{},'warning'),fontFamily:family,objectType:'font',reasonKey:'PPTX_EMBEDDED_FONT_LOCAL_ONLY'});
              continue;
            }
          }
          const {bytes:fontBytes,meta}=cached.get(target),permission=decideFontEmbedding(meta,{forEditing:true,subset:false});
          if(!permission.allowed)reject(permission.reason,'이 내장 글꼴은 편집용으로 사용할 수 없습니다.');
          if(meta.family!==family||meta.weight!==weight||meta.italic!==(style==='italic'))reject('FONT_STYLE_MISMATCH','PPT 글꼴 이름·스타일과 포함 파일이 다릅니다.');
          const asset=resolveAsset('embedded-font.ttf',[{path:'embedded-font.ttf',mime:'font/ttf',bytes:fontBytes}]);
          if(!model.assets.some(a=>a.id===asset.id))model.assets.push(asset);
          model.fonts.push({id:`font-${sha256(JSON.stringify([asset.sha256,family.toLowerCase(),weight,style]))}`,cssFamily:family,family,weight,style,assetId:asset.id,origin:'pptx'});
        }
      }
      if(model.fonts.length)model.version=2;else delete model.fonts;
    }
    for(const [index,ref] of slideRefs.entries()){
      currentContext={slideId:`slide-${index+1}`,slideNumber:index+1};
      const {file,doc,rels,tree,slide,inherited,masterPart,themeDoc,masterDoc,colorMap,colorScheme}=preparePptxSlide(ref,index,{presRels,xml,relationships,relationRecords,parts,model});
      let activeNodeId=null;
      const issue=(message,nodeId=null,explicitReason=null,styleDetail=null)=>{
        nodeId??=activeNodeId;
        const reasonKey=/명시적 글꼴·크기가 없는 상속/.test(message)?'PPTX_INHERITED_TEXT_UNRESOLVED':/테마·시스템 색/.test(message)?'PPTX_THEME_COLOR_UNSUPPORTED':/상속된 도형 채움·선·효과/.test(message)?'PPTX_INHERITED_SHAPE_STYLE_UNSUPPORTED':/테마 배경|단색 외의 PPT 배경/.test(message)?'PPTX_THEME_BACKGROUND_UNSUPPORTED':/마스터·레이아웃/.test(message)?'PPTX_MASTER_ARTWORK_UNSUPPORTED':/타이밍·전환/.test(message)?'PPTX_ANIMATION_UNSUPPORTED':/회전·뒤집힌/.test(message)?'PPTX_OBJECT_ROTATION_UNSUPPORTED':/객체 위치/.test(message)?'PPTX_OBJECT_POSITION_UNSUPPORTED':/그룹/.test(message)?'PPTX_GROUP_TRANSFORM_UNSUPPORTED':/그림/.test(message)?'PPTX_IMAGE_UNSUPPORTED':/도형·채움·효과|도형 채움/.test(message)?'PPTX_SHAPE_EFFECT_UNSUPPORTED':/글꼴|텍스트|글자|문단|글머리/.test(message)?'PPTX_TEXT_STYLE_UNSUPPORTED':'PPTX_OBJECT_UNSUPPORTED';
        diagnostics.push({...diagnostic('UNSUPPORTED_OBJECT',message,{slideId:slide.id,nodeId}),slideNumber:index+1,...(objectContexts.get(`${slide.id}:${nodeId}`)??{objectType:'background'}),reasonKey:explicitReason??reasonKey,recoverable:true,...(styleDetail?{styleDetail}:{})});
      };
      if(all(doc,'timing').length||all(doc,'transition').length){
        const transitions=all(doc,'transition');
        if(transitions.length){const preserved=readTransition(transitions[0]);if(preserved)slide.transition=preserved;}
        if(allowPartial){slide.animationOmitted=true;diagnostics.push({...diagnostic('PPTX_ANIMATION_OMITTED','애니메이션과 화면 전환은 정적 HTML에서 재현되지 않습니다.',{slideId:slide.id},'warning'),slideNumber:index+1,objectType:'background',reasonKey:'PPTX_ANIMATION_UNSUPPORTED'});}
        else issue('타이밍·전환 효과는 정적 HTML로 보존할 수 없습니다.');
      }
      const color=(fill,fallback='transparent')=>{
        if(!fill)return fallback;
        const c=elements(fill)[0];let source=c;
        if(c?.localName==='schemeClr'){
          const key=attr(colorMap,attr(c,'val'),attr(c,'val'));
          source=elements(child(colorScheme,key))[0];
        }
        const presetColors={black:'000000',white:'FFFFFF'};
        const hex=source?.localName==='sysClr'?attr(source,'lastClr',''):source?.localName==='srgbClr'?attr(source,'val',''):source?.localName==='prstClr'?(presetColors[attr(source,'val')]??''):'';
        if(source===c&&source?.localName==='srgbClr'&&!/^[a-f\d]{6}$/i.test(hex))throw new Error('잘못된 RGB 색입니다.');
        if(!/^[a-f\d]{6}$/i.test(hex)){issue('테마·시스템 색은 명시적 RGB로 바꿔 주세요.');return fallback;}
        let rgb=[0,2,4].map(i=>parseInt(hex.slice(i,i+2),16)/255),alpha=1;
        for(const effect of [...elements(source),...(source===c?[]:elements(c))]){
          const v=num(effect,'val',NaN)/100000;
          if(!Number.isFinite(v)){issue('색 변환 효과는 지원하지 않습니다.');continue;}
          if(effect.localName==='alpha'){alpha=Math.max(0,Math.min(1,v));continue;}
          if(effect.localName==='shade'){rgb=rgb.map(n=>n*v);continue;}
          if(effect.localName==='tint'){rgb=rgb.map(n=>n+(1-n)*(1-v));continue;}
          if(['lumMod','lumOff'].includes(effect.localName)){
            const max=Math.max(...rgb),min=Math.min(...rgb),l=(max+min)/2,d=max-min;
            const sat=d===0?0:d/(1-Math.abs(2*l-1));
            let h=0;if(d){if(max===rgb[0])h=((rgb[1]-rgb[2])/d)%6;else if(max===rgb[1])h=(rgb[2]-rgb[0])/d+2;else h=(rgb[0]-rgb[1])/d+4;h=(h*60+360)%360;}
            const light=Math.max(0,Math.min(1,effect.localName==='lumMod'?l*v:l+v));
            const chroma=(1-Math.abs(2*light-1))*sat,x=chroma*(1-Math.abs((h/60)%2-1)),m=light-chroma/2;
            rgb=(h<60?[chroma,x,0]:h<120?[x,chroma,0]:h<180?[0,chroma,x]:h<240?[0,x,chroma]:h<300?[x,0,chroma]:[chroma,0,x]).map(n=>n+m);continue;
          }
          issue('색 변환 효과는 지원하지 않습니다.');
        }
        const channels=rgb.map(n=>Math.round(Math.max(0,Math.min(1,n))*255));
        return alpha===1?`#${channels.map(n=>n.toString(16).padStart(2,'0')).join('')}`:`rgba(${channels.join(',')},${alpha})`;
      };
      const backgrounds=[doc,...[...inherited].map(part=>xml(part))].map(d=>child(child(d.documentElement,'cSld'),'bg')).filter(Boolean);
      const bg=child(backgrounds[0],'bgPr'),bgRef=child(backgrounds[0],'bgRef');
      if(bg){
        slide.background.color=color(child(bg,'solidFill'),'#ffffff');
        if(elements(bg).some(n=>!['solidFill','effectLst'].includes(n.localName)))issue('단색 외의 PPT 배경은 지원하지 않습니다.');
      }else if(bgRef){
        const master=[...inherited].find(part=>part.startsWith('ppt/slideMasters/'));
        const theme=master&&[...relationRecords(master).values()].find(r=>r.type===`${R}/theme`)?.target;
        const layouts=[...inherited].filter(part=>part.startsWith('ppt/slideLayouts/'));
        const resolved=theme&&parts[theme]?themeBackground(bgRef,xml(theme),xml(master),[doc,...layouts.map(part=>xml(part))]):null;
        if(resolved)slide.background.color=resolved;else issue('이 테마 배경은 아직 지원하지 않습니다.');
      }
      const approximatedParagraphs=new Set();
      const textApproximations=new Set();
      const {placeholderContext}=preparePlaceholderContext({inherited,xml,masterDoc});
      const readText=(tx,id)=>readSlideText(tx,id,{issue,slide,diagnostics,textApproximations,approximatedParagraphs,selectedFont,allowPartial,color,tableStylesDoc,masterDoc,presentation,themeDoc,placeholderContext});
      const visit=createPptxVisitor({file,doc,rels,relationships,allowPartial,objectContexts,slide,index,model,diagnostics,relationRecords,parts,validateMathPng,renderChartImages,chartPackages,budget,contentTypeFor,workbooks,usedCharts,usedWorkbooks,blockedChartWorkbooks,unverifiedChartTargets,pendingChartImages,chartPointCounter,colorMap,colorScheme,themeDoc,issue,placeholderContext,readText,color,setActiveNodeId:value=>{activeNodeId=value;},setCurrentContext:value=>{currentContext=value;}});
      visitPptxSlideLayers({doc,inherited,xml,masterPart,masterDoc,tree},visit);
      currentContext={};
    }
    const blockedWorkbookPaths=new Set(blockedChartWorkbooks.values());
    for(const name of Object.keys(parts)){
      if(!name.endsWith('/')&&(/(?:^|\/)embeddings\//i.test(name)||contentTypeFor(name)==='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')&&!usedWorkbooks.has(name)&&!blockedWorkbookPaths.has(name))throw new Error(unverifiedChartTargets.has(name)?'차트 오류로 내장 문서의 연결을 확인하지 못했습니다.':'현재 차트에 연결되지 않은 내장 문서는 지원하지 않습니다.');
    }
    for(const [owner,records] of relationCache)for(const relation of records.values())if(relation.type===`${R}/package`&&(!usedCharts.has(owner)||!usedWorkbooks.has(relation.target))&&blockedChartWorkbooks.get(owner)!==relation.target)throw new Error('검증한 차트 workbook 외의 패키지 관계는 지원하지 않습니다.');
    if(pendingChartImages.length){
      const blocked=diagnostics.some(d=>d.severity==='error'&&(!allowPartial||d.recoverable!==true));
      const reportFailure=(item,code)=>{const reasonKey=({POWERPOINT_UNAVAILABLE:'PPTX_CHART_IMAGE_UNAVAILABLE',POWERPOINT_OPEN_FAILED:'PPTX_CHART_IMAGE_OPEN_FAILED',POWERPOINT_SECURITY_FAILED:'PPTX_CHART_IMAGE_SECURITY_FAILED',POWERPOINT_CHART_NOT_FOUND:'PPTX_CHART_IMAGE_NOT_FOUND',CHART_IMAGE_BLOCKED:'PPTX_CHART_IMAGE_BLOCKED',POWERPOINT_TIMEOUT:'PPTX_CHART_IMAGE_TIMEOUT'})[code]??'PPTX_CHART_IMAGE_FAILED';diagnostics.push({...diagnostic('UNSUPPORTED_OBJECT','차트 그림을 추출하지 못했습니다.',{slideId:item.slideId,nodeId:item.nodeId}),slideNumber:item.slideNumber,objectName:item.objectName,objectType:'chart',reasonKey,blocking:true});};
      try{
        if(blocked)throw Object.assign(new Error('다른 안전성 또는 구조 오류로 차트 그림 추출을 실행하지 않았습니다.'),{code:'CHART_IMAGE_BLOCKED'});
        const rendered=await renderChartImages(bytes,pendingChartImages.map(({slideNumber,shapeId})=>({slideNumber,shapeId})));
        if(!Array.isArray(rendered)||rendered.length!==pendingChartImages.length)throw new Error('차트 그림 추출 결과의 개수가 다릅니다.');
        for(const item of pendingChartImages){
          const matches=rendered.filter(r=>r.slideNumber===item.slideNumber&&r.shapeId===item.shapeId);
          if(matches.length!==1||matches[0].errorCode||!(matches[0].bytes instanceof Uint8Array)){reportFailure(item,matches[0]?.errorCode);continue;}
          try{validatePngStructure(matches[0].bytes);const asset=resolveAsset('chart.png',[{path:'chart.png',mime:'image/png',bytes:matches[0].bytes}]);
            if(!model.assets.some(a=>a.id===asset.id))model.assets.push(asset);
            item.node.kind='image';item.node.content={assetId:asset.id};
            diagnostics.push({...diagnostic('PPTX_CHART_IMAGE_FALLBACK','원본 차트를 그림으로 보존했습니다. 이 차트는 편집할 수 없습니다.',{slideId:item.slideId,nodeId:item.nodeId},'warning'),slideNumber:item.slideNumber,objectName:item.objectName,objectType:'chart',reasonKey:'PPTX_CHART_IMAGE_FALLBACK'});
          }catch{reportFailure(item,'POWERPOINT_EXPORT_FAILED');}
        }
      }catch(error){
        for(const item of pendingChartImages)reportFailure(item,error.code);
      }
    }
    if(selectedFont){
      for(const slide of model.slides){
        const replaceStyle=style=>{if(style&&typeof style.fontFamily==='string')style.fontFamily=selectedFont;};
        for(const node of slide.nodes){
          replaceStyle(node.style);for(const run of node.content.runs??[])replaceStyle(run.style);
          for(const cell of node.content.cells??[]){replaceStyle(cell.style);for(const run of cell.runs??[])replaceStyle(run.style);}
        }
        diagnostics.push({...diagnostic('PPTX_FONT_REPLACED','선택한 글꼴로 텍스트를 바꿨습니다. 설치 여부와 줄바꿈·배치를 확인해 주세요.',{slideId:slide.id},'warning'),slideNumber:model.slides.indexOf(slide)+1,objectType:'font',fontFamily:selectedFont,reasonKey:'PPTX_FONT_SELECTED'});
      }
    }
    if(allowPartial){
      // Recover only known representation limitations. Parser/security/limit and
      // malformed-object errors are never downgraded.
      for(const slide of model.slides){
        const recoverable=diagnostics.filter(d=>d.slideId===slide.id&&d.code==='UNSUPPORTED_OBJECT'&&d.severity==='error'&&d.recoverable===true);
        const bad=new Set(recoverable.map(d=>d.nodeId).filter(Boolean));let changed=true;
        while(changed){changed=false;for(const n of slide.nodes)if(bad.has(n.parentId)&&!bad.has(n.id)){bad.add(n.id);changed=true;}}
        const oldNodes=new Map(slide.nodes.map(n=>[n.id,n]));slide.nodes=slide.nodes.filter(n=>!bad.has(n.id));
        const groups=new Map();for(const d of recoverable){const key=d.nodeId??'slide-background';if(!groups.has(key))groups.set(key,[]);groups.get(key).push(d);}
        let offset=0;for(const [key,items] of groups){
          const old=oldNodes.get(key),bounds=objectContexts.get(`${slide.id}:${key}`)?.bounds;
          // Unsupported charts/tables may have no model node. Reuse only their
          // already-validated native coordinates, never an invented location.
          const hasBounds=bounds&&['x','y','w','h'].every(k=>Number.isFinite(bounds[k]))&&bounds.x>=0&&bounds.y>=0&&bounds.w>0&&bounds.h>0&&bounds.x+bounds.w<=1&&bounds.y+bounds.h<=1;
          const candidate=old?.boxCssPx??(hasBounds?{x:bounds.x*slide.widthCssPx,y:bounds.y*slide.heightCssPx,w:bounds.w*slide.widthCssPx,h:bounds.h*slide.heightCssPx}:null);
          const valid=candidate&&Number.isFinite(candidate.x)&&Number.isFinite(candidate.y);
          const size=Math.min(20,slide.widthCssPx,slide.heightCssPx);
          const x=valid?candidate.x:16,y=valid?candidate.y:16+(offset/80)*22;
          const box={x:Math.max(0,Math.min(x,slide.widthCssPx-size)),y:Math.max(0,Math.min(y,slide.heightCssPx-size)),w:size,h:size};offset+=80;
          const text='!',description='[변환하지 못한 '+(items[0].objectType==='background'?'배경/효과':'객체')+'] '+(items[0].objectName?items[0].objectName+': ':'')+items.map(d=>d.message).filter((x,i,a)=>a.indexOf(x)===i).join(' ');
          let id=`partial-placeholder-${offset}-${slide.nodes.length}`;while(slide.nodes.some(n=>n.id===id))id+='-';
          const style={fontFamily:selectedFont??'Malgun Gothic',fontSize:14,fontWeight:700,fontStyle:'normal',color:'#b91c1c',fill:'#ffffff',stroke:'#b91c1c',strokeWidth:1,opacity:1,lineHeight:'18px',whiteSpace:'pre-wrap',padding:{left:0,right:0,top:0,bottom:0},align:'center',valign:'mid',letterSpacing:0,textDecoration:'none'};
          slide.nodes.push({id,kind:'text',boxCssPx:box,zOrder:10000+offset,content:{text,placeholderDescription:description,runs:[{text,style:{...style}}]},style,parentId:null,sourceRef:{slideId:slide.id,elementId:id}});
          for(const d of items){d.code='PPTX_OBJECT_PLACEHOLDER';d.severity='warning';}
        }
      }
    }
    for(const e of validateModel(model).errors)diagnostics.push({...diagnostic('INPUT_INVALID',`${e.path}: ${e.message}`),reasonKey:'PPTX_MODEL_INVALID',blocking:true});
  }catch(e){
    const code=typeof e.code==='string'&&(/^(FONT_|CHART_)/.test(e.code)||e.code==='LIMIT_EXCEEDED')?e.code:'INPUT_INVALID';
    diagnostics.push({...diagnostic(code,e.message,currentContext),...currentContext,reasonKey:fatalReasons.get(e.message)??(code==='INPUT_INVALID'?'PPTX_PARSE_FATAL':undefined),blocking:true,error:safeError(e)});
  }
  // Preserve the locations collected before a later fatal package error as
  // well as those from a successful parse. The model itself remains private.
  for(const d of diagnostics){if(d.slideId&&!d.slideNumber){const index=model.slides.findIndex(s=>s.id===d.slideId);if(index>=0)d.slideNumber=index+1;}if(d.nodeId)Object.assign(d,objectContexts.get(`${d.slideId}:${d.nodeId}`)??{});}
  return done();
}
