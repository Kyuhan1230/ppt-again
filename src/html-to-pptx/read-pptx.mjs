import path from 'node:path';
import {strFromU8} from 'fflate';
import {DOMParser} from '@xmldom/xmldom';
import {diagnostic,validateModel} from './model.mjs';
import {resolveAsset,sha256} from './assets.mjs';
import {EMU_PER_CSS_PX} from './canvas.mjs';
import {unwrapPptxFont} from './font-embedding.mjs';
import {decideFontEmbedding} from './font-policy.mjs';
import {readNativeTable} from './read-table.mjs';
import {createPackageBudget,unzipBudgeted} from './package-budget.mjs';
import {readChartPackage} from './read-chart.mjs';
import {readMathTableFallback,stretchImageFragment,readMathTransform} from './math-table-fallback.mjs';
import {applyPngDuotone,duotoneColors,validatePngStructure} from './png-duotone.mjs';
import {validPresetVector,presetVectorMarkup} from './preset-vector.mjs';
import {snippedTextGeometry,snippedTextSvg} from './snip-text.mjs';
import {simpleOuterShadow} from './shadow.mjs';
import {nearlyEllipticalGeometry} from './ellipse-geometry.mjs';

const elements=n=>[...(n?.childNodes??[])].filter(c=>c.nodeType===1);
const child=(n,name)=>elements(n).find(c=>c.localName===name);
const all=(n,name)=>n?[...n.getElementsByTagNameNS('*',name)]:[];
const attr=(n,k,fallback=null)=>n?.hasAttribute(k)?n.getAttribute(k):fallback;
const num=(n,k,fallback=0)=>Number(attr(n,k,fallback));
const px=n=>n/EMU_PER_CSS_PX;
const mimeByExt={'.png':'image/png','.jpg':'image/jpeg','.jpeg':'image/jpeg','.svg':'image/svg+xml'};
const P='http://schemas.openxmlformats.org/presentationml/2006/main',R='http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const REL='http://schemas.openxmlformats.org/package/2006/relationships',CT='http://schemas.openxmlformats.org/package/2006/content-types';
const C='http://schemas.openxmlformats.org/drawingml/2006/chart',A='http://schemas.openxmlformats.org/drawingml/2006/main';
const MC='http://schemas.openxmlformats.org/markup-compatibility/2006',MATH='http://schemas.openxmlformats.org/officeDocument/2006/math';

// Only bounded, solid-color primitive geometry is converted to inline SVG.
function rotatedPresetArtwork(object,pr,xfrm,box,map,resolveColor,colorScheme,colorMap,themeDoc,uniqueKey){
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

// Resolve only solid, unmodified theme backgrounds. More complex theme fills
// stay unsupported rather than being replaced with a guessed color.
function themeBackground(ref,theme,master,overrides){
  if(theme.documentElement?.namespaceURI!==A||theme.documentElement.localName!=='theme')return null;
  const fillList=all(theme,'bgFillStyleLst')[0],index=num(ref,'idx')-1001;
  if(!Number.isInteger(index)||index<0||!fillList)return null;
  const fill=elements(fillList)[index];
  if(fill?.namespaceURI!==A||fill.localName!=='solidFill'||elements(fill).length!==1)return null;
  let mapping=all(master,'clrMap')[0];
  for(const doc of overrides){
    const override=all(doc,'clrMapOvr')[0];if(!override)continue;
    const chosen=elements(override)[0];
    if(chosen?.localName==='overrideClrMapping')mapping=chosen;
    else if(chosen?.localName!=='masterClrMapping')return null;
    break;
  }
  const scheme=all(theme,'clrScheme')[0];
  const resolve=(node,placeholder=null)=>{
    if(!node||node.namespaceURI!==A||elements(node).length)return null;
    if(node.localName==='srgbClr'||node.localName==='sysClr'){
      const hex=attr(node,node.localName==='sysClr'?'lastClr':'val','');
      return /^[a-f\d]{6}$/i.test(hex)?`#${hex}`:null;
    }
    if(node.localName!=='schemeClr')return null;
    let key=attr(node,'val');if(key==='phClr')return placeholder;
    key=attr(mapping,key,key);const slot=child(scheme,key);
    return slot&&elements(slot).length===1?resolve(elements(slot)[0]):null;
  };
  const colors=elements(ref);if(colors.length!==1)return null;
  return resolve(elements(fill)[0],resolve(colors[0]));
}

// Strict subset reader. No Office process, document code, network or disk extraction.
export async function readPptx(bytes,{fontPolicy='original',allowPartial=false}={}){
  const selectedFont=({'malgun-gothic':'Malgun Gothic',pretendard:'Pretendard'})[fontPolicy];
  const objectContexts=new Map();
  let currentContext={};
  const diagnostics=[],model={version:1,sourceDigest:sha256(bytes??''),slides:[],assets:[]};
  const fail=(code,message,slideId=null,nodeId=null)=>diagnostics.push(diagnostic(code,message,{slideId,nodeId}));
  const done=()=>{const failed=diagnostics.some(d=>d.severity==='error');return {status:failed?'failed':'parsed',model:failed?null:model,diagnostics};};
  const fatalReasons=new Map([
    ['객체 좌표·크기·회전값이 잘못되었습니다.','PPTX_INVALID_COORDINATES'],
    ['그룹 좌표·배율이 잘못되었습니다.','PPTX_INVALID_COORDINATES'],
    ['그림 자르기 좌표가 잘못되었습니다.','PPTX_INVALID_CROP'],
    ['현재 차트에 연결되지 않은 내장 문서는 지원하지 않습니다.','PPTX_UNUSED_EMBEDDED_PACKAGE'],
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
    // Every connected XLSX consumes the same package-wide expansion budget.
    const budget=createPackageBudget(),parts=unzipBudgeted(bytes,{budget});
    const validatedMathPngs=new Set();
    const validateMathPng=target=>{if(!validatedMathPngs.has(target)){validatePngStructure(parts[target]);validatedMathPngs.add(target);}};
    const xml=file=>{
      if(!parts[file])throw new Error('PPTX 필수 부분이 없습니다.');
      const s=strFromU8(parts[file]);if(/<!DOCTYPE|<!ENTITY/i.test(s))throw new Error('DTD·엔터티는 지원하지 않습니다.');
      try{return new DOMParser({onError:()=>{throw new Error('잘못된 PPTX XML입니다.');}}).parseFromString(s,'application/xml');}catch{throw new Error('잘못된 PPTX XML입니다.');}
    };
    const relationCache=new Map();
    const relationRecords=file=>{
      if(relationCache.has(file))return relationCache.get(file);
      const rfile=path.posix.join(path.posix.dirname(file),'_rels',path.posix.basename(file)+'.rels');
      const map=new Map();relationCache.set(file,map);if(!parts[rfile])return map;
      const doc=xml(rfile);if(doc.documentElement.namespaceURI!==REL||doc.documentElement.localName!=='Relationships')throw new Error('PPTX 관계 목록이 잘못되었습니다.');
      for(const r of elements(doc.documentElement)){
        const id=attr(r,'Id');if(r.namespaceURI!==REL||!id||map.has(id))throw new Error('PPTX 관계 ID 또는 네임스페이스가 잘못되었습니다.');
        if(r.localName!=='Relationship'||!['Internal',null].includes(attr(r,'TargetMode')))throw new Error('외부 참조가 있는 PPTX는 지원하지 않습니다.');
        const target=attr(r,'Target','');
        if(!target||/^[a-z][a-z\d+.-]*:|\\|[\u0000-\u001f?#]/i.test(target))throw new Error('외부 PPTX 경로입니다.');
        const resolved=target.startsWith('/')?target.slice(1):path.posix.normalize(path.posix.join(path.posix.dirname(file),target));
        if(resolved.startsWith('../'))throw new Error('PPTX 경로가 범위를 벗어났습니다.');
        map.set(id,{target:resolved,type:attr(r,'Type')});
      }
      return map;
    };
    const relationships=file=>new Map([...relationRecords(file)].map(([id,r])=>[id,r.target]));
    const types=xml('[Content_Types].xml');if(types.documentElement.namespaceURI!==CT||types.documentElement.localName!=='Types')throw new Error('PPTX ContentType 목록이 잘못되었습니다.');
    const defaults=new Map(),overrides=new Map();
    for(const node of elements(types.documentElement)){
      if(node.namespaceURI!==CT||!['Default','Override'].includes(node.localName))throw new Error('PPTX ContentType 선언이 잘못되었습니다.');
      const isDefault=node.localName==='Default',key=isDefault?attr(node,'Extension','').toLowerCase():attr(node,'PartName',''),map=isDefault?defaults:overrides;
      if(!key||map.has(key))throw new Error('PPTX ContentType 선언이 중복되거나 없습니다.');map.set(key,attr(node,'ContentType'));
    }
    const contentTypeFor=part=>overrides.get('/'+part)??defaults.get(path.posix.extname(part).slice(1).toLowerCase());
    const workbooks=new Map(),chartPackages=new Map(),usedWorkbooks=new Set(),usedCharts=new Set();let chartPoints=0;
    // Inspect every relationship, including unused external/OLE/script packages.
    for(const name of Object.keys(parts)){
      if(!name.endsWith('/')&&/vbaProject|activeX/i.test(name)||/macroEnabled|vbaProject|activeX|oleObject/i.test(contentTypeFor(name)??''))throw new Error('매크로·ActiveX·OLE 문서는 지원하지 않습니다.');
      if(name.endsWith('.rels')){
        if(name!=='_rels/.rels'&&path.posix.basename(path.posix.dirname(name))!=='_rels')throw new Error('PPTX 관계 부분의 경로가 잘못되었습니다.');
        const owner=name==='_rels/.rels'?'':path.posix.join(path.posix.dirname(path.posix.dirname(name)),path.posix.basename(name,'.rels'));
        for(const r of relationRecords(owner).values())if(/\/(?:oleObject|vbaProject|activeX|control)$/i.test(r.type??''))throw new Error('실행 가능한 내장 객체는 지원하지 않습니다.');
      }
    }
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
      const file=presRels.get(attr(ref,'r:id')),doc=xml(file),rels=relationships(file),tree=all(doc,'spTree')[0];
      const slide={id:`slide-${index+1}`,widthCssPx:px(model.pptxSizeEmu.width),heightCssPx:px(model.pptxSizeEmu.height),background:{color:'#ffffff',image:'none'},nodes:[]};model.slides.push(slide);
      let activeNodeId=null;
      const issue=(message,nodeId=null,explicitReason=null)=>{
        nodeId??=activeNodeId;
        const reasonKey=/명시적 글꼴·크기가 없는 상속/.test(message)?'PPTX_INHERITED_TEXT_UNRESOLVED':/테마·시스템 색/.test(message)?'PPTX_THEME_COLOR_UNSUPPORTED':/상속된 도형 채움·선·효과/.test(message)?'PPTX_INHERITED_SHAPE_STYLE_UNSUPPORTED':/테마 배경|단색 외의 PPT 배경/.test(message)?'PPTX_THEME_BACKGROUND_UNSUPPORTED':/마스터·레이아웃/.test(message)?'PPTX_MASTER_ARTWORK_UNSUPPORTED':/타이밍·전환/.test(message)?'PPTX_ANIMATION_UNSUPPORTED':/회전·뒤집힌/.test(message)?'PPTX_OBJECT_ROTATION_UNSUPPORTED':/객체 위치/.test(message)?'PPTX_OBJECT_POSITION_UNSUPPORTED':/그룹/.test(message)?'PPTX_GROUP_TRANSFORM_UNSUPPORTED':/그림/.test(message)?'PPTX_IMAGE_UNSUPPORTED':/도형·채움·효과|도형 채움/.test(message)?'PPTX_SHAPE_EFFECT_UNSUPPORTED':/글꼴|텍스트|글자|문단|글머리/.test(message)?'PPTX_TEXT_STYLE_UNSUPPORTED':'PPTX_OBJECT_UNSUPPORTED';
        diagnostics.push({...diagnostic('UNSUPPORTED_OBJECT',message,{slideId:slide.id,nodeId}),slideNumber:index+1,...(objectContexts.get(`${slide.id}:${nodeId}`)??{objectType:'background'}),reasonKey:explicitReason??reasonKey,recoverable:true});
      };
      const inherited=new Set([...rels.values()].filter(p=>p.startsWith('ppt/slideLayouts/')));
      for(const layout of [...inherited])for(const target of relationships(layout).values())if(target.startsWith('ppt/slideMasters/'))inherited.add(target);
      const masterPart=[...inherited].find(part=>part.startsWith('ppt/slideMasters/'));
      const themePart=masterPart&&[...relationRecords(masterPart).values()].find(r=>r.type===`${R}/theme`)?.target;
      const themeDoc=themePart&&parts[themePart]?xml(themePart):null;
      const masterDoc=masterPart?xml(masterPart):null;
      let colorMap=all(masterDoc,'clrMap')[0];
      for(const part of [doc,...[...inherited].filter(p=>p.startsWith('ppt/slideLayouts/')).map(p=>xml(p))]){
        const override=elements(all(part,'clrMapOvr')[0])[0];
        if(override?.localName==='overrideClrMapping'){colorMap=override;break;}
        if(override?.localName==='masterClrMapping')break;
      }
      const colorScheme=all(themeDoc,'clrScheme')[0];
      if(all(doc,'timing').length||all(doc,'transition').length){
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
      const readText=(tx,id)=>{
      const body=child(tx,'bodyPr'),paras=elements(tx).filter(n=>n.localName==='p'),runs=[];
      const approximate=(reasonKey,message)=>{const key=`${id}:${reasonKey}`;if(!textApproximations.has(key)){textApproximations.add(key);diagnostics.push({...diagnostic('PARAGRAPH_SPACING_APPROXIMATED',message,{slideId:slide.id,nodeId:id},'warning'),reasonKey});}};
      // Text-box defaults are inherited per paragraph level; placeholders can
      // also inherit the corresponding layout/master placeholder's text style.
      // Ambiguous placeholder matches are not guessed.
      const owner=tx?.parentNode?.localName==='sp'?tx.parentNode:null;
      const placeholder=sp=>child(child(child(sp,'nvSpPr'),'nvPr'),'ph');
      const ph=placeholder(owner),layouts=ph?[...inherited].filter(p=>p.startsWith('ppt/slideLayouts/')).map(p=>xml(p)):[];
      const findPlaceholder=(container,byIndex,type,index)=>{
        const found=all(container,'sp').filter(sp=>{const p=placeholder(sp);return p&&(byIndex?attr(p,'idx','0')===index:attr(p,'type','obj')===type);});
        return found.length===1?found[0]:null;
      };
      const layoutShape=ph?layouts.map(d=>findPlaceholder(d,true,null,attr(ph,'idx','0'))).filter(Boolean):[];
      const matchedLayout=layoutShape.length===1?layoutShape[0]:null;
      const phType=attr(placeholder(matchedLayout),'type',attr(ph,'type','obj'));
      const masterShape=ph?(findPlaceholder(masterDoc,false,phType,null)??(phType==='ctrTitle'?findPlaceholder(masterDoc,false,'title',null):null)):null;
      const bodyProperty=(name,fallback)=>attr(body,name)??attr(child(child(matchedLayout,'txBody'),'bodyPr'),name)??attr(child(child(masterShape,'txBody'),'bodyPr'),name)??fallback;
      const paragraphAttribute=(pPr,name,fallback)=>{
        const level=Number(attr(pPr,'lvl','0'));if(!Number.isInteger(level)||level<0||level>8)return attr(pPr,name,fallback);
        const fromList=list=>attr(child(list,`lvl${level+1}pPr`),name)??attr(child(list,'defPPr'),name);
        const fromPlaceholder=sp=>{
          const text=child(sp,'txBody');if(!text)return null;
          return fromList(child(text,'lstStyle'))??attr(child(elements(text).find(n=>n.localName==='p'),'pPr'),name);
        };
        const styleName=ph?(['title','ctrTitle'].includes(phType)?'titleStyle':['body','subTitle','obj'].includes(phType)?'bodyStyle':'otherStyle'):'otherStyle';
        return attr(pPr,name)??fromList(child(tx,'lstStyle'))??fromPlaceholder(matchedLayout)??fromPlaceholder(masterShape)??fromList(child(child(masterDoc?.documentElement,'txStyles'),styleName))??fromList(child(presentation.documentElement,'defaultTextStyle'))??fallback;
      };
      const paragraphAlignment=pPr=>paragraphAttribute(pPr,'algn','l');
      const paragraphChoice=(pPr,names)=>{
        const select=node=>elements(node).find(n=>names.includes(n.localName));
        const direct=select(pPr);if(direct)return direct;
        const level=Number(attr(pPr,'lvl','0'));if(!Number.isInteger(level)||level<0||level>8)return null;
        const fromList=list=>select(child(list,`lvl${level+1}pPr`))??select(child(list,'defPPr'));
        const fromPlaceholder=sp=>{
          const text=child(sp,'txBody');if(!text)return null;
          return fromList(child(text,'lstStyle'))??select(child(elements(text).find(n=>n.localName==='p'),'pPr'));
        };
        const own=fromList(child(tx,'lstStyle'));if(own||!owner)return own;
        const styleName=ph?(['title','ctrTitle'].includes(phType)?'titleStyle':['body','subTitle','obj'].includes(phType)?'bodyStyle':'otherStyle'):'otherStyle';
        return fromPlaceholder(matchedLayout)??fromPlaceholder(masterShape)??fromList(child(child(masterDoc?.documentElement,'txStyles'),styleName))??fromList(child(presentation.documentElement,'defaultTextStyle'));
      };
      const paragraphSetting=(pPr,name)=>paragraphChoice(pPr,[name]);
      const paragraphOwnSetting=(pPr,name)=>{
        const direct=child(pPr,name);if(direct)return direct;
        const level=Number(attr(pPr,'lvl','0'));if(!Number.isInteger(level)||level<0||level>8)return null;
        const list=child(tx,'lstStyle');
        return child(child(list,`lvl${level+1}pPr`),name)??child(child(list,'defPPr'),name);
      };
      const inheritedSize=(pPr)=>{
        if(!owner)return null;
        const level=Number(attr(pPr,'lvl','0'));if(!Number.isInteger(level)||level<0||level>8)return null;
        const fromList=list=>attr(child(child(list,`lvl${level+1}pPr`),'defRPr'),'sz',attr(child(child(list,'defPPr'),'defRPr'),'sz'));
        const placeholderSize=sp=>{
          const text=child(sp,'txBody');if(!text)return null;
          const listed=fromList(child(text,'lstStyle'));if(listed!==null)return listed;
          const values=[...all(text,'defRPr'),...all(text,'rPr'),...all(text,'endParaRPr')].map(n=>attr(n,'sz')).filter(v=>v!==null);
          return new Set(values).size===1?values[0]:null;
        };
        const styleName=ph?(['title','ctrTitle'].includes(phType)?'titleStyle':['body','subTitle','obj'].includes(phType)?'bodyStyle':'otherStyle'):'otherStyle';
        const styles=child(child(masterDoc?.documentElement,'txStyles'),styleName);
        return fromList(child(tx,'lstStyle'))??placeholderSize(matchedLayout)??placeholderSize(masterShape)??fromList(styles)??fromList(child(presentation.documentElement,'defaultTextStyle'));
      };
      const inheritedFill=(pPr)=>{
        if(!owner)return null;
        const level=Number(attr(pPr,'lvl','0'));if(!Number.isInteger(level)||level<0||level>8)return null;
        const fromList=list=>child(child(child(list,`lvl${level+1}pPr`),'defRPr'),'solidFill')??child(child(child(list,'defPPr'),'defRPr'),'solidFill');
        const fromPlaceholder=sp=>{
          const text=child(sp,'txBody');if(!text)return null;
          const listed=fromList(child(text,'lstStyle'));if(listed)return listed;
          const fills=[...all(text,'defRPr'),...all(text,'rPr'),...all(text,'endParaRPr')].map(n=>child(n,'solidFill')).filter(Boolean);
          return new Set(fills.map(n=>n.toString())).size===1?fills[0]:null;
        };
        const styleName=ph?(['title','ctrTitle'].includes(phType)?'titleStyle':['body','subTitle','obj'].includes(phType)?'bodyStyle':'otherStyle'):'otherStyle';
        const fontReference=child(child(owner,'style'),'fontRef');
        return fromList(child(tx,'lstStyle'))??(!ph&&elements(fontReference).length?fontReference:null)??fromPlaceholder(matchedLayout)??fromPlaceholder(masterShape)??fromList(child(child(masterDoc?.documentElement,'txStyles'),styleName))??fromList(child(presentation.documentElement,'defaultTextStyle'));
      };
      const inheritedTypeface=(pPr,kind)=>{
        if(!owner)return null;
        const level=Number(attr(pPr,'lvl','0'));if(!Number.isInteger(level)||level<0||level>8)return null;
        const fromList=list=>attr(child(child(child(list,`lvl${level+1}pPr`),'defRPr'),kind),'typeface')??attr(child(child(child(list,'defPPr'),'defRPr'),kind),'typeface');
        const fromPlaceholder=sp=>{
          const text=child(sp,'txBody');if(!text)return null;
          const listed=fromList(child(text,'lstStyle'));if(listed)return listed;
          const values=[...all(text,'defRPr'),...all(text,'rPr'),...all(text,'endParaRPr')].map(n=>attr(child(n,kind),'typeface')).filter(Boolean);
          return new Set(values).size===1?values[0]:null;
        };
        const styleName=ph?(['title','ctrTitle'].includes(phType)?'titleStyle':['body','subTitle','obj'].includes(phType)?'bodyStyle':'otherStyle'):'otherStyle';
        const referenced=attr(child(child(owner,'style'),'fontRef'),'idx');
        const suffix=kind==='ea'?'ea':'lt';
        return fromList(child(tx,'lstStyle'))??fromPlaceholder(matchedLayout)??fromPlaceholder(masterShape)??fromList(child(child(masterDoc?.documentElement,'txStyles'),styleName))??fromList(child(presentation.documentElement,'defaultTextStyle'))??(referenced==='major'?`+mj-${suffix}`:referenced==='minor'?`+mn-${suffix}`:null);
      };
      const resolveTypeface=(value,script)=>{
        if(!value)return null;
        const alias=/^\+(mj|mn)-(lt|ea|cs)$/.exec(value);
        if(!alias)return value.startsWith('+')?null:value;
        const group=child(all(themeDoc,'fontScheme')[0],alias[1]==='mj'?'majorFont':'minorFont');
        const kind=({lt:'latin',ea:'ea',cs:'cs'})[alias[2]];
        const direct=attr(child(group,kind),'typeface');if(direct)return direct;
        if(kind==='ea'){
          const languageScript=script==='hangul'?'Hang':script==='japanese'?'Jpan':script==='han'?'Hans':null;
          const match=elements(group).find(n=>n.localName==='font'&&attr(n,'script')===languageScript);
          const specified=attr(match,'typeface');if(specified)return specified;
        }
        return null;
      };
      const scriptOf=character=>/\p{Script=Hangul}/u.test(character)?'hangul':/[\p{Script=Hiragana}\p{Script=Katakana}]/u.test(character)?'japanese':/\p{Script=Han}/u.test(character)?'han':'latin';
      const splitScripts=text=>{
        const pieces=[];for(const char of text){const script=scriptOf(char),last=pieces.at(-1);if(last&&last.script===script)last.text+=char;else pieces.push({text:char,script});}return pieces;
      };
      const spacingValue=node=>{
        const raw=attr(node,'val','').trim();
        if(node?.localName==='spcPct'&&/^\+?(?:\d+(?:\.\d+)?|\.\d+)%$/.test(raw))return Number(raw.slice(0,-1))*1000;
        return /^\+?\d+$/.test(raw)&&Number.isSafeInteger(Number(raw))?Number(raw):NaN;
      };
      const checkedSpacing=spacing=>{
        const parts=elements(spacing),item=parts[0];
        if(parts.length!==1||item.namespaceURI!==A||!['spcPct','spcPts'].includes(item.localName)||elements(item).length)throw new Error('문단의 간격 구조가 올바르지 않습니다.');
        const value=spacingValue(item),max=item.localName==='spcPct'?13200000:158400;
        if(!Number.isFinite(value)||value<0||value>max)throw new Error('문단의 간격 값이 올바르지 않습니다.');
        return {kind:item.localName,value};
      };
      const spacingPx=(pPr,name)=>{
        const spacing=paragraphOwnSetting(pPr,name);if(!spacing)return 0;
        const {kind,value}=checkedSpacing(spacing);
        if(kind==='spcPct'&&value===0)return 0;
        return kind==='spcPts'?value/100*4/3:null;
      };
      const paragraphLineHeight=pPr=>{
        const spacing=paragraphSetting(pPr,'lnSpc');if(!spacing)return '1.2';
        const {kind,value}=checkedSpacing(spacing);if(value<=0)return null;
        return kind==='spcPts'?`${value/100*4/3}px`:`${Number((value*1.2/100000).toFixed(6))}`;
      };
      // Keep individual paragraph layout only where its spacing can be read
      // without guessing. Empty paragraphs still occupy a line in PowerPoint.
      const paragraphLayout=paras.map(p=>{
        const pPr=child(p,'pPr'),tableCell=tx?.parentNode?.localName==='tc';
        for(const key of ['marL','marR','indent'])if(pPr?.hasAttribute(key)){
          const raw=pPr.getAttribute(key).trim(),value=Number(raw);
          if(!/^[+-]?\d+$/.test(raw)||!Number.isSafeInteger(value)||key!=='indent'&&value<0)throw new Error('문단의 들여쓰기 값이 올바르지 않습니다.');
        }
        for(const name of ['lnSpc','spcBef','spcAft']){const spacing=child(pPr,name);if(spacing)checkedSpacing(spacing);}
        // Office defaults hanging punctuation to true when it is unspecified.
        const hanging=tableCell?paragraphAttribute(pPr,'hangingPunct','1'):null;
        if(tableCell&&!['0','1','false','true'].includes(hanging))throw new Error('문단의 문장부호 배치 값이 올바르지 않습니다.');
        return {beforePx:spacingPx(pPr,'spcBef'),afterPx:spacingPx(pPr,'spcAft'),leftPx:px(num(pPr,'marL')),rightPx:px(num(pPr,'marR')),indentPx:px(num(pPr,'indent')),lineHeight:paragraphLineHeight(pPr),align:({l:'left',ctr:'center',r:'right',just:'justify'})[paragraphAlignment(pPr)],...(tableCell?{hangingPunct:['1','true'].includes(hanging)}:{})};
      });
      const useParagraphLayout=Boolean((owner||tx?.parentNode?.localName==='tc')&&paras.length>0&&paras.every((p,i)=>{
        const layout=paragraphLayout[i];return layout.beforePx!==null&&layout.afterPx!==null&&layout.lineHeight&&layout.align&&[layout.leftPx,layout.rightPx].every(n=>Number.isFinite(n)&&n>=0&&n<=5000)&&Number.isFinite(layout.indentPx)&&Math.abs(layout.indentPx)<=5000;
      })&&(paras.some(p=>{const pPr=child(p,'pPr');return ['marL','marR','indent'].some(k=>num(pPr,k));})||
        paras.length>1&&(paras.some(p=>['lnSpc','spcBef','spcAft'].some(k=>paragraphOwnSetting(child(p,'pPr'),k)))||paragraphLayout.some(p=>p.align!==paragraphLayout[0].align))||
        ['1','true'].includes(attr(body,'spcFirstLastPara','0'))&&paras.some(p=>['spcBef','spcAft'].some(k=>paragraphOwnSetting(child(p,'pPr'),k)))));
      const paragraphRuns=[];
      let inheritedTextColorEntry=null,inheritedTextColorAmbiguous=false;
      const cell=tx?.parentNode,row=cell?.parentNode,table=row?.parentNode,tableProps=child(table,'tblPr');
      if(cell?.localName==='tc'&&row?.localName==='tr'&&table?.localName==='tbl'&&tableStylesDoc){
        const styleId=child(tableProps,'tableStyleId')?.textContent,style=elements(tableStylesDoc.documentElement).find(n=>n.localName==='tblStyle'&&attr(n,'styleId')===styleId);
        const isFirstRow=elements(table).filter(n=>n.localName==='tr')[0]===row&&['1','true'].includes(attr(tableProps,'firstRow','0'));
        const rows=elements(table).filter(n=>n.localName==='tr'),cells=elements(row).filter(n=>n.localName==='tc'),col=cells.indexOf(cell);
        const enabled=key=>['1','true'].includes(attr(tableProps,key,'0'));
        const mixedCell=col<0||enabled('firstCol')&&col===0||enabled('lastCol')&&col===cells.length-1||enabled('lastRow')&&rows.length===1;
        if(style&&isFirstRow){
          const textStyle=child(child(style,'firstRow'),'tcTxStyle')??child(child(style,'wholeTbl'),'tcTxStyle');
          const entry=elements(textStyle).find(n=>['schemeClr','srgbClr','sysClr'].includes(n.localName));
          if(entry){inheritedTextColorEntry=entry;inheritedTextColorAmbiguous=mixedCell;}
        }
      }
      if(num(body,'rot')!==0||attr(body,'vert','horz')!=='horz'||num(body,'numCol',1)!==1||num(body,'rtlCol')!==0||num(body,'fromWordArt')!==0||child(body,'prstTxWarp')&&attr(child(body,'prstTxWarp'),'prst')!=='textNoShape')issue('회전·세로쓰기·다단·워드아트 텍스트는 아직 지원하지 않습니다.',id);
      const runStyle=(rPr,pPr,script='latin')=>{
        const def=child(pPr,'defRPr'),get=k=>attr(rPr,k,attr(def,k)),kind=script==='latin'?'latin':'ea';
        const baseline=Number(get('baseline')??0);
        if(!Number.isInteger(baseline)||Math.abs(baseline)>100000)issue('지원 범위를 벗어난 위첨자·아래첨자 위치입니다.',id);
        const highlighted=child(rPr,'highlight')??child(def,'highlight');
        let highlight=highlighted?color(highlighted):null;
        if(highlight&&!/^#[a-f\d]{6}$/i.test(highlight)){
          highlight=null;
          if(allowPartial)approximate('PPTX_TEXT_HIGHLIGHT_APPROXIMATED','반투명 형광펜 배경 없이 글자를 표시했습니다.');else issue('반투명 글자 형광펜 표시는 아직 지원하지 않습니다.',id);
        }
        if(![null,'noStrike'].includes(get('strike'))||![null,'none'].includes(get('cap'))||![null,'none','sng'].includes(get('u'))||elements(rPr).some(n=>n.localName==='effectLst'?elements(n).length>0:['effectDag','gradFill','blipFill','uFill','uLn','hlinkClick','hlinkMouseOver'].includes(n.localName)))issue('취소선·대소문자 변환·첨자·효과 등 이 글자 서식은 아직 지원하지 않습니다.',id);
        const outline=child(rPr,'ln')??child(def,'ln');let textOutline;
        const outlinePaint=child(outline,'solidFill'),paintColors=elements(outlinePaint),paintEffects=elements(paintColors[0]);
        const invisibleOutline=outline&&(outline.hasAttribute('w')&&num(outline,'w')===0||paintColors.length===1&&paintEffects.length===1&&paintEffects[0].localName==='alpha'&&num(paintEffects[0],'val',NaN)===0);
        if(outline&&!child(outline,'noFill')&&!invisibleOutline){
          if(!child(outline,'solidFill')||elements(outline).some(n=>n.localName!=='solidFill')||all(outline,'alpha').some(n=>num(n,'val')!==100000))issue('불투명 단색 외의 글자 외곽선은 지원하지 않습니다.',id);
          textOutline={color:color(child(outline,'solidFill')),width:px(num(outline,'w'))};
        }
        const declaredFont=attr(child(rPr,kind),'typeface')??attr(child(def,kind),'typeface')??inheritedTypeface(pPr,kind);
        const font=resolveTypeface(declaredFont,script)??(kind==='ea'?resolveTypeface(attr(child(rPr,'latin'),'typeface')??attr(child(def,'latin'),'typeface')??inheritedTypeface(pPr,'latin'),'latin'):null);
        const declaredSize=get('sz'),inherited=declaredSize===null?inheritedSize(pPr):null;
        let sizePt=Number(declaredSize??inherited)/100;
        if(declaredSize===null&&inherited===null&&allowPartial&&(selectedFont||font&&!font.startsWith('+'))){sizePt=18;approximate('PPTX_INHERITED_SIZE_APPROXIMATED','상속된 글자 크기를 확인하지 못해 18pt로 표시했습니다.');}
        if((!selectedFont&&(!font||font.startsWith('+')))||!Number.isFinite(sizePt)||sizePt<=0)issue('명시적 글꼴·크기가 없는 상속 텍스트는 아직 지원하지 않습니다.',id);
        const kerningValue=get('kern'),kerning=kerningValue===null?null:Number(kerningValue.trim());
        const validKerning=kerningValue===null||/^\+?\d+$/.test(kerningValue.trim())&&Number.isInteger(kerning)&&kerning<=400000;
        if(!validKerning)issue('글자 커닝 기준 크기가 올바르지 않습니다.',id);
        const disableKerning=kerning!==null&&validKerning&&(kerning===0||sizePt*100<kerning);
        const textFill=child(rPr,'solidFill')??child(def,'solidFill')??inheritedFill(pPr);
        if(!textFill&&inheritedTextColorEntry&&inheritedTextColorAmbiguous){
          if(allowPartial)approximate('PPTX_TABLE_TEXT_COLOR_APPROXIMATED','표 교차 셀의 상속 글자색은 원본과 다를 수 있습니다.');
          else issue('표 교차 셀의 상속 글자색을 확정할 수 없습니다.',id);
        }
        const fontReference=child(child(owner,'style'),'fontRef');
        const textColor=textFill?color(textFill,'#000000'):inheritedTextColorEntry?color({childNodes:[inheritedTextColorEntry]},'#000000'):fontReference&&elements(fontReference).length?color(fontReference,'#000000'):'#000000';
        return {fontFamily:selectedFont??font??'',fontSize:sizePt*4/3,fontWeight:get('b')==='1'?700:400,fontStyle:get('i')==='1'?'italic':'normal',color:textColor,letterSpacing:Number(get('spc')??0)/100*4/3,textDecoration:get('u')&&get('u')!=='none'?'underline':'none',...(disableKerning?{disableKerning:true}:{}),...(baseline&&Number.isInteger(baseline)&&Math.abs(baseline)<=100000?{baseline}:{}),...(textOutline?{outline:textOutline}:{}),...(highlight?{highlight}:{})};
      };
      const bulletStyle=(p,pPr,characterBullet=false)=>{
        const firstRun=elements(p).find(n=>['r','fld'].includes(n.localName));
        const style=runStyle(child(firstRun,'rPr')??child(p,'endParaRPr'),pPr);
        delete style.highlight;delete style.outline;delete style.baseline;style.textDecoration='none';
        // PowerPoint renders character bullets with their own font face's
        // default weight, independently of the first text run's bold setting.
        if(characterBullet)style.fontWeight=400;
        const font=paragraphChoice(pPr,['buFont','buFontTx']);
        if(!selectedFont&&font?.localName==='buFont'){
          const family=resolveTypeface(attr(font,'typeface'),'latin');
          if(family)style.fontFamily=family;else approximate('PPTX_PARAGRAPH_INDENT_APPROXIMATED','글머리표 글꼴을 확인하지 못해 본문 글꼴로 표시했습니다.');
        }
        const size=paragraphChoice(pPr,['buSzTx','buSzPct','buSzPts']);
        if(size&&size.localName!=='buSzTx'){
          const value=num(size,'val',NaN),scaled=size.localName==='buSzPct'?style.fontSize*value/100000:value/100*4/3;
          if(Number.isFinite(value)&&value>0&&Number.isFinite(scaled)&&scaled>0&&scaled<=5000)style.fontSize=scaled;
          else approximate('PPTX_PARAGRAPH_INDENT_APPROXIMATED','글머리표 크기를 확인하지 못해 본문 크기로 표시했습니다.');
        }
        const bulletColor=paragraphChoice(pPr,['buClr','buClrTx']);if(bulletColor?.localName==='buClr')style.color=color(bulletColor,style.color);
        return style;
      };
      let previousNumber=null;
      for(const [pi,p] of paras.entries()){
        const pPr=child(p,'pPr'),bullet=child(pPr,'buChar'),bulletChar=attr(bullet,'char',''),auto=child(pPr,'buAutoNum');
        let numberLabel=null;
        if(auto&&allowPartial&&attr(auto,'type')==='arabicPeriod'){
          const start=attr(auto,'startAt'),level=Number(attr(pPr,'lvl','0')),parsed=start===null?null:Number(start);
          if(Number.isInteger(level)&&level>=0&&level<=8&&(parsed===null||Number.isInteger(parsed)&&parsed>=1&&parsed<=32767)){
            const count=parsed??(previousNumber?.level===level?previousNumber.count+1:1);
            if(count<=32767){numberLabel=`${count}.`;previousNumber={level,count};approximate('PPTX_AUTONUM_APPROXIMATED','자동 번호를 텍스트로 옮겼습니다. 들여쓰기와 번호 배치는 원본과 다를 수 있습니다.');}
          }
        }
        if(!auto)previousNumber=null;
        const hasInheritedBulletProperties=elements(pPr).some(n=>/^bu/.test(n.localName))&&!bullet&&!child(pPr,'buNone');
        const unsupportedBullet=auto&&numberLabel===null||child(pPr,'buBlip')||hasInheritedBulletProperties&&!auto||(bullet&&(!bulletChar||/[\u0000-\u001f]/.test(bulletChar)||[...bulletChar].length!==1));
        if(unsupportedBullet)issue('이 글머리표는 아직 지원하지 않습니다.',id);
        const usesTab=all(p,'t').some(t=>t.textContent.includes('\t'))||elements(p).some(n=>n.localName==='tab');
        if(usesTab){if(allowPartial)approximate('PPTX_TEXT_TAB_APPROXIMATED','기본 탭 간격으로 표시했으며 개별 탭 위치는 원본과 다를 수 있습니다.');else issue('문단 탭은 아직 지원하지 않습니다.',id);}
        if((bullet&&!unsupportedBullet||num(pPr,'marL')||num(pPr,'marR')||num(pPr,'indent'))&&!approximatedParagraphs.has(id)){
          diagnostics.push({...diagnostic('PARAGRAPH_SPACING_APPROXIMATED','문단 들여쓰기와 글머리 배치는 근사했습니다.',{slideId:slide.id,nodeId:id},'warning'),reasonKey:'PPTX_PARAGRAPH_INDENT_APPROXIMATED'});approximatedParagraphs.add(id);
        }
        if(!useParagraphLayout&&['spcBef','spcAft'].some(k=>{const spacing=paragraphOwnSetting(pPr,k);return spacing&&checkedSpacing(spacing).value!==0;}))diagnostics.push(diagnostic('PARAGRAPH_SPACING_APPROXIMATED','문단 앞뒤 간격은 그대로 옮기지 못했습니다.',{slideId:slide.id,nodeId:id},'warning'));
        if(!paragraphLineHeight(pPr)&&paragraphSetting(pPr,'lnSpc')){
          if(allowPartial)approximate('PPTX_LINE_SPACING_APPROXIMATED','줄 간격 0은 기본 간격으로 표시했습니다.');else issue('줄 간격 0은 아직 지원하지 않습니다.',id);
        }
        if(!useParagraphLayout&&pi&&paragraphAlignment(pPr)!==paragraphAlignment(child(paras[0],'pPr'))){
          if(allowPartial&&['l','ctr','r','just'].includes(paragraphAlignment(pPr)))approximate('PPTX_PARAGRAPH_ALIGNMENT_APPROXIMATED','서로 다른 문단 정렬은 첫 문단의 정렬로 표시했습니다.');
          else issue('서로 다른 문단 정렬은 아직 지원하지 않습니다.',id);
        }
        if(pi)runs.push({text:'\n',style:runs.at(-1)?.style??runStyle(child(p,'endParaRPr'),pPr)});
        const paragraphStart=runs.length;
        if(bullet&&!unsupportedBullet&&all(p,'t').some(t=>t.textContent.trim()))runs.push({text:`${bulletChar} `,style:bulletStyle(p,pPr,true),marker:true});
        if(numberLabel&&all(p,'t').some(t=>t.textContent.trim()))runs.push({text:`${numberLabel} `,style:bulletStyle(p,pPr),marker:true});
        for(const r of elements(p))if(['r','fld','br','tab'].includes(r.localName)){
          const text=r.localName==='br'?'\n':r.localName==='tab'?'\t':all(r,'t').map(t=>t.textContent).join('');
          if(r.localName==='tab'||r.localName==='br'){runs.push({text,style:runs.at(-1)?.style??runStyle(child(p,'endParaRPr'),pPr)});continue;}
          const start=runs.length;
          for(const part of selectedFont?[{text,script:'latin'}]:splitScripts(text)){
            const style=runStyle(child(r,'rPr'),pPr,part.script),last=runs.at(-1);
            if(runs.length>start&&JSON.stringify(last.style)===JSON.stringify(style))last.text+=part.text;
            else runs.push({text:part.text,style});
          }
        }
        if(useParagraphLayout){
          const current=runs.slice(paragraphStart);
          if(!current.length){
            const firstRun=elements(p).find(n=>['r','fld'].includes(n.localName));
            const properties=child(firstRun,'rPr')??child(p,'endParaRPr');
            current.push({text:'',style:properties?runStyle(properties,pPr):runs.at(-1)?.style??runStyle(null,pPr)});
          }
          paragraphRuns.push(current);
        }
      }
      const content={text:runs.map(r=>r.text).join(''),runs};
      if(useParagraphLayout){
        const edgeSpacing=['1','true'].includes(attr(body,'spcFirstLastPara','0'));
        content.paragraphs=paragraphLayout.map((layout,index)=>({...layout,beforePx:index===0&&!edgeSpacing?0:layout.beforePx,afterPx:index===paras.length-1&&!edgeSpacing?0:layout.afterPx,runs:paragraphRuns[index]}));
      }
      const pPr=child(paras[0],'pPr');
      // PowerPoint's percentage is relative to its single-line advance,
      // which is 1.2 times the point size in the supported text bodies.
      const lineHeight=paragraphLineHeight(pPr)??'normal';
      const textRun=runs.find(run=>!run.marker&&run.text.trim());
      const style=Object.assign({},textRun?.style??runs[0]?.style??runStyle(child(paras[0],'endParaRPr'),pPr),{padding:{left:px(Number(bodyProperty('lIns',91440))),right:px(Number(bodyProperty('rIns',91440))),top:px(Number(bodyProperty('tIns',45720))),bottom:px(Number(bodyProperty('bIns',45720)))},align:({l:'left',ctr:'center',r:'right',just:'justify'})[paragraphAlignment(pPr)],valign:({t:'top',ctr:'mid',b:'bottom'})[bodyProperty('anchor','t')],lineHeight,whiteSpace:bodyProperty('wrap','square')==='none'?'pre':'pre-wrap'});
      if(runs.some(run=>run.text.includes('\t'))){
        const raw=attr(pPr,'defTabSz'),tabSize=raw===null?96:px(Number(raw));
        if(Number.isFinite(tabSize)&&tabSize>0&&tabSize<=5000)style.tabSize=tabSize;
      }
      // Auto-sized Office text can extend slightly into the right inset before
      // breaking a line. Preserve the actual inset separately for re-export.
      if(owner&&child(body,'spAutoFit')&&style.whiteSpace==='pre-wrap')style.wrapAllowance=Math.min(style.padding.right/2,style.fontSize/4);
        return {content,style};
      };
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
          activeNodeId=id;
          currentContext={slideId:slide.id,slideNumber:index+1,nodeId:id};
          objectContexts.set(`${slide.id}:${id}`,{objectName:base,objectType:object.localName==='pic'?'image':object.localName==='grpSp'?'group':object.localName==='graphicFrame'?(all(object,'tbl').length?'table':'chart'):all(object,'t').length?'text':'shape'});
          Object.assign(currentContext,objectContexts.get(`${slide.id}:${id}`));
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
              objectContexts.set(`${slide.id}:${id}`,{objectName:base,objectType:'image'});
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
              objectContexts.set(`${slide.id}:${id}`,{objectName:base,objectType:'image'});
              slide.nodes.push({id,kind:'image',boxCssPx:box,zOrder:slide.nodes.length,content:{assetId:asset.id,...(crop?{crop}:{})},style:{opacity:1},sourceRef:{slideId:slide.id,elementId:id},parentId});
              diagnostics.push({...diagnostic('PPTX_MATH_IMAGE_FALLBACK','수식은 그림으로 보존되며 수식을 개별 편집할 수 없습니다. 그림 배치는 원본과 다를 수 있습니다.',{slideId:slide.id,nodeId:id},'warning'),slideNumber:index+1,objectName:base,objectType:'image',reasonKey:'PPTX_MATH_IMAGE_FALLBACK'});
              continue;
            }
          }
          if(!['sp','cxnSp','pic','grpSp','graphicFrame'].includes(object.localName)){issue('이 객체는 아직 지원하지 않습니다.',id);continue;}
          const group=object.localName==='grpSp',pr=child(object,group?'grpSpPr':'spPr'),xfrm=object.localName==='graphicFrame'?child(object,'xfrm'):child(pr,'xfrm');
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
              diagnostics.push({...diagnostic('PPTX_SHAPE_IMAGE_APPROXIMATED','HTML에서는 SVG 도형을 편집할 수 있지만, 다시 만든 PPT에서는 그림으로 저장되어 개별 편집할 수 없습니다.',{slideId:slide.id,nodeId:id},'warning'),slideNumber:index+1,...objectContexts.get(`${slide.id}:${id}`),reasonKey:'PPTX_SHAPE_IMAGE_APPROXIMATED'});
              continue;
            }
          }
          if(object.localName==='graphicFrame'){
            const graphic=child(object,'graphic'),data=child(graphic,'graphicData');
            if(attr(data,'uri')===C){
              const refs=elements(data),chart=refs[0],rid=chart?.getAttributeNS(R,'id'),relation=relationRecords(ownerFile).get(rid);
              if(graphic.namespaceURI!==A||data.namespaceURI!==A||refs.length!==1||chart?.namespaceURI!==C||chart.localName!=='chart'||relation?.type!==`${R}/chart`||!relation.target.startsWith('ppt/charts/')||!parts[relation.target])throw new Error('차트 관계 또는 내장 부분이 잘못되었습니다.');
              const chartPath=relation.target;
              if(!chartPackages.has(chartPath))chartPackages.set(chartPath,readChartPackage(parts,chartPath,{budget,contentTypeFor,workbooks}));
              const result=chartPackages.get(chartPath);
              // Workbook ownership is a validated package fact, independent of
              // whether this chart's visual presentation can be represented.
              if(result.verifiedWorkbookPath){usedCharts.add(chartPath);usedWorkbooks.add(result.verifiedWorkbookPath);}
              diagnostics.push(...result.diagnostics.map(d=>{
                const blocking=d.severity==='error'&&d.recoverable!==true;
                const reasonKey=d.reasonKey??(blocking?({CHART_DATA_MISMATCH:'PPTX_CHART_DATA_MISMATCH',INPUT_INVALID:'PPTX_CHART_STRUCTURE_INVALID',LIMIT_EXCEEDED:'PPTX_CHART_LIMIT_EXCEEDED',UNSUPPORTED_OBJECT:'PPTX_CHART_UNSUPPORTED'})[d.code]??'PPTX_CHART_VALIDATION_FAILED':undefined);
                return {...d,slideId:slide.id,nodeId:id,slideNumber:index+1,objectName:base,objectType:'chart',...(blocking?{blocking:true}:{}),...(reasonKey?{reasonKey}:{})};
              }));
              if(result.content){
                if((chartPoints+=result.content.categories.length*result.content.series.length)>10000)throw Object.assign(new Error('문서의 차트 데이터가 10,000점을 초과했습니다.'),{code:'LIMIT_EXCEEDED'});
                usedCharts.add(chartPath);usedWorkbooks.add(result.workbookPath);
                slide.nodes.push({...node,kind:'chart',content:structuredClone(result.content),style:{...node.style,...result.style}});model.version=2;
              }
              continue;
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
            }continue;
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
            if(!target||!parts[target]){issue('내장 그림 자산이 없습니다.',id);continue;}
            const extension=path.posix.extname(target).toLowerCase(),declaredMime=mimeByExt[extension];if(!declaredMime){issue('이 그림 형식은 지원하지 않습니다.',id);continue;}
            // Some PowerPoint files label JPEG bytes as PNG (or vice versa).
            // Normalize only between these two existing raster formats; the
            // normal asset byte, size, and SVG checks still run below.
            const bytes=parts[target],raster=['image/png','image/jpeg'].includes(declaredMime);
            const jpeg=raster&&bytes.length>=3&&bytes[0]===0xff&&bytes[1]===0xd8&&bytes[2]===0xff;
            const png=raster&&bytes.length>=8&&[137,80,78,71,13,10,26,10].every((v,i)=>bytes[i]===v);
            const mime=jpeg?'image/jpeg':png?'image/png':declaredMime;
            const assetName='image'+(jpeg?'.jpg':png?'.png':extension);
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
          }else{
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
            }else if(!['rect','roundRect','ellipse'].includes(shape))issue('지원 도형은 사각형·둥근 사각형·타원입니다.',id);
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
      // PowerPoint layers master below layout below slide. A slide can hide
      // both inherited layers; a layout can independently hide the master.
      const showShapes=d=>!['0','false'].includes(attr(d?.documentElement,'showMasterSp','1'));
      const layoutPart=[...inherited].find(part=>part.startsWith('ppt/slideLayouts/'));
      if(showShapes(doc)){
        const layoutDoc=layoutPart?xml(layoutPart):null;
        if(masterPart&&(!layoutDoc||showShapes(layoutDoc)))visit(all(masterDoc,'spTree')[0],null,undefined,0,masterPart,'master');
        if(layoutDoc)visit(all(layoutDoc,'spTree')[0],null,undefined,0,layoutPart,'layout');
      }
      visit(tree);
      currentContext={};
    }
    for(const name of Object.keys(parts)){
      if(!name.endsWith('/')&&(/(?:^|\/)embeddings\//i.test(name)||contentTypeFor(name)==='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')&&!usedWorkbooks.has(name))throw new Error('현재 차트에 연결되지 않은 내장 문서는 지원하지 않습니다.');
    }
    for(const [owner,records] of relationCache)for(const relation of records.values())if(relation.type===`${R}/package`&&(!usedCharts.has(owner)||!usedWorkbooks.has(relation.target)))throw new Error('검증한 차트 workbook 외의 패키지 관계는 지원하지 않습니다.');
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
    for(const d of diagnostics){if(d.slideId&&!d.slideNumber){const index=model.slides.findIndex(s=>s.id===d.slideId);if(index>=0)d.slideNumber=index+1;}if(d.nodeId)Object.assign(d,objectContexts.get(`${d.slideId}:${d.nodeId}`)??{});}
    for(const e of validateModel(model).errors)diagnostics.push({...diagnostic('INPUT_INVALID',`${e.path}: ${e.message}`),reasonKey:'PPTX_MODEL_INVALID',blocking:true});
  }catch(e){
    const code=typeof e.code==='string'&&(/^(FONT_|CHART_)/.test(e.code)||e.code==='LIMIT_EXCEEDED')?e.code:'INPUT_INVALID';
    diagnostics.push({...diagnostic(code,e.message,currentContext),...currentContext,reasonKey:fatalReasons.get(e.message)??(code==='INPUT_INVALID'?'PPTX_PARSE_FATAL':undefined),blocking:true});
  }
  return done();
}
