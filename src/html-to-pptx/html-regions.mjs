import {sha256} from './assets.mjs';
import {applyFallbackFont} from './fallback-font.mjs';
import {installedNativeAliases} from './font-profiles.mjs';
import {scanHtmlRegions} from './html-region-scan.mjs';

const invalid=code=>Object.assign(new Error(code),{code});
const digest=value=>sha256(JSON.stringify(value));
const token=(kind,snapshot,index)=>`${kind}-${digest([snapshot,kind,index]).slice(0,24)}`;
const validViewport=v=>v&&Number.isSafeInteger(v.width)&&Number.isSafeInteger(v.height)&&v.width>0&&v.height>0&&v.width<=8192&&v.height<=8192;

// Inspection is a separate operation from conversion: it can succeed even
// when the current measurement engine cannot convert a proposed scope.
export async function inspectInput(input,runtime,{viewport={width:1280,height:960}}={}){
  if(!input?.source||!runtime||!validViewport(viewport))throw invalid('INPUT_INVALID');
  if(input.diagnostics?.some(d=>d.severity==='error'))throw invalid('INPUT_INVALID');
  if(input.embeddedSlides){
    if(!input.embeddedSlides.length||input.embeddedSlides.length>200)throw invalid('INPUT_LIMIT');
    const pages=[];
    for(const child of input.embeddedSlides)pages.push(await inspectInput(child,runtime,{viewport}));
    const snapshotId=`s-${digest({version:1,source:input.source.sha256,providedAssetsHash:input.source.providedAssetsHash,viewport,pages:pages.map(p=>p.snapshotId)})}`;
    const regions=pages.map((page,i)=>{
      const preferred=page.choices.find(c=>c.evidence.includes('declared-slide')||c.evidence.includes('validated-svg'));
      const root=preferred&&page.regions.find(r=>r.id===preferred.regionIds[0]);
      const bounds=root?.bounds??page.choices.find(c=>c.scope==='document')?.bounds??null;
      return {id:token('r',snapshotId,i),bounds,visibility:root?.visibility??'visible',parentId:null,overlaps:[],pageIndex:i,evidence:['validated-embedded-slide']};
    });
    const reasons=['conversion-selection-not-connected'];
    const choices=regions.map((region,i)=>({id:token('c',snapshotId,`page:${i}`),scope:'elements',regionIds:[region.id],evidence:['validated-embedded-slide'],outsideContentCount:0,outsideContentCountAccuracy:'estimate',convertible:false,reasons}));
    choices.push({id:token('c',snapshotId,'pages'),scope:'elements',regionIds:regions.map(r=>r.id),evidence:['validated-embedded-slides'],outsideContentCount:0,outsideContentCountAccuracy:'estimate',convertible:false,reasons});
    return {snapshotId,regions,choices,uncertainties:[...new Set(['embedded-documents-have-separate-viewports',...pages.flatMap(p=>p.uncertainties)])].sort()};
  }
  let session;
  try{
    session=await runtime.render(input,{viewport});
    const fallback=input.profile==='static-slides-v1'?await applyFallbackFont(session.page,installedNativeAliases()):[];
    await session.page.evaluate(()=>document.fonts.ready);
    let observed;
    const scanOptions={selector:input.source.slideSelector,svgPaths:input.profile==='svg-slides-v1'?input.svgRootPaths:[]};
    try{observed=await session.page.evaluate(scanHtmlRegions,scanOptions);}
    catch(error){if(/INPUT_LIMIT/.test(error.message))throw invalid('INPUT_LIMIT');if(/INPUT_INVALID/.test(error.message))throw invalid('INPUT_INVALID');throw error;}
    await new Promise(resolve=>setTimeout(resolve,60));
    let repeated;
    try{repeated=await session.page.evaluate(scanHtmlRegions,scanOptions);}
    catch(error){if(/INPUT_LIMIT/.test(error.message))throw invalid('INPUT_LIMIT');if(/INPUT_INVALID/.test(error.message))throw invalid('INPUT_INVALID');throw error;}
    if(digest(observed)!==digest(repeated))observed.uncertainties.push('layout-may-change');
    if(observed.fingerprint.images.some(([complete,width,height])=>!complete||!width||!height)||observed.fingerprint.fonts.some(f=>f[3]!=='loaded'))observed.uncertainties.push('resource-not-ready');
    // Element boxes alone do not identify native font glyphs. Include a
    // bounded image of the rendered document in the private fingerprint.
    const bounds=observed.documentBounds;
    if(bounds.x<0||bounds.y<0||bounds.width>8192||bounds.height>8192||bounds.width*bounds.height>8000000){
      observed.uncertainties.push('render-pixels-not-fully-observed');
    }else{
      try{
        const image=await session.page.screenshot({fullPage:true,animations:'disabled',scale:'css',timeout:runtime.timeoutMs});
        if(image.byteLength>16*1024*1024)observed.uncertainties.push('render-pixels-not-fully-observed');
        else observed.fingerprint.renderHash=sha256(image);
      }catch{observed.uncertainties.push('render-pixels-not-fully-observed');}
    }
    if(runtime.limitExceeded)throw invalid('INPUT_LIMIT');
    if(session.sourceHadScripts)observed.uncertainties.push('script-content-not-observed');
    if(session.blocked?.length)observed.uncertainties.push('external-resources-blocked');
    observed.uncertainties=[...new Set(observed.uncertainties)].sort();
    // Paths and document details contribute to the hash, but are never returned.
    const snapshotId=`s-${digest({version:1,source:input.source.sha256,providedAssetsHash:input.source.providedAssetsHash,profile:input.profile,selector:input.source.slideSelector,allowTableOverlay:input.allowTableOverlay,assets:input.assets.map(a=>[a.path,a.mime,a.sha256]).sort(),references:input.assetRefs.map(r=>[r.url,r.basePath,r.assetId,r.usage]),viewport,environment:runtime.environment,fallback:fallback.map(f=>f.sha256),observed})}`;
    const ids=observed.regions.map((_,i)=>token('r',snapshotId,i));
    const regions=observed.regions.map(r=>({id:ids[r.index],bounds:r.bounds,visibility:r.visibility,parentId:r.parent===null?null:ids[r.parent],overlaps:r.overlaps.map(i=>ids[i]),evidence:r.evidence}));
    const reasons=['conversion-selection-not-connected'];
    const choices=[];
    const equivalent=new Map(),leader=observed.regions.map((_,i)=>i);
    const find=i=>{while(leader[i]!==i)i=leader[i];return i;};
    for(const r of observed.regions){
      if(!r.directParent||r.parent===null||!r.bounds)continue;
      const parent=observed.regions[r.parent];
      if(!parent.neutralWrapper||!parent.bounds||JSON.stringify(parent.counts)!==JSON.stringify(r.counts)||parent.outsideCount!==r.outsideCount)continue;
      if(!['x','y','width','height'].every(key=>Math.abs(parent.bounds[key]-r.bounds[key])<.01))continue;
      leader[find(r.index)]=find(parent.index);
    }
    const components=new Map();for(const r of observed.regions){const key=find(r.index);if(!components.has(key))components.set(key,[]);components.get(key).push(r.index);}
    const duplicate=new Set();
    for(const indexes of components.values()){
      if(indexes.length<2)continue;
      const authoritative=indexes.filter(i=>observed.regions[i].evidence.some(e=>['explicit-selector','declared-slide','validated-svg'].includes(e)));
      // Distinct explicit slide declarations must remain independently selectable.
      if(authoritative.length>1)continue;
      const representative=authoritative[0]??indexes[0],rest=indexes.filter(i=>i!==representative);
      equivalent.set(representative,rest);for(const i of rest)duplicate.add(i);
    }
    for(const r of observed.regions){
      if(duplicate.has(r.index))continue;
      choices.push({id:token('c',snapshotId,`region:${r.index}`),scope:'elements',regionIds:[ids[r.index]],...(equivalent.has(r.index)?{equivalentRegionIds:equivalent.get(r.index).map(i=>ids[i])}:{}),evidence:r.evidence,outsideContentCount:r.outsideCount,outsideContentCountAccuracy:'estimate',convertible:false,reasons});
    }
    for(const [groupIndex,indexes] of observed.groups.entries()){
      choices.push({id:token('c',snapshotId,`group:${indexes.join(',')}`),scope:'elements',regionIds:indexes.map(i=>ids[i]),evidence:['repeated-sibling'],outsideContentCount:observed.groupOutsideCounts[groupIndex],outsideContentCountAccuracy:'estimate',convertible:false,reasons});
    }
    choices.sort((a,b)=>{const priority=c=>c.evidence.includes('explicit-selector')?0:c.evidence.includes('declared-slide')||c.evidence.includes('validated-svg')?1:c.evidence.includes('repeated-sibling')?2:3;return priority(a)-priority(b);});
    choices.push({id:token('c',snapshotId,'document'),scope:'document',regionIds:[],bounds:observed.documentBounds,evidence:['document'],outsideContentCount:0,convertible:false,reasons});
    return {snapshotId,regions,choices,uncertainties:observed.uncertainties};
  }finally{await session?.context.close();}
}

// Only the IDs are accepted from the caller. The candidate is reconstructed
// from a fresh rendering; coordinates and selectors from a caller are ignored.
export async function verifySelection(input,runtime,{snapshotId,choiceId,viewport}={}){
  if(typeof snapshotId!=='string'||typeof choiceId!=='string')throw invalid('STALE_INSPECTION');
  const inspection=await inspectInput(input,runtime,{viewport});
  if(inspection.uncertainties.some(reason=>['layout-may-change','resource-not-ready','render-pixels-not-fully-observed'].includes(reason)))throw invalid('STALE_INSPECTION');
  const choice=inspection.choices.find(c=>c.id===choiceId);
  if(inspection.snapshotId!==snapshotId||!choice)throw invalid('STALE_INSPECTION');
  return {snapshotId:inspection.snapshotId,choice};
}
