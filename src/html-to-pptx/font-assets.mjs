import {sha256} from './assets.mjs';
import {inspectFont,decideFontEmbedding} from './font-policy.mjs';
import {diagnostic} from './model.mjs';
import {chartText} from './charts.mjs';

const weightOf=value=>value==='normal'?400:value==='bold'?700:Number(value??400);
const familyKey=value=>String(value??'').toLowerCase();
function unicodeRanges(value){
  if(!value)return [[0,0x10FFFF]];
  return value.split(',').map(part=>{
    const match=/^U\+([\dA-F?]{1,6})(?:-([\dA-F]{1,6}))?$/i.exec(part.trim());
    if(!match||match[1].includes('?')&&(match[2]||!/^\w*\?+$/.test(match[1])))throw new Error('unicode-range');
    const start=parseInt(match[1].replaceAll('?','0'),16),end=parseInt(match[2]??match[1].replaceAll('?','F'),16);
    if(start>end||end>0x10FFFF)throw new Error('unicode-range');
    return [start,end];
  });
}

// Match only provided faces whose bytes cover the current run. CSS fallback
// lists and local() are not evidence that a particular file drew those glyphs.
export function collectFontAssets(input,model){
  const faces=[],usages=[],diagnostics=[],resolved=new Map(),seenAssets=new Set(),reportedFormats=new Set();let totalBytes=0;
  const issue=(code,message,location={},severity='error')=>diagnostics.push(diagnostic(code,message,location,severity));
  const declared=input.fontFaces??[],assets=new Map((input.assets??[]).map(a=>[a.id,a]));
  for(const slide of model?.slides??[])for(const node of slide.nodes??[]){
    if(!['text','table','chart'].includes(node.kind))continue;
    const sources=node.kind==='table'?node.content.cells:node.kind==='chart'?chartText(node.content).map(text=>({text,style:node.style})):[{...node.content,style:node.style}];
    const runs=sources.flatMap(source=>source.runs?.length?source.runs.map(run=>({...run,style:{...source.style,...run.style}})):[{text:source.text,style:source.style}]);
    for(const run of runs){
      if(!run.text)continue;
      const style={...node.style,...run.style},weight=weightOf(style.fontWeight),fontStyle=style.fontStyle??'normal';
      const family=String(style.fontFamily??''),key=JSON.stringify([familyKey(family),weight,fontStyle]);
      const loc={slideId:slide.id,nodeId:node.id,sourceRef:node.sourceRef};
      if(!resolved.has(key)){
        resolved.set(key,null);
        const matchingFamily=declared.filter(face=>familyKey(face.family)===familyKey(family));
        if(!matchingFamily.length)continue;
        const supplied=matchingFamily.flatMap(face=>face.assetIds??[]).map(id=>assets.get(id)).filter(Boolean);
        // ponytail: webfont normalization is F01.3. Keep the existing verified
        // installed-font mapping usable, without claiming these bytes embed.
        if(supplied.length&&supplied.every(asset=>asset.mime!=='font/ttf')){
          if(!reportedFormats.has(familyKey(family))){issue('FONT_FORMAT_UNSUPPORTED','이 웹폰트 형식의 PPT 포함은 아직 지원하지 않습니다.',matchingFamily[0],'warning');reportedFormats.add(familyKey(family));}
          continue;
        }
        const candidates=matchingFamily.filter(face=>weightOf(face.weight)===weight&&(face.style??'normal')===fontStyle);
        if(!candidates.length){issue('FONT_STYLE_MISMATCH','사용한 굵기·기울임과 같은 제공 폰트가 없습니다.',loc);continue;}
        if(candidates.length!==1){issue('FONT_SUBSET_UNSUPPORTED','같은 스타일의 여러 글꼴 선언·Unicode 조각은 아직 지원하지 않습니다.',loc);continue;}
        const candidate=candidates[0],ids=candidate.assetIds??[],asset=assets.get(ids[0]);
        if(candidate.hasLocal||ids.length!==1||!asset){issue('FONT_SOURCE_UNSUPPORTED','사용된 제공 글꼴 파일을 한 개로 확정할 수 없습니다.',loc);continue;}
        if(candidate.adjusted||!['normal','100%'].includes(candidate.stretch??'normal')){issue('FONT_STYLE_UNSUPPORTED','글꼴 폭 또는 메트릭을 변경하는 선언은 아직 지원하지 않습니다.',loc);continue;}
        let meta,ranges;
        try{meta=inspectFont(asset.bytes);}
        catch(error){issue(error.code??'FONT_INVALID',error.message,loc);continue;}
        const decision=decideFontEmbedding(meta,{forEditing:true,subset:false});
        if(!decision.allowed){issue(decision.reason,'편집용으로 포함할 수 없는 글꼴입니다.',loc);continue;}
        if(meta.weight!==weight||(meta.italic?'italic':'normal')!==fontStyle){issue('FONT_STYLE_MISMATCH','CSS 스타일과 실제 글꼴 파일의 스타일이 다릅니다.',loc);continue;}
        try{ranges=unicodeRanges(candidate.unicodeRange);}
        catch{issue('FONT_SUBSET_UNSUPPORTED','글꼴 Unicode 범위를 확인할 수 없습니다.',loc);continue;}
        const hash=sha256(asset.bytes);
        if(hash!==asset.sha256){issue('FONT_INVALID','글꼴 파일의 지문이 입력 자산과 다릅니다.',loc);continue;}
        const face={id:`font-${sha256(JSON.stringify([hash,familyKey(candidate.family),meta.weight,fontStyle]))}`,cssFamily:candidate.family,family:meta.family,
          weight:meta.weight,style:fontStyle,assetId:asset.id,origin:'supplied'};
        resolved.set(key,{face,meta,ranges,asset});
      }
      const selected=resolved.get(key);if(!selected)continue;
      const {face,meta,ranges,asset}=selected;
      const missing=[...run.text].find(character=>{
        const cp=character.codePointAt(0);if([9,10,13].includes(cp))return false;
        return !meta.glyphCoverage.some(([start,end])=>cp>=start&&cp<=end)||!ranges.some(([start,end])=>cp>=start&&cp<=end);
      });
      if(missing!==undefined){issue('FONT_GLYPH_MISSING','제공 글꼴이 현재 문구의 일부 문자를 포함하지 않습니다.',loc);continue;}
      if(!faces.some(item=>item.id===face.id)){
        if(!seenAssets.has(asset.id)){totalBytes+=asset.bytes.length;seenAssets.add(asset.id);}
        if(faces.length>=32||totalBytes>32*1024*1024){issue('FONT_ASSET_LIMIT','사용 글꼴의 개수 또는 전체 크기 제한을 초과했습니다.',loc);return {faces:[],usages:[],diagnostics};}
        faces.push(face);
      }
      usages.push({fontId:face.id,text:run.text,slideId:slide.id,nodeId:node.id});
    }
  }
  return {faces,usages,diagnostics};
}
