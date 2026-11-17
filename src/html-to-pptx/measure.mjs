import {diagnostic,validateModel} from './model.mjs';
import {resolveAsset} from './assets.mjs';
import {identifyFontMappings,installedNativeAliases} from './font-profiles.mjs';
import {captureDecorations} from './decorations.mjs';
import {measureSvgInput,measureSvgFragments} from './svg-slides.mjs';
import {collectFontAssets} from './font-assets.mjs';
import {normalizeWoff2Fonts} from './woff2-fonts.mjs';
import {applyFallbackFont} from './fallback-font.mjs';
import {renderChartSvg} from './charts.mjs';
import {officeCompatiblePresetGradient,presetVectorMarkup} from './preset-vector.mjs';
import {safeError} from './error-log.mjs';
import {parse} from 'parse5';
import {attribute,walkHtml} from './input.mjs';
import {scanHtmlRegions} from './html-region-scan.mjs';
import {measurePage} from './html-object-readers.mjs';
import {normalizeSlideSizes} from './slide-size-policy.mjs';

export function scopeEmbeddedModelErrors(errors,slides){
  return errors.map(error=>{const match=/^slides\[(\d+)\]\./.exec(error.path),slide=match?slides[Number(match[1])]:null;return diagnostic('INPUT_INVALID',`${error.path}: ${error.message}`,slide?{slideId:slide.id}:{});});
}

export async function measureInput(input,runtime,options={}) {
  const result={model:null,measurements:[],environment:{...runtime.environment},diagnostics:[...input.diagnostics],previews:[]};
  const normalize=()=>{
    const fitted=normalizeSlideSizes(result.model,result.measurements,input.sizePolicy);
    Object.assign(result,fitted);
    if(fitted.sizePolicy.originalAspectMismatch)result.diagnostics.push({...diagnostic('UNSUPPORTED_OBJECT','원본 PPT 크기 정보와 모든 장표의 비율이 다릅니다. 사용자가 지정한 출력 크기에 맞추었습니다.',{},'warning'),reasonKey:'SIZE_ORIGINAL_ASPECT_OVERRIDE'});
  };
  let fallbackFonts=[];
  const attachFonts=async()=>{
    const normalized=await normalizeWoff2Fonts(input);
    result.model.assets=[...result.model.assets,...normalized.assets.filter(a=>!result.model.assets.some(existing=>existing.id===a.id))];
    const fonts=collectFontAssets(normalized,result.model);result.diagnostics.push(...fonts.diagnostics);
    if(fonts.faces.length){result.model.version=2;result.model.fonts=fonts.faces;result.environment.fontUsages=fonts.usages;}
    if(fallbackFonts.length){
      const used=new Set(),usages=[];
      for(const slide of result.model.slides){let substituted=false;for(const node of slide.nodes){
        const styles=node.kind==='table'?node.content.cells.flatMap(c=>[c.style,...c.runs.map(r=>r.style)]):node.kind==='text'?[node.style,...node.content.runs.map(r=>r.style)]:node.kind==='chart'?[node.style]:[];
        for(const style of styles)if(style?.fontFamily==='PPTAgainFallback'){const weight=Number(style.fontWeight)>=600?700:400;used.add(weight);substituted=true;usages.push({fontId:`fallback-pretendard-${weight}`,slideId:slide.id,nodeId:node.id,text:''});}
      }if(substituted)result.diagnostics.push(diagnostic('FONT_SUBSTITUTED','원래 글꼴 대신 Pretendard를 사용합니다.',{slideId:slide.id},'warning'));}
      if(used.size){result.model.version=2;result.model.fonts??=[];result.environment.fontUsages??=[];for(const f of fallbackFonts.filter(f=>used.has(f.weight))){if(!result.model.assets.some(a=>a.id===f.id))result.model.assets.push({id:f.id,sha256:f.sha256,mime:f.mime,bytes:f.bytes});result.model.fonts.push({id:`fallback-pretendard-${f.weight}`,cssFamily:'PPTAgainFallback',family:'Pretendard',weight:f.weight,style:'normal',assetId:f.id,origin:'supplied'});}result.environment.fontUsages.push(...usages);}
    }
    for(const m of result.measurements)if(fonts.diagnostics.some(d=>d.severity==='error'&&(!d.slideId||d.slideId===m.slideId)))m.fontsReady=false;
  };
  if(result.diagnostics.some(d=>d.severity==='error'))return result;
  if(input.embeddedSlides){
    const slides=[],assets=new Map(),fonts=new Map(),mapped=new Map(),available=new Map(),usages=[];let version=1,assetBytes=0;
    let batch=[];
    for(const [index,child] of input.embeddedSlides.entries()){
      if(index%2===0){const settled=await Promise.allSettled(input.embeddedSlides.slice(index,index+2).map(page=>measureInput(page,runtime,{deferSizeNormalization:true})));const failed=settled.find(item=>item.status==='rejected');if(failed)throw failed.reason;batch=settled.map(item=>item.value);}
      const measured=batch[index%2],ids=new Map();
      if(measured.model&&measured.model.slides.length!==1)result.diagnostics.push(diagnostic('SLIDE_ROOT_REQUIRED','내장 데이터의 각 페이지는 정확히 한 장표여야 합니다.',{slideId:`slide-${index+1}`}));
      for(const slide of measured.model?.slides??[]){
        const id=`slide-${slides.length+1}`;
        ids.set(slide.id,id);slide.id=id;for(const node of slide.nodes){node.sourceRef.slideId=id;if(node.pptxOrigin)node.pptxOrigin.slideId=id;}slides.push(slide);
      }
      result.diagnostics.push(...measured.diagnostics.map(d=>({...d,slideId:ids.get(d.slideId)??`slide-${index+1}`})));
      result.measurements.push(...measured.measurements.map(m=>({...m,slideId:ids.get(m.slideId)??m.slideId})));
      result.previews.push(...measured.previews.map(p=>({...p,slideId:ids.get(p.slideId)??p.slideId})));
      for(const asset of measured.model?.assets??[])if(!assets.has(asset.id)){
        if((assetBytes+=asset.bytes.length)>64*1024*1024){result.diagnostics.push(diagnostic('LIMIT_EXCEEDED','내장 장표 자산의 합계 크기 한도를 초과했습니다.'));break;}
        assets.set(asset.id,asset);
      }
      if(assetBytes>64*1024*1024)break;
      for(const font of measured.model?.fonts??[]){
        const other=[...fonts.values()].find(f=>f.cssFamily===font.cssFamily&&f.weight===font.weight&&f.style===font.style&&f.assetId!==font.assetId);
        if(other)result.diagnostics.push(diagnostic('FONT_STYLE_MISMATCH','장표마다 같은 글꼴 이름에 서로 다른 파일이 연결되어 있습니다.'));
        fonts.set(font.id,font);
      }
      for(const mapping of measured.environment.fontMappings??[])for(const slideId of ids.values()){
        const scoped={...mapping,slideId};mapped.set(JSON.stringify(scoped),scoped);
      }
      for(const font of measured.environment.fonts??[])available.set(font.family,font);
      usages.push(...(measured.environment.fontUsages??[]).map(u=>({...u,slideId:ids.get(u.slideId)??u.slideId})));
      version=Math.max(version,measured.model?.version??1);
    }
    result.model={version,sourceDigest:input.source.sha256,slides,assets:[...assets.values()],...(fonts.size?{fonts:[...fonts.values()]}:{})};
    Object.assign(result.environment,{inputProfile:'embedded-static-documents-v1',sourceScriptsExecuted:false,fontMappings:[...mapped.values()],fonts:[...available.values()],...(usages.length?{fontUsages:usages}:{})});
    if(slides.length&&!result.diagnostics.some(d=>d.severity==='error'))try{normalize();}catch(error){result.diagnostics.push({...diagnostic(error.code==='INPUT_INVALID'?'INPUT_INVALID':'RENDER_INCOMPLETE',error.message),...(error.reasonKey?{reasonKey:error.reasonKey}:{})});}
    result.diagnostics.push(...scopeEmbeddedModelErrors(validateModel(result.model).errors,slides));
    return result;
  }
  let session;
  try {
    session=options.session??await runtime.render(input,{viewport:options.viewport});
    if(input.profile==='static-slides-v1')fallbackFonts=options.fallbackFonts??await applyFallbackFont(session.page,installedNativeAliases());
    if(input.profile==='svg-slides-v1'){
      result.environment.viewport=session.page.viewportSize();
      await measureSvgInput(input,runtime,session,result);await attachFonts();
      if(result.model?.slides.length&&!result.diagnostics.some(d=>d.severity==='error')){
        if(!options.deferSizeNormalization)normalize();
        for(const e of validateModel(result.model).errors)result.diagnostics.push(diagnostic('INPUT_INVALID',`${e.path}: ${e.message}`));
      }
      result.environment.previewKind='source';return result;
    }
    let selector=input.profile==='project-data-v1'?'[data-slide]':input.source.slideSelector??'[data-slide]';
    // A fixed-size candidate is only evidence for inspection; without an
    // explicit or declared slide it must never become a conversion decision.
    let automaticallyDetected=false;
    let selectionRequired=false;
    if(input.profile==='static-slides-v1'&&!input.source.slideSelector){
      selectionRequired=!await session.page.evaluate(()=>Boolean(document.querySelector('[data-slide]')));
      const detected=await session.page.evaluate(()=>{
        if(document.querySelector('[data-slide]'))return null;
        const candidates=[...document.querySelectorAll('section,article,main,div')].filter(el=>{
          const style=getComputedStyle(el),width=el.computedStyleMap().get('width')?.toString(),height=el.computedStyleMap().get('height')?.toString();
          return el.getClientRects().length>0&&style.display!=='none'&&style.visibility==='visible'&&/^\d+(?:\.\d+)?px$/.test(width??'')&&/^\d+(?:\.\d+)?px$/.test(height??'')&&parseFloat(width)>=300&&parseFloat(height)>=200;
        });
        if(candidates.length!==1)return null;
        let el=candidates[0],parts=[];while(el&&el!==document.body){const siblings=[...el.parentElement.children].filter(child=>child.tagName===el.tagName);parts.unshift(el.tagName.toLowerCase()+':nth-of-type('+(siblings.indexOf(el)+1)+')');el=el.parentElement;}
        return el?'body > '+parts.join(' > '):null;
      });
      if(detected){selector=detected;automaticallyDetected=true;}
    }
    if(selectionRequired)result.diagnostics.push(diagnostic('SELECTION_REQUIRED','장표 범위를 선택해야 합니다. 후보를 조사하거나 선택자를 지정해 주세요.'));
    if(selectionRequired&&!automaticallyDetected)return result;
    const inspectOutside=selector=>{
        let roots;try{roots=[...document.querySelectorAll(selector)];}catch{return {outside:false,hidden:false};}
        if(!roots.length)return {outside:false,hidden:false};
        const rootSet=new Set(roots),insideSet=new WeakSet(),ancestorSet=new WeakSet(),noText=new WeakSet(),noElement=new WeakSet();
        for(const el of document.querySelectorAll('*')){
          const parent=el.parentElement;
          if(rootSet.has(el)||parent&&insideSet.has(parent))insideSet.add(el);
          if(['script','style','svg'].includes(el.localName)||parent&&noText.has(parent))noText.add(el);
          if(['script','style'].includes(el.localName)||parent&&noText.has(parent))noElement.add(el);
        }
        for(const root of roots)for(let p=root.parentElement;p&&!ancestorSet.has(p);p=p.parentElement)ancestorSet.add(p);
        const inside=el=>insideSet.has(el),ancestor=el=>ancestorSet.has(el);
        const paintableCache=new WeakMap();
        const paintable=el=>{
          const pending=[];let p=el;
          while(p&&!paintableCache.has(p)){pending.push(p);p=p.parentElement;}
          let ok=p?paintableCache.get(p):true;
          while(pending.length){const item=pending.pop(),s=getComputedStyle(item);ok=ok&&s.display!=='none'&&Number(s.opacity)!==0&&s.contentVisibility!=='hidden';paintableCache.set(item,ok);}
          return ok;
        };
        const rendered=el=>getComputedStyle(el).visibility==='visible'&&paintable(el);
        const clipCache=new WeakMap();
        const clippedBy=el=>{
          const pending=[];let p=el;
          while(p&&!clipCache.has(p)){pending.push(p);p=p.parentElement;}
          let bounds=p?clipCache.get(p):{left:-Infinity,right:Infinity,top:-Infinity,bottom:Infinity};
          while(pending.length){
            const item=pending.pop(),s=getComputedStyle(item),documentRoot=item===document.documentElement||item===document.body;
            const clips=value=>!['inline','contents'].includes(s.display)&&(['hidden','clip'].includes(value)||(!documentRoot&&['scroll','auto'].includes(value)));
            if(clips(s.overflowX)||clips(s.overflowY)){
              const r=item.getBoundingClientRect(),sx=item.offsetWidth?r.width/item.offsetWidth:1,sy=item.offsetHeight?r.height/item.offsetHeight:1;
              const x=r.left+item.clientLeft*sx,y=r.top+item.clientTop*sy;
              bounds={left:clips(s.overflowX)?Math.max(bounds.left,x):bounds.left,right:clips(s.overflowX)?Math.min(bounds.right,x+item.clientWidth*sx):bounds.right,top:clips(s.overflowY)?Math.max(bounds.top,y):bounds.top,bottom:clips(s.overflowY)?Math.min(bounds.bottom,y+item.clientHeight*sy):bounds.bottom};
            }
            clipCache.set(item,bounds);
          }
          return bounds;
        };
        const visibleRect=(rect,el,includeSelf=false)=>{
          const bounds=clippedBy(includeSelf?el:el.parentElement);
          return Math.min(rect.right,bounds.right)>Math.max(rect.left,bounds.left)&&Math.min(rect.bottom,bounds.bottom)>Math.max(rect.top,bounds.top);
        };
        const rootRects=roots.map(root=>root.getBoundingClientRect());
        const exposed=(rect,el)=>{
          const clip=clippedBy(el.parentElement);
          let pieces=[{left:Math.max(rect.left,clip.left),right:Math.min(rect.right,clip.right),top:Math.max(rect.top,clip.top),bottom:Math.min(rect.bottom,clip.bottom)}].filter(r=>r.right>r.left&&r.bottom>r.top);
          for(const root of rootRects){
            const next=[];
            for(const r of pieces){
              const left=Math.max(r.left,root.left),right=Math.min(r.right,root.right),top=Math.max(r.top,root.top),bottom=Math.min(r.bottom,root.bottom);
              if(right<=left||bottom<=top){next.push(r);continue;}
              next.push({left:r.left,right:r.right,top:r.top,bottom:top},{left:r.left,right:r.right,top:bottom,bottom:r.bottom},{left:r.left,right:left,top,bottom},{left:right,right:r.right,top,bottom});
            }
            pieces=next.filter(r=>r.right>r.left&&r.bottom>r.top);
            if(!pieces.length)break;
            if(pieces.length>2000)return null;
          }
          return pieces.length>0;
        };
        const transparent=color=>color==='transparent'||/^rgba\([^)]*[, /]\s*0(?:\.0+)?\s*\)$/.test(color);
        let outside=false,hidden=false,incomplete=false;
        const walker=document.createTreeWalker(document.body,NodeFilter.SHOW_TEXT);let node;
        while((node=walker.nextNode())){
          const el=node.parentElement;
          if(!node.data.trim()||!el||noText.has(el))continue;
          if(!rendered(el)){hidden=true;continue;}
          if(inside(el))continue;
          const textStyle=getComputedStyle(el);
          if(transparent(textStyle.webkitTextFillColor)&&(!parseFloat(textStyle.webkitTextStrokeWidth)||transparent(textStyle.webkitTextStrokeColor))&&textStyle.textShadow==='none'){hidden=true;continue;}
          const range=document.createRange();range.selectNodeContents(node);
          if([...range.getClientRects()].some(rect=>visibleRect(rect,el,true)))outside=true;
          else hidden=true;
        }
        const elements=[document.documentElement,document.body,...document.body.querySelectorAll('*')];
        for(const el of elements){
          if(noElement.has(el))continue;
          const media=el.matches('img,svg,video,canvas,iframe,object,embed');
          const s=getComputedStyle(el);
          if(!inside(el)&&(s.clipPath!=='none'||s.maskImage!=='none'||s.filter!=='none'||s.mixBlendMode!=='normal'))hidden=true;
          const documentRoot=el===document.body||el===document.documentElement;
          const background=s.backgroundImage!=='none'||!transparent(s.backgroundColor)&&!(documentRoot&&s.backgroundColor==='rgb(255, 255, 255)');
          const border=['Top','Right','Bottom','Left'].some(side=>parseFloat(s[`border${side}Width`])>0&&s[`border${side}Style`]!=='none'&&!transparent(s[`border${side}Color`]));
          if(!rendered(el)){if(media||background||border)hidden=true;continue;}
          if(inside(el))continue;
          const rects=[...el.getClientRects()];
          if(s.boxShadow!=='none'||s.outlineStyle!=='none'&&parseFloat(s.outlineWidth)>0&&!transparent(s.outlineColor)){hidden=true;incomplete=true;}
          if(ancestor(el)){
            // A body background can also propagate to the viewport beyond its
            // own box. Its exact painted extent is not always measurable here.
            if(documentRoot&&background){hidden=true;incomplete=true;}
            for(const rect of rects){
              const strips=background?[rect]:['Top','Right','Bottom','Left'].filter(side=>parseFloat(s[`border${side}Width`])>0&&s[`border${side}Style`]!=='none'&&!transparent(s[`border${side}Color`])).map(side=>{
                const width=parseFloat(s[`border${side}Width`]);
                return {left:rect.left+(side==='Right'?rect.width-width:0),right:rect.right-(side==='Left'?rect.width-width:0),top:rect.top+(side==='Bottom'?rect.height-width:0),bottom:rect.bottom-(side==='Top'?rect.height-width:0)};
              });
              for(const strip of strips){const result=exposed(strip,el);if(result===null){hidden=true;incomplete=true;}else if(result)outside=true;}
            }
            continue;
          }
          if(media||background||border){
            if(rects.some(rect=>visibleRect(rect,el)))outside=true;
            else if(media)hidden=true;
          }
        }
        // Pseudo-elements have no DOM rect API. Do not alter the source DOM:
        // structural selectors can change layout even during a temporary probe.
        pseudoScan: for(const el of elements){
          if(outside)break;
          if(inside(el)||noText.has(el)||!paintable(el))continue;
          for(const pseudo of ['::before','::after']){
            const ps=getComputedStyle(el,pseudo),content=ps.content;
            if(['none','normal'].includes(content)||ps.display==='none'||ps.visibility!=='visible'||Number(ps.opacity)===0)continue;
            let value;try{value=JSON.parse(content);}catch{value='\u00a0';}
            const hasText=typeof value!=='string'||value.length>0;
            const textVisible=hasText&&(!transparent(ps.webkitTextFillColor)||parseFloat(ps.webkitTextStrokeWidth)>0&&!transparent(ps.webkitTextStrokeColor)||ps.textShadow!=='none');
            const decoration=ps.backgroundImage!=='none'||!transparent(ps.backgroundColor)||['Top','Right','Bottom','Left'].some(side=>parseFloat(ps[`border${side}Width`])>0&&ps[`border${side}Style`]!=='none'&&!transparent(ps[`border${side}Color`]));
            if(!textVisible&&!decoration){if(hasText)hidden=true;continue;}
            const host=getComputedStyle(el),r=el.getBoundingClientRect(),clip=clippedBy(el);
            const unclipped=clip.left<=r.left&&clip.right>=r.right&&clip.top<=r.top&&clip.bottom>=r.bottom;
            const simple=el.childNodes.length===0&&ps.position==='static'&&ps.display==='inline'&&ps.transform==='none'&&ps.cssFloat==='none'&&host.transform==='none'&&parseFloat(host.textIndent)===0&&parseFloat(ps.fontSize)>0&&unclipped;
            if(simple&&r.width>0&&r.height>0)outside=true;
            else {hidden=true;incomplete=true;}
          }
        }
        return {outside,hidden,incomplete};
    };
    const inspectScope=input.profile==='static-slides-v1'&&(automaticallyDetected||input.source.slideSelector);
    const initialObserved=inspectScope?await session.page.evaluate(inspectOutside,selector):null;
    const inspectExplicitSafety=async()=>{try{return await session.page.evaluate(scanHtmlRegions,{selector,onlyExplicit:true});}catch(error){if(/INPUT_LIMIT/.test(error.message))throw Object.assign(new Error('INPUT_LIMIT'),{code:'INPUT_LIMIT'});throw error;}};
    const initialScopeSafety=input.profile==='static-slides-v1'&&input.source.slideSelector?await inspectExplicitSafety():null;
    // Measure in the source viewport: any profile can contain responsive
    // descendants, including HTML exported from a PowerPoint document.
    result.environment.viewport=session.page.viewportSize();
    if(options.validateSelection)await options.validateSelection(session.page,selector);
    // The viewport does not bound the scan, but ancestor overflow does bound
    // visible content.
    if(inspectScope){
      const finalObserved=await session.page.evaluate(inspectOutside,selector);
      const observed={outside:initialObserved.outside||finalObserved.outside,hidden:initialObserved.hidden||finalObserved.hidden,incomplete:initialObserved.incomplete||finalObserved.incomplete};
      const executableScript=[...walkHtml(parse(input.source.html))].some(n=>n.tagName==='script'&&/^(?:|module|(?:text|application)\/(?:x-)?(?:javascript(?:1\.[0-5])?|ecmascript|jscript|livescript)(?:\s*;.*)?)$/i.test((attribute(n,'type')??'').trim()));
      if(observed.outside)result.diagnostics.push(diagnostic('SLIDE_CONTENT_OUTSIDE','장표 영역 밖에 표시된 내용이 있습니다. 스크립트로 생성되거나 숨겨진 내용은 확인하지 못합니다.',{},input.source.slideSelector?'warning':'error'));
      if(observed.incomplete&&!input.source.slideSelector)result.diagnostics.push(diagnostic('RENDER_INCOMPLETE','장표 바깥 내용의 검사를 마치지 못했습니다.'));
      if(observed.hidden||observed.incomplete||executableScript)result.diagnostics.push(diagnostic('SLIDE_CONTENT_UNOBSERVED','스크립트로 생성되거나 숨겨진 내용과 일부 CSS 합성은 확인하지 못했습니다. 원본과 선택 범위를 확인해 주세요.',{},'warning'));
    }
    if(selectionRequired)return result;
    if(input.profile==='static-slides-v1'&&input.source.slideSelector){
      const finalScopeSafety=await inspectExplicitSafety();
      const scopeRegions=[...initialScopeSafety.regions,...finalScopeSafety.regions];
      if(scopeRegions.some(region=>region.rootSizeMismatch)){
        result.diagnostics.push(diagnostic('UNSUPPORTED_OBJECT','장표 루트의 content-box에 패딩이나 테두리가 있습니다. box-sizing:border-box를 사용하거나 패딩과 테두리를 제거해 주세요.'));
        return result;
      }
      if(scopeRegions.some(region=>region.rootAppearance||region.ancestorAppearance||region.externalOverlap)){
        result.diagnostics.push(diagnostic('UNSUPPORTED_OBJECT','선택 영역의 배경·테두리나 겹친 요소를 보존할 수 없습니다. 범위와 효과를 확인해 주세요.'));
        return result;
      }
      if(scopeRegions.some(region=>region.externalComplex)&&!result.diagnostics.some(d=>d.code==='SLIDE_CONTENT_UNOBSERVED'))result.diagnostics.push(diagnostic('SLIDE_CONTENT_UNOBSERVED','선택 영역 밖의 일부 CSS 합성이 화면에 미치는 영향을 확인하지 못했습니다.',{},'warning'));
    }
    if(input.charts?.length){
      const frames=await session.page.evaluate(({selector,ids})=>{
        let roots;try{roots=[...document.querySelectorAll(selector)];}catch{return [];}
        return [...document.querySelectorAll('figure[data-chart-id]')].filter(el=>ids.includes(el.dataset.chartId)).map(el=>{
          const st=getComputedStyle(el),width=parseFloat(st.width),height=parseFloat(st.height);
          return {id:el.dataset.chartId,width,height,fontFamily:st.fontFamily.split(',')[0].trim().replace(/^["']|["']$/g,''),fontSize:parseFloat(st.fontSize),fixed:el.computedStyleMap?.().get('width')?.toString()!=='auto'&&el.computedStyleMap?.().get('height')?.toString()!=='auto',roots:roots.filter(root=>root===el||root.contains(el)).length};
        });
      },{selector,ids:input.charts.map(c=>c.id)});
      const displays=[];
      for(const chart of input.charts){
        const frame=frames.find(f=>f.id===chart.id);
        if(!frame||frame.roots!==1||!frame.fixed||![frame.width,frame.height].every(v=>Number.isFinite(v)&&v>0))result.diagnostics.push(diagnostic('INPUT_INVALID','차트는 한 장표 안에서 너비와 높이를 지정해 주세요.',chart));
        else displays.push({...frame,svg:renderChartSvg(chart.content,frame)});
      }
      if(result.diagnostics.some(d=>d.severity==='error'))return result;
      await session.page.evaluate(displays=>{
        for(const display of displays){
          const figure=[...document.querySelectorAll('figure[data-chart-id]')].find(el=>el.dataset.chartId===display.id),host=document.createElement('div');
          // Author styles cannot change the generated marks. The shadow tree
          // still shares the document's FontFaceSet; no font bytes are copied.
          const shadow=host.attachShadow({mode:'closed'});
          shadow.innerHTML=`<style>:host{all:initial!important;display:block!important;width:${display.width}px!important;height:${display.height}px!important;direction:ltr!important;unicode-bidi:normal!important}:host::before,:host::after{content:none!important}svg{display:block;overflow:hidden}</style>${display.svg}`;
          figure.replaceChildren(host);
        }
      },displays);
    }
    let timer;
    const measured=await Promise.race([
      session.page.evaluate(measurePage,{selector,profile:input.profile,timeoutMs:runtime.timeoutMs,sourceHadScripts:session.sourceHadScripts,allowTableOverlay:input.allowTableOverlay===true,nativeAliases:installedNativeAliases(),charts:input.charts??[]}),
      new Promise((_,reject)=>{timer=setTimeout(()=>reject(new Error('측정 제한 시간 초과')),Math.max(15000,runtime.timeoutMs));})
    ]).finally(()=>clearTimeout(timer));
    if(measured.layoutSnapshot){
      await new Promise(resolve=>setTimeout(resolve,60));
      const after=await session.page.evaluate(selector=>[...new Set([...document.querySelectorAll(selector)].flatMap(root=>[root,...root.querySelectorAll('*')]))].map(e=>{const r=e.getBoundingClientRect();return [r.x,r.y,r.width,r.height];}),selector);
      if(JSON.stringify(after)!==JSON.stringify(measured.layoutSnapshot)){measured.diagnostics.push(diagnostic('RENDER_INCOMPLETE','시간 간격을 둔 검사에서 배치가 바뀌었습니다.'));for(const m of measured.measurements)m.layoutStable=false;}
    }
    const slots=measured.slides.flatMap(slide=>slide.nodes.filter(n=>n.kind==='image'&&n.content.inlineSvg&&n.sourceRef.path).map(n=>({slideId:slide.id,path:n.sourceRef.path,renderScale:measured.measurements.find(m=>m.slideId===slide.id)?.renderScale})));
    if(slots.length){
      const svg=await measureSvgFragments(session,{selector,timeoutMs:runtime.timeoutMs,slots});
      measured.diagnostics.push(...svg.diagnostics);
      for(const font of svg.fonts){
        const prior=measured.fonts.find(f=>f.family===font.family);
        if(prior){prior.available=prior.available&&font.available;prior.custom=prior.custom||font.custom;}else measured.fonts.push(font);
        if(!font.available)measured.diagnostics.push(diagnostic('FONT_UNAVAILABLE',`글꼴을 사용할 수 없습니다: ${font.family}`));
      }
      for(const slide of measured.slides){
        const measurement=measured.measurements.find(m=>m.slideId===slide.id),used=new Set(slide.nodes.map(n=>n.id)),merged=[];
        for(const slot of slide.nodes){
          const part=svg.fragments.find(f=>f.slideId===slide.id&&f.svgPath===slot.sourceRef.path);
          if(!part){merged.push(slot);continue;}
          const ids=new Map();
          for(const node of part.nodes){let id=`${slot.id}::${node.id}`,i=1;while(used.has(id))id=`${slot.id}::${node.id}::${++i}`;used.add(id);ids.set(node.id,id);}
          for(const node of [...part.nodes].sort((a,b)=>a.zOrder-b.zOrder))merged.push({...node,id:ids.get(node.id),parentId:node.parentId?ids.get(node.parentId):null});
          measurement.objects=measurement.objects.filter(o=>o.nodeId!==slot.id);
          measurement.objects.push(...part.objects.map(o=>({...o,nodeId:ids.get(o.nodeId)})));
        }
        slide.nodes=merged;slide.nodes.forEach((node,i)=>{node.zOrder=i;});
        if(svg.fonts.some(font=>!font.available))measurement.fontsReady=false;
      }
    }
    result.diagnostics.push(...measured.diagnostics);result.measurements=measured.measurements;
    const mappings=identifyFontMappings(input);
    const selfLocalFamilies=[...new Set(input.fontFaces.filter(f=>f.selfLocal).map(f=>f.family))].filter(family=>input.fontFaces.filter(f=>f.family===family).every(f=>f.selfLocal));
    const availableNative=await session.page.evaluate(async names=>{const available=[];for(const name of names){try{await new FontFace('__native_probe',`local(${JSON.stringify(name)})`).load();available.push(name);}catch{}}return available;},[...new Set([...mappings.map(m=>m.nativeLocalName),...selfLocalFamilies])]);
    result.environment.fonts=(measured.fonts??[]).map(font=>selfLocalFamilies.includes(font.family)&&availableNative.includes(font.family)?{...font,custom:false}:font);
    result.environment.fontMappings=mappings.filter(m=>availableNative.includes(m.nativeLocalName));
    result.model={version:measured.slides.some(s=>s.nodes.some(n=>['table','chart'].includes(n.kind)))?2:1,sourceDigest:input.source.sha256,slides:measured.slides,assets:[...input.assets]};
    if(['editable-slides-v1','editable-slides-v2'].includes(input.profile))result.model.pptxSizeEmu=input.pptxSizeEmu;
    for(const slide of result.model.slides) {
      const source=input.project?.slides.find(s=>s.id===slide.id);
      for(const n of slide.nodes) {
        const original=source?.nodes.find(o=>o.id===n.id);
        if(original) {
          n.sourceRef={slideId:slide.id,elementId:n.id};
          n.boxCssPx={x:original.x,y:original.y,w:original.w,h:original.h};
          // Preserve source semantics; browser measurements remain independent evidence.
          Object.assign(n.style,{dash:original.dash??false,stroke:original.stroke??n.style.stroke,strokeWidth:original.strokeWidth??n.style.strokeWidth,fill:original.fill??n.style.fill,radius:original.radius??n.style.radius});
          if(n.style.linePaint?.kind==='bar')Object.assign(n.style.linePaint,{thickness:Math.max(1,original.strokeWidth??1),horizontal:original.w>original.h});
          if(n.kind==='image')n.content.assetId=input.assetRefs.find(r=>r.slideId===slide.id&&r.nodeId===n.id)?.assetId??null;
        }
        if(n.content.inlineSvg) {
          try{const a=resolveAsset(`data:image/svg+xml;base64,${Buffer.from(officeCompatiblePresetGradient(n.content.inlineSvg)).toString('base64')}`);if(!result.model.assets.some(x=>x.id===a.id))result.model.assets.push(a);const candidate=input.presetVectors?.find(v=>v.slideId===slide.id&&v.nodeId===n.id)?.vector,geometry=candidate&&presetVectorMarkup(candidate),native=candidate&&['ellipse','downArrow','rightArrow'].includes(candidate.shape)&&Math.abs(geometry.width-n.boxCssPx.w)<.1&&Math.abs(geometry.height-n.boxCssPx.h)<.1;n.content={assetId:a.id,...(n.content.artworkOnly?{artworkOnly:true}:{}),...(native?{editableVector:candidate}:{})};if(native)result.diagnostics=result.diagnostics.filter(d=>!(d.code==='PPTX_SHAPE_IMAGE_APPROXIMATED'&&d.slideId===slide.id&&d.nodeId===n.id));}
          catch(e){result.diagnostics.push(diagnostic('UNSUPPORTED_OBJECT',e.message,{slideId:slide.id,nodeId:n.id,sourceRef:n.sourceRef}));}
        }
        if(n.content.inlinePng){try{const a=resolveAsset(n.content.inlinePng);if(!result.model.assets.some(x=>x.id===a.id))result.model.assets.push(a);n.content={assetId:a.id,artworkOnly:true};}catch(e){result.diagnostics.push(diagnostic('UNSUPPORTED_OBJECT',e.message,{slideId:slide.id,nodeId:n.id,sourceRef:n.sourceRef}));}}
      }
      if(source)for(const groupId of new Set(source.nodes.map(n=>n.group).filter(Boolean))) {
        const children=slide.nodes.filter(n=>source.nodes.find(o=>o.id===n.id)?.group===groupId);if(!children.length)continue;
        const x=Math.min(...children.map(n=>n.boxCssPx.x)),y=Math.min(...children.map(n=>n.boxCssPx.y));
        const id=`group-${groupId}`;for(const n of children)n.parentId=id;
        slide.nodes.push({id,kind:'group',boxCssPx:{x,y,w:Math.max(...children.map(n=>n.boxCssPx.x+n.boxCssPx.w))-x,h:Math.max(...children.map(n=>n.boxCssPx.y+n.boxCssPx.h))-y},zOrder:Math.min(...children.map(n=>n.zOrder)),content:{},style:{},sourceRef:{slideId:slide.id,elementId:groupId},parentId:null});
      }
    }
    await attachFonts();
    if(!result.diagnostics.some(d=>d.severity==='error')) {
      for(const slide of result.model.slides)for(const node of slide.nodes.filter(n=>n.kind==='image'&&n.content.tableBackground)){
        const target=session.page.locator('[data-source-path='+JSON.stringify(node.sourceRef.path)+']');
        const originalRootStyle=await session.page.evaluate(({selector,path})=>{const el=[...document.querySelectorAll('[data-source-path]')].find(e=>e.dataset.sourcePath===path),root=[...document.querySelectorAll(selector)].find(r=>r.contains(el));if(!root)throw new Error('Table slide not found');const old=root.getAttribute('style'),marks=[],rules=[];for(const cell of el.querySelectorAll('td,th'))for(const pseudo of ['::before','::after']){const ps=getComputedStyle(cell,pseudo);let glyph='';try{glyph=JSON.parse(ps.content);}catch{}if(!/^[→←↑↓❯❮]$/u.test(glyph))continue;let mark=marks.find(m=>m.el===cell);if(!mark){mark={el:cell,old:cell.getAttribute('data-pptagain-pseudo'),id:String(marks.length)};marks.push(mark);cell.setAttribute('data-pptagain-pseudo',mark.id);}rules.push(`[data-pptagain-mask] [data-pptagain-pseudo="${mark.id}"]${pseudo}{color:${ps.color}!important;-webkit-text-fill-color:${ps.color}!important}`);}root.style.setProperty('position','fixed','important');root.style.setProperty('left','0','important');root.style.setProperty('top','0','important');root.style.setProperty('margin','0','important');el.setAttribute('data-pptagain-mask','');const style=document.createElement('style');style.setAttribute('data-pptagain-mask-style','');style.textContent='[data-pptagain-mask], [data-pptagain-mask] * { color: transparent !important; -webkit-text-fill-color: transparent !important; text-shadow: none !important; }'+rules.join('');document.head.append(style);return {old,marks:marks.map(m=>({path:m.el.dataset.sourcePath,old:m.old}))};},{selector,path:node.sourceRef.path});
        let bytes;try{const clip=await target.boundingBox();if(!clip||clip.x<0||clip.y<0||clip.width<=0||clip.height<=0)throw new Error('Table outside slide');bytes=await session.page.screenshot({clip,animations:'disabled',scale:'css',omitBackground:true});}finally{await session.page.evaluate(({selector,path,state})=>{const el=[...document.querySelectorAll('[data-source-path]')].find(e=>e.dataset.sourcePath===path),root=[...document.querySelectorAll(selector)].find(r=>r.contains(el));el.removeAttribute('data-pptagain-mask');for(const m of state.marks){const cell=[...el.querySelectorAll('[data-source-path]')].find(e=>e.dataset.sourcePath===m.path);if(cell){if(m.old===null)cell.removeAttribute('data-pptagain-pseudo');else cell.setAttribute('data-pptagain-pseudo',m.old);}}document.querySelector('[data-pptagain-mask-style]')?.remove();if(state.old===null)root.removeAttribute('style');else root.setAttribute('style',state.old);},{selector,path:node.sourceRef.path,state:originalRootStyle});}
        const asset=resolveAsset('data:image/png;base64,'+bytes.toString('base64'));if(!result.model.assets.some(a=>a.id===asset.id))result.model.assets.push(asset);node.content={assetId:asset.id};
      }
      await captureDecorations(session,result.model,result.measurements,selector);
    }
    if(runtime.capture&&!result.diagnostics.some(d=>d.code==='SLIDE_ROOT_REQUIRED')) {
      for(const [index,m] of result.measurements.entries()) {
        const root=session.page.locator(selector).nth(index),bytes=await root.screenshot({animations:'disabled',scale:'css'});
        const width=bytes.readUInt32BE(16),height=bytes.readUInt32BE(20);
        result.previews.push({slideId:m.slideId,bytes:new Uint8Array(bytes),width,height});
      }
    }
    if(!result.diagnostics.some(d=>d.severity==='error')&&result.model?.slides.length){
      if(!options.deferSizeNormalization)normalize();
      const validation=validateModel(result.model);
      for(const e of validation.errors)result.diagnostics.push(diagnostic('INPUT_INVALID',`${e.path}: ${e.message}`));
    }
    result.environment.previewKind='source';
    return result;
  } catch(e) {
    if(e.code==='STALE_INSPECTION'||e.code==='REGION_UNSUPPORTED'||e.code==='INPUT_LIMIT')throw e;
    result.diagnostics.push({...diagnostic(e.code==='INPUT_INVALID'?'INPUT_INVALID':'RENDER_INCOMPLETE',e.code==='INPUT_INVALID'?e.message:`브라우저 측정이 완료되지 않았습니다: ${e.message}`),...(e.reasonKey?{reasonKey:e.reasonKey}:{}),error:safeError(e)});return result;
  } finally {await session?.context.close();}
}
