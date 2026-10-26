import {diagnostic,validateModel} from './model.mjs';
import {resolveAsset} from './assets.mjs';
import {identifyFontMappings,installedNativeAliases} from './font-profiles.mjs';
import {captureDecorations} from './decorations.mjs';
import {measureSvgInput} from './svg-slides.mjs';

// Runs inside a script-disabled browser. This is repository-owned measurement code.
async function measurePage({selector,profile,timeoutMs,sourceHadScripts,nativeAliases=[]}) {
  const editable=profile==='editable-slides-v1',declared=editable||profile==='project-data-v1';
  const diagnostics=[];
  const issue=(code,message,slideId=null,nodeId=null,sourceRef=null)=>diagnostics.push({code,severity:'error',slideId,nodeId,sourceRef,message});
  const timeout=(promise)=>Promise.race([promise,new Promise((_,reject)=>setTimeout(()=>reject(new Error('loading timeout')),timeoutMs))]);
  try { await timeout(document.fonts.ready); }catch {issue('RENDER_INCOMPLETE','글꼴 로딩이 끝나지 않았습니다.');}
  try { await timeout(Promise.all([...document.images].map(img=>img.decode()))); }catch {issue('RENDER_INCOMPLETE','이미지 디코딩이 완료되지 않았습니다.');}
  const backgroundUrls=new Set();
  for(const el of document.querySelectorAll('*'))for(const match of getComputedStyle(el).backgroundImage.matchAll(/url\((?:"([^"]*)"|'([^']*)'|([^)]*))\)/g))backgroundUrls.add(match[1]??match[2]??match[3]);
  for(const url of backgroundUrls){try{const img=new Image();img.src=url;await timeout(img.decode());}catch{issue('RENDER_INCOMPLETE','CSS 배경 이미지 디코딩이 완료되지 않았습니다.');}}
  let roots;
  try{roots=[...document.querySelectorAll(selector)];}catch{issue('SLIDE_ROOT_REQUIRED','올바른 장표 선택자가 필요합니다.');return {slides:[],measurements:[],diagnostics};}
  if(!roots.length){issue(sourceHadScripts&&!document.body.innerText.trim()?'SCRIPT_REQUIRED':'SLIDE_ROOT_REQUIRED','장표 영역을 자동으로 찾지 못했습니다. 고정 크기의 HTML 장표 또는 SVG 장표인지 확인해 주세요.');return {slides:[],measurements:[],diagnostics};}
  const round=v=>Math.round(v*10000)/10000;
  const visibleColor=c=>c!=='transparent'&&!/^rgba\(.*[, /]0\)$/.test(c);
  const fontAvailability=new Map(),customFonts=new Set();
  const familyName=s=>s.split(',')[0].trim().replace(/^["']|["']$/g,'');
  const fontAvailable=async(family)=>{
    if(fontAvailability.has(family))return fontAvailability.get(family);
    let ok=['serif','sans-serif','monospace','system-ui','cursive','fantasy'].includes(family);
    if(!ok){
      const faces=[...document.fonts].filter(f=>f.family.replace(/^["']|["']$/g,'')===family);
      if(faces.length){customFonts.add(family);try{await timeout(document.fonts.load(`20px "${family.replaceAll('"','\\"')}"`));ok=faces.every(f=>f.status==='loaded');}catch{ok=false;}}
      else{
        for(const localName of [family,...nativeAliases.filter(a=>a.family===family).map(a=>a.localName)]){
          try{const face=new FontFace('__availability_probe',`local(${JSON.stringify(localName)})`);await timeout(face.load());ok=true;break;}catch{ok=false;}
        }
      }
    }
    fontAvailability.set(family,ok);return ok;
  };
  const slides=[],measurements=[];
  for(const [slideIndex,root] of roots.entries()) {
    const id=(declared?root.dataset.slideId:null)||root.id||`slide-${slideIndex+1}`;
    const rootStyle=getComputedStyle(root),before=root.getBoundingClientRect();
    const width=parseFloat(rootStyle.width),height=parseFloat(rootStyle.height);
    const widthValue=root.computedStyleMap?.().get('width')?.toString(),heightValue=root.computedStyleMap?.().get('height')?.toString();
    if(!width||!height||(widthValue==='auto'||heightValue==='auto')&&rootStyle.aspectRatio==='auto') {issue('SLIDE_ROOT_REQUIRED','장표의 고정 너비와 높이를 확인할 수 없습니다.',id,null,{slideId:id,path:root.dataset.sourcePath??selector});continue;}
    if(roots.some(r=>r!==root&&r.contains(root))){issue('SLIDE_ROOT_REQUIRED','장표 루트가 중첩됩니다.',id);continue;}
    if(rootStyle.transform!=='none') {const m=new DOMMatrix(rootStyle.transform);if(m.b!==0||m.c!==0||m.a<=0||Math.abs(m.a-m.d)>.0001){issue('UNSUPPORTED_OBJECT','장표 루트의 비균일/회전 변환은 지원하지 않습니다.',id);continue;}}
    const sourceTransformScale=round(before.width/width);
    if(slides.length&&(Math.abs(slides[0].widthCssPx-width)>.01||Math.abs(slides[0].heightCssPx-height)>.01))issue('UNSUPPORTED_OBJECT','한 PPT 안의 서로 다른 장표 크기는 지원하지 않습니다.',id);
    const renderScale=1280/Math.max(width,height),referenceSize={width:Math.round(width*renderScale),height:Math.round(height*renderScale)};
    // The responsive frame is presentation chrome, not a clipping mask on the
    // logical slide. Its dimensions must not crop a normalized measurement.
    if(editable&&root.parentElement?.classList.contains('slide-view')){
      root.parentElement.style.width=`${referenceSize.width}px`;root.parentElement.style.maxWidth='none';
      root.parentElement.style.height=`${referenceSize.height}px`;root.parentElement.style.overflow='visible';
    }
    // Render at the declared output size from the start; never rescale captures.
    root.style.setProperty('transform',`scale(${renderScale})`,'important');root.style.setProperty('transform-origin','0 0','important');
    const rr=root.getBoundingClientRect();
    const boxOf=el=>{const r=el.getBoundingClientRect();return {x:round((r.x-rr.x)/renderScale),y:round((r.y-rr.y)/renderScale),w:round(r.width/renderScale),h:round(r.height/renderScale)};};
    const nodes=[],objects=[],seen=new Set(),textOwners=new Set();let order=0;
    const elements=declared?[...root.querySelectorAll('[data-source-id]')]:[root,...root.querySelectorAll('*')];
    if(editable&&[...root.children].some(e=>!e.hasAttribute('data-source-id')))issue('UNSUPPORTED_OBJECT','객체 표식이 없는 장표 콘텐츠가 있습니다.',id);
    const hiddenByOpacity=el=>{for(let p=el;p;p=p.parentElement)if(Number(getComputedStyle(p).opacity)===0)return true;return false;};
    const plainStyle=st=>({fontFamily:familyName(st.fontFamily),fontSize:parseFloat(st.fontSize),fontWeight:Number(st.fontWeight)||400,color:st.color,align:st.textAlign,lineHeight:st.lineHeight,whiteSpace:st.whiteSpace,
      fill:st.backgroundColor,backgroundImage:st.backgroundImage,stroke:st.borderTopColor,strokeWidth:parseFloat(st.borderTopWidth)||0,borderStyle:st.borderTopStyle,radius:parseFloat(st.borderTopLeftRadius)||0,opacity:Number(st.opacity),padding:{top:parseFloat(st.paddingTop),right:parseFloat(st.paddingRight),bottom:parseFloat(st.paddingBottom),left:parseFloat(st.paddingLeft)}});
    const textRuns=el=>{
      const runs=[],walker=document.createTreeWalker(el,NodeFilter.SHOW_TEXT|NodeFilter.SHOW_ELEMENT);let child;
      while((child=walker.nextNode())){
        if(child.nodeType!==3&&child.tagName!=='BR')continue;
        const parent=child.nodeType===3?child.parentElement:child,cs=getComputedStyle(parent);
        if(hiddenByOpacity(parent)||cs.visibility==='hidden'||!parent.getClientRects().length)continue;
        let text=child.nodeType===3?child.data:'\n';
        if(child.nodeType===3&&!['pre','pre-wrap','break-spaces'].includes(cs.whiteSpace))text=text.replace(cs.whiteSpace==='pre-line'?/[\t\r\f ]+/g:/[\t\r\n\f ]+/g,' ');
        const style={fontFamily:familyName(cs.fontFamily),fontSize:parseFloat(cs.fontSize),fontWeight:Number(cs.fontWeight)||400,fontStyle:cs.fontStyle,color:cs.color,textDecoration:cs.textDecorationLine,letterSpacing:cs.letterSpacing};
        if(parseFloat(cs.webkitTextStrokeWidth)>0)style.outline={color:cs.webkitTextStrokeColor,width:parseFloat(cs.webkitTextStrokeWidth)};
        if(text)runs.push({text,style});
      }
      if(profile!=='project-data-v1'&&!['pre','pre-wrap','break-spaces'].includes(getComputedStyle(el).whiteSpace)){
        for(let i=1;i<runs.length;i++)if(runs[i-1].text.endsWith(' '))runs[i].text=runs[i].text.replace(/^ +/,'');
        if(runs.length){runs[0].text=runs[0].text.replace(/^ +/,'');runs.at(-1).text=runs.at(-1).text.replace(/ +$/,'');}
      }
      return runs.filter(r=>r.text.length);
    };
    const initialLayout=elements.map(boxOf);
    const lineMetrics=el=>{
      const walker=document.createTreeWalker(el,NodeFilter.SHOW_TEXT);const chars=[];let tn;
      while((tn=walker.nextNode())) {
        if(tn.parentElement.closest('svg'))continue;
        for(let i=0;i<tn.length;i++){
          const range=document.createRange();range.setStart(tn,i);range.setEnd(tn,i+1);
          const r=range.getBoundingClientRect(),char=tn.data[i];
          if(r.height>0&&r.width>0&&char!=='\n'&&char!=='\r')chars.push({text:char,x:round((r.x-rr.x)/renderScale),y:round((r.y-rr.y)/renderScale),w:round(r.width/renderScale),h:round(r.height/renderScale)});
        }
      }
      const lines=[];
      for(const c of chars){let line=lines.find(l=>Math.abs(l.y-c.y)<.5);if(!line){line={text:'',x:c.x,y:c.y,w:0,h:c.h};lines.push(line);}line.text+=c.text;line.w=round(Math.max(line.x+line.w,c.x+c.w)-Math.min(line.x,c.x));line.x=Math.min(line.x,c.x);line.h=Math.max(line.h,c.h);}
      const marker=document.createElement('span');marker.style.cssText='display:inline-block;width:0;height:0;vertical-align:baseline;padding:0;border:0;margin:0';el.prepend(marker);
      const baseline=round((marker.getBoundingClientRect().top-rr.y)/renderScale);marker.remove();
      return {lines,baseline,baselineMethod:'zero-size inline baseline marker',lineAdvance:lines.length>1?round(lines[1].y-lines[0].y):null};
    };
    for(const [index,el] of elements.entries()) {
      if(['STYLE','SCRIPT','LINK','META','BR'].includes(el.tagName)||el.closest('svg')&&el.tagName.toLowerCase()!=='svg')continue;
      const st=getComputedStyle(el);if(st.display==='none'||st.visibility==='hidden'||hiddenByOpacity(el))continue;
      const sourceId=declared?el.dataset.sourceId:null;
      const nodeId=sourceId||el.id||`node-${index+1}`,sourceRef={slideId:id,...(sourceId?{elementId:sourceId}:{path:el.dataset.sourcePath??`node[${index}]`})};
      const at=(code,msg)=>issue(code,msg,id,nodeId,sourceRef);
      if(['CANVAS','IFRAME','VIDEO','AUDIO','OBJECT','EMBED','TABLE'].includes(el.tagName)){at('UNSUPPORTED_OBJECT',`${el.tagName} 객체는 지원하지 않습니다.`);continue;}
      if(profile!=='project-data-v1'&&((el!==root&&st.transform!=='none')||st.filter!=='none'||st.perspective!=='none'||st.clipPath!=='none'||st.animationName!=='none'))at('UNSUPPORTED_OBJECT','객체의 CSS 변환·필터·클리핑·애니메이션을 지원하지 않습니다.');
      if(Number(st.opacity)<1&&el.children.length)at('UNSUPPORTED_OBJECT','여러 자식을 합성하는 컨테이너의 투명도는 지원하지 않습니다.');
      if(st.backgroundClip==='text')at('UNSUPPORTED_OBJECT','글자 윤곽으로 자르는 배경은 지원하지 않습니다.');
      for(const pseudo of ['::before','::after'])if(!['none','normal','""'].includes(getComputedStyle(el,pseudo).content))at('UNSUPPORTED_OBJECT',`CSS ${pseudo}의 생성 콘텐츠는 명시적인 HTML 요소로 넣어 주세요.`);
      let kind=declared?el.dataset.kind:null;
      let box=boxOf(el);
      if(editable){
        if(!['text','shape','line','image','group'].includes(kind)){at('UNSUPPORTED_OBJECT','지원하지 않는 편집 객체입니다.');continue;}
        if(kind!=='text'&&(el.textContent.trim()||kind!=='line'&&el.children.length))at('UNSUPPORTED_OBJECT','이 객체 안에 추가된 콘텐츠를 편집 가능한 객체로 읽을 수 없습니다.');
        if(kind==='text'&&[...el.querySelectorAll('*')].some(c=>!['SPAN','BR','B','I','STRONG','EM','U'].includes(c.tagName)))at('UNSUPPORTED_OBJECT','글자 객체 안의 비텍스트 콘텐츠는 지원하지 않습니다.');
        const fields={x:'left',y:'top',w:'width',h:'height'},authored={};
        // CSSOM stringifies numeric values to fewer significant digits. Read
        // current inline px literals, and independently compare visible bounds.
        const literals=new Map([...el.getAttribute('style').matchAll(/(?:^|;)\s*(left|top|width|height)\s*:\s*(-?(?:\d+\.?\d*|\.\d+)px)\s*(?:!important\s*)?(?=;|$)/gi)].map(m=>[m[1].toLowerCase(),m[2]]));
        for(const [key,prop] of Object.entries(fields)){
          const value=literals.get(prop)??el.style.getPropertyValue(prop);
          if(!/^-?(?:\d+\.?\d*|\.\d+)px$/.test(value)){at('INPUT_INVALID','편집 객체의 논리 좌표는 px 값이어야 합니다.');continue;}
          authored[key]=parseFloat(value);
          if(Math.abs(authored[key]-box[key])>.05)at('UNSUPPORTED_OBJECT','현재 표시 배치와 객체의 논리 좌표가 다릅니다.');
        }
        if(st.position!=='absolute'||el.parentElement!==root)at('UNSUPPORTED_OBJECT','편집 객체는 장표에 직접 배치되어야 합니다.');
        box=authored;
      }
      if((box.w<=0||box.h<=0)&&!(editable&&kind==='line'))continue;
      const isSvg=el.tagName.toLowerCase()==='svg';
      const children=[...el.children].filter(c=>!['BR','STYLE'].includes(c.tagName));
      const containsBlock=children.some(c=>!['inline','inline-block','contents'].includes(getComputedStyle(c).display));
      const text=declared?el.textContent:el.innerText;
      const hasText=typeof text==='string'&&text.trim().length>0;
      const inlineParent=[...textOwners].some(owner=>owner.contains(el));
      if(hasText&&containsBlock&&[...el.childNodes].some(n=>n.nodeType===3&&n.textContent.trim()))at('UNSUPPORTED_OBJECT','블록과 직접 텍스트가 섞인 컨테이너를 개별 텍스트 요소로 나누어 주세요.');
      if(!kind){
        if(el.tagName==='IMG'||isSvg)kind='image';
        else if(el.tagName==='HR')kind='line';
        else if(hasText&&!containsBlock&&!inlineParent)kind='text';
        else if(el!==root&&!inlineParent&&(visibleColor(st.backgroundColor)||parseFloat(st.borderTopWidth)>0||st.backgroundImage!=='none'))kind='shape';
        else continue;
      }
      if(seen.has(nodeId)){at('INPUT_INVALID','중복 객체 ID입니다.');continue;}seen.add(nodeId);
      const style=plainStyle(st);
      if(editable){
        style.valign=st.alignContent==='center'?'mid':st.alignContent==='end'?'bottom':'top';
        if(kind==='shape'){style.svgGeometry=el.dataset.geometry==='centerline';style.shape=st.borderTopLeftRadius==='50%'?'ellipse':'rect';}
        if(kind==='line'){
          const line=el.querySelector('svg > line');
          if(!line||el.querySelectorAll('svg > *').length!==1){at('UNSUPPORTED_OBJECT','편집 선은 SVG line 한 개여야 합니다.');continue;}
          const ls=getComputedStyle(line);Object.assign(style,{svgGeometry:true,stroke:ls.stroke,strokeWidth:parseFloat(ls.strokeWidth),endpoints:{start:{x:box.x+line.x1.baseVal.value,y:box.y+line.y1.baseVal.value},end:{x:box.x+line.x2.baseVal.value,y:box.y+line.y2.baseVal.value}}});
          if(ls.strokeDasharray!=='none'||ls.markerStart!=='none'||ls.markerEnd!=='none')at('UNSUPPORTED_OBJECT','점선·화살표 선은 아직 지원하지 않습니다.');
        }
      }
      // A background-painted rule has rectangular filled ends. Retain that
      // paint primitive before authored JSON fields restore their semantic style.
      if(kind==='line'&&visibleColor(st.backgroundColor)&&st.backgroundImage==='none'&&['Top','Right','Bottom','Left'].every(side=>parseFloat(st[`border${side}Width`])===0)){
        style.linePaint={kind:'bar',color:st.backgroundColor,thickness:Math.min(box.w,box.h),horizontal:box.w>=box.h};
      }
      const stacking=[];for(let parent=el;parent&&parent!==root;parent=parent.parentElement){const ps=getComputedStyle(parent),container=parent.parentElement&&getComputedStyle(parent.parentElement);if(ps.zIndex!=='auto'&&(ps.position!=='static'||/^(inline-)?(grid|flex)$/.test(container?.display)))stacking.push(Number(ps.zIndex));if(ps.order!=='0')at('UNSUPPORTED_OBJECT','CSS order를 지정한 항목은 원본 DOM 순서로 정리해 주세요.');}
      if(stacking.length>1)at('UNSUPPORTED_OBJECT','중첩된 CSS 겹침 문맥은 지원하지 않습니다.');
      style.cssZIndex=stacking.at(-1)??0;
      let runs;
      if(kind==='text') {
        textOwners.add(el);runs=textRuns(el);
        for(const family of new Set([style.fontFamily,...runs.map(r=>r.style.fontFamily)]))if(!await fontAvailable(family))at('FONT_UNAVAILABLE',`글꼴을 사용할 수 없습니다: ${family}`);
        if(runs.map(r=>r.text).join('')!==text)at('UNSUPPORTED_OBJECT','화면 문구와 텍스트 조각의 공백·줄바꿈을 일치시킬 수 없습니다.');
        if(el!==root&&(style.backgroundImage!=='none'||visibleColor(style.fill)||style.strokeWidth>0)) {
          nodes.push({id:`${nodeId}::background`,kind:'shape',boxCssPx:box,zOrder:order++,content:{},style:{...style},sourceRef,parentId:null});
        }
      }
      const content=kind==='text'?{text,runs}:kind==='image'?{assetId:el.dataset.assetId??null}:{};
      if(isSvg)content.inlineSvg=new XMLSerializer().serializeToString(el);
      nodes.push({id:nodeId,kind,boxCssPx:box,zOrder:order++,content,style,sourceRef,parentId:editable?(el.dataset.parentId??null):null});
      objects.push({nodeId,kind,boxCssPx:box,...(kind==='text'?lineMetrics(el):{})});
    }
    const layoutStable=JSON.stringify(initialLayout)===JSON.stringify(elements.map(boxOf));
    if(profile!=='project-data-v1'&&!nodes.length&&sourceHadScripts)issue('SCRIPT_REQUIRED','스크립트 없이 표시되는 장표 콘텐츠가 없습니다. 렌더된 정적 HTML을 제공해 주세요.',id);
    if(!layoutStable)issue('RENDER_INCOMPLETE','측정 중 레이아웃 경계가 바뀌었습니다.',id);
    nodes.sort((a,b)=>a.style.cssZIndex-b.style.cssZIndex||a.zOrder-b.zOrder);nodes.forEach((n,i)=>{n.zOrder=i;});
    slides.push({id,widthCssPx:width,heightCssPx:height,background:{color:rootStyle.backgroundColor,image:rootStyle.backgroundImage},nodes});
    measurements.push({slideId:id,referenceSize,renderScale,sourceTransformScale,layoutStable,objects,assetsReady:diagnostics.every(d=>d.code!=='RENDER_INCOMPLETE'),fontsReady:diagnostics.every(d=>d.code!=='FONT_UNAVAILABLE')});
  }
  const layoutSnapshot=[...document.querySelectorAll(`${selector},${selector} *`)].map(e=>{const r=e.getBoundingClientRect();return [r.x,r.y,r.width,r.height];});
  return {slides,measurements,diagnostics,layoutSnapshot,fonts:[...fontAvailability].map(([family,available])=>({family,available,custom:customFonts.has(family)}))};
}

export async function measureInput(input,runtime) {
  const result={model:null,measurements:[],environment:{...runtime.environment},diagnostics:[...input.diagnostics],previews:[]};
  if(result.diagnostics.some(d=>d.severity==='error'))return result;
  let session;
  try {
    session=await runtime.render(input);
    if(input.profile==='svg-slides-v1')return await measureSvgInput(input,runtime,session,result);
    const selector=input.profile==='project-data-v1'?'[data-slide]':input.source.slideSelector??'[data-slide]';
    const dimensions=selector=>{try{return [...document.querySelectorAll(selector)].map(r=>{const st=getComputedStyle(r);return [parseFloat(st.width),parseFloat(st.height)];});}catch{return [];}};
    const initialDimensions=await session.page.evaluate(dimensions,selector);
    await session.page.setViewportSize({width:1151,height:877});
    const alternateDimensions=await session.page.evaluate(dimensions,selector);
    await session.page.setViewportSize({width:1280,height:960});
    if(initialDimensions.some((size,i)=>size.some((value,j)=>Math.abs(value-(alternateDimensions[i]?.[j]??NaN))>.01))){
      result.diagnostics.push(diagnostic('SLIDE_ROOT_REQUIRED','화면 크기에 따라 장표 크기가 달라집니다. 고정 너비와 높이를 지정해 주세요.'));return result;
    }
    const size=await session.page.evaluate(selector=>{try{const r=document.querySelector(selector);if(!r)return null;const st=getComputedStyle(r),w=parseFloat(st.width),h=parseFloat(st.height);return w>0&&h>0?{width:Math.round(w*1280/Math.max(w,h)),height:Math.round(h*1280/Math.max(w,h))}:null;}catch{return null;}},selector);
    if(size&&size.width>0&&size.height>0)await session.page.setViewportSize(size);
    result.environment.viewport=session.page.viewportSize();
    let timer;
    const measured=await Promise.race([
      session.page.evaluate(measurePage,{selector,profile:input.profile,timeoutMs:runtime.timeoutMs,sourceHadScripts:session.sourceHadScripts,nativeAliases:installedNativeAliases()}),
      new Promise((_,reject)=>{timer=setTimeout(()=>reject(new Error('측정 제한 시간 초과')),Math.max(15000,runtime.timeoutMs));})
    ]).finally(()=>clearTimeout(timer));
    if(measured.layoutSnapshot){
      await new Promise(resolve=>setTimeout(resolve,60));
      const after=await session.page.evaluate(selector=>[...document.querySelectorAll(`${selector},${selector} *`)].map(e=>{const r=e.getBoundingClientRect();return [r.x,r.y,r.width,r.height];}),selector);
      if(JSON.stringify(after)!==JSON.stringify(measured.layoutSnapshot)){measured.diagnostics.push(diagnostic('RENDER_INCOMPLETE','시간 간격을 둔 검사에서 배치가 바뀌었습니다.'));for(const m of measured.measurements)m.layoutStable=false;}
    }
    result.diagnostics.push(...measured.diagnostics);result.measurements=measured.measurements;result.environment.fonts=measured.fonts??[];
    const mappings=identifyFontMappings(input);
    const availableNative=await session.page.evaluate(async names=>{const available=[];for(const name of names){try{await new FontFace('__native_probe',`local("${name}")`).load();available.push(name);}catch{}}return available;},[...new Set(mappings.map(m=>m.nativeLocalName))]);
    result.environment.fontMappings=mappings.filter(m=>availableNative.includes(m.nativeLocalName));
    result.model={version:1,sourceDigest:input.source.sha256,slides:measured.slides,assets:[...input.assets]};
    if(input.profile==='editable-slides-v1')result.model.pptxSizeEmu=input.pptxSizeEmu;
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
          try{const a=resolveAsset(`data:image/svg+xml;base64,${Buffer.from(n.content.inlineSvg).toString('base64')}`);if(!result.model.assets.some(x=>x.id===a.id))result.model.assets.push(a);n.content={assetId:a.id};}
          catch(e){result.diagnostics.push(diagnostic('UNSUPPORTED_OBJECT',e.message,{slideId:slide.id,nodeId:n.id,sourceRef:n.sourceRef}));}
        }
      }
      if(source)for(const groupId of new Set(source.nodes.map(n=>n.group).filter(Boolean))) {
        const children=slide.nodes.filter(n=>source.nodes.find(o=>o.id===n.id)?.group===groupId);if(!children.length)continue;
        const x=Math.min(...children.map(n=>n.boxCssPx.x)),y=Math.min(...children.map(n=>n.boxCssPx.y));
        const id=`group-${groupId}`;for(const n of children)n.parentId=id;
        slide.nodes.push({id,kind:'group',boxCssPx:{x,y,w:Math.max(...children.map(n=>n.boxCssPx.x+n.boxCssPx.w))-x,h:Math.max(...children.map(n=>n.boxCssPx.y+n.boxCssPx.h))-y},zOrder:Math.min(...children.map(n=>n.zOrder)),content:{},style:{},sourceRef:{slideId:slide.id,elementId:groupId},parentId:null});
      }
    }
    if(!result.diagnostics.some(d=>d.severity==='error')) {
      await captureDecorations(session,result.model,result.measurements,selector);
      const validation=validateModel(result.model);
      for(const e of validation.errors)result.diagnostics.push(diagnostic('INPUT_INVALID',`${e.path}: ${e.message}`));
    }
    if(runtime.capture&&!result.diagnostics.some(d=>d.code==='SLIDE_ROOT_REQUIRED')) {
      for(const [index,m] of result.measurements.entries()) {
        await session.page.setViewportSize(m.referenceSize);
        const root=session.page.locator(selector).nth(index);
        const bytes=await root.screenshot({animations:'disabled',scale:'css'});
        if(bytes.readUInt32BE(16)!==m.referenceSize.width||bytes.readUInt32BE(20)!==m.referenceSize.height){result.diagnostics.push(diagnostic('RENDER_INCOMPLETE','캡처 크기가 선언한 기준 화면 크기와 다릅니다.',{slideId:m.slideId}));continue;}
        result.previews.push({slideId:m.slideId,bytes:new Uint8Array(bytes),...m.referenceSize});
      }
    }
    return result;
  } catch(e) {
    result.diagnostics.push(diagnostic('RENDER_INCOMPLETE',`브라우저 측정이 완료되지 않았습니다: ${e.message}`));return result;
  } finally {await session?.context.close();}
}
