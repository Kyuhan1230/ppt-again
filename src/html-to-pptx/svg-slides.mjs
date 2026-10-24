import {diagnostic,validateModel} from './model.mjs';
import {resolveAsset} from './assets.mjs';
import {identifyFontMappings} from './font-profiles.mjs';

// Repository-owned code in the script-disabled browser. SVG viewBox coordinates
// define slides independently of the responsive page or its navigation scripts.
async function measureSvgPage({paths,timeoutMs}) {
  const diagnostics=[],slides=[],measurements=[],fonts=new Map();
  const issue=(message,slideId=null,nodeId=null,severity='error')=>diagnostics.push({code:'UNSUPPORTED_OBJECT',severity,slideId,nodeId,sourceRef:null,message});
  const wait=p=>Promise.race([p,new Promise((_,reject)=>setTimeout(()=>reject(new Error('SVG resource timeout')),timeoutMs))]);
  await wait(document.fonts.ready);
  const roots=[...document.querySelectorAll('svg')].filter(e=>paths.includes(e.dataset.sourcePath));
  const visible=el=>{for(let p=el;p;p=p.parentElement){const s=getComputedStyle(p);if(s.display==='none'||s.visibility==='hidden'||Number(s.opacity)===0)return false;}return true;};
  // A document containing SVG charts and prose is not silently reduced to charts.
  // Controls and closed dialogs are page UI; their contents are not slide content.
  const controls=el=>Boolean(el.closest('nav,button,[role="navigation"],[role="toolbar"],dialog'))||Boolean(el.closest('footer')?.querySelector('button'));
  const walker=document.createTreeWalker(document.body,NodeFilter.SHOW_TEXT);let tn;
  while((tn=walker.nextNode())){
    const p=tn.parentElement;
    if(!tn.data.trim()||p.closest('style,script,svg')||roots.some(r=>r.contains(p))||controls(p)||!visible(p))continue;
    // A sibling label inside a toolbar with buttons is also navigation chrome.
    if(p.parentElement?.querySelector('button')&&!roots.some(r=>p.parentElement.contains(r)))continue;
    issue('장표 SVG 밖에 본문 글자가 있습니다. 해당 내용까지 포함한 장표 HTML이 필요합니다.');break;
  }
  for(const media of document.querySelectorAll('img,video,canvas,iframe,object,embed'))if(!roots.some(r=>r.contains(media))&&!controls(media)&&visible(media))issue('장표 SVG 밖의 그림·동영상은 아직 함께 변환할 수 없습니다.');
  const props=['fill','fill-opacity','fill-rule','stroke','stroke-opacity','stroke-width','stroke-dasharray','stroke-linecap','stroke-linejoin','stroke-miterlimit','stop-color','stop-opacity','font-family','font-size','font-weight','font-style','font-kerning','font-variant','font-feature-settings','letter-spacing','word-spacing','text-anchor','text-decoration','dominant-baseline','alignment-baseline','white-space','visibility','display','opacity','color','filter','clip-path','mask','transform','transform-origin','vector-effect','marker-start','marker-mid','marker-end'];
  const fragment=value=>value.replace(/url\(["']?[^)"']*#([^)'" ]+)["']?\)/g,'url(#$1)');
  const copies=roots.map(root=>{
    const copy=root.cloneNode(true),originals=[root,...root.querySelectorAll('*')],clones=[copy,...copy.querySelectorAll('*')];
    originals.forEach((el,i)=>{const st=getComputedStyle(el);
      if(st.animationName!=='none')issue('움직이는 SVG 객체는 지원하지 않습니다. 애니메이션을 제거한 장표를 사용해 주세요.');
      if(['g','svg'].includes(el.localName)&&Number(st.opacity)<1&&el.children.length>1)issue('여러 SVG 객체를 합성하는 그룹 투명도는 아직 지원하지 않습니다.');
      if(st.filter!=='none'||st.clipPath!=='none'||st.maskImage&&st.maskImage!=='none')issue('SVG 객체와 그룹의 필터·마스크·클리핑은 아직 지원하지 않습니다.');
      if(['markerStart','markerMid','markerEnd'].some(k=>st[k]&&st[k]!=='none'))issue('SVG marker 화살표는 아직 지원하지 않습니다. 일반 경로로 된 화살표를 사용해 주세요.');
      for(const p of props)clones[i].style.setProperty(p,fragment(st.getPropertyValue(p)));
    });
    return copy;
  });
  // Preserve computed styles before removing the page shell. SVG <style> rules
  // are now redundant and could otherwise affect other cloned slides.
  for(const c of copies)for(const s of c.querySelectorAll('style'))s.remove();
  // Retain font-face definitions from the sanitized document, without executing JS.
  const fontRules=[];for(const sheet of document.styleSheets){try{for(const rule of sheet.cssRules)if(rule.type===CSSRule.FONT_FACE_RULE)fontRules.push(rule.cssText);}catch{}}
  document.body.replaceChildren();document.body.style.cssText='margin:0!important;padding:0!important;display:block!important;overflow:visible!important;min-width:0!important;height:auto!important;';
  document.documentElement.style.cssText='margin:0!important;padding:0!important;overflow:visible!important;';
  const fontStyle=document.createElement('style');fontStyle.textContent=fontRules.join('\n');document.body.append(fontStyle);
  copies.forEach((copy,i)=>{
    const v=copy.viewBox.baseVal,scale=1280/Math.max(v.width,v.height);copy.dataset.measuredSvg=String(i);
    copy.style.cssText+=`;display:block!important;position:relative!important;left:0!important;top:0!important;margin:0!important;padding:0!important;border:0!important;transform:none!important;overflow:hidden!important;width:${Math.round(v.width*scale)}px!important;height:${Math.round(v.height*scale)}px!important;`;
    document.body.append(copy);
  });
  await wait(document.fonts.ready);
  const round=n=>Math.round(n*10000)/10000,family=s=>s.split(',')[0].trim().replace(/^["']|["']$/g,'');
  const fontCheck=async name=>{
    if(fonts.has(name))return;
    const faces=[...document.fonts].filter(f=>f.family.replace(/^["']|["']$/g,'')===name);let available=false;
    try{if(faces.length){await wait(document.fonts.load(`20px "${name.replaceAll('"','\\"')}"`));available=faces.every(f=>f.status==='loaded');}else{await wait(new FontFace('__svg_font_probe',`local("${name.replaceAll('"','\\"')}")`).load());available=true;}}catch{}
    fonts.set(name,{family:name,available,custom:faces.length>0});
  };
  for(const [slideIndex,root] of copies.entries()){
    const id=`slide-${slideIndex+1}`,v=root.viewBox.baseVal,width=v.width,height=v.height,renderScale=1280/Math.max(width,height),nodes=[],objects=[];
    if(slides.length&&(slides[0].widthCssPx!==width||slides[0].heightCssPx!==height))issue('한 PPT 안의 서로 다른 장표 크기는 지원하지 않습니다.',id);
    const matrix=el=>root.getScreenCTM().inverse().multiply(el.getScreenCTM()),point=(m,x,y)=>({x:round(m.a*x+m.c*y+m.e-v.x),y:round(m.b*x+m.d*y+m.f-v.y)});
    const box=(el,padding=0)=>{const b=el.getBBox(),m=matrix(el),ps=[[b.x-padding,b.y-padding],[b.x+b.width+padding,b.y-padding],[b.x-padding,b.y+b.height+padding],[b.x+b.width+padding,b.y+b.height+padding]].map(([x,y])=>point(m,x,y));return {x:Math.min(...ps.map(p=>p.x)),y:Math.min(...ps.map(p=>p.y)),w:round(Math.max(...ps.map(p=>p.x))-Math.min(...ps.map(p=>p.x))),h:round(Math.max(...ps.map(p=>p.y))-Math.min(...ps.map(p=>p.y)))};};
    const opacity=el=>{let a=1;for(let p=el;p&&p!==root.parentElement;p=p.parentElement)a*=Number(getComputedStyle(p).opacity);return a;};
    const textStyle=el=>{const s=getComputedStyle(el),m=matrix(el);return {fontFamily:family(s.fontFamily),fontSize:parseFloat(s.fontSize)*m.a,fontWeight:Number(s.fontWeight)||400,fontStyle:s.fontStyle,color:s.fill==='none'?'transparent':s.fill,align:'left',lineHeight:'normal',letterSpacing:s.letterSpacing==='normal'?'normal':`${parseFloat(s.letterSpacing)*m.a}px`,textDecoration:s.textDecorationLine,opacity:opacity(el)*Number(s.fillOpacity),padding:{top:0,right:0,bottom:0,left:0}};};
    const names=new Set(),groupIds=new Map();
    const unique=base=>{let n=base,i=1;while(names.has(n))n=`${base}-${++i}`;names.add(n);return n;};
    const elements=[...root.querySelectorAll('*')];
    for(const [index,el] of elements.entries()){
      const tag=el.localName;
      if(el.closest('defs,clipPath,mask,marker,pattern')||['g','a','title','desc','style','tspan'].includes(tag)||!visible(el))continue;
      const nodeId=unique(el.id||`${tag}-${index+1}`),s=getComputedStyle(el),m=matrix(el),sourceRef={slideId:id,path:el.dataset.sourcePath};
      if(['foreignObject','use','image','svg','textPath','switch'].includes(tag)){issue(`SVG ${tag} 객체는 아직 편집 가능한 객체로 변환할 수 없습니다.`,id,nodeId);continue;}
      if(!['rect','circle','ellipse','line','text','path','polygon','polyline'].includes(tag)){issue(`SVG ${tag} 객체는 지원하지 않습니다.`,id,nodeId);continue;}
      if(s.filter!=='none'||s.clipPath!=='none'||s.maskImage&&s.maskImage!=='none'){issue('SVG 필터·마스크·클리핑은 아직 지원하지 않습니다.',id,nodeId);continue;}
      const axisAligned=Math.abs(m.b)<1e-6&&Math.abs(m.c)<1e-6&&m.a>0&&m.d>0,uniform=axisAligned&&Math.abs(m.a-m.d)<1e-6;
      let kind,content={},style={},b=box(el),measurement={};
      if(tag==='text'){
        const text=el.textContent.replace(/[\t\r\n ]+/g,' ').trim();if(!text)continue;
        if(!uniform||[el,...el.querySelectorAll('*')].some(e=>['dx','dy','rotate','textLength'].some(a=>e.hasAttribute(a)))||[...el.querySelectorAll('tspan')].some(e=>e.hasAttribute('x')||e.hasAttribute('y'))||el.querySelector('textPath')||el.getAttribute('x')?.trim().split(/[ ,]+/).length>1||el.getAttribute('y')?.trim().split(/[ ,]+/).length>1||s.dominantBaseline!=='auto'||s.alignmentBaseline!=='auto'){
          issue('회전·개별 글자 위치·여러 줄 위치가 지정된 SVG 글자는 아직 지원하지 않습니다.',id,nodeId);continue;
        }
        if(s.fill.includes('url(')||s.stroke!=='none'||parseFloat(s.wordSpacing)!==0&&s.wordSpacing!=='normal'){issue('SVG 글자의 그라데이션·외곽선·단어 간격은 아직 지원하지 않습니다.',id,nodeId);continue;}
        kind='text';style=textStyle(el);const runs=[],tw=document.createTreeWalker(el,NodeFilter.SHOW_TEXT);let child;
        while((child=tw.nextNode())){let value=child.data.replace(/[\t\r\n ]+/g,' ');if(runs.at(-1)?.text.endsWith(' '))value=value.replace(/^ +/,'');if(value)runs.push({text:value,style:textStyle(child.parentElement)});}
        if(runs.length){runs[0].text=runs[0].text.trimStart();runs.at(-1).text=runs.at(-1).text.trimEnd();}
        for(const run of runs){await fontCheck(run.style.fontFamily);if(run.style.color.includes('url('))issue('SVG 글자의 그라데이션 색은 아직 지원하지 않습니다.',id,nodeId);}
        content={text,runs:runs.filter(r=>r.text)};
        const baseline=point(m,el.getStartPositionOfChar(0).x,el.getStartPositionOfChar(0).y).y;
        measurement={lines:[{text,...b}],baseline,baselineMethod:'SVG getStartPositionOfChar',lineAdvance:null};
      }else{
        const gradient=s.fill.includes('url(')||s.stroke.includes('url('),marker=['markerStart','markerMid','markerEnd'].some(k=>s[k]&&s[k]!=='none');
        if(['rect','circle','ellipse'].includes(tag)&&uniform&&!gradient&&!marker&&s.strokeDasharray==='none'){
          kind='shape';style={svgGeometry:true,shape:tag==='rect'?'rect':'ellipse',fill:s.fill==='none'?'transparent':s.fill,stroke:s.stroke==='none'?'transparent':s.stroke,strokeWidth:s.stroke==='none'?0:parseFloat(s.strokeWidth)*m.a,radius:tag==='rect'?el.rx.baseVal.value*m.a:0,opacity:opacity(el),fillOpacity:Number(s.fillOpacity),strokeOpacity:Number(s.strokeOpacity)};
        }else if(tag==='line'&&uniform&&!gradient&&!marker&&s.strokeDasharray==='none'){
          kind='line';const start=point(m,el.x1.baseVal.value,el.y1.baseVal.value),end=point(m,el.x2.baseVal.value,el.y2.baseVal.value);
          style={svgGeometry:true,endpoints:{start,end},stroke:s.stroke==='none'?'transparent':s.stroke,strokeWidth:parseFloat(s.strokeWidth)*m.a,opacity:opacity(el)*Number(s.strokeOpacity)};
        }else{
          // Complex paths and gradients remain separate artwork, never a slide
          // screenshot containing text. Text always follows the native path above.
          kind='image';const pad=s.stroke==='none'?0:parseFloat(s.strokeWidth)/2;b=box(el,pad);if(b.w<=0||b.h<=0)continue;
          const svg=document.createElementNS('http://www.w3.org/2000/svg','svg');svg.setAttribute('xmlns','http://www.w3.org/2000/svg');svg.setAttribute('width',String(b.w));svg.setAttribute('height',String(b.h));svg.setAttribute('viewBox',`${b.x+v.x} ${b.y+v.y} ${b.w} ${b.h}`);
          for(const defs of root.querySelectorAll('defs'))svg.append(defs.cloneNode(true));
          const copy=el.cloneNode(true);copy.removeAttribute('transform');copy.style.setProperty('transform','none');copy.style.setProperty('opacity',String(opacity(el)));const g=document.createElementNS(svg.namespaceURI,'g');g.setAttribute('transform',`matrix(${m.a} ${m.b} ${m.c} ${m.d} ${m.e} ${m.f})`);g.append(copy);svg.append(g);
          // CSSOM quotes url(#id); XMLSerializer entity-escapes those quotes.
          // Use the equivalent unquoted local fragment accepted by asset checks.
          content={inlineSvg:new XMLSerializer().serializeToString(svg).replace(/url\(&quot;(#[^&<>"\s)]+)&quot;\)/g,'url($1)'),artworkOnly:true};style={opacity:1};
        }
      }
      if(kind!=='line'&&(b.w<=0||b.h<=0))continue;
      let topGroup=null;for(let p=el.parentElement;p&&p!==root;p=p.parentElement)if(p.localName==='g')topGroup=p;
      if(topGroup&&!groupIds.has(topGroup))groupIds.set(topGroup,unique(topGroup.id||`group-${groupIds.size+1}`));
      nodes.push({id:nodeId,kind,boxCssPx:b,zOrder:nodes.length,content,style,sourceRef,parentId:topGroup?groupIds.get(topGroup):null});
      objects.push({nodeId,kind,boxCssPx:b,...measurement});
    }
    for(const [g,gid] of groupIds){const children=nodes.filter(n=>n.parentId===gid);if(!children.length)continue;const x=Math.min(...children.map(n=>n.boxCssPx.x)),y=Math.min(...children.map(n=>n.boxCssPx.y));nodes.push({id:gid,kind:'group',boxCssPx:{x,y,w:Math.max(1,Math.max(...children.map(n=>n.boxCssPx.x+n.boxCssPx.w))-x),h:Math.max(1,Math.max(...children.map(n=>n.boxCssPx.y+n.boxCssPx.h))-y)},zOrder:children[0].zOrder,content:{},style:{},sourceRef:{slideId:id,path:g.dataset.sourcePath},parentId:null});if(g.querySelector('g'))issue('중첩 SVG 그룹은 상위 그룹 한 개로 묶었습니다.',id,gid,'warning');}
    slides.push({id,widthCssPx:width,heightCssPx:height,background:{color:'#ffffff',image:'none'},nodes});
    measurements.push({slideId:id,referenceSize:{width:Math.round(width*renderScale),height:Math.round(height*renderScale)},renderScale,sourceTransformScale:1,layoutStable:false,fontsReady:[...fonts.values()].every(f=>f.available),assetsReady:true,objects});
  }
  const layoutSnapshot=[...document.querySelectorAll('[data-measured-svg], [data-measured-svg] *')].map(e=>{const r=e.getBoundingClientRect();return [r.x,r.y,r.width,r.height];});
  return {slides,measurements,diagnostics,layoutSnapshot,fonts:[...fonts.values()]};
}

export async function measureSvgInput(input,runtime,session,result){
  const measured=await session.page.evaluate(measureSvgPage,{paths:input.svgRootPaths,timeoutMs:runtime.timeoutMs});
  await new Promise(resolve=>setTimeout(resolve,60));
  const after=await session.page.evaluate(()=>[...document.querySelectorAll('[data-measured-svg], [data-measured-svg] *')].map(e=>{const r=e.getBoundingClientRect();return [r.x,r.y,r.width,r.height];}));
  const layoutStable=JSON.stringify(after)===JSON.stringify(measured.layoutSnapshot);
  for(const m of measured.measurements)m.layoutStable=layoutStable;
  if(!layoutStable)measured.diagnostics.push(diagnostic('RENDER_INCOMPLETE','시간 간격을 둔 검사에서 SVG 배치가 바뀌었습니다.'));
  result.diagnostics.push(...measured.diagnostics);result.measurements=measured.measurements;result.environment.fonts=measured.fonts;
  const mappings=identifyFontMappings(input),available=await session.page.evaluate(async names=>{const ok=[];for(const name of names){try{await new FontFace('__native_probe',`local("${name}")`).load();ok.push(name);}catch{}}return ok;},[...new Set(mappings.map(m=>m.nativeLocalName))]);
  result.environment.fontMappings=mappings.filter(m=>available.includes(m.nativeLocalName));
  result.environment.inputProfile='svg-slides-v1';result.environment.slideDetection='outer SVG viewBox at least 300 by 200 units';
  result.model={version:1,sourceDigest:input.source.sha256,slides:measured.slides,assets:[...input.assets]};
  for(const slide of result.model.slides)for(const n of slide.nodes)if(n.content.inlineSvg){const asset=resolveAsset(`data:image/svg+xml;base64,${Buffer.from(n.content.inlineSvg).toString('base64')}`);if(!result.model.assets.some(a=>a.id===asset.id))result.model.assets.push(asset);n.content={assetId:asset.id,artworkOnly:true};}
  for(const font of measured.fonts)if(!font.available)result.diagnostics.push(diagnostic('FONT_UNAVAILABLE',`글꼴을 사용할 수 없습니다: ${font.family}`));
  const validation=validateModel(result.model);for(const e of validation.errors)result.diagnostics.push(diagnostic('INPUT_INVALID',`${e.path}: ${e.message}`));
  if(runtime.capture&&!result.diagnostics.some(d=>d.severity==='error'))for(const [index,m] of measured.measurements.entries()){
    await session.page.setViewportSize(m.referenceSize);const bytes=await session.page.locator(`[data-measured-svg="${index}"]`).screenshot({animations:'disabled',scale:'css'});
    if(bytes.readUInt32BE(16)!==m.referenceSize.width||bytes.readUInt32BE(20)!==m.referenceSize.height)result.diagnostics.push(diagnostic('RENDER_INCOMPLETE','SVG 캡처 크기가 기준 화면 크기와 다릅니다.',{slideId:m.slideId}));
    else result.previews.push({slideId:m.slideId,bytes:new Uint8Array(bytes),...m.referenceSize});
  }
  return result;
}
