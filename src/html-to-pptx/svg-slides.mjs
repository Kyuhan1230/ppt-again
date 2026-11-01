import {diagnostic,validateModel} from './model.mjs';
import {resolveAsset} from './assets.mjs';
import {identifyFontMappings} from './font-profiles.mjs';

// Repository-owned code in the script-disabled browser. SVG viewBox coordinates
// define slides independently of the responsive page or its navigation scripts.
async function measureSvgPage({paths,timeoutMs,fragments=false,selector='[data-slide]',slots=[]}) {
  const diagnostics=[],slides=[],measurements=[],fonts=new Map();let artworkBytes=0;
  const issue=(message,slideId=null,nodeId=null,severity='error',code='UNSUPPORTED_OBJECT')=>diagnostics.push({code,severity,slideId,nodeId,sourceRef:null,message});
  const wait=p=>Promise.race([p,new Promise((_,reject)=>setTimeout(()=>reject(new Error('SVG resource timeout')),timeoutMs))]);
  await wait(document.fonts.ready);
  const roots=[...document.querySelectorAll('svg')].filter(e=>paths.includes(e.dataset.sourcePath));
  if(fragments)for(const root of roots){
    const st=getComputedStyle(root),visibleBackground=st.backgroundColor!=='rgba(0, 0, 0, 0)'&&st.backgroundColor!=='transparent';
    if(visibleBackground||st.backgroundImage!=='none'||st.boxShadow!=='none'||['Top','Right','Bottom','Left'].some(side=>parseFloat(st[`border${side}Width`])>0||parseFloat(st[`padding${side}`])>0))issue('SVG 루트의 CSS 배경·테두리·여백·그림자는 아직 함께 변환할 수 없습니다.');
  }
  const visible=el=>{for(let p=el;p;p=p.parentElement){const s=getComputedStyle(p);if(s.display==='none'||s.visibility==='hidden'||Number(s.opacity)===0)return false;}return true;};
  // A document containing SVG charts and prose is not silently reduced to charts.
  // Controls and closed dialogs are page UI; their contents are not slide content.
  const controls=el=>Boolean(el.closest('nav,button,[role="navigation"],[role="toolbar"],dialog'))||Boolean(el.closest('footer')?.querySelector('button'));
  if(!fragments){
    const walker=document.createTreeWalker(document.body,NodeFilter.SHOW_TEXT);let tn;
    while((tn=walker.nextNode())){
      const p=tn.parentElement;
      if(!tn.data.trim()||p.closest('style,script,svg')||roots.some(r=>r.contains(p))||controls(p)||!visible(p))continue;
      // A sibling label inside a toolbar with buttons is also navigation chrome.
      if(p.parentElement?.querySelector('button')&&!roots.some(r=>p.parentElement.contains(r)))continue;
      issue('장표 SVG 밖에 본문 글자가 있습니다. 해당 내용까지 포함한 장표 HTML이 필요합니다.');break;
    }
    for(const media of document.querySelectorAll('img,video,canvas,iframe,object,embed'))if(!roots.some(r=>r.contains(media))&&!controls(media)&&visible(media))issue('장표 SVG 밖의 그림·동영상은 아직 함께 변환할 수 없습니다.');
  }
  const props=['fill','fill-opacity','fill-rule','stroke','stroke-opacity','stroke-width','stroke-dasharray','stroke-linecap','stroke-linejoin','stroke-miterlimit','stop-color','stop-opacity','font-family','font-size','font-weight','font-style','font-kerning','font-variant','font-feature-settings','letter-spacing','word-spacing','text-anchor','text-decoration','dominant-baseline','alignment-baseline','white-space','visibility','display','opacity','color','filter','clip-path','mask','transform','transform-origin','vector-effect','marker-start','marker-mid','marker-end'];
  const fragment=value=>value.replace(/url\(["']?[^)"']*#([^)'" ]+)["']?\)/g,'url(#$1)');
  const inlineComputed=(original,clone)=>{
    const originals=[original,...original.querySelectorAll('*')],clones=[clone,...clone.querySelectorAll('*')];
    originals.forEach((item,i)=>{const st=getComputedStyle(item);for(const p of props)clones[i].style.setProperty(p,fragment(st.getPropertyValue(p)));});
  };
  const copies=roots.map(root=>{
    const copy=fragments?root:root.cloneNode(true),originals=[root,...root.querySelectorAll('*')],clones=fragments?[]:[copy,...copy.querySelectorAll('*')];
    originals.forEach((el,i)=>{const st=getComputedStyle(el);
      if(st.animationName!=='none')issue('움직이는 SVG 객체는 지원하지 않습니다. 애니메이션을 제거한 장표를 사용해 주세요.');
      if(['g','svg'].includes(el.localName)&&Number(st.opacity)<1&&el.children.length>1)issue('여러 SVG 객체를 합성하는 그룹 투명도는 아직 지원하지 않습니다.');
      if(st.filter!=='none'||st.clipPath!=='none'||st.maskImage&&st.maskImage!=='none')issue('SVG 객체와 그룹의 필터·마스크·클리핑은 아직 지원하지 않습니다.');
      if(['markerStart','markerMid','markerEnd'].some(k=>st[k]&&st[k]!=='none')&&!(fragments&&['polyline','line','path'].includes(el.localName)&&st.markerStart==='none'&&st.markerMid==='none'))issue('SVG marker 화살표는 아직 지원하지 않습니다. 일반 경로로 된 화살표를 사용해 주세요.');
      if(!fragments)for(const p of props)clones[i].style.setProperty(p,fragment(st.getPropertyValue(p)));
    });
    return copy;
  });
  // Preserve computed styles before removing the page shell. SVG <style> rules
  // are now redundant and could otherwise affect other cloned slides.
  if(!fragments)for(const c of copies)for(const s of c.querySelectorAll('style'))s.remove();
  // Retain font-face definitions from the sanitized document, without executing JS.
  const fontRules=[];for(const sheet of document.styleSheets){try{for(const rule of sheet.cssRules)if(rule.type===CSSRule.FONT_FACE_RULE)fontRules.push(rule.cssText);}catch{}}
  if(!fragments){
    document.body.replaceChildren();document.body.style.cssText='margin:0!important;padding:0!important;display:block!important;overflow:visible!important;min-width:0!important;height:auto!important;';
    document.documentElement.style.cssText='margin:0!important;padding:0!important;overflow:visible!important;';
    const fontStyle=document.createElement('style');fontStyle.textContent=fontRules.join('\n');document.body.append(fontStyle);
    copies.forEach((copy,i)=>{
      const v=copy.viewBox.baseVal,scale=1280/Math.max(v.width,v.height);copy.dataset.measuredSvg=String(i);
      copy.style.cssText+=`;display:block!important;position:relative!important;left:0!important;top:0!important;margin:0!important;padding:0!important;border:0!important;transform:none!important;overflow:hidden!important;width:${Math.round(v.width*scale)}px!important;height:${Math.round(v.height*scale)}px!important;`;
      document.body.append(copy);
    });
  }
  await wait(document.fonts.ready);
  const round=n=>Math.round(n*10000)/10000,family=s=>s.split(',')[0].trim().replace(/^["']|["']$/g,'');
  const fontCheck=async name=>{
    if(fonts.has(name))return;
    const faces=[...document.fonts].filter(f=>f.family.replace(/^["']|["']$/g,'')===name);let available=false;
    try{if(faces.length){await wait(document.fonts.load(`20px "${name.replaceAll('"','\\"')}"`));available=faces.every(f=>f.status==='loaded');}else{await wait(new FontFace('__svg_font_probe',`local("${name.replaceAll('"','\\"')}")`).load());available=true;}}catch{}
    fonts.set(name,{family:name,available,custom:faces.length>0});
  };
  for(const [slideIndex,root] of copies.entries()){
    const slot=fragments?slots.find(s=>s.path===root.dataset.sourcePath):null,slideRoot=fragments?root.closest(selector):null;
    if(fragments&&(!slot||!slideRoot)){issue('SVG가 속한 장표를 찾을 수 없습니다.');continue;}
    const id=slot?.slideId??`slide-${slideIndex+1}`,v=fragments?{x:0,y:0}:root.viewBox.baseVal,width=fragments?parseFloat(getComputedStyle(slideRoot).width):v.width,height=fragments?parseFloat(getComputedStyle(slideRoot).height):v.height,renderScale=slot?.renderScale??1280/Math.max(width,height),nodes=[],objects=[];
    if(!fragments&&slides.length&&(slides[0].widthCssPx!==width||slides[0].heightCssPx!==height))issue('한 PPT 안의 서로 다른 장표 크기는 지원하지 않습니다.',id);
    const rr=slideRoot?.getBoundingClientRect(),slideMatrix=fragments?new DOMMatrix([1/renderScale,0,0,1/renderScale,-rr.x/renderScale,-rr.y/renderScale]):null;
    const bounds=fragments?root.getBoundingClientRect():null,svgBox=fragments?{x:(bounds.x-rr.x)/renderScale,y:(bounds.y-rr.y)/renderScale,w:bounds.width/renderScale,h:bounds.height/renderScale}:null,rootStyle=getComputedStyle(root);
    const matrix=el=>fragments?slideMatrix.multiply(el.getScreenCTM()):root.getScreenCTM().inverse().multiply(el.getScreenCTM()),point=(m,x,y)=>({x:round(m.a*x+m.c*y+m.e-v.x),y:round(m.b*x+m.d*y+m.f-v.y)});
    const box=(el,padding=0)=>{const b=el.getBBox(),m=matrix(el),ps=[[b.x-padding,b.y-padding],[b.x+b.width+padding,b.y-padding],[b.x-padding,b.y+b.height+padding],[b.x+b.width+padding,b.y+b.height+padding]].map(([x,y])=>point(m,x,y));return {x:Math.min(...ps.map(p=>p.x)),y:Math.min(...ps.map(p=>p.y)),w:round(Math.max(...ps.map(p=>p.x))-Math.min(...ps.map(p=>p.x))),h:round(Math.max(...ps.map(p=>p.y))-Math.min(...ps.map(p=>p.y)))};};
    const opacity=el=>{let a=1;for(let p=el;p&&p!==root.parentElement;p=p.parentElement)a*=Number(getComputedStyle(p).opacity);return a;};
    const textStyle=el=>{const s=getComputedStyle(el),m=matrix(el);return {fontFamily:family(s.fontFamily),fontSize:parseFloat(s.fontSize)*m.a,fontWeight:Number(s.fontWeight)||400,fontStyle:s.fontStyle,color:s.fill==='none'?'transparent':s.fill,align:'left',lineHeight:'normal',letterSpacing:s.letterSpacing==='normal'?'normal':`${parseFloat(s.letterSpacing)*m.a}px`,textDecoration:s.textDecorationLine,opacity:opacity(el)*Number(s.fillOpacity),padding:{top:0,right:0,bottom:0,left:0}};};
    const names=new Set(),groupIds=new Map();
    const unique=base=>{let n=base,i=1;while(names.has(n))n=`${base}-${++i}`;names.add(n);return n;};
    const groupFor=el=>{let top=null;for(let p=el.parentElement;p&&p!==root;p=p.parentElement)if(p.localName==='g')top=p;if(top&&!groupIds.has(top))groupIds.set(top,unique(top.id||`group-${groupIds.size+1}`));return top?groupIds.get(top):null;};
    const elements=[...root.querySelectorAll('*')],separatedNested=new Set();
    for(const [index,el] of elements.entries()){
      const tag=el.localName;
      if(el.closest('defs,clipPath,mask,marker,pattern')||['g','a','title','desc','style','tspan'].includes(tag)||!visible(el))continue;
      if(fragments&&el.parentElement?.closest('svg')!==root){const nested=el.parentElement?.closest('svg');if(!separatedNested.has(nested)||tag!=='text')continue;}
      const nodeId=unique(el.id||`${tag}-${index+1}`),s=getComputedStyle(el),m=matrix(el),sourceRef={slideId:id,path:el.dataset.sourcePath};
      if(fragments&&tag==='svg'){
        // getBoundingClientRect() is the artwork bounds, not the nested viewport.
        const parentMatrix=matrix(el.parentElement),x=el.x.baseVal.value,y=el.y.baseVal.value,w=el.width.baseVal.value,h=el.height.baseVal.value,p=point(parentMatrix,x,y);
        const b={x:p.x,y:p.y,w:round(w*parentMatrix.a),h:round(h*parentMatrix.d)};
        const localPaint=value=>{
          if(!value?.includes('url('))return true;
          const ref=value.match(/^url\(["']?(?:https:\/\/html-input\.invalid\/document)?#([^"'()\s]+)["']?\)$/)?.[1];
          const matches=ref?[...el.querySelectorAll('[id]')].filter(item=>item.id===ref):[];
          return matches.length===1&&['linearGradient','radialGradient'].includes(matches[0].localName)&&document.getElementById(ref)===matches[0];
        };
        const styled=[el,...el.querySelectorAll('*')],unsafe=styled.some(item=>{const st=getComputedStyle(item);return [...item.attributes].some(a=>a.name==='href'||a.name==='xlink:href')||![st.fill,st.stroke].every(localPaint)||[st.filter,st.clipPath,st.maskImage,st.markerStart,st.markerMid,st.markerEnd].some(v=>v&&v.includes('url('))||st.mixBlendMode!=='normal';});
        let blended=false;for(let p=el.parentElement;p&&p!==root;p=p.parentElement)if(getComputedStyle(p).mixBlendMode!=='normal')blended=true;
        const cssSize=['x','y','width','height'].every(k=>s[k]==='auto'||/^-?\d+(?:\.\d+)?px$/.test(s[k])&&Math.abs(parseFloat(s[k])-el[k].baseVal.value)<.001);
        const outsideRoot=['hidden','clip','scroll','auto'].includes(rootStyle.overflowX)&&(b.x<svgBox.x-.01||b.x+b.w>svgBox.x+svgBox.w+.01)||['hidden','clip','scroll','auto'].includes(rootStyle.overflowY)&&(b.y<svgBox.y-.01||b.y+b.h>svgBox.y+svgBox.h+.01);
        const nestedTexts=[...el.querySelectorAll('text')],nestedRect=el.getBoundingClientRect(),paintOrder=[...el.querySelectorAll('*')];
        const invalidText=nestedTexts.some(text=>{const r=text.getBoundingClientRect();if(text.closest('defs')||r.left<nestedRect.left-.01||r.top<nestedRect.top-.01||r.right>nestedRect.right+.01||r.bottom>nestedRect.bottom+.01)return true;const later=paintOrder.slice(paintOrder.indexOf(text)+1).filter(item=>!text.contains(item)&&['rect','circle','ellipse','line','path','polygon','polyline'].includes(item.localName)&&!item.closest('defs'));return later.some(item=>{const q=item.getBoundingClientRect(),paint=getComputedStyle(item),pad=paint.stroke==='none'?0:(parseFloat(paint.strokeWidth)||0)*Math.max(Math.abs(m.a),Math.abs(m.d))/2*renderScale;return q.right+pad>r.left&&q.left-pad<r.right&&q.bottom+pad>r.top&&q.top-pad<r.bottom;});});
        if(unsafe||blended||!cssSize||outsideRoot||invalidText||el.hasAttribute('transform')||s.transform!=='none'||el.querySelector('svg,textPath,image,use,foreignObject,script')||!el.viewBox.baseVal.width||!el.viewBox.baseVal.height||b.w<=0||b.h<=0||Math.abs(parentMatrix.b)>=1e-6||Math.abs(parentMatrix.c)>=1e-6||parentMatrix.a<=0||Math.abs(parentMatrix.a-parentMatrix.d)>=1e-6||Math.abs(m.b)>=1e-6||Math.abs(m.c)>=1e-6||m.a<=0||Math.abs(m.a-m.d)>=1e-6||s.overflowX!=='hidden'||s.overflowY!=='hidden'||s.backgroundColor!=='rgba(0, 0, 0, 0)'&&s.backgroundColor!=='transparent'||s.backgroundImage!=='none'||s.boxShadow!=='none'||['Top','Right','Bottom','Left'].some(side=>parseFloat(s[`border${side}Width`])>0||parseFloat(s[`padding${side}`])>0)){
          issue('중첩 SVG의 글자·참조·효과 또는 뷰포트는 아직 보존할 수 없습니다.',id,nodeId);continue;
        }
        const copy=el.cloneNode(true);inlineComputed(el,copy);for(const style of copy.querySelectorAll('style'))style.remove();for(const text of copy.querySelectorAll('text'))text.remove();if(nestedTexts.length)separatedNested.add(el);
        copy.setAttribute('xmlns','http://www.w3.org/2000/svg');copy.removeAttribute('x');copy.removeAttribute('y');copy.removeAttribute('transform');copy.setAttribute('width',String(b.w));copy.setAttribute('height',String(b.h));
        copy.style.setProperty('width',`${b.w}px`,'important');copy.style.setProperty('height',`${b.h}px`,'important');copy.style.setProperty('transform','none','important');copy.style.setProperty('overflow',s.overflowX,'important');copy.style.setProperty('opacity',String(opacity(el)));
        const markup=new XMLSerializer().serializeToString(copy).replace(/url\(&quot;(#[^&<>"\s)]+)&quot;\)/g,'url($1)'),bytes=new TextEncoder().encode(markup).length;
        if(bytes>16*1024*1024||(artworkBytes+=bytes)>64*1024*1024){issue('SVG 그림 자산의 크기 한도를 초과했습니다.',id,nodeId,'error','LIMIT_EXCEEDED');break;}
        nodes.push({id:nodeId,kind:'image',boxCssPx:b,zOrder:nodes.length,content:{inlineSvg:markup,artworkOnly:true},style:{opacity:1},sourceRef,parentId:groupFor(el)});
        objects.push({nodeId,kind:'image',boxCssPx:b});continue;
      }
      if(tag==='image'){
        const assetId=el.dataset.assetId,aspect=el.getAttribute('preserveAspectRatio')||'xMidYMid meet',rawBox=box(el);
        if(!assetId||!/^asset-[a-f0-9]{64}$/.test(assetId)||!/^data:image\/(png|jpeg);base64,/i.test(el.getAttribute('href')||'')||!(/^(none|x(?:Min|Mid|Max)Y(?:Min|Mid|Max)(?: meet)?)$/.test(aspect))||s.filter!=='none'||s.clipPath!=='none'||s.maskImage&&s.maskImage!=='none'||s.mixBlendMode!=='normal') {issue('SVG 그림의 형식·회전·자르기는 아직 보존할 수 없습니다.',id,nodeId);continue;}
        const image=new Image();image.src=el.getAttribute('href');try{await wait(image.decode());}catch{issue('SVG 그림을 읽을 수 없습니다.',id,nodeId);continue;}
        if(Math.abs(m.b)>=1e-6||Math.abs(m.c)>=1e-6||m.a<=0||m.d<=0){
          const blended=[];for(let parent=el.parentElement;parent&&parent!==slideRoot;parent=parent.parentElement){const ps=getComputedStyle(parent);if(ps.mixBlendMode!=='normal')blended.push(ps.mixBlendMode);}
          const crop=fragments?{x:Math.max(0,svgBox.x),y:Math.max(0,svgBox.y),w:Math.min(width,svgBox.x+svgBox.w)-Math.max(0,svgBox.x),h:Math.min(height,svgBox.y+svgBox.h)-Math.max(0,svgBox.y)}:null;
          let blendBackground=null;
          if(fragments&&blended.length===1&&blended[0]==='lighten'){
            const base=getComputedStyle(slideRoot),rgb=base.backgroundColor.match(/^rgb\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)$/);
            let safeBackground=Boolean(rgb)&&base.backgroundImage==='none'&&Number(base.opacity)===1;
            for(let parent=root;safeBackground&&parent&&parent!==slideRoot;parent=parent.parentElement){const ps=getComputedStyle(parent),rect=parent.getBoundingClientRect(),sameBackground=ps.backgroundColor===base.backgroundColor&&rect.left<=rr.x+crop.x*renderScale+.01&&rect.top<=rr.y+crop.y*renderScale+.01&&rect.right>=rr.x+(crop.x+crop.w)*renderScale-.01&&rect.bottom>=rr.y+(crop.y+crop.h)*renderScale-.01;if(ps.backgroundImage!=='none'||!['transparent','rgba(0, 0, 0, 0)'].includes(ps.backgroundColor)&&!sameBackground||Number(ps.opacity)!==1)safeBackground=false;}
            const overlap=rect=>rect.right>rr.x+crop.x*renderScale&&rect.left<rr.x+(crop.x+crop.w)*renderScale&&rect.bottom>rr.y+crop.y*renderScale&&rect.top<rr.y+(crop.y+crop.h)*renderScale;
            const transparent=color=>color==='transparent'||/^rgba\([^)]*[, /]0(?:\.0+)?\)$/.test(color);
            for(const other of slideRoot.querySelectorAll('*')){
              if(!safeBackground)break;if(other.contains(root)||root.contains(other)||!visible(other))continue;
              const os=getComputedStyle(other),rect=other.getBoundingClientRect();
              if(['::before','::after'].some(p=>!['none','normal','""'].includes(getComputedStyle(other,p).content))){safeBackground=false;break;}
              for(const child of other.childNodes)if(child.nodeType===3&&child.data.trim()){const range=document.createRange();range.selectNode(child);if([...range.getClientRects()].some(overlap)){safeBackground=false;break;}}
              if(!safeBackground||!overlap(rect))continue;
              const painted=!transparent(os.backgroundColor)||os.backgroundImage!=='none'||['Top','Right','Bottom','Left'].some(side=>parseFloat(os[`border${side}Width`])>0&&!transparent(os[`border${side}Color`]))||['img','svg','canvas','video'].includes(other.localName);
              if(painted){safeBackground=false;break;}
            }
            if(safeBackground)blendBackground=rgb.slice(1).map(Number);
          }
          const safe=fragments&&(!blended.length||blendBackground)&&rootStyle.overflowX==='hidden'&&rootStyle.overflowY==='hidden'&&el.parentElement===root&&[m.a,m.b,m.c,m.d,m.e,m.f].every(Number.isFinite);
          const cw=Math.ceil(crop?.w*renderScale),ch=Math.ceil(crop?.h*renderScale);
          if(!safe||cw<=0||ch<=0||cw*ch>16000000){issue('회전한 SVG 그림의 잘림을 안전하게 보존할 수 없습니다.',id,nodeId);continue;}
          let x=el.x.baseVal.value,y=el.y.baseVal.value,w=el.width.baseVal.value,h=el.height.baseVal.value;
          if(aspect!=='none'){const k=Math.min(w/image.naturalWidth,h/image.naturalHeight),iw=image.naturalWidth*k,ih=image.naturalHeight*k;x+=(w-iw)*(aspect.includes('xMid')?.5:aspect.includes('xMax')?1:0);y+=(h-ih)*(aspect.includes('YMid')?.5:aspect.includes('YMax')?1:0);w=iw;h=ih;}
          const canvas=document.createElement('canvas');canvas.width=cw;canvas.height=ch;const ctx=canvas.getContext('2d');ctx.globalAlpha=opacity(el);ctx.setTransform(m.a*renderScale,m.b*renderScale,m.c*renderScale,m.d*renderScale,(m.e-crop.x)*renderScale,(m.f-crop.y)*renderScale);ctx.drawImage(image,x,y,w,h);
          if(blendBackground){const pixels=ctx.getImageData(0,0,cw,ch);for(let i=0;i<pixels.data.length;i+=4)if(pixels.data[i+3])for(let c=0;c<3;c++)pixels.data[i+c]=Math.max(pixels.data[i+c],blendBackground[c]);ctx.putImageData(pixels,0,0);}
          const inlinePng=canvas.toDataURL('image/png');if(inlinePng.length>24*1024*1024){issue('SVG 그림 자산의 크기 한도를 초과했습니다.',id,nodeId,'error','LIMIT_EXCEEDED');continue;}
          nodes.push({id:nodeId,kind:'image',boxCssPx:crop,zOrder:nodes.length,content:{inlinePng,artworkOnly:true},style:{opacity:1},sourceRef,parentId:groupFor(el)});objects.push({nodeId,kind:'image',boxCssPx:crop});continue;
        }
        let b=rawBox;if(aspect!=='none'){
          const scale=Math.min(rawBox.w/image.naturalWidth,rawBox.h/image.naturalHeight),w=image.naturalWidth*scale,h=image.naturalHeight*scale;
          b={x:round(rawBox.x+(rawBox.w-w)*(aspect.includes('xMid')?.5:aspect.includes('xMax')?1:0)),y:round(rawBox.y+(rawBox.h-h)*(aspect.includes('YMid')?.5:aspect.includes('YMax')?1:0)),w:round(w),h:round(h)};
        }
        if(b.w<=0||b.h<=0)continue;
        if(fragments&&(['hidden','clip','scroll','auto'].includes(rootStyle.overflowX)&&(b.x<svgBox.x-.01||b.x+b.w>svgBox.x+svgBox.w+.01)||['hidden','clip','scroll','auto'].includes(rootStyle.overflowY)&&(b.y<svgBox.y-.01||b.y+b.h>svgBox.y+svgBox.h+.01))){issue('SVG 영역 밖으로 잘리는 그림은 아직 보존할 수 없습니다.',id,nodeId);continue;}
        nodes.push({id:nodeId,kind:'image',boxCssPx:b,zOrder:nodes.length,content:{assetId,artworkOnly:true},style:{opacity:opacity(el)},sourceRef,parentId:groupFor(el)});objects.push({nodeId,kind:'image',boxCssPx:b});continue;
      }
      if(['foreignObject','use','svg','textPath','switch'].includes(tag)){issue(`SVG ${tag} 객체는 아직 편집 가능한 객체로 변환할 수 없습니다.`,id,nodeId);continue;}
      if(!['rect','circle','ellipse','line','text','path','polygon','polyline'].includes(tag)){issue(`SVG ${tag} 객체는 지원하지 않습니다.`,id,nodeId);continue;}
      if(s.filter!=='none'||s.clipPath!=='none'||s.maskImage&&s.maskImage!=='none'){issue('SVG 필터·마스크·클리핑은 아직 지원하지 않습니다.',id,nodeId);continue;}
      const axisAligned=Math.abs(m.b)<1e-6&&Math.abs(m.c)<1e-6&&m.a>0&&m.d>0,uniform=axisAligned&&Math.abs(m.a-m.d)<1e-6;
      let kind,content={},style={},b=box(el),measurement={};
      const miter=fragments&&['path','polygon','polyline'].includes(tag)&&s.strokeLinejoin==='miter'?Math.max(1,parseFloat(s.strokeMiterlimit)||4):1;
      const strokePadding=s.stroke==='none'||Number(s.strokeOpacity)===0?0:(parseFloat(s.strokeWidth)||0)*miter/2,painted=strokePadding?box(el,strokePadding):b;
      if(fragments&&(['hidden','clip','scroll','auto'].includes(rootStyle.overflowX)&&(painted.x<svgBox.x-.01||painted.x+painted.w>svgBox.x+svgBox.w+.01)||['hidden','clip','scroll','auto'].includes(rootStyle.overflowY)&&(painted.y<svgBox.y-.01||painted.y+painted.h>svgBox.y+svgBox.h+.01)))issue('SVG 영역 밖으로 잘리는 객체는 아직 편집 가능한 형태로 변환할 수 없습니다.',id,nodeId);
      if(tag==='text'){
        if([...el.querySelectorAll('*')].some(child=>child.textContent.trim()&&!visible(child))){issue('숨겨진 SVG 글자는 편집 가능한 텍스트로 안전하게 변환할 수 없습니다.',id,nodeId);continue;}
        const text=el.textContent.replace(/[\t\r\n ]+/g,' ').trim();if(!text)continue;
        const onePosition=(element,attr)=>!element.hasAttribute(attr)||element[attr]?.baseVal?.numberOfItems===1;
        const spans=[...el.querySelectorAll('tspan')],span=spans[0],singlePositionedSpan=spans.length===1&&span.parentElement===el&&!span.children.length&&span.hasAttribute('x')&&span.hasAttribute('dy')&&!span.hasAttribute('y')&&[...el.childNodes].every(n=>n===span||n.nodeType===3&&!n.data.trim())&&['x','dy'].every(a=>onePosition(span,a));
        const stretchedText=el.hasAttribute('textLength')&&!el.children.length&&el.getAttribute('lengthAdjust')==='spacingAndGlyphs'&&/^[가-힣a-zA-Z0-9 .,!?·:;()\-]+$/.test(text)&&text.length>1;
        const scaledText=!uniform&&axisAligned&&!el.children.length&&!el.hasAttribute('textLength')&&/^[가-힣a-zA-Z0-9 .,!?·:;()→\-]+$/.test(text)&&text.length>0;
        const lineSpans=[...el.children];
        const positionedLines=lineSpans.length>1&&[...el.childNodes].every(n=>n.nodeType===3&&!n.data.trim()||n.nodeType===1&&n.localName==='tspan')&&!['dx','dy','rotate','textLength'].some(a=>el.hasAttribute(a))&&['x','y'].every(a=>onePosition(el,a))&&lineSpans.every(t=>t.hasAttribute('x')&&t.hasAttribute('y')!==t.hasAttribute('dy')&&!['dx','rotate','textLength'].some(a=>t.hasAttribute(a))&&['x','y','dy'].every(a=>onePosition(t,a))&&[...t.querySelectorAll('*')].every(c=>c.localName==='tspan'&&!['x','y','dx','dy','rotate','textLength'].some(a=>c.hasAttribute(a))));
        if(positionedLines&&uniform&&['auto','central'].includes(s.dominantBaseline)&&s.alignmentBaseline==='auto'){
          const parentId=groupFor(el);
          for(const [lineIndex,line] of lineSpans.entries()){
            const value=line.textContent.replace(/[\t\r\n ]+/g,' ').trim();if(!value)continue;
            const ls=getComputedStyle(line),lm=matrix(line),lineBox=box(line);
            if(Math.abs(lm.b)>=1e-6||Math.abs(lm.c)>=1e-6||lm.a<=0||Math.abs(lm.a-lm.d)>=1e-6||[line,...line.querySelectorAll('*')].some(part=>{const ps=getComputedStyle(part);return !['auto','central'].includes(ps.dominantBaseline)||ps.alignmentBaseline!=='auto'||ps.fill.includes('url(')||ps.stroke!=='none'||parseFloat(ps.wordSpacing)!==0&&ps.wordSpacing!=='normal';})){issue('위치가 지정된 SVG 글자 줄의 변환·기준선·색은 아직 지원하지 않습니다.',id,nodeId);continue;}
            if(lineBox.w<=0||lineBox.h<=0)continue;
            const lineId=unique(`${nodeId}-line-${lineIndex+1}`),lineStyle=textStyle(line),position=line.getStartPositionOfChar(0),baseline=point(lm,position.x,position.y).y;
            const runs=[],walker=document.createTreeWalker(line,NodeFilter.SHOW_TEXT);let child;
            while((child=walker.nextNode())){let text=child.data.replace(/[\t\r\n ]+/g,' ');if(runs.at(-1)?.text.endsWith(' '))text=text.replace(/^ +/,'');if(text)runs.push({text,style:textStyle(child.parentElement)});}
            if(runs.length){runs[0].text=runs[0].text.trimStart();runs.at(-1).text=runs.at(-1).text.trimEnd();}
            for(const run of runs)await fontCheck(run.style.fontFamily);
            nodes.push({id:lineId,kind:'text',boxCssPx:lineBox,zOrder:nodes.length,content:{text:value,runs:runs.filter(r=>r.text)},style:lineStyle,sourceRef:{slideId:id,path:line.dataset.sourcePath??el.dataset.sourcePath},parentId});
            objects.push({nodeId:lineId,kind:'text',boxCssPx:lineBox,lines:[{text:value,...lineBox}],baseline,baselineMethod:'SVG getStartPositionOfChar',lineAdvance:null});
          }
          continue;
        }
        if(!uniform&&!scaledText||[el,...el.querySelectorAll('*')].some(e=>['dx','dy','rotate','textLength'].some(a=>e.hasAttribute(a)&&!(singlePositionedSpan&&e===span&&a==='dy')&&!(stretchedText&&e===el&&a==='textLength')))||spans.some(e=>(e.hasAttribute('x')||e.hasAttribute('y'))&&!(singlePositionedSpan&&e===span))||el.querySelector('textPath')||!['x','y'].every(a=>onePosition(el,a))||!['auto','central'].includes(s.dominantBaseline)||s.alignmentBaseline!=='auto'){
          issue('회전·개별 글자 위치·여러 줄 위치가 지정된 SVG 글자는 아직 지원하지 않습니다.',id,nodeId);continue;
        }
        const solidOutline=!spans.length&&s.stroke!=='none'&&!s.stroke.includes('url(')&&parseFloat(s.strokeWidth)>0&&Number(s.strokeOpacity)===1&&Number(s.fillOpacity)===1&&opacity(el)===1&&s.strokeDasharray==='none';
        if([el,...spans].some(item=>{const textPaint=getComputedStyle(item);return textPaint.fill.includes('url(')||textPaint.stroke!=='none'&&!solidOutline||parseFloat(textPaint.wordSpacing)!==0&&textPaint.wordSpacing!=='normal'||!['auto','central'].includes(textPaint.dominantBaseline)||textPaint.alignmentBaseline!=='auto';})){issue('SVG 글자의 그라데이션·외곽선·단어 간격·기준선은 아직 지원하지 않습니다.',id,nodeId);continue;}
        kind='text';style=textStyle(el);const runs=[],tw=document.createTreeWalker(el,NodeFilter.SHOW_TEXT);let child;
        if(stretchedText){const natural=el.cloneNode(true);natural.removeAttribute('textLength');natural.removeAttribute('lengthAdjust');natural.removeAttribute('id');el.after(natural);const width=natural.getComputedTextLength();natural.remove();const wanted=el.getComputedTextLength(),spacing=(s.letterSpacing==='normal'?0:parseFloat(s.letterSpacing))*m.a+(wanted-width)*m.a/(text.length-1);if(!Number.isFinite(spacing)){issue('SVG 글자 간격을 측정할 수 없습니다.',id,nodeId);continue;}style.letterSpacing=`${spacing}px`;issue('글자 폭을 글자 간격으로 근사하여 원본 모양과 다를 수 있습니다.',id,nodeId,'warning','SVG_TEXT_APPROXIMATED');}
        if(scaledText){style.fontSize=parseFloat(s.fontSize)*m.d;const spacing=(s.letterSpacing==='normal'?0:parseFloat(s.letterSpacing))*m.d+(text.length>1?el.getComputedTextLength()*(m.a-m.d)/(text.length-1):0);if(!Number.isFinite(spacing)){issue('SVG 글자 간격을 측정할 수 없습니다.',id,nodeId);continue;}style.letterSpacing=`${spacing}px`;issue('글자 폭을 글자 간격으로 근사하여 원본 모양과 다를 수 있습니다.',id,nodeId,'warning','SVG_TEXT_APPROXIMATED');}
        while((child=tw.nextNode())){let value=child.data.replace(/[\t\r\n ]+/g,' ');if(runs.at(-1)?.text.endsWith(' '))value=value.replace(/^ +/,'');if(value)runs.push({text:value,style:textStyle(child.parentElement)});}
        if(solidOutline){const outline={color:s.stroke,width:parseFloat(s.strokeWidth)*m.a};style.outline=outline;for(const run of runs)run.style.outline=outline;}
        if(stretchedText||scaledText)for(const run of runs){run.style.letterSpacing=style.letterSpacing;if(scaledText)run.style.fontSize=style.fontSize;}
        if(runs.length){runs[0].text=runs[0].text.trimStart();runs.at(-1).text=runs.at(-1).text.trimEnd();}
        for(const run of runs){await fontCheck(run.style.fontFamily);if(run.style.color.includes('url('))issue('SVG 글자의 그라데이션 색은 아직 지원하지 않습니다.',id,nodeId);}
        content={text,runs:runs.filter(r=>r.text)};
        const baseline=point(m,el.getStartPositionOfChar(0).x,el.getStartPositionOfChar(0).y).y;
        measurement={lines:[{text,...b}],baseline,baselineMethod:'SVG getStartPositionOfChar',lineAdvance:null};
      }else{
        const gradient=s.fill.includes('url(')||s.stroke.includes('url('),marker=['markerStart','markerMid','markerEnd'].some(k=>s[k]&&s[k]!=='none');let markerBounds;
        if(marker){
          const ref=s.markerEnd.match(/^url\(["']?#([^)'"\s]+)["']?\)$/)?.[1],definition=ref&&document.getElementById(ref);
          const parts=definition?[definition,...definition.querySelectorAll('*')]:[];
          let slideClip=false,ancestorsSafe=true,rootClipSafe=true;
          if(fragments)for(let parent=root.parentElement;parent&&parent!==slideRoot;parent=parent.parentElement){
            const ps=getComputedStyle(parent),r=parent.getBoundingClientRect(),clipped=ps.overflowX==='hidden'&&ps.overflowY==='hidden';
            if(ps.clipPath!=='none'||ps.maskImage&&ps.maskImage!=='none'||ps.filter!=='none'||ps.clip!=='auto'||ps.contain.includes('paint'))ancestorsSafe=rootClipSafe=false;
            if(clipped){
              const simple=ps.transform==='none'&&!['TopLeft','TopRight','BottomRight','BottomLeft'].some(k=>parseFloat(ps[`border${k}Radius`])!==0)&&!['Top','Right','Bottom','Left'].some(k=>parseFloat(ps[`border${k}Width`])!==0);
              if(!simple||['left','top','right','bottom'].some(k=>Math.abs(r[k]-rr[k])>.01))ancestorsSafe=false;else slideClip=true;
              if(!simple||r.left>bounds.left+.01||r.top>bounds.top+.01||r.right<bounds.right-.01||r.bottom<bounds.bottom-.01)rootClipSafe=false;
            }else if(ps.overflowX!=='visible'||ps.overflowY!=='visible')ancestorsSafe=rootClipSafe=false;
          }
          const rootMatrix=matrix(root),rootClipped=rootStyle.overflowX==='hidden'&&rootStyle.overflowY==='hidden'&&rootClipSafe&&rootStyle.clip==='auto'&&!rootStyle.contain.includes('paint')&&!['TopLeft','TopRight','BottomRight','BottomLeft'].some(k=>parseFloat(rootStyle[`border${k}Radius`])!==0)&&Math.abs(rootMatrix.b)<1e-6&&Math.abs(rootMatrix.c)<1e-6&&rootMatrix.a>0&&rootMatrix.d>0,slideClipped=rootStyle.overflowX==='visible'&&rootStyle.overflowY==='visible'&&slideClip&&ancestorsSafe;
          const safe=fragments&&['polyline','line','path'].includes(tag)&&!el.children.length&&s.markerStart==='none'&&s.markerMid==='none'&&(rootClipped||slideClipped)&&definition?.localName==='marker'&&parts.every(item=>{
            const st=getComputedStyle(item);return (item===definition||item.localName==='path')&&![...item.attributes].some(a=>a.name.startsWith('on')||a.name==='href'||a.name==='xlink:href')&&![st.fill,st.stroke,st.filter,st.clipPath,st.maskImage].some(v=>v&&v.includes('url('))&&st.filter==='none'&&st.clipPath==='none'&&(!st.maskImage||st.maskImage==='none')&&st.mixBlendMode==='normal';
          });
          if(!safe){issue('SVG marker 화살표의 그림 참조 또는 잘림은 아직 보존할 수 없습니다.',id,nodeId);continue;}
          markerBounds=slideClipped?{x:0,y:0,w:width,h:height}:{...svgBox};
        }
        if(['rect','circle','ellipse'].includes(tag)&&uniform&&!gradient&&!marker&&s.strokeDasharray==='none'){
          kind='shape';style={svgGeometry:true,shape:tag==='rect'?'rect':'ellipse',fill:s.fill==='none'?'transparent':s.fill,stroke:s.stroke==='none'?'transparent':s.stroke,strokeWidth:s.stroke==='none'?0:parseFloat(s.strokeWidth)*m.a,radius:tag==='rect'?el.rx.baseVal.value*m.a:0,opacity:opacity(el),fillOpacity:Number(s.fillOpacity),strokeOpacity:Number(s.strokeOpacity)};
        }else if(tag==='line'&&uniform&&!gradient&&!marker&&s.strokeDasharray==='none'){
          kind='line';const start=point(m,el.x1.baseVal.value,el.y1.baseVal.value),end=point(m,el.x2.baseVal.value,el.y2.baseVal.value);
          style={svgGeometry:true,endpoints:{start,end},stroke:s.stroke==='none'?'transparent':s.stroke,strokeWidth:parseFloat(s.strokeWidth)*m.a,opacity:opacity(el)*Number(s.strokeOpacity)};
        }else{
          // Complex paths and gradients remain separate artwork, never a slide
          // screenshot containing text. Text always follows the native path above.
          kind='image';b=marker?markerBounds:box(el,strokePadding);if(b.w<=0||b.h<=0)continue;
          const svg=document.createElementNS('http://www.w3.org/2000/svg','svg');svg.setAttribute('xmlns','http://www.w3.org/2000/svg');svg.setAttribute('width',String(b.w));svg.setAttribute('height',String(b.h));svg.setAttribute('viewBox',`${b.x+v.x} ${b.y+v.y} ${b.w} ${b.h}`);
          const copy=el.cloneNode(true);
          if(fragments){
            // Each source document can define shared gradients outside this SVG.
            // Copy only the references Chromium actually resolves in that document.
            const refs=[],seenRefs=new Set(),collect=element=>{
              for(const item of [element,...element.querySelectorAll('*')])for(const attr of item.attributes){
                for(const match of attr.value.matchAll(/url\(["']?#([^)'"\s]+)["']?\)/g))refs.push(match[1]);
                if((attr.name==='href'||attr.name==='xlink:href')&&attr.value.startsWith('#'))refs.push(attr.value.slice(1));
              }
            };
            inlineComputed(el,copy);collect(copy);
            const defs=document.createElementNS(svg.namespaceURI,'defs');
            while(refs.length){
              const refId=refs.shift();if(seenRefs.has(refId))continue;seenRefs.add(refId);
              const original=document.getElementById(refId);
              if(!original||original.namespaceURI!==svg.namespaceURI){issue('SVG 그림의 참조 정의를 찾을 수 없습니다.',id,nodeId);continue;}
              if(['text','tspan','textPath'].includes(original.localName)||original.querySelector('text,tspan,textPath')){issue('SVG 그림의 참조 정의 안에 있는 글자는 아직 편집 가능한 객체로 변환할 수 없습니다.',id,nodeId);continue;}
              const clone=original.cloneNode(true);inlineComputed(original,clone);collect(clone);defs.append(clone);
            }
            if(defs.children.length)svg.append(defs);
          }else for(const defs of root.querySelectorAll('defs'))svg.append(defs.cloneNode(true));
          copy.removeAttribute('transform');copy.style.setProperty('transform','none');copy.style.setProperty('opacity',String(opacity(el)));const g=document.createElementNS(svg.namespaceURI,'g');g.setAttribute('transform',`matrix(${m.a} ${m.b} ${m.c} ${m.d} ${m.e} ${m.f})`);g.append(copy);svg.append(g);
          // CSSOM quotes url(#id); XMLSerializer entity-escapes those quotes.
          // Use the equivalent unquoted local fragment accepted by asset checks.
          const markup=new XMLSerializer().serializeToString(svg).replace(/url\(&quot;(#[^&<>"\s)]+)&quot;\)/g,'url($1)');
          if(fragments){const bytes=new TextEncoder().encode(markup).length;if(bytes>16*1024*1024||(artworkBytes+=bytes)>64*1024*1024){issue('SVG 그림 자산의 크기 한도를 초과했습니다.',id,nodeId,'error','LIMIT_EXCEEDED');break;}}
          content={inlineSvg:markup,artworkOnly:true};style={opacity:1};
        }
      }
      if(kind!=='line'&&(b.w<=0||b.h<=0))continue;
      nodes.push({id:nodeId,kind,boxCssPx:b,zOrder:nodes.length,content,style,sourceRef,parentId:groupFor(el)});
      objects.push({nodeId,kind,boxCssPx:b,...measurement});
    }
    for(const [g,gid] of groupIds){const children=nodes.filter(n=>n.parentId===gid);if(!children.length)continue;const x=Math.min(...children.map(n=>n.boxCssPx.x)),y=Math.min(...children.map(n=>n.boxCssPx.y));nodes.push({id:gid,kind:'group',boxCssPx:{x,y,w:Math.max(1,Math.max(...children.map(n=>n.boxCssPx.x+n.boxCssPx.w))-x),h:Math.max(1,Math.max(...children.map(n=>n.boxCssPx.y+n.boxCssPx.h))-y)},zOrder:children[0].zOrder,content:{},style:{},sourceRef:{slideId:id,path:g.dataset.sourcePath},parentId:null});if(g.querySelector('g'))issue('중첩 SVG 그룹은 상위 그룹 한 개로 묶었습니다.',id,gid,'warning');}
    slides.push({id,...(fragments?{svgPath:root.dataset.sourcePath}:{}),widthCssPx:width,heightCssPx:height,background:{color:'#ffffff',image:'none'},nodes});
    measurements.push({slideId:id,referenceSize:{width:Math.round(width*renderScale),height:Math.round(height*renderScale)},renderScale,sourceTransformScale:1,layoutStable:false,fontsReady:[...fonts.values()].every(f=>f.available),assetsReady:true,objects});
  }
  const layoutSnapshot=[...document.querySelectorAll('[data-measured-svg], [data-measured-svg] *')].map(e=>{const r=e.getBoundingClientRect();return [r.x,r.y,r.width,r.height];});
  return {slides,measurements,diagnostics,layoutSnapshot,fonts:[...fonts.values()]};
}

export async function measureSvgFragments(session,{selector,timeoutMs,slots}){
  const paths=await session.page.evaluate(paths=>[...document.querySelectorAll('svg')].filter(el=>paths.includes(el.dataset.sourcePath)&&el.querySelector('text,image')).map(el=>el.dataset.sourcePath),slots.map(s=>s.path));
  if(!paths.length)return {fragments:[],diagnostics:[],fonts:[]};
  const measured=await session.page.evaluate(measureSvgPage,{paths,timeoutMs,fragments:true,selector,slots});
  return {fragments:measured.slides.map((slide,i)=>({slideId:slide.id,svgPath:slide.svgPath,nodes:slide.nodes,objects:measured.measurements[i].objects})),diagnostics:measured.diagnostics,fonts:measured.fonts};
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
