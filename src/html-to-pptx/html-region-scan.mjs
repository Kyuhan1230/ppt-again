// This function is serialized by Playwright and runs entirely in the page.
// It does not mutate the source DOM or return text, selectors or paths.
export function scanHtmlRegions({selector,svgPaths=[],deckPaths=[],maxRegions=500,maxRelations=50000,onlyExplicit=false}={}){
  const fail=()=>{throw new Error('INPUT_LIMIT');};
  const body=document.body,viewport={width:innerWidth,height:innerHeight};
  const all=[...body.querySelectorAll('*')];if(all.length>100000)fail();
  const structural=new Set(['main','article','section','div','aside','header','footer','nav','figure','table','ul','ol','li']);
  const textBlocks=new Set(['p','h1','h2','h3','h4','h5','h6','pre','blockquote']);
  const ignored=new Set(['script','style','link','meta','template','noscript']);
  let explicit=[];
  if(selector){try{explicit=[...document.querySelectorAll(selector)];}catch{throw new Error('INPUT_INVALID');}}
  const explicitSet=new Set(explicit),allSet=new Set(all);
  if(explicit.some(el=>el!==body&&el!==document.documentElement&&!allSet.has(el)))throw new Error('INPUT_INVALID');
  const declared=new Set(document.querySelectorAll('[data-slide]'));
  const validatedSvg=new Set(all.filter(el=>el.localName==='svg'&&svgPaths.includes(el.dataset.sourcePath)));
  const stats=new WeakMap(),state=new WeakMap(),children=new Map(),uncertainties=new Set();
  const sources=[document.documentElement,body,...all];
  const domIndex=new WeakMap(sources.map((el,i)=>[el,i])),subtreeEnd=sources.map((_,i)=>i),footprints=[],complexPaint=[];
  for(let i=sources.length-1;i>0;i--){const parent=domIndex.get(sources[i].parentElement);if(parent!==undefined)subtreeEnd[parent]=Math.max(subtreeEnd[parent],subtreeEnd[i]);}
  const clippedRect=(r,c)=>{const left=Math.max(r.left,c.left),right=Math.min(r.right,c.right),top=Math.max(r.top,c.top),bottom=Math.min(r.bottom,c.bottom);return right>left&&bottom>top?{left,top,right,bottom}:null;};
  const transparent=color=>color==='transparent'||/^rgba\([^)]*[, /]\s*0(?:\.0+)?\s*\)$/.test(color);
  const sides=['Top','Right','Bottom','Left'];
  const contentBoxExpands=s=>s.boxSizing==='content-box'&&sides.some(side=>parseFloat(s['padding'+side])>0||parseFloat(s['border'+side+'Width'])>0);
  const square=s=>['TopLeft','TopRight','BottomRight','BottomLeft'].every(side=>parseFloat(s['border'+side+'Radius'])===0);
  const plainClip=s=>[s.overflowX,s.overflowY].every(value=>['hidden','clip','visible'].includes(value))&&s.overflowClipMargin==='0px'&&square(s);
  const baseGradient=/(?:^|,\s*)linear-gradient\((?:[\d.]+deg, )?rgb\(\d+, \d+, \d+\)(?: [\d.]+%)?(?:, rgb\(\d+, \d+, \d+\)(?: [\d.]+%)?)+\)$/;
  const opaqueBackground=s=>{
    const solid=/^rgb\(/.test(s.backgroundColor);
    const gradient=baseGradient.test(s.backgroundImage)&&s.backgroundSize.split(',').every(v=>v.trim()==='auto')&&
      s.backgroundPosition.split(',').every(v=>v.trim()==='0% 0%')&&s.backgroundAttachment.split(',').every(v=>v.trim()==='scroll')&&
      s.backgroundOrigin.split(',').every(v=>['padding-box','border-box'].includes(v.trim()))&&
      sides.every(side=>parseFloat(s['border'+side+'Width'])===0);
    return (solid||gradient)&&s.backgroundClip.split(',').every(value=>value.trim()==='border-box')&&square(s)&&s.opacity==='1'&&
      s.filter==='none'&&s.maskImage==='none'&&s.clipPath==='none'&&s.mixBlendMode==='normal';
  };
  const pseudoBackground=(el,s,pseudos)=>{
    // A positioned, clipped backdrop with all real children painted above it.
    // Copy computed paint only, never generated text or the original subtree.
    if(!pseudos.length||s.display!=='block'||s.position==='static'||s.transform!=='none'||!opaqueBackground(s)||
      !plainClip(s)||!['hidden','clip'].includes(s.overflowX)||!['hidden','clip'].includes(s.overflowY)||el.scrollLeft!==0||el.scrollTop!==0||
      [...el.childNodes].some(n=>n.nodeType===Node.TEXT_NODE&&n.data.trim())||
      [...el.children].some(child=>{const c=getComputedStyle(child);return c.display!=='none'&&(c.display==='contents'||c.position==='static'||!(Number(c.zIndex)>0));}))return null;
    const pixel=value=>/^-?\d+(?:\.\d+)?px$/.test(value);
    const props=['top','left','width','height','box-sizing','background-color','border-radius',...sides.flatMap(side=>['width','style','color'].map(p=>'border-'+side.toLowerCase()+'-'+p))];
    if(pseudos.some(([,p])=>p.content!=='""'||p.position!=='absolute'||p.display!=='block'||p.visibility!=='visible'||
      ![p.width,p.height].every(v=>pixel(v)&&parseFloat(v)>0&&parseFloat(v)<=8192)||
      ![p.top,p.left].every(v=>pixel(v)&&Math.abs(parseFloat(v))<=8192)||
      p.opacity!=='1'||p.zIndex!=='auto'||p.transform!=='none'||p.translate!=='none'||p.rotate!=='none'||p.scale!=='none'||
      p.filter!=='none'||p.backdropFilter!=='none'||p.maskImage!=='none'||p.clipPath!=='none'||p.mixBlendMode!=='normal'||
      p.boxShadow!=='none'||p.outlineStyle!=='none'||p.borderImageSource!=='none'||p.clip!=='auto'||p.backgroundImage!=='none'||p.backgroundClip!=='border-box'||
      p.animationName!=='none'||p.transitionDuration.split(',').some(v=>parseFloat(v)>0)||
      sides.some(side=>parseFloat(p['padding'+side])||parseFloat(p['margin'+side])||!['none','solid'].includes(p['border'+side+'Style']))))return null;
    return pseudos.map(([,p])=>Object.fromEntries(props.map(key=>[key,p.getPropertyValue(key)])));
  };
  let totalText=0,textRectBudget=0;
  for(const el of sources){
    const s=getComputedStyle(el),parent=el.parentElement,ps=parent&&state.get(parent);
    const blocked=Boolean(ps?.blocked||s.display==='none'||Number(s.opacity)===0||s.contentVisibility==='hidden');
    // visibility may be explicitly restored by a descendant; display and
    // opacity cannot be restored inside an invisible ancestor.
    const hidden=blocked||s.visibility!=='visible';
    const documentRoot=el===body||el===document.documentElement;
    const clips=value=>['hidden','clip'].includes(value)||(!documentRoot&&['auto','scroll'].includes(value));
    const clip=Boolean(ps?.clip||clips(s.overflowX)||clips(s.overflowY));
    let clipBounds=ps?.clipBounds??{left:-Infinity,top:-Infinity,right:Infinity,bottom:Infinity};
    if(clips(s.overflowX)||clips(s.overflowY)){
      const r=el.getBoundingClientRect(),sx=el.offsetWidth?r.width/el.offsetWidth:1,sy=el.offsetHeight?r.height/el.offsetHeight:1,x=r.left+el.clientLeft*sx,y=r.top+el.clientTop*sy;
      clipBounds={left:clips(s.overflowX)?Math.max(clipBounds.left,x):clipBounds.left,right:clips(s.overflowX)?Math.min(clipBounds.right,x+el.clientWidth*sx):clipBounds.right,top:clips(s.overflowY)?Math.max(clipBounds.top,y):clipBounds.top,bottom:clips(s.overflowY)?Math.min(clipBounds.bottom,y+el.clientHeight*sy):clipBounds.bottom};
    }
    const inSvg=Boolean(ps?.inSvg||el.localName==='svg');
    state.set(el,{hidden,blocked,clip,clipBounds,inSvg,withinDeclared:declared.has(el)||Boolean(ps?.withinDeclared),display:s.display,fixedDescendant:false,position:s.position});
    if(hidden)uncertainties.add('hidden-content-not-observed');
    if(clip)uncertainties.add('clipped-or-scroll-content-not-observed');
    if(s.clipPath!=='none'||s.maskImage!=='none'||s.filter!=='none'||s.mixBlendMode!=='normal'||Number(s.opacity)<1)uncertainties.add('css-compositing-not-fully-observed');
    if(s.boxShadow!=='none'||s.outlineStyle!=='none'&&parseFloat(s.outlineWidth)>0)uncertainties.add('css-compositing-not-fully-observed');
    if(s.animationName!=='none'||s.transitionDuration.split(',').some(x=>parseFloat(x)>0))uncertainties.add('layout-may-change');
    const pseudos=['::before','::after'].map(name=>[name,getComputedStyle(el,name)]).filter(([,p])=>p.content!=='none'&&p.content!=='normal'&&p.display!=='none');
    const hasPseudo=pseudos.length>0;
    if(hasPseudo)uncertainties.add('generated-content-not-fully-observed');
    const rootDecorations=hasPseudo?pseudoBackground(el,s,pseudos):null;
    state.get(el).rootDecorations=rootDecorations;
    const coloredBorder=['Top','Right','Bottom','Left'].some(side=>parseFloat(s[`border${side}Width`])>0&&s[`border${side}Style`]!=='none'&&!transparent(s[`border${side}Color`]));
    const ownBackground=s.backgroundImage!=='none'||!transparent(s.backgroundColor)&&s.backgroundColor!=='rgb(255, 255, 255)';
    const ownEffects=Number(s.opacity)<1||s.transform!=='none'||s.translate!=='none'||s.rotate!=='none'||s.scale!=='none'||s.perspective!=='none'||
      s.filter!=='none'||s.backdropFilter!=='none'||s.maskImage!=='none'||s.clipPath!=='none'||s.mixBlendMode!=='normal'||
      s.boxShadow!=='none'||s.outlineStyle!=='none'&&parseFloat(s.outlineWidth)>0||coloredBorder||hasPseudo||
      (clips(s.overflowX)||clips(s.overflowY))&&(!plainClip(s)||el.scrollLeft!==0||el.scrollTop!==0);
    state.get(el).ancestorEffects=Boolean(ps?.ancestorEffects||ownEffects);
    state.get(el).ancestorBackground=Boolean(ps?.ancestorBackground||ownBackground);
    const paintedBackground=s.backgroundImage!=='none'||!transparent(s.backgroundColor);
    const rounded=['borderTopLeftRadius','borderTopRightRadius','borderBottomRightRadius','borderBottomLeftRadius'].some(key=>parseFloat(s[key])!==0);
    const clippedBackground=!s.backgroundClip.split(',').every(value=>value.trim()==='border-box');
    state.get(el).rootAppearance=Boolean(paintedBackground&&(rounded||clippedBackground)||coloredBorder||hasPseudo&&!rootDecorations||Number(s.opacity)!==1||s.filter!=='none'||s.maskImage!=='none'||s.clipPath!=='none'||s.mixBlendMode!=='normal'||s.boxShadow!=='none'||s.outlineStyle!=='none'&&parseFloat(s.outlineWidth)>0);
    if(!hidden&&(hasPseudo||s.boxShadow!=='none'||s.outlineStyle!=='none'&&parseFloat(s.outlineWidth)>0||s.filter!=='none'||s.maskImage!=='none'||s.clipPath!=='none'||s.mixBlendMode!=='normal'))complexPaint.push(domIndex.get(el));
    if(ignored.has(el.localName)||inSvg&&el.localName!=='svg'){stats.set(el,{count:0,text:0,media:0,paint:0});continue;}
    const textVisible=!transparent(s.webkitTextFillColor)||parseFloat(s.webkitTextStrokeWidth)>0&&!transparent(s.webkitTextStrokeColor)||s.textShadow!=='none';
    let text=0;for(const n of el.childNodes)if(n.nodeType===Node.TEXT_NODE&&n.data.trim()){
      totalText+=n.data.length;if(n.data.length>100000||totalText>2000000)fail();
      if(textVisible&&!hidden){const range=document.createRange();range.selectNodeContents(n);const rects=range.getClientRects();if((textRectBudget+=rects.length)>50000)fail();const visible=[...rects].map(r=>clippedRect(r,clipBounds)).filter(Boolean);if(visible.length){text++;footprints.push({owner:domIndex.get(el),kind:'text',rects:visible});}}
    }
    const ownRect=el.getBoundingClientRect(),parentClip=ps?.clipBounds??{left:-Infinity,top:-Infinity,right:Infinity,bottom:Infinity};
    state.get(el).area=ownRect.width*ownRect.height;
    state.get(el).offscreen=ownRect.right<=0||ownRect.bottom<=0||ownRect.left>=innerWidth||ownRect.top>=innerHeight;
    const boxVisible=ownRect.width>0&&ownRect.height>0&&ownRect.right>parentClip.left&&ownRect.left<parentClip.right&&ownRect.bottom>parentClip.top&&ownRect.top<parentClip.bottom;
    const media=boxVisible&&el.matches('img,svg,video,canvas,iframe,object,embed')?1:0;
    const paint=boxVisible&&(s.backgroundImage!=='none'||!transparent(s.backgroundColor)&&!(el===body&&s.backgroundColor==='rgb(255, 255, 255)')||['Top','Right','Bottom','Left'].some(side=>parseFloat(s[`border${side}Width`])>0&&s[`border${side}Style`]!=='none'&&!transparent(s[`border${side}Color`])))?1:0;
    if(!hidden&&(media||paint))footprints.push({owner:domIndex.get(el),kind:'paint',rects:[clippedRect(ownRect,parentClip)]});
    state.get(el).neutral=!(text||media||paint||hasPseudo)&&s.opacity==='1'&&s.transform==='none'&&s.filter==='none'&&s.maskImage==='none'&&s.clipPath==='none'&&s.mixBlendMode==='normal'&&s.boxShadow==='none'&&s.outlineStyle==='none'&&s.overflowX==='visible'&&s.overflowY==='visible';
    stats.set(el,{count:hidden?0:text+media+paint,text:hidden?0:text,media:hidden?0:media,paint:hidden?0:paint});
    if(parent){if(!children.has(parent))children.set(parent,[]);children.get(parent).push(el);}
  }
  for(let i=sources.length-1;i>0;i--){
    const el=sources[i],p=el.parentElement,a=stats.get(el),b=stats.get(p),child=state.get(el);
    if(child.fixedDescendant||child.position==='fixed')state.get(p).fixedDescendant=true;
    if(a&&b){b.count+=a.count;b.text+=a.text;b.media+=a.media;b.paint+=a.paint;}
  }
  const selected=[];
  for(const el of all){
    if(ignored.has(el.localName)||state.get(el).inSvg&&el.localName!=='svg')continue;
    const st=state.get(el),stat=stats.get(el),isMarked=explicitSet.has(el)||declared.has(el)||validatedSvg.has(el);
    const significantText=textBlocks.has(el.localName)&&st.area>=innerWidth*innerHeight*.05;
    // Descendants of a declared slide are objects in a known slide, not
    // additional guessed slides. An explicit selector or declaration still
    // creates a candidate inside that root.
    const declaredObject=st.withinDeclared&&!isMarked;
    const candidate=onlyExplicit?explicitSet.has(el):isMarked||!declaredObject&&(((structural.has(el.localName)||el.parentElement===body||significantText)&&(stat.count||st.hidden||st.display==='contents'))||(!st.hidden&&(el.matches('img,svg,video,canvas')||stat.paint&&(children.get(el)?.length??0)===0)));
    if(candidate){selected.push(el);if(selected.length>maxRegions)fail();}
    else if(stat.count&&(declaredObject||textBlocks.has(el.localName)))uncertainties.add(declaredObject?'objects-inside-declared-slide-not-separate-candidates':'small-elements-not-proposed');
  }
  if(explicitSet.has(body))selected.unshift(body);
  if(explicitSet.has(document.documentElement))selected.unshift(document.documentElement);
  if(selected.length>maxRegions)fail();
  const index=new Map(selected.map((el,i)=>[el,i]));
  const childNumber=new WeakMap(),childCount=new WeakMap();
  for(const el of sources){const p=el.parentElement;if(p){const n=(childCount.get(p)??0)+1;childCount.set(p,n);childNumber.set(el,p===document.documentElement?[...p.children].indexOf(el)+1:n);}}
  let selectorWork=0;
  const internalSelector=el=>{const parts=[];for(let p=el;p;p=p.parentElement){if(++selectorWork>200000)fail();parts.unshift(CSS.escape(p.localName)+':nth-child('+(childNumber.get(p)??1)+')');}return parts.join(' > ');};
  const nearest=new WeakMap(),parentIndex=new WeakMap();
  for(const el of sources){const parent=el.parentElement?nearest.get(el.parentElement)??null:null;parentIndex.set(el,parent);nearest.set(el,index.has(el)?index.get(el):parent);}
  const raw=selected.map((el,i)=>{
    const s=getComputedStyle(el),st=state.get(el),stat=stats.get(el),r=el.getBoundingClientRect();
    const clip=state.get(el.parentElement)?.clipBounds??{left:-Infinity,top:-Infinity,right:Infinity,bottom:Infinity};
    const clipped=r.right<=clip.left||r.left>=clip.right||r.bottom<=clip.top||r.top>=clip.bottom;
    const box=!st.hidden&&!clipped&&s.display!=='contents'&&r.width>0&&r.height>0?{x:r.x+scrollX,y:r.y+scrollY,width:r.width,height:r.height}:null;
    const visibility=st.hidden||clipped?'hidden':!box?'unmeasured':r.right<=0||r.bottom<=0||r.left>=innerWidth||r.top>=innerHeight?'offscreen':'visible';
    const evidence=[];if(explicitSet.has(el))evidence.push('explicit-selector');if(declared.has(el))evidence.push('declared-slide');if(deckPaths.includes(el.dataset.sourcePath))evidence.push('declared-deck-frame');if(validatedSvg.has(el))evidence.push('validated-svg');if(stat.text||stat.media)evidence.push('visible-content');if(stat.paint)evidence.push('background');if(st.hidden)evidence.push('hidden');if(s.display==='contents')evidence.push('display-contents');
    if(st.clip)uncertainties.add('clipped-or-scroll-content-not-observed');
    const features=[];let layoutComparable=true;
    if(box){
      const localPoint=rect=>[Math.round((rect.left-r.left)/r.width*4),Math.round((rect.top-r.top)/r.height*4)];
      const relevant=(children.get(el)??[]).filter(c=>!state.get(c).hidden&&stats.get(c).count>0);
      if(relevant.length>32){uncertainties.add('repeated-layout-not-fully-observed');layoutComparable=false;}
      else{
        for(const child of relevant){const q=stats.get(child),cr=child.getBoundingClientRect();if(cr.width&&cr.height)features.push([Boolean(q.text),Boolean(q.media),Boolean(q.paint),...localPoint(cr)]);}
        for(const node of el.childNodes)if(node.nodeType===Node.TEXT_NODE&&node.data.trim()){const range=document.createRange();range.selectNodeContents(node);const rect=range.getClientRects()[0];if(rect)features.push(['text',...localPoint(rect)]);}
      }
    }
    const signature=box&&layoutComparable&&features.length<=32?JSON.stringify(features):null;
    const ancestor=state.get(el.parentElement);
    const ownClip=ancestor?.clip&&s.display==='block'&&plainClip(s)&&['hidden','clip'].includes(s.overflowX)&&['hidden','clip'].includes(s.overflowY)&&
      s.position!=='static'&&el.scrollLeft===0&&el.scrollTop===0&&r.left>=clip.left&&r.top>=clip.top&&r.right<=clip.right&&r.bottom<=clip.bottom&&
      !st.fixedDescendant;
    const ancestorAppearance=Boolean(ancestor?.ancestorEffects||ancestor?.clip&&!ownClip||ancestor?.ancestorBackground&&!opaqueBackground(s));
    const specifiedWidth=el.computedStyleMap?.().get('width')?.toString(),specifiedHeight=el.computedStyleMap?.().get('height')?.toString();
    const fixed=Boolean(box&&/^\d+(?:\.\d+)?px$/.test(specifiedWidth??'')&&/^\d+(?:\.\d+)?px$/.test(specifiedHeight??'')&&s.transform==='none');
    const rootSizeMismatch=contentBoxExpands(s);
    const centeredFrame=fixed&&ancestor?.clip&&getComputedStyle(el.parentElement).position==='fixed'&&
      Math.abs(r.left+r.right-innerWidth)<.01&&Math.abs(r.top+r.bottom-innerHeight)<.01&&
      (r.left<0||r.top<0)&&r.width>=300&&r.height>=200&&Number.isInteger(r.width)&&Number.isInteger(r.height);
    return {index:i,selector:internalSelector(el),fixed,centeredFrame,transformed:s.transform!=='none',rootSizeMismatch,ancestorAppearance,rootAppearance:st.rootAppearance,rootDecorations:st.rootDecorations,sourcePath:el.dataset.sourcePath,parent:parentIndex.get(el),directParent:index.has(el.parentElement),neutralWrapper:st.neutral&&(children.get(el)?.length??0)===1,bounds:box,visibility,evidence,counts:{...stat},overlaps:[],tag:el.localName,signature};
  });
  const ancestors=raw.map(r=>{const found=new Set();for(let p=r.parent;p!==null;p=raw[p].parent)found.add(p);return found;});
  const isAncestor=(a,b)=>ancestors[b].has(a);
  let relations=0;
  for(let i=0;i<raw.length;i++)for(let j=i+1;j<raw.length;j++){
    const a=raw[i],b=raw[j];if(!a.bounds||!b.bounds||isAncestor(i,j)||isAncestor(j,i))continue;
    if(a.bounds.x<b.bounds.x+b.bounds.width&&b.bounds.x<a.bounds.x+a.bounds.width&&a.bounds.y<b.bounds.y+b.bounds.height&&b.bounds.y<a.bounds.y+a.bounds.height){
      if(++relations>maxRelations)fail();a.overlaps.push(j);b.overlaps.push(i);uncertainties.add('overlapping-elements');
    }
  }
  const groups=[];
  for(const siblings of children.values()){
    const grouped=new Map();
    for(const el of siblings){const i=index.get(el);if(i===undefined||state.get(el).hidden||raw[i].signature===null)continue;const q=stats.get(el);const key=JSON.stringify([el.localName,Boolean(q.text),Boolean(q.media),Boolean(q.paint),raw[i].signature]);if(!grouped.has(key))grouped.set(key,[]);grouped.get(key).push(i);}
    for(const indexes of grouped.values())if(indexes.length>1)groups.push(indexes);
  }
  if(groups.length>maxRegions)fail();
  if((raw.length+groups.length)*footprints.length>2000000)fail();
  const insideCandidate=(footprint,i)=>{
    const root=selected[i],start=domIndex.get(root),end=subtreeEnd[start],b=raw[i].bounds;
    if(!b||footprint.owner<start||footprint.owner>end)return false;
    const fullyInside=footprint.rects.every(rect=>rect.left+scrollX>=b.x-.01&&rect.top+scrollY>=b.y-.01&&rect.right+scrollX<=b.x+b.width+.01&&rect.bottom+scrollY<=b.y+b.height+.01);
    if(!fullyInside)uncertainties.add('descendant-extends-outside-candidate');
    return fullyInside;
  };
  if(raw.length*complexPaint.length>2000000)fail();
  for(const region of raw){
    region.outsideCount=footprints.filter(f=>!insideCandidate(f,region.index)).length;
    const root=selected[region.index],start=domIndex.get(root),end=subtreeEnd[start],box=region.bounds;
    const inRoot=owner=>owner>=start&&owner<=end;
    const isAncestorOwner=owner=>owner<start&&subtreeEnd[owner]>=start;
    region.externalOverlap=Boolean(box&&footprints.some(f=>!inRoot(f.owner)&&(!isAncestorOwner(f.owner)||f.kind==='text')&&f.rects.some(r=>r.right+scrollX>box.x&&r.left+scrollX<box.x+box.width&&r.bottom+scrollY>box.y&&r.top+scrollY<box.y+box.height)));
    region.externalComplex=complexPaint.some(owner=>!inRoot(owner)&&!isAncestorOwner(owner));
  }
  const groupOutsideCounts=groups.map(group=>footprints.filter(f=>!group.some(i=>insideCandidate(f,i))).length);
  const textRects=[];let rectCount=0,minX=0,minY=0,maxX=Math.max(document.documentElement.scrollWidth,body.scrollWidth,innerWidth),maxY=Math.max(document.documentElement.scrollHeight,body.scrollHeight,innerHeight);
  for(const r of raw)if(r.bounds){minX=Math.min(minX,r.bounds.x);minY=Math.min(minY,r.bounds.y);maxX=Math.max(maxX,r.bounds.x+r.bounds.width);maxY=Math.max(maxY,r.bounds.y+r.bounds.height);}
  const walker=document.createTreeWalker(body,NodeFilter.SHOW_TEXT);let node;
  while((node=walker.nextNode())){
    if(!node.data.trim()||ignored.has(node.parentElement?.localName))continue;
    const range=document.createRange();range.selectNodeContents(node);
    const rects=range.getClientRects();if((rectCount+=rects.length)>50000)fail();
    textRects.push([...rects].map(r=>{const x=r.x+scrollX,y=r.y+scrollY;if(!state.get(node.parentElement)?.hidden){minX=Math.min(minX,x);minY=Math.min(minY,y);maxX=Math.max(maxX,x+r.width);maxY=Math.max(maxY,y+r.height);}return [x,y,r.width,r.height];}));
  }
  const images=[...document.images].map(img=>[img.complete,img.naturalWidth,img.naturalHeight]);
  const fonts=[...document.fonts].map(f=>[f.family,f.style,f.weight,f.status]).sort((a,b)=>JSON.stringify(a).localeCompare(JSON.stringify(b)));
  const bodyStyle=getComputedStyle(body),bodyWidth=body.computedStyleMap?.().get('width')?.toString(),bodyHeight=body.computedStyleMap?.().get('height')?.toString(),bodyRect=body.getBoundingClientRect();
  const documentOpaque=opaqueBackground(bodyStyle);
  const documentAncestor=state.get(body.parentElement);
  const documentRoot={selector:internalSelector(body),rootSizeMismatch:contentBoxExpands(bodyStyle),fixed:/^\d+(?:\.\d+)?px$/.test(bodyWidth??'')&&/^\d+(?:\.\d+)?px$/.test(bodyHeight??'')&&bodyStyle.transform==='none',transformed:bodyStyle.transform!=='none',rootAppearance:state.get(body).rootAppearance,ancestorAppearance:Boolean(documentAncestor?.ancestorEffects||documentAncestor?.clip||documentAncestor?.ancestorBackground&&!documentOpaque),bounds:{x:bodyRect.x+scrollX,y:bodyRect.y+scrollY,width:bodyRect.width,height:bodyRect.height}};
  return {viewport,regions:raw,groups,groupOutsideCounts,documentRoot,documentBounds:{x:minX,y:minY,width:maxX-minX,height:maxY-minY},documentCount:stats.get(body).count,uncertainties:[...uncertainties].sort(),fingerprint:{textRects,images,fonts}};
}
