import {EMU_PER_CSS_PX} from './canvas.mjs';

const elements=n=>[...(n?.childNodes??[])].filter(c=>c.nodeType===1);
const child=(n,name)=>elements(n).find(c=>c.localName===name);
const all=(n,name)=>n?[...n.getElementsByTagNameNS('*',name)]:[];
const attr=(n,k,fallback=null)=>n?.hasAttribute(k)?n.getAttribute(k):fallback;
const px=n=>n/EMU_PER_CSS_PX;
const P='http://schemas.openxmlformats.org/presentationml/2006/main',R='http://schemas.openxmlformats.org/officeDocument/2006/relationships',A='http://schemas.openxmlformats.org/drawingml/2006/main';
const num=(n,k,fallback=0)=>Number(attr(n,k,fallback));

// Prepare a slide and its inherited theme with shared package lookup state.
export function preparePptxSlide(ref,index,{presRels,xml,relationships,relationRecords,parts,model}){
  const file=presRels.get(attr(ref,'r:id')),doc=xml(file),rels=relationships(file),tree=all(doc,'spTree')[0];
  const slide={id:`slide-${index+1}`,widthCssPx:px(model.pptxSizeEmu.width),heightCssPx:px(model.pptxSizeEmu.height),background:{color:'#ffffff',image:'none'},nodes:[]};model.slides.push(slide);
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
  return {file,doc,rels,tree,slide,inherited,masterPart,themeDoc,masterDoc,colorMap,colorScheme};
}

export function preparePlaceholderContext({inherited,xml,masterDoc}){
  const layoutDocs=[...inherited].filter(p=>p.startsWith('ppt/slideLayouts/')).map(p=>xml(p));
  const placeholder=sp=>{
    const nv=child(sp,'nvSpPr'),props=child(nv,'nvPr'),ph=child(props,'ph');
    return [sp,nv,props,ph].every(n=>n?.namespaceURI===P)?ph:null;
  };
  const placeholderMatches=(container,predicate)=>all(container,'sp').filter(sp=>{const p=placeholder(sp),parent=sp.parentNode;return parent?.namespaceURI===P&&parent.localName==='spTree'&&p&&predicate(p);});
  const placeholderContext=owner=>{
    const ph=placeholder(owner),index=attr(ph,'idx','0');
    const layoutMatches=ph?layoutDocs.flatMap(doc=>placeholderMatches(doc,p=>attr(p,'idx','0')===index)):[];
    const matchedLayout=layoutMatches.length===1?layoutMatches[0]:null;
    const phType=attr(placeholder(matchedLayout),'type',attr(ph,'type','obj'));
    let masterMatches=ph?placeholderMatches(masterDoc,p=>attr(p,'type','obj')===phType):[];
    if(ph&&phType==='ctrTitle'&&!masterMatches.length)masterMatches=placeholderMatches(masterDoc,p=>attr(p,'type','obj')==='title');
    return {ph,matchedLayout,phType,masterShape:masterMatches.length===1?masterMatches[0]:null,ambiguousLayout:layoutMatches.length>1};
  };
  return {layoutDocs,placeholderContext};
}

export function visitPptxSlideLayers({doc,inherited,xml,masterPart,masterDoc,tree},visit){
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
}

// Resolve only solid, unmodified theme backgrounds. More complex theme fills
// stay unsupported rather than being replaced with a guessed color.
export function themeBackground(ref,theme,master,overrides){
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
