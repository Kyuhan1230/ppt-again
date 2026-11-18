const declaredSlides='[data-slide]';

// Inspect only explicitly declared absolute siblings in one fixed frame.
// A hidden tab, template, or unrelated element is not a deck frame.
export async function findSlideDeck(page,selector=declaredSlides){
  if(selector!==declaredSlides)return null;
  const deck=await page.evaluate(()=>{
    const roots=[...document.querySelectorAll('[data-slide]')];
    if(roots.length<2||roots.length>200)return null;
    const parent=roots[0].parentElement;
    const fixed=(el,key)=>/^\d+(?:\.\d+)?px$/.test(el.computedStyleMap().get(key)?.toString());
    if(getComputedStyle(parent).position!=='relative'||!fixed(parent,'width')||!fixed(parent,'height'))return null;
    for(let el=parent;el;el=el.parentElement){
      const s=getComputedStyle(el);
      if(s.display==='none'||s.contentVisibility==='hidden'||s.visibility!=='visible'||s.opacity!=='1'||
         s.transform!=='none'||s.perspective!=='none'||s.filter!=='none'||s.clipPath!=='none'||s.maskImage!=='none')return null;
    }
    const frames=roots.map((root,index)=>{
      const s=getComputedStyle(root),path=root.dataset.sourcePath;
      if(!path||root.parentElement!==parent||root.localName!==roots[0].localName||
         root.closest('[hidden],[aria-hidden="true"],[inert]')||s.visibility!=='visible'||
         s.position!=='absolute'||s.opacity!=='1'||s.transform!=='none'||s.contentVisibility!=='visible'||
         s.animationName!=='none'||s.transitionDuration.split(',').some(value=>parseFloat(value)>0)||
         !fixed(root,'width')||!fixed(root,'height')||!['block','none'].includes(s.display))return null;
      return {index,path,selector:'[data-source-path='+JSON.stringify(path)+']',hidden:s.display==='none'};
    });
    if(frames.includes(null)||new Set(frames.map(f=>f.path)).size!==frames.length)return null;
    const visible=frames.filter(f=>!f.hidden);
    if(visible.length!==1)return null;
    const bounds=el=>{const r=el.getBoundingClientRect();return [r.x,r.y,r.width,r.height];};
    const frameBounds=bounds(roots[visible[0].index]);
    if(frameBounds[2]<=0||frameBounds[3]<=0)return null;
    return {active:visible[0],frames,frameBounds,parentBounds:bounds(parent)};
  });
  if(!deck)return null;
  for(const frame of deck.frames){
    if(!frame.hidden)continue;
    const matched=await withSlideFrame(page,deck,frame,()=>true);
    if(!matched)return null;
  }
  return deck;
}

export async function withSlideFrame(page,deck,frame,work){
  if(!frame.hidden)return work();
  const state=await page.evaluate(({active,frame,frameBounds,parentBounds})=>{
    const current=document.querySelector(active.selector),target=document.querySelector(frame.selector);
    const boundsMatch=(el,expected)=>{const r=el.getBoundingClientRect();return [r.x,r.y,r.width,r.height].every((v,i)=>Math.abs(v-expected[i])<.01);};
    const styles=[current,target].map(el=>el.getAttribute('style'));
    current.style.setProperty('display','none','important');
    target.style.setProperty('display','block','important');
    const matched=!current.getClientRects().length&&getComputedStyle(target).display==='block'&&
      getComputedStyle(target).visibility==='visible'&&boundsMatch(target,frameBounds)&&boundsMatch(target.parentElement,parentBounds);
    return {styles,matched};
  },{active:deck.active,frame,frameBounds:deck.frameBounds,parentBounds:deck.parentBounds});
  try{return state.matched?await work():null;}
  finally{
    await page.evaluate(({selectors,styles})=>{
      selectors.forEach((selector,i)=>{
        const el=document.querySelector(selector);
        if(styles[i]===null)el.removeAttribute('style');else el.setAttribute('style',styles[i]);
      });
    },{selectors:[deck.active.selector,frame.selector],styles:state.styles});
  }
}
