import {sha256} from './assets.mjs';
import {inspectInput} from './html-regions.mjs';
import {measureInput} from './measure.mjs';

const failure=code=>Object.assign(new Error(code),{code});
async function selectedFingerprint(page,selectors,timeoutMs){
  const state=await page.evaluate(selectors=>{let totalNodes=0,totalRects=0;return selectors.map(selector=>{
    const matches=[...document.querySelectorAll(selector)];
    if(matches.length!==1)throw new Error('STALE_INSPECTION');
    const root=matches[0],base=root.getBoundingClientRect();
    if(!base.width||!base.height||base.width>8192||base.height>8192||base.width*base.height>8000000)throw new Error('STALE_INSPECTION');
    const nodes=[root,...root.querySelectorAll('*')];if(nodes.length>10000||(totalNodes+=nodes.length)>100000)throw new Error('INPUT_LIMIT');
    const props=['display','visibility','opacity','position','overflow','transform','clipPath','maskImage','filter','mixBlendMode','font','fontFamily','fontWeight','fontStyle','fontSize','lineHeight','letterSpacing','whiteSpace','color','background','border','boxShadow','textShadow'];
    return nodes.map(el=>{
      const r=el.getBoundingClientRect(),s=getComputedStyle(el),style=props.map(p=>s[p]);
      const texts=[];for(const n of el.childNodes)if(n.nodeType===Node.TEXT_NODE&&n.data.trim()){const range=document.createRange();range.selectNodeContents(n);const rects=range.getClientRects();if((totalRects+=rects.length)>50000)throw new Error('INPUT_LIMIT');texts.push([...rects].map(r=>[r.x-base.x,r.y-base.y,r.width,r.height]));}
      const pseudo=['::before','::after'].map(p=>{const s=getComputedStyle(el,p);return [s.content,s.display,s.visibility,s.color,s.background,s.position,s.transform];});
      return [r.x-base.x,r.y-base.y,r.width,r.height,style,texts,pseudo];
    });
  });},selectors).catch(error=>{throw failure(/INPUT_LIMIT/.test(error.message)?'INPUT_LIMIT':'STALE_INSPECTION');});
  if(state.reduce((sum,nodes)=>sum+nodes.length,0)>100000||state.reduce((sum,nodes)=>sum+(nodes[0]?.[2]??0)*(nodes[0]?.[3]??0),0)>32000000)throw failure('INPUT_LIMIT');
  const pictures=[];
  try{for(const selector of selectors){const image=await page.locator(selector).screenshot({animations:'disabled',scale:'css',timeout:timeoutMs});if(image.length>16*1024*1024)throw failure('STALE_INSPECTION');pictures.push(sha256(image));}}
  catch(error){if(error.code==='INPUT_LIMIT')throw error;throw failure('STALE_INSPECTION');}
  return sha256(JSON.stringify({state,pictures}));
}

// Rebuild selection inside the same session used for the actual measurement.
// The caller supplies IDs only; selectors never cross the worker boundary.
export async function measureSelectedInput(input,runtime,{snapshotId,choiceId,viewport={width:1280,height:960}}={}){
  if(typeof snapshotId!=='string'||typeof choiceId!=='string')throw failure('STALE_INSPECTION');
  if(input.profile!=='static-slides-v1'||input.embeddedSlides)throw failure('REGION_UNSUPPORTED');
  let session;
  try{
    session=await runtime.render(input,{viewport});
    const inspected=await inspectInput(input,runtime,{viewport,session,includePrivate:true});
    if(inspected.snapshotId!==snapshotId||inspected.uncertainties.some(reason=>['layout-may-change','resource-not-ready','render-pixels-not-fully-observed'].includes(reason)))throw failure('STALE_INSPECTION');
    const choice=inspected.choices.find(c=>c.id===choiceId),selectors=inspected.private.choiceSelectors[choiceId];
    if(!choice)throw failure('STALE_INSPECTION');
    if(!choice.convertible||!selectors?.length)throw failure('REGION_UNSUPPORTED');
    const fingerprint=await selectedFingerprint(session.page,selectors,runtime.timeoutMs);
    const selectedInput={...input,source:{...input.source,slideSelector:selectors.join(',')}};
    return await measureInput(selectedInput,runtime,{session,viewport,fallbackFonts:inspected.private.fallbackFonts,validateSelection:async(page)=>{
      if(fingerprint!==await selectedFingerprint(page,selectors,runtime.timeoutMs))throw failure('STALE_INSPECTION');
    }});
  }finally{await session?.context.close().catch(()=>{});}
}
