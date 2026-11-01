import fs from 'node:fs/promises';
import {sha256} from './assets.mjs';
import {inspectFont,decideFontEmbedding} from './font-policy.mjs';

const entries=[
  {file:'Pretendard-Regular.ttf',weight:400,hash:'6d0af5258997aec7354a6e340fc2325ba321c410ca48b3af858c8c3d6e92a324'},
  {file:'Pretendard-Bold.ttf',weight:700,hash:'c16b88c670d23e83fa1170c954cbc4822d3b8dad3c3cde15d798a94b43d97985'}
];
let cached;
async function loadFallback(){
  return cached??=Promise.all(entries.map(async e=>{
    const bytes=await fs.readFile(new URL(`../../assets/fonts/pretendard/${e.file}`,import.meta.url));
    const meta=inspectFont(bytes);if(sha256(bytes)!==e.hash||meta.family!=='Pretendard'||meta.weight!==e.weight||!decideFontEmbedding(meta,{forEditing:true,subset:false}).allowed)throw new Error('Fallback font verification failed');
    return {id:`asset-${e.hash}`,sha256:e.hash,mime:'font/ttf',bytes:new Uint8Array(bytes),weight:e.weight};
  }));
}

// Substitute only unavailable, visible HTML fonts; document scripts remain disabled.
export async function applyFallbackFont(page,nativeAliases=[]){
  const missing=await page.evaluate(async(nativeAliases)=>{
    await document.fonts.ready;
    const generic=new Set(['serif','sans-serif','monospace','system-ui','cursive','fantasy']);
    const elements=[...document.body.querySelectorAll('*')].filter(el=>!el.closest('svg')&&el.getClientRects().length&&getComputedStyle(el).visibility==='visible');
    const candidates=elements.map(el=>{const st=getComputedStyle(el);return {el,family:st.fontFamily.split(',')[0].trim().replace(/^['"]|['"]$/g,''),style:st.fontStyle,weight:Number(st.fontWeight)||400};});
    const families=[...new Set(candidates.map(c=>c.family))],missing=[];
    for(const family of families){if(generic.has(family))continue;const faces=[...document.fonts].filter(f=>f.family.replace(/^['"]|['"]$/g,'')===family);let available=false;
      if(faces.length){try{const loaded=await document.fonts.load(`20px ${JSON.stringify(family)}`);available=loaded.length>0&&loaded.every(f=>f.status==='loaded');}catch{}}
      else{for(const name of [family,...nativeAliases.filter(a=>a.family===family).map(a=>a.localName)]){try{await new FontFace('__fallback_availability',`local(${JSON.stringify(name)})`).load();available=true;break;}catch{}}}
      if(!available)missing.push(family);
    }
    return missing;
  },nativeAliases);
  if(!missing.length)return [];
  const fonts=await loadFallback();
  const applied=await page.evaluate(async({missing,fonts})=>{
    if([...document.fonts].some(face=>face.family.replace(/^['"]|['"]$/g,'')==='PPTAgainFallback'))throw new Error('Fallback font name already used by input');
    for(const font of fonts){const face=new FontFace('PPTAgainFallback',`url(data:font/ttf;base64,${font.base64})`,{weight:String(font.weight),style:'normal'});await face.load();document.fonts.add(face);}
    let changed=0;const candidates=[...document.body.querySelectorAll('*')].filter(el=>!el.closest('svg')&&el.getClientRects().length&&getComputedStyle(el).visibility==='visible').map(el=>{const st=getComputedStyle(el);return {el,family:st.fontFamily.split(',')[0].trim().replace(/^['"]|['"]$/g,''),style:st.fontStyle,weight:Number(st.fontWeight)||400};});
    for(const c of candidates)if(missing.includes(c.family)&&c.style==='normal'){c.el.style.setProperty('font-family','PPTAgainFallback','important');c.el.style.setProperty('font-weight',c.weight>=600?'700':'400','important');changed++;}
    await document.fonts.ready;return changed;
  },{missing,fonts:fonts.map(f=>({weight:f.weight,base64:Buffer.from(f.bytes).toString('base64')}))});
  return applied?fonts:[];
}
