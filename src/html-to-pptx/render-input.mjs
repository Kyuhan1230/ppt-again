import {parse, serialize} from 'parse5';
import * as cssTree from 'css-tree';
import {attribute, walkHtml} from './input.mjs';
import {assetDataUrl} from './assets.mjs';
const esc = v => String(v??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
export function renderInput(input,{fontAssetUrl=null}={}) {
  const assets=new Map(input.assets.map(a=>[a.id,a]));
  const rewriteCss=(css,basePath=null)=>{
    try { const ast=cssTree.parse(css);cssTree.walk(ast,n=>{if(n.type==='Url'&&!n.value.startsWith('#')){const ref=input.assetRefs.find(r=>r.url===n.value&&r.basePath===basePath),asset=ref&&assets.get(ref.assetId);n.value=asset?(ref.usage==='font'&&fontAssetUrl?fontAssetUrl(asset):null)??assetDataUrl(asset):'data:,';}});return cssTree.generate(ast); }
    catch { return ''; }
  };
  if(input.profile==='project-data-v1'&&input.project) {
    const p=input.project;
    const markup=p.slides.map(s=>`<section data-slide data-slide-id="${esc(s.id)}" class="${esc(s.theme??'')}" style="width:${p.width}px;height:${p.height}px;background-color:${esc(s.background??'transparent')}">${s.nodes.map(n=>{
      const base=`position:absolute;left:${n.x}px;top:${n.y}px;width:${n.w}px;height:${n.h}px;box-sizing:border-box;opacity:${n.opacity??1};`;
      let style=base,content='';
      if(n.type==='text'){style+=`font-size:${n.fontSize}px;font-family:${esc(n.fontFamily??'sans-serif')};font-weight:${n.fontWeight??400};color:${esc(n.color??'#000')};text-align:${esc(n.align??'left')};line-height:${n.lineHeight??1.24};white-space:pre-wrap;padding:0;font-synthesis:none;`;content=esc(n.text);}
      else if(['svg','image'].includes(n.type)){const ref=input.assetRefs.find(r=>r.slideId===s.id&&r.nodeId===n.id);content=ref?`<img alt="" style="display:block;width:100%;height:100%" src="${assetDataUrl(assets.get(ref.assetId))}">`:'';}
      else if(n.type==='line')style+=`background:${esc(n.stroke??'#000')};${n.w>n.h?'height':'width'}:${Math.max(1,n.strokeWidth??1)}px;`;
      else style+=`background:${esc(n.fill??'transparent')};border:${n.strokeWidth??0}px ${n.dash?'dashed':'solid'} ${esc(n.stroke??'transparent')};border-radius:${n.radius??0}px;`;
      return `<div data-source-id="${esc(n.id)}" data-kind="${n.type==='svg'?'image':esc(n.type)}" style="${style}">${content}</div>`;
    }).join('')}</section>`).join('');
    return {html:`<!doctype html><html><head><meta charset="utf-8"><style>${input.styles.filter(s=>s.sourceRef.path.startsWith('style@')||s.basePath).map(s=>rewriteCss(s.css,s.basePath)).join('\n')}</style><style>html,body{margin:0!important;padding:0!important;overflow:visible!important;min-width:0!important;display:block!important}section[data-slide]{position:relative;overflow:hidden;box-sizing:border-box;flex-shrink:0}</style></head><body>${markup}</body></html>`,sourceHadScripts:true};
  }
  const tree=parse(input.source.html,{sourceCodeLocationInfo:true});
  const sourceHadScripts=[...walkHtml(tree)].some(n=>n.tagName==='script');
  for(const node of [...walkHtml(tree)]) {
    if(!node.tagName)continue;
    const sourcePath=`${node.tagName}@${node.sourceCodeLocation?.startOffset??0}`;
    if(['script','base','meta'].includes(node.tagName)) { if(node.parentNode)node.parentNode.childNodes=node.parentNode.childNodes.filter(n=>n!==node);continue; }
    const add=(name,value)=>{node.attrs=node.attrs.filter(a=>a.name!==name);node.attrs.push({name,value});};
    node.attrs=node.attrs.filter(a=>!a.name.startsWith('on')&&!['srcdoc','formaction','action','autofocus','autoplay'].includes(a.name));
    add('data-source-path',sourcePath);
    if(node.tagName==='style')node.childNodes=[{nodeName:'#text',value:rewriteCss(node.childNodes.map(n=>n.value??'').join('')),parentNode:node}];
    if(node.tagName==='link') {
      const style=input.styles.find(s=>s.sourceRef.path===sourcePath);
      if(style){node.tagName='style';node.nodeName='style';node.attrs=[];node.childNodes=[{nodeName:'#text',value:rewriteCss(style.css,style.basePath),parentNode:node}];}
      else if(node.parentNode)node.parentNode.childNodes=node.parentNode.childNodes.filter(n=>n!==node);
      continue;
    }
    const inline=attribute(node,'style');if(inline){const rewritten=rewriteCss(`x{${inline}}`);add('style',rewritten.slice(2,-1));}
    if(node.tagName==='figure'&&(input.charts??[]).some(c=>c.id===attribute(node,'data-chart-id')))node.childNodes=[];
    if(node.tagName==='img') {
      const ref=input.assetRefs.find(r=>r.sourceRef?.path===sourcePath&&r.usage==='image');
      add('src',ref?assetDataUrl(assets.get(ref.assetId)):'data:,');
      if(ref)add('data-asset-id',ref.assetId);
      node.attrs=node.attrs.filter(a=>!['srcset','sizes','loading'].includes(a.name));
    } else if(node.tagName==='image'&&node.namespaceURI==='http://www.w3.org/2000/svg'){
      const ref=input.assetRefs.find(r=>r.sourceRef?.path===sourcePath&&r.usage==='image');
      node.attrs=node.attrs.filter(a=>a.name!=='href');
      if(ref){add('href',assetDataUrl(assets.get(ref.assetId)));add('data-asset-id',ref.assetId);}
    } else node.attrs=node.attrs.filter(a=>!['src','href'].includes(a.name)||a.value.startsWith('#'));
  }
  return {html:serialize(tree),sourceHadScripts};
}
