import {resolveAsset} from './assets.mjs';
// A separate blank page receives only computed backgrounds. Source elements,
// text, scripts, pseudo-elements and source stylesheets never enter this page.
export async function captureDecorations(session,model,measurements,selector){
  let surface;const cache=new Map();
  try{
  for(const [index,slide] of model.slides.entries()){
    const m=measurements.find(m=>m.slideId===slide.id);if(!m)continue;
    const targets=[...(slide.background?.image&&slide.background.image!=='none'?[{background:slide.background,box:{w:slide.widthCssPx,h:slide.heightCssPx},sourceRef:null}]:[]),...slide.nodes.filter(n=>n.kind==='shape'&&n.style.backgroundImage&&n.style.backgroundImage!=='none').map(n=>({background:n.style,box:n.boxCssPx,sourceRef:n.sourceRef}))];
    for(const target of targets){
      const properties=await session.page.evaluate(({selector,index,ref})=>{
        const root=document.querySelectorAll(selector)[index],source=ref?[root,...root.querySelectorAll('*')].find(e=>ref.elementId?e.dataset.sourceId===ref.elementId:e.dataset.sourcePath===ref.path):root;
        if(!source)throw new Error('Decorative background source missing');
        const st=getComputedStyle(source);
        return Object.fromEntries(['background-color','background-image','background-size','background-position','background-repeat','background-origin','background-clip','background-blend-mode','border-radius','padding-top','padding-right','padding-bottom','padding-left','border-top-width','border-right-width','border-bottom-width','border-left-width'].map(p=>[p,st.getPropertyValue(p)]));
      },{selector,index,ref:target.sourceRef});
      const key=JSON.stringify([properties,target.box.w,target.box.h,m.renderScale]);
      let captured=cache.get(key);
      if(!captured){
        surface??=await session.context.newPage();await surface.setViewportSize(m.referenceSize);
        await surface.setContent('<!doctype html><html><head><style>html,body{margin:0;padding:0;background:transparent}</style></head><body></body></html>');
        const handle=await surface.evaluateHandle(({properties,width,height,scale})=>{
          const el=document.createElement('div');el.style.cssText=`display:block;position:absolute;left:0;top:0;margin:0;box-sizing:border-box;width:${width}px;height:${height}px;transform:scale(${scale});transform-origin:0 0`;
          for(const [property,value] of Object.entries(properties))el.style.setProperty(property,value);
          el.style.borderStyle='solid';el.style.borderColor='transparent';document.body.append(el);return el;
        },{properties,width:target.box.w,height:target.box.h,scale:m.renderScale});
        const bytes=await handle.asElement().screenshot({animations:'disabled',scale:'css',omitBackground:true});await handle.dispose();
        captured={asset:resolveAsset(`data:image/png;base64,${bytes.toString('base64')}`),capture:{width:bytes.readUInt32BE(16),height:bytes.readUInt32BE(20),renderScale:m.renderScale,content:'computed-background-only',sourceChildrenCopied:0}};cache.set(key,captured);
        if(!model.assets.some(a=>a.id===captured.asset.id))model.assets.push({...captured.asset,role:'background-art'});
      }
      target.background.assetId=captured.asset.id;target.background.capture=captured.capture;
    }
  }
  }finally{await surface?.close();}
}
