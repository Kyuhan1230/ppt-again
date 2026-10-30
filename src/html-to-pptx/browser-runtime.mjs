import {chromium} from 'playwright-core';
import os from 'node:os';
import {renderInput} from './render-input.mjs';
import {sha256} from './assets.mjs';

export async function createBrowserRuntime({channel='msedge',executablePath,capture=true,timeoutMs=5000,checkLimits=false,onBrowserProcess}={}) {
  const options={headless:true,...(executablePath?{executablePath}:{channel})};
  const owned=onBrowserProcess?await chromium.launchServer({...options,host:'127.0.0.1'}):null;
  if(owned)onBrowserProcess(owned.process().pid);
  let browser;
  try{browser=owned?await chromium.connect(owned.wsEndpoint()):await chromium.launch(options);}catch(error){await owned?.close();throw error;}
  let closing,limitExceeded=false;
  const environment={node:process.version,os:{platform:os.platform(),release:os.release()},browser:browser.version(),channel:executablePath?'explicit executable':channel,dpr:1,sourceScriptsExecuted:false,networkPolicy:'document and verified supplied fonts only; external requests blocked'};
  return {capture,timeoutMs,environment,get limitExceeded(){return limitExceeded;},
    async render(input){
      const fontUrls=new Map();
      for(const ref of input.assetRefs.filter(r=>r.usage==='font')){
        const asset=input.assets.find(a=>a.id===ref.assetId);
        if(!asset||!['font/woff','font/woff2','font/ttf','font/otf'].includes(asset.mime)||asset.id!==`asset-${asset.sha256}`||sha256(asset.bytes)!==asset.sha256)throw new Error('글꼴 자산의 형식 또는 무결성을 확인할 수 없습니다.');
        fontUrls.set(`https://html-input.invalid/fonts/${asset.id}`,asset);
      }
      const rendered=renderInput(input,{fontAssetUrl:asset=>fontUrls.has(`https://html-input.invalid/fonts/${asset.id}`)?`https://html-input.invalid/fonts/${asset.id}`:null});
      const context=await browser.newContext({viewport:{width:1280,height:960},deviceScaleFactor:1,javaScriptEnabled:false,serviceWorkers:'block'});
      const page=await context.newPage();
      const blocked=[];
      await context.route('**/*',route=>{
        if(route.request().url()==='https://html-input.invalid/document'&&route.request().isNavigationRequest()&&route.request().frame()===page.mainFrame()) return route.fulfill({status:200,contentType:'text/html; charset=utf-8',headers:{'Content-Security-Policy':"default-src 'none'; style-src 'unsafe-inline'; img-src data:; font-src 'self' data:; script-src 'none'; connect-src 'none'; frame-src 'none'; object-src 'none'; base-uri 'none'"},body:rendered.html});
        const asset=fontUrls.get(route.request().url());
        if(asset&&route.request().method()==='GET'&&route.request().resourceType()==='font')return route.fulfill({status:200,contentType:asset.mime,headers:{'cache-control':'no-store','x-content-type-options':'nosniff'},body:Buffer.from(asset.bytes)});
        blocked.push(route.request().url());return route.abort();
      });
      try{
        await page.goto('https://html-input.invalid/document',{waitUntil:'load',timeout:timeoutMs});
        if(checkLimits){
          const exceeds=await page.evaluate(({selector,svgPaths})=>{
            const elements=document.body.querySelectorAll('*');if(elements.length>100000)return true;
            let roots;try{roots=svgPaths?[...document.querySelectorAll('svg')].filter(root=>svgPaths.includes(root.dataset.sourcePath)):document.querySelectorAll(selector);}catch{return false;}
            if(roots.length>200||[...roots].some(root=>root.querySelectorAll('*').length>10000))return true;
            const walker=document.createTreeWalker(document.body,NodeFilter.SHOW_TEXT);let node,total=0;
            while((node=walker.nextNode())){if(['SCRIPT','STYLE'].includes(node.parentElement?.tagName))continue;const length=node.textContent.length;if(length>100000||(total+=length)>2000000)return true;}
            return false;
          },{selector:input.source.slideSelector??'[data-slide]',svgPaths:input.profile==='svg-slides-v1'?input.svgRootPaths:null});
          if(exceeds){limitExceeded=true;throw Object.assign(new Error('INPUT_LIMIT'),{code:'INPUT_LIMIT'});}
        }
        return {page,context,sourceHadScripts:rendered.sourceHadScripts,blocked};
      }
      catch(e){await context.close();throw e;}
    },
    close:()=>closing??=(owned?owned.close():browser.close())
  };
}
