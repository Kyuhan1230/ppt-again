import {chromium} from 'playwright-core';
import os from 'node:os';
import {renderInput} from './render-input.mjs';

export async function createBrowserRuntime({channel='msedge',executablePath,capture=true,timeoutMs=5000}={}) {
  const browser=await chromium.launch({headless:true,...(executablePath?{executablePath}:{channel})});
  const environment={node:process.version,os:{platform:os.platform(),release:os.release()},browser:browser.version(),channel:executablePath?'explicit executable':channel,dpr:1,sourceScriptsExecuted:false,networkPolicy:'document fulfillment only; external requests blocked'};
  return {capture,timeoutMs,environment,
    async render(input){
      const rendered=renderInput(input);
      const context=await browser.newContext({viewport:{width:1280,height:960},deviceScaleFactor:1,javaScriptEnabled:false,serviceWorkers:'block'});
      const page=await context.newPage();
      const blocked=[];
      await context.route('**/*',route=>{
        if(route.request().url()==='https://html-input.invalid/document'&&route.request().isNavigationRequest()&&route.request().frame()===page.mainFrame()) return route.fulfill({status:200,contentType:'text/html; charset=utf-8',headers:{'Content-Security-Policy':"default-src 'none'; style-src 'unsafe-inline'; img-src data:; font-src data:; script-src 'none'; connect-src 'none'; frame-src 'none'; object-src 'none'; base-uri 'none'"},body:rendered.html});
        blocked.push(route.request().url());return route.abort();
      });
      try{await page.goto('https://html-input.invalid/document',{waitUntil:'load',timeout:timeoutMs});return {page,context,sourceHadScripts:rendered.sourceHadScripts,blocked};}
      catch(e){await context.close();throw e;}
    },
    close:()=>browser.close()
  };
}
