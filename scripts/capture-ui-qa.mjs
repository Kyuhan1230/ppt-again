// Synthetic UI-only visual QA. No user files, credentials or production documents.
import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import assert from 'node:assert/strict';
import {chromium} from 'playwright-core';
import {createHtmlServer} from '../src/html-to-pptx/server.mjs';
const output=process.env.QA_SCREENSHOT_DIR??path.join(os.tmpdir(),'ppt-again-ui-qa');await fs.mkdir(output,{recursive:true});
const app=await createHtmlServer({port:0});let browser;
const preview=`<svg xmlns="http://www.w3.org/2000/svg" width="960" height="540"><rect width="960" height="540" fill="#ffffff"/><rect x="0" width="960" height="14" fill="#2563eb"/><text x="55" y="90" font-family="Arial" font-size="34" fill="#17263c">Synthetic slide · UI QA</text><text x="55" y="140" font-family="Arial" font-size="20" fill="#64748b">HTML preview, not an original PowerPoint screenshot</text><rect x="96" y="216" width="288" height="162" rx="12" fill="#e0e7ff"/><text x="125" y="300" font-family="Arial" font-size="22" fill="#4338ca">Known issue region</text><rect x="480" y="216" width="380" height="162" rx="12" fill="#f1f5f9"/><text x="510" y="300" font-family="Arial" font-size="22" fill="#334155">Supported content</text></svg>`;
const types=[['PPTX_TEXT_STYLE_UNSUPPORTED','일부 글자 서식을 옮기지 못합니다.','text'],['PPTX_SHAPE_GRADIENT_UNSUPPORTED','도형의 그라데이션 채우기를 옮기지 못합니다.','shape'],['PPTX_OBJECT_UNSUPPORTED','원본 확인이 필요한 표현입니다.','text'],['PPTX_MASTER_ARTWORK_UNSUPPORTED','마스터 그림을 옮기지 못합니다.','background']];
const diagnostics=Array.from({length:640},(_,i)=>{const [reasonKey,reason,objectType]=types[i%4];return {code:'PPTX_OBJECT_PLACEHOLDER',severity:'warning',message:'지원하지 않는 내용을 표시 상자로 남겼습니다.',slideNumber:i%29+1,reasonKey,reason,objectType,objectName:objectType==='background'?undefined:`Sample object ${Math.floor(i/4)%12+1}`,bounds:objectType==='background'?undefined:{x:.1,y:.4,w:.3,h:.3}};});
const slides=Array.from({length:29},(_,i)=>({id:`slide-${i+1}`,number:i+1,widthCssPx:960,heightCssPx:540,objectCount:12,previewUrl:`/api/jobs/ui-fixture/preview/${i+1}.png`}));
const job={fileName:'synthetic-ui-qa.pptx',id:'ui-fixture',sourceType:'pptx',targetType:'html',status:'analyzed',slideCount:29,objectCount:348,diagnostics,exportAvailable:true,partialExportAvailable:false,slides,visualQuality:'unmeasured'};
try{
 browser=await chromium.launch({channel:'msedge',headless:true});
 const context=await browser.newContext({viewport:{width:1440,height:1000},deviceScaleFactor:1});const page=await context.newPage();
 await page.route('**/api/analyze',route=>route.fulfill({status:200,contentType:'application/json',body:JSON.stringify(job)}));
 await page.route('**/api/jobs/ui-fixture/preview/*.png',route=>route.fulfill({status:200,contentType:'image/svg+xml',body:preview}));
 await page.goto(app.url);await page.screenshot({path:path.join(output,'01-desktop-idle.png'),fullPage:true});
 await page.locator('#html-file').setInputFiles({name:'synthetic-ui-qa.pptx',mimeType:'application/vnd.openxmlformats-officedocument.presentationml.presentation',buffer:Buffer.from('synthetic UI fixture only')});
 await page.locator('#diagnostics').waitFor({state:'visible'});await page.locator('#preview-image').evaluate(img=>img.decode());
 assert.equal(await page.locator('#diagnostic-list details[open]').count(),0);
 assert.ok(await page.evaluate(()=>document.documentElement.scrollWidth<=innerWidth));
 await page.screenshot({path:path.join(output,'02-desktop-640-diagnostics.png'),fullPage:true});
 const details=page.locator('#diagnostic-list details').first();if(await details.count()){await details.locator('summary').click();await details.locator('.diagnostic-objects > li').nth(11).waitFor();await page.screenshot({path:path.join(output,'03-desktop-detail.png'),fullPage:true});const jump=details.locator('.diagnostic-jump').first();if(await jump.count()){await jump.click();await page.locator('#problem-overlay').waitFor({state:'visible'});const imageBox=await page.locator('#preview-image').boundingBox(),overlayBox=await page.locator('#problem-overlay').boundingBox();assert.ok(Math.abs(overlayBox.x-imageBox.x-imageBox.width*.1)<3&&Math.abs(overlayBox.width-imageBox.width*.3)<3,'Issue overlay must align to rendered image bounds');await page.screenshot({path:path.join(output,'05-desktop-known-location.png'),fullPage:true});}}
 await page.setViewportSize({width:390,height:844});
 await page.locator('#diagnostic-list details[open]').evaluateAll(nodes=>nodes.forEach(n=>n.open=false));
 assert.ok(await page.evaluate(()=>document.documentElement.scrollWidth<=innerWidth),'Mobile must not overflow horizontally');
 await page.screenshot({path:path.join(output,'04-mobile-640-diagnostics.png'),fullPage:true});
 await details.locator('summary').click();await details.locator('.diagnostic-objects > li').nth(11).waitFor();await page.screenshot({path:path.join(output,'06-mobile-detail.png'),fullPage:true});
 await fs.writeFile(path.join(output,'validation.json'),JSON.stringify({synthetic:true,userDataUsed:false,diagnosticCount:640,slideCount:29,desktopWidth:1440,mobileWidth:390,checks:['initial details collapsed','desktop no horizontal overflow','mobile no horizontal overflow','known issue overlay matches rendered image bounds'],screenshots:['01-desktop-idle.png','02-desktop-640-diagnostics.png','03-desktop-detail.png','04-mobile-640-diagnostics.png','05-desktop-known-location.png','06-mobile-detail.png']},null,2));
}finally{await browser?.close();await app.close();}
