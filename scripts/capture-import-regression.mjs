// Real synthetic PPTX import through the HTTP app. No user or company content.
import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import assert from 'node:assert/strict';
import PptxGenJS from 'pptxgenjs';
import {unzipSync,zipSync,strFromU8,strToU8} from 'fflate';
import {chromium} from 'playwright-core';
import {createHtmlServer} from '../src/html-to-pptx/server.mjs';
import {addNativeChart} from '../src/html-to-pptx/write-chart.mjs';
const output=process.env.QA_SCREENSHOT_DIR??path.join(os.tmpdir(),'ppt-again-import-qa');await fs.mkdir(output,{recursive:true});
const ppt=new PptxGenJS();ppt.layout='LAYOUT_WIDE';
for(let i=0;i<12;i++){
 const slide=ppt.addSlide();slide.background={color:'FFFFFF'};
 slide.addText(`Production QA — safe content ${i+1}`,{x:.5,y:.3,w:11,h:.6,fontFace:'Arial',fontSize:24,objectName:`safe-${i+1}`});
 slide.addText(`Inherited typography ${i+1}`,{x:.5,y:1,w:6,h:.5,fontFace:'Arial',fontSize:18,objectName:`inherited-${i+1}`});
 if(i===0)addNativeChart(slide,{id:'chart-qa',kind:'chart',boxCssPx:{x:96,y:190,w:700,h:300},style:{fontFamily:'Arial',fontSize:16},content:{type:'bar',categories:['A','B'],series:[{name:'Synthetic',values:[10,20],color:'#1976d2'}],title:'Synthetic chart',legend:true,axes:{}}},{pointScale:.75,resolveFont:()=>({fontFace:'Arial'})});
}
const parts=unzipSync(new Uint8Array(await ppt.write({outputType:'uint8array'})));
for(let i=1;i<=12;i++){const key=`ppt/slides/slide${i}.xml`;let xml=strFromU8(parts[key]);const shape=xml.match(/<p:sp>[\s\S]*?<\/p:sp>/g)[1];xml=xml.replace(shape,shape.replace(/\s+sz="[^"]*"/g,'').replace(/typeface="Arial"/g,'typeface="+mn-lt"'));parts[key]=strToU8(xml);}
const chart=Object.keys(parts).find(p=>/^ppt\/charts\/chart\d+\.xml$/.test(p));parts[chart]=strToU8(strFromU8(parts[chart]).replace('<c:chart>','<c:style val="9"/><c:chart>'));
const last='ppt/slides/slide12.xml';let lastXml=strFromU8(parts[last]);const lastShape=lastXml.match(/<p:sp>[\s\S]*?<\/p:sp>/g)[1];parts[last]=strToU8(lastXml.replace(lastShape,lastShape.replace(/(<a:ext\s+cx=")[^"]+/, '$10')));
const valid=zipSync(parts),fatalParts={...parts};const zeroXml=strFromU8(parts[last]),zeroShape=zeroXml.match(/<p:sp>[\s\S]*?<\/p:sp>/g)[1];fatalParts[last]=strToU8(zeroXml.replace(zeroShape,zeroShape.replace('cx="0"','cx="-1"')));const fatal=zipSync(fatalParts);
await fs.writeFile(path.join(output,'synthetic-partial.pptx'),valid);await fs.writeFile(path.join(output,'synthetic-fatal.pptx'),fatal);
if(process.argv.includes('--fixtures-only'))process.exit(0);
const app=await createHtmlServer({port:0});let browser;
try{
 browser=await chromium.launch({channel:'msedge',headless:true});const context=await browser.newContext({viewport:{width:1440,height:1000},acceptDownloads:true});const page=await context.newPage();page.setDefaultTimeout(120000);
 await page.goto(app.url);await page.locator('#html-file').setInputFiles({name:'synthetic-partial.pptx',mimeType:'application/vnd.openxmlformats-officedocument.presentationml.presentation',buffer:Buffer.from(valid)});
 await page.waitForFunction(()=>['analyzed','failed'].includes(document.getElementById('status-badge').dataset.state),null,{timeout:180000});
 assert.equal(await page.locator('#download-button').isEnabled(),true,'Valid unsupported presentation styles must not block safe HTML');
 await page.locator('#preview-image').waitFor({state:'visible'});await page.locator('#preview-image').evaluate(img=>img.decode());await page.screenshot({path:path.join(output,'07-real-partial-import.png'),fullPage:true});
 const downloadEvent=page.waitForEvent('download');await page.locator('#download-button').click();const download=await downloadEvent;const htmlPath=path.join(output,'synthetic-partial.html');await download.saveAs(htmlPath);assert.match(await fs.readFile(htmlPath,'utf8'),/Production QA/);
 await page.locator('#html-file').setInputFiles({name:'synthetic-fatal.pptx',mimeType:'application/vnd.openxmlformats-officedocument.presentationml.presentation',buffer:Buffer.from(fatal)});await page.waitForFunction(()=>document.getElementById('status-badge').dataset.state==='failed');
 assert.equal(await page.locator('#download-button').isDisabled(),true);assert.match(await page.locator('#status-message').innerText(),/입력 분석 중단/);await page.screenshot({path:path.join(output,'08-real-fatal-import.png'),fullPage:true});
 await fs.writeFile(path.join(output,'real-import-validation.json'),JSON.stringify({synthetic:true,userDataUsed:false,slides:12,validPartialHtmlDownloaded:true,actualPreviewRendered:true,invalidGeometryBlocked:true},null,2));
}finally{await browser?.close();await app.close();}
