import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import {execFile} from 'node:child_process';
import {promisify} from 'node:util';
import {chromium} from 'playwright-core';
import {createHtmlServer} from '../src/html-to-pptx/server.mjs';
import {inspectPptxIndependently} from './inspect-pptx.mjs';

test('UI and CLI export the same mixed-size slides under the chosen wide policy',async t=>{
 const html='<!doctype html><html><body style="margin:0"><section data-slide style="position:relative;width:800px;height:600px;background:red"><p id="first-text" style="position:absolute;left:100px;top:100px;width:200px;height:60px;margin:0;font:20px/30px Arial">First slide</p></section><section data-slide style="position:relative;width:960px;height:540px;background:blue"><p id="second-text" style="position:absolute;left:96px;top:54px;width:200px;height:60px;margin:0;font:20px/30px Arial">Second slide</p></section></body></html>';
 const dir=await fs.mkdtemp(path.join(os.tmpdir(),'ppt-size-e2e-')),input=path.join(dir,'mixed.html'),cliOutput=path.join(dir,'cli.pptx'),uiOutput=path.join(dir,'ui.pptx');
 t.after(async()=>{for(const file of [input,cliOutput,uiOutput])await fs.unlink(file).catch(()=>{});await fs.rmdir(dir).catch(()=>{});});
 await fs.writeFile(input,html);
 const app=await createHtmlServer({port:0}),browser=await chromium.launch({channel:'msedge',headless:true});
 t.after(async()=>{await browser.close();await app.close();});
 const page=await browser.newPage({viewport:{width:390,height:844}});page.setDefaultTimeout(120000);
 await page.goto(app.url);await page.locator('#html-options').evaluate(el=>el.open=true);
 await page.locator('#slide-size-policy').focus();await page.keyboard.press('ArrowDown');assert.equal(await page.locator('#slide-size-policy').inputValue(),'wide');
 await page.locator('#html-file').setInputFiles(input);
 await page.waitForFunction(()=>['analyzed','failed','selection-required'].includes(document.querySelector('#status-badge').dataset.state));
 assert.equal(await page.locator('#status-badge').getAttribute('data-state'),'analyzed',await page.locator('#status-message').innerText());
 assert.match(await page.locator('#slide-size-report').innerText(),/13\.33.*7\.5.*비율 유지.*여백 흰색/);
 assert.match(await page.locator('#quality-note').innerText(),/입력 HTML 미리보기/);
 assert.equal(await page.evaluate(()=>document.documentElement.scrollWidth<=innerWidth),true);
 await page.locator('#slide-size-details summary').click();assert.equal(await page.locator('#slide-size-list li').count(),2);
 if(process.env.PPT_SIZE_SCREENSHOTS==='1'){
  await page.screenshot({path:path.join(os.tmpdir(),'ppt-size-real-mobile.png'),fullPage:true});
  await page.setViewportSize({width:1440,height:900});
  assert.equal(await page.evaluate(()=>document.documentElement.scrollWidth<=innerWidth),true);
  await page.screenshot({path:path.join(os.tmpdir(),'ppt-size-real-desktop.png'),fullPage:true});
 }
 const downloadPromise=page.waitForEvent('download');await page.locator('#download-button').click();const download=await downloadPromise;await download.saveAs(uiOutput);
 const {stdout}=await promisify(execFile)(process.execPath,['src/html-to-pptx/cli.mjs','convert',input,'--output',cliOutput,'--slide-size','16:9','--json'],{timeout:120000,maxBuffer:1024*1024});
 const report=JSON.parse(stdout);assert.equal(report.status,'generated');assert.equal(report.sizePolicy.source,'custom');assert.equal(report.sizePolicy.slides.length,2);
 for(const file of [uiOutput,cliOutput]){
  const ppt=await inspectPptxIndependently(await fs.readFile(file));assert.deepEqual([ppt.widthEmu,ppt.heightEmu],[12192000,6858000]);assert.equal(ppt.slides.length,2);
  const first=ppt.slides[0].objects.find(o=>o.name==='first-text'),second=ppt.slides[1].objects.find(o=>o.name==='second-text');
  assert.match(first?.text??'',/First slide/);assert.match(second?.text??'',/Second slide/);
  assert.ok(Math.abs(first.boxEmu.x-280*9525)<=20,JSON.stringify(first));
  assert.ok(Math.abs(second.boxEmu.x-128*9525)<=20,JSON.stringify(second));
 }
});
