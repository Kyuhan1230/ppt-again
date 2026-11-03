import test from 'node:test';
import assert from 'node:assert/strict';
import {readFile} from 'node:fs/promises';
import PptxGenJS from 'pptxgenjs';
import {chromium} from 'playwright-core';
import {readPptx} from '../src/html-to-pptx/read-pptx.mjs';
import {exportHtml} from '../src/html-to-pptx/export-html.mjs';

test('bold remains visible when the chosen face has only a regular font file',async()=>{
  const ppt=new PptxGenJS(),slide=ppt.addSlide();slide.background={color:'FFFFFF'};
  slide.addText('Bold example',{x:1,y:1,w:4,h:1,fontFace:'SingleFace',fontSize:30,bold:true,margin:0,objectName:'bold'});
  const parsed=await readPptx(new Uint8Array(await ppt.write({outputType:'uint8array'})),{fontPolicy:'original'});
  assert.equal(parsed.status,'parsed',JSON.stringify(parsed.diagnostics));
  const regular=await readFile(new URL('../assets/fonts/pretendard/Pretendard-Regular.ttf',import.meta.url));
  const html=exportHtml(parsed.model).html.replace('</style>',`@font-face{font-family:SingleFace;src:url(data:font/ttf;base64,${regular.toString('base64')});font-weight:400}</style>`);
  const browser=await chromium.launch({channel:'msedge',headless:true});
  try{
    const page=await browser.newPage();await page.setContent(html);await page.evaluate(()=>document.fonts.ready);
    const bold=page.locator('[data-source-id="bold"]');
    assert.equal(await bold.locator('span').last().evaluate(el=>getComputedStyle(el).fontWeight),'700');
    const generated=await bold.screenshot();
    await page.addStyleTag({content:'*{font-synthesis:none!important}'});
    const withoutSynthesis=await bold.screenshot();
    assert.notDeepEqual(generated,withoutSynthesis,'bold must not render as the regular face');
  }finally{await browser.close();}
});
