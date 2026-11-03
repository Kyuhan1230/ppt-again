import test from 'node:test';
import assert from 'node:assert/strict';
import PptxGenJS from 'pptxgenjs';
import {chromium} from 'playwright-core';
import {readPptx} from '../src/html-to-pptx/read-pptx.mjs';
import {exportHtml} from '../src/html-to-pptx/export-html.mjs';
import {readInput} from '../src/html-to-pptx/input.mjs';
import {createBrowserRuntime} from '../src/html-to-pptx/browser-runtime.mjs';
import {measureInput} from '../src/html-to-pptx/measure.mjs';
import {exportPptx} from '../src/html-to-pptx/export.mjs';

test('PPT text renders without colour fringes and remains editable',async()=>{
  const ppt=new PptxGenJS(),slide=ppt.addSlide();slide.background={color:'FFFFFF'};
  slide.addText('A black text example',{x:1,y:1,w:5,h:1,fontFace:'Arial',fontSize:30,color:'000000',margin:0,objectName:'text'});
  const source=await readPptx(new Uint8Array(await ppt.write({outputType:'uint8array'})),{fontPolicy:'original'});
  assert.equal(source.status,'parsed',JSON.stringify(source.diagnostics));
  const html=exportHtml(source.model).html,browser=await chromium.launch({channel:'msedge',headless:true});
  try{
    const page=await browser.newPage({viewport:{width:1920,height:1200}});await page.setContent(html);
    const screenshot=await page.locator('section[data-slide]').screenshot();
    const pixels=await page.evaluate(async encoded=>{
      const img=new Image();img.src='data:image/png;base64,'+encoded;await img.decode();
      const canvas=document.createElement('canvas');canvas.width=img.width;canvas.height=img.height;
      const context=canvas.getContext('2d');context.drawImage(img,0,0);
      const data=context.getImageData(0,0,img.width,img.height).data;let ink=0,coloured=0;
      for(let i=0;i<data.length;i+=4){const r=data[i],g=data[i+1],b=data[i+2];if(Math.min(r,g,b)<240){ink++;if(Math.max(r,g,b)-Math.min(r,g,b)>5)coloured++;}}
      return {ink,coloured};
    },screenshot.toString('base64'));
    assert.ok(pixels.ink>100,JSON.stringify(pixels));assert.equal(pixels.coloured,0,JSON.stringify(pixels));
  }finally{await browser.close();}
  const runtime=await createBrowserRuntime({capture:false});
  try{
    const measured=await measureInput(await readInput({html}),runtime),output=await exportPptx(measured);
    assert.equal(output.status,'generated',JSON.stringify(output.diagnostics));
    assert.equal(measured.model.slides[0].nodes.find(node=>node.id==='text').content.text,'A black text example');
  }finally{await runtime.close();}
});
