import test from 'node:test';
import assert from 'node:assert/strict';
import {readInput} from '../src/html-to-pptx/input.mjs';
import {measureInput} from '../src/html-to-pptx/measure.mjs';
import {createBrowserRuntime} from '../src/html-to-pptx/browser-runtime.mjs';

test('source previews preserve the selected viewport and fractional slide bounds', async () => {
  const runtime = await createBrowserRuntime();
  try {
    const html = '<!doctype html><style>body{margin:0}section{width:1040.5px;height:720.25px;background:rgb(17,34,51)}.narrow{display:none}@media(max-width:900px){section{width:720.25px;height:1040.5px}.wide{display:none}.narrow{display:inline}}</style><section data-slide><p style="margin:0;color:white;font:20px Arial"><span class="wide">Wide viewport</span><span class="narrow">Narrow viewport</span></p></section>';
    const input = await readInput({html});
    for (const [viewport, expectedText] of [[{width:1280,height:960},'Wide viewport'],[{width:800,height:1200},'Narrow viewport']]) {
      const source = await runtime.render(input,{viewport});
      try {
        const bounds = await source.page.locator('[data-slide]').boundingBox();
        const referenceText = await source.page.locator('[data-slide]').innerText();
        assert.ok(bounds);assert.equal(referenceText,expectedText);
        assert.ok(!Number.isInteger(bounds.width)||!Number.isInteger(bounds.height),'the source must retain fractional CSS dimensions');
        const measured = await measureInput(input, runtime, {viewport});
        assert.deepEqual(measured.environment.viewport,viewport);
        assert.equal(measured.environment.previewKind,'source');
        assert.equal(measured.previews.length,1,JSON.stringify(measured.diagnostics));
        assert.ok(JSON.stringify(measured.model).includes(expectedText),JSON.stringify(measured.diagnostics));
        const png=Buffer.from(measured.previews[0].bytes),width=png.readUInt32BE(16),height=png.readUInt32BE(20);
        assert.ok(Math.abs(width-bounds.width)<=1,JSON.stringify({width,bounds}));
        assert.ok(Math.abs(height-bounds.height)<=1,JSON.stringify({height,bounds}));
        const pixels=await source.page.evaluate(async url=>{const image=new Image();image.src=url;await image.decode();const canvas=document.createElement('canvas');canvas.width=image.width;canvas.height=image.height;const ctx=canvas.getContext('2d');ctx.drawImage(image,0,0);return [[2,2],[image.width-2,image.height-2]].map(([x,y])=>[...ctx.getImageData(x,y,1,1).data]);},`data:image/png;base64,${png.toString('base64')}`);
        assert.deepEqual(pixels,[[17,34,51,255],[17,34,51,255]],'the source preview must include both corners of the actual slide');
      } finally { await source.context.close(); }
    }
  } finally { await runtime.close(); }
});
