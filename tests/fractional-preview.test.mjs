import test from 'node:test';
import assert from 'node:assert/strict';
import {readInput} from '../src/html-to-pptx/input.mjs';
import {measureInput} from '../src/html-to-pptx/measure.mjs';
import {createBrowserRuntime} from '../src/html-to-pptx/browser-runtime.mjs';

test('preview is generated for slide dimensions that scale to fractional pixels', async () => {
  const runtime = await createBrowserRuntime();
  try {
    for (const [width,height,expectedWidth,expectedHeight] of [[1040,720,1280,886],[720,1040,886,1280]]) {
      const html = `<section data-slide style="width:${width}px;height:${height}px;background:#fff"><p style="font-family:Arial">Preview</p></section>`;
      const measured = await measureInput(await readInput({html}), runtime);
      assert.equal(measured.previews.length, 1, JSON.stringify(measured.diagnostics));
      const png = Buffer.from(measured.previews[0].bytes);
      assert.equal(png.readUInt32BE(16), expectedWidth);
      assert.equal(png.readUInt32BE(20), expectedHeight);
    }
  } finally { await runtime.close(); }
});
