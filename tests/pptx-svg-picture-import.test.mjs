import test from 'node:test';
import assert from 'node:assert/strict';
import PptxGenJS from 'pptxgenjs';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { readPptx } from '../src/html-to-pptx/read-pptx.mjs';
import { exportHtml } from '../src/html-to-pptx/export-html.mjs';

const svg = '<svg xmlns="http://www.w3.org/2000/svg" width="40" height="40"><rect width="40" height="40" fill="#4499ff"/></svg>';
const png = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+j1ioAAAAASUVORK5CYII=', 'base64');

async function fixture(artwork = svg) {
  const ppt = new PptxGenJS();
  const slide = ppt.addSlide();
  slide.background = { color: 'FFFFFF' };
  slide.addImage({ data: `data:image/svg+xml;base64,${Buffer.from(artwork).toString('base64')}`, x: 1, y: 1, w: 1, h: 1, objectName: 'vector-picture' });
  return unzipSync(new Uint8Array(await ppt.write({ outputType: 'uint8array' })));
}

function patch(parts, file, change) { parts[file] = strToU8(change(strFromU8(parts[file]))); }

test('imports the linked SVG when a PowerPoint SVG picture has mislabeled identical fallback bytes', async () => {
  const result = await readPptx(zipSync(await fixture()), { allowPartial: true });
  assert.equal(result.status, 'parsed', JSON.stringify(result.diagnostics));
  const image = result.model.slides[0].nodes.find(node => node.id === 'vector-picture');
  const asset = result.model.assets.find(item => item.id === image.content.assetId);
  assert.equal(asset.mime, 'image/svg+xml');
  assert.equal(Buffer.from(asset.bytes).toString(), svg);
  assert.match(exportHtml(result.model).html, /data:image\/svg\+xml;base64,/);
});

test('continues using a valid raster fallback when an SVG alternative is present', async () => {
  const parts = await fixture();
  parts['ppt/media/image-1-1.png'] = png;
  const result = await readPptx(zipSync(parts), { allowPartial: true });
  assert.equal(result.status, 'parsed', JSON.stringify(result.diagnostics));
  const image = result.model.slides[0].nodes.find(node => node.id === 'vector-picture');
  assert.equal(result.model.assets.find(item => item.id === image.content.assetId).mime, 'image/png');
});

test('never recovers a mislabeled raster from an unrelated or unvalidated SVG', async () => {
  for (const change of [
    parts => { parts['ppt/media/image-1-2.svg'] = strToU8(svg.replace('#4499ff', '#ff0000')); },
    parts => patch(parts, 'ppt/slides/_rels/slide1.xml.rels', xml => xml.replace('relationships/image" Target="../media/image-1-2.svg"', 'relationships/other" Target="../media/image-1-2.svg"')),
    parts => patch(parts, '[Content_Types].xml', xml => xml.replace('Extension="svg" ContentType="image/svg+xml"', 'Extension="svg" ContentType="application/xml"')),
    parts => patch(parts, 'ppt/slides/slide1.xml', xml => xml.replace('http://schemas.microsoft.com/office/drawing/2016/SVG/main', 'urn:untrusted')),
  ]) {
    const parts = await fixture();
    change(parts);
    const result = await readPptx(zipSync(parts), { allowPartial: true });
    assert.equal(result.status, 'failed');
    assert.ok(result.diagnostics.some(item => item.severity === 'error'));
  }
});

test('still rejects scripts and external references in the linked SVG', async () => {
  for (const content of ['<script>bad()</script>', '<image href="https://example.invalid/tracker.png"/>']) {
    const unsafe = `<svg xmlns="http://www.w3.org/2000/svg" width="40" height="40">${content}</svg>`;
    const result = await readPptx(zipSync(await fixture(unsafe)), { allowPartial: true });
    assert.equal(result.status, 'failed');
    assert.ok(result.diagnostics.some(item => item.severity === 'error'));
  }
});
