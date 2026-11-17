import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import {execFile} from 'node:child_process';
import {promisify} from 'node:util';
import {strFromU8, unzipSync} from 'fflate';
import {chromium} from 'playwright-core';
import {createHtmlServer} from '../src/html-to-pptx/server.mjs';
import {inspectPptxIndependently} from './inspect-pptx.mjs';

const run = promisify(execFile);
const documentWith = content => `<!doctype html><html><body style="box-sizing:border-box;margin:0;padding:16px 24px;width:640px">${content}</body></html>`;
const words = Array.from({length: 56}, (_, index) => [
  `Item${String(index + 1).padStart(2, '0')}`, 'repeat',
]).flat();
const longDocument = documentWith(
  `<p style="width:150px;margin:0;font:20px/30px Arial;color:#123456">${words.join(' ')}</p>`,
);
const image = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADElEQVR4nGOQTDkBAAHfAUZ76qvRAAAAAElFTkSuQmCC';
const pageFlags = ['--scope', 'document', '--page-mode', 'paginate'];

async function fixture(t, html) {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), 'ppt-pagination-'));
  const input = path.join(dir, 'document.html');
  const output = path.join(dir, 'document.pptx');
  t.after(async () => {
    await fs.unlink(input).catch(() => {});
    await fs.unlink(output).catch(() => {});
    await fs.rmdir(dir);
  });
  await fs.writeFile(input, html);
  return {input, output};
}

async function convert(input, output, args = []) {
  try {
    const {stdout} = await run(process.execPath,
      ['src/html-to-pptx/cli.mjs', 'convert', input, '--output', output, ...args, '--json'],
      {timeout: 220000, maxBuffer: 1024 * 1024});
    return {exitCode: 0, report: JSON.parse(stdout)};
  } catch (error) {
    if (typeof error.code !== 'number' || !error.stdout) throw error;
    return {exitCode: error.code, report: JSON.parse(error.stdout)};
  }
}

function checkPageRanges(report, slideCount) {
  assert.equal(report.mode, 'paginate');
  assert.equal(report.oversizePolicy, 'stop');
  assert.equal(report.textEditable, true);
  assert.equal(report.textObjectGranularity, 'line');
  assert.equal(report.sizeSource, 'default-16:9');
  assert.equal(report.sourceWidth, 640);
  assert.equal(report.pageHeight, 360);
  assert.equal(report.pages.length, slideCount);
  assert.ok(slideCount >= 3, 'the fixture must put a single long paragraph on several slides');
  let end = 0;
  for (const [index, page] of report.pages.entries()) {
    assert.equal(page.number, index + 1);
    assert.equal(page.fromY, end, 'adjacent source ranges must meet in one place');
    assert.ok(page.toY > page.fromY && page.toY - page.fromY <= report.pageHeight + 0.001);
    end = page.toY;
  }
  assert.equal(end, report.sourceHeight, 'the last range must reach the end of the document');
}

async function inspectTextSlides(bytes) {
  const ppt = await inspectPptxIndependently(bytes);
  assert.deepEqual([ppt.widthEmu, ppt.heightEmu], [12192000, 6858000]);
  assert.equal(ppt.pictureCount, 0, 'the source text must remain editable in PowerPoint');
  const editable = ppt.slides.flatMap(slide => slide.objects.filter(object => object.kind === 'text'));
  assert.ok(ppt.slides.every(slide => slide.objects.some(object => object.kind === 'text')));
  assert.deepEqual(editable.flatMap(object => object.text.match(/\S+/g)), words,
    'numbered words and the repeated text between them must appear once in source order');
  for (const object of editable) {
    assert.equal(object.tag, 'sp');
    assert.ok(Math.abs(object.boxEmu.x - 24 * 2 * 9525) <= 9525,
      `the left inset should be preserved on every slide: ${JSON.stringify(object.boxEmu)}`);
    assert.ok(object.boxEmu.y >= 0 && object.boxEmu.y + object.boxEmu.h <= ppt.heightEmu + 20,
      `text box crosses the page edge: ${JSON.stringify(object.boxEmu)}`);
    const visibleRuns = object.runs.filter(run => run.text.trim());
    assert.ok(visibleRuns.length);
    for (const line of visibleRuns) {
      assert.equal(line.font, 'Arial');
      assert.equal(line.fontSizePt, 30);
      assert.equal(line.color, '123456');
    }
  }
  return ppt.slides.map(slide => slide.objects.filter(object => object.kind === 'text').map(object => ({
    text: object.text, boxEmu: object.boxEmu, bodyPr: object.bodyPr,
  })));
}

async function checkPagePlacement(browser, report, slides) {
  const sourcePage = await browser.newPage({viewport: {width: 1280, height: 960}});
  try {
    await sourcePage.setContent(longDocument);
    const sourceLines = await sourcePage.evaluate(() => {
      const root = document.body.getBoundingClientRect();
      const paragraph = document.querySelector('p');
      const range = document.createRange();
      range.selectNodeContents(paragraph.firstChild);
      return [...range.getClientRects()].filter(rect => rect.width > 0)
        .map(rect => ({x: rect.x - root.x, y: rect.y - root.y}));
    });
    const emuPerSourcePixel = 12192000 / report.sourceWidth;
    for (const [index, {fromY, toY}] of report.pages.entries()) {
      const lines = sourceLines.filter(line => line.y >= fromY && line.y < toY);
      assert.equal(slides[index].length, lines.length, `slide ${index + 1} must have one editable text box per measured line`);
      for (const [lineIndex, line] of lines.entries()) {
        const box = slides[index][lineIndex].boxEmu;
        assert.ok(Math.abs(box.x / emuPerSourcePixel - line.x) <= 1);
        assert.ok(Math.abs(box.y / emuPerSourcePixel - (line.y - fromY)) <= 1,
          `slide ${index + 1} line ${lineIndex + 1} moved vertically: ${JSON.stringify({box, line, fromY})}`);
      }
    }
    assert.equal(slides.flat().length, sourceLines.length);
  } finally {
    await sourcePage.close();
  }
}

test('explicit document pagination generates equivalent editable PowerPoint files from CLI and the web UI', async t => {
  const {input, output} = await fixture(t, longDocument);
  const cli = await convert(input, output, pageFlags);
  assert.equal(cli.exitCode, 0, JSON.stringify(cli.report));
  assert.equal(cli.report.status, 'generated');
  const cliSlides = await inspectTextSlides(await fs.readFile(output));
  checkPageRanges(cli.report.pagination, cliSlides.length);

  const app = await createHtmlServer({port: 0});
  let browser;
  t.after(async () => {
    try { await browser?.close(); }
    finally { await app.close(); }
  });
  browser = await chromium.launch({channel: 'msedge', headless: true});
  await checkPagePlacement(browser, cli.report.pagination, cliSlides);
  const page = await browser.newPage();
  page.setDefaultTimeout(220000);
  await page.goto(app.url);
  const options = page.locator('#html-options');
  await options.evaluate(details => { details.open = true; });
  await page.locator('#html-page-mode').selectOption('paginate');
  const analyzed = page.waitForResponse(response => response.url().endsWith('/api/analyze') && response.request().method() === 'POST');
  await page.locator('#html-file').setInputFiles(input);
  const response = await analyzed;
  const job = await response.json();
  assert.equal(response.request().postDataJSON().pageMode, 'paginate');
  assert.equal(response.status(), 200, JSON.stringify(job));
  assert.equal(job.status, 'analyzed', JSON.stringify(job.diagnostics));
  assert.equal(job.exportAvailable, true);
  checkPageRanges(job.pagination, cliSlides.length);
  assert.deepEqual(job.pagination.pages, cli.report.pagination.pages);
  await page.waitForFunction(() => document.querySelector('#status-badge').dataset.state === 'analyzed');
  assert.match(await page.locator('#html-pagination-report').innerText(), new RegExp(`${cliSlides.length}장.*편집 가능한 텍스트`));
  await page.locator('#html-pagination-details summary').click();
  assert.equal(await page.locator('#html-pagination-list li').count(), cliSlides.length);
  const downloaded = page.waitForEvent('download');
  await page.locator('#download-button').click();
  const file = await downloaded;
  const uiSlides = await inspectTextSlides(await fs.readFile(await file.path()));
  assert.deepEqual(uiSlides, cliSlides);
});

test('a long unmarked document needs an explicit choice and an unrepresentable fixed overlay stops pagination', async t => {
  const ordinary = await fixture(t, longDocument);
  const implicit = await convert(ordinary.input, ordinary.output);
  assert.equal(implicit.exitCode, 1, JSON.stringify(implicit.report));
  assert.equal(implicit.report.code, 'SELECTION_REQUIRED');
  await assert.rejects(fs.access(ordinary.output), {code: 'ENOENT'});

  const fixedOverlay = longDocument.replace('</body>',
    '<p style="position:fixed;left:0;top:0;background:white;font:20px Arial">Pinned content</p></body>');
  const unsupported = await fixture(t, fixedOverlay);
  await expectPaginationFailure(unsupported);
});

async function expectPaginationFailure({input, output}, reasonKey) {
  const result = await convert(input, output, pageFlags);
  const {report} = result;
  assert.equal(result.exitCode, 1, JSON.stringify(report));
  assert.equal(report.code, 'PAGINATION_UNSUPPORTED', JSON.stringify(report));
  assert.ok(report.diagnostics.some(diagnostic => diagnostic.code === 'PAGINATION_UNSUPPORTED' &&
    diagnostic.severity === 'error' && (reasonKey === undefined || diagnostic.reasonKey === reasonKey)), JSON.stringify(report));
  await assert.rejects(fs.access(output), {code: 'ENOENT'});
}

test('pagination keeps a rowspan intact and exports the table fragments and image as native objects', async t => {
  const rowLabels = Array.from({length: 14}, (_, index) => `Row ${index + 1}`);
  const rows = rowLabels.map((label, index) => {
    const firstCell = index === 5 ? '<td rowspan="2">Merged</td>' : index === 6 ? '' : `<td>Key ${index + 1}</td>`;
    return `<tr>${firstCell}<td>${label}</td></tr>`;
  }).join('');
  const html = documentWith(`<style>tr{height:60px;border:0}td{padding:0;border:0}</style>
    <table style="width:240px;border:0;border-collapse:collapse;border-spacing:0;table-layout:fixed;font:16px/20px Arial;color:black">${rows}</table>
    <img src="${image}" alt="" style="display:block;width:64px;height:64px;margin-top:12px">`);
  const {input, output} = await fixture(t, html);
  const result = await convert(input, output, pageFlags);
  assert.equal(result.exitCode, 0, JSON.stringify(result.report));
  const bytes = await fs.readFile(output);
  const ppt = await inspectPptxIndependently(bytes);
  checkPageRanges(result.report.pagination, ppt.slides.length);
  assert.ok(Math.abs(result.report.pagination.pages[0].toY - 316) <= 0.01,
    'with a 360px page cap the cut must back up to the start of the rowspan');
  const tables = ppt.slides.flatMap(slide => slide.objects.filter(object => object.kind === 'table'));
  assert.ok(tables.length > 1);
  assert.equal(result.report.pagination.editableTableCount, tables.length);
  assert.deepEqual(tables.flatMap(object => object.table.rows.map(row => row.cells.at(-1))), rowLabels);
  assert.ok(tables.every(object => object.table.rows.every(row => Math.abs(row.height - 60 * 2 * 9525) <= 20)));
  const xmlParts = unzipSync(bytes);
  const tableXml = ppt.slides.filter(slide => slide.objects.some(object => object.kind === 'table'))
    .map(slide => strFromU8(xmlParts[slide.file]));
  const mergedSlide = tableXml.find(xml => xml.includes('>Merged</a:t>'));
  assert.ok(mergedSlide);
  assert.match(mergedSlide, /rowSpan="2"/);
  assert.match(mergedSlide, /vMerge="1"/);
  assert.match(mergedSlide, />Row 6<\/a:t>/);
  assert.match(mergedSlide, />Row 7<\/a:t>/);
  assert.equal(tableXml.filter(xml => xml.includes('>Merged</a:t>')).length, 1);
  assert.equal(ppt.pictureCount, 1);
  assert.equal(result.report.pagination.imageCount, 1);
  assert.equal(ppt.fullSlideRasterReplacementCount, 0);
  const picture = ppt.slides.flatMap(slide => slide.objects).find(object => object.kind === 'image');
  assert.ok(Math.abs(picture.boxEmu.w - 64 * 2 * 9525) <= 20);
  assert.ok(Math.abs(picture.boxEmu.h - 64 * 2 * 9525) <= 20);
});

test('a picture taller than the requested slide stops instead of producing a clipped file', async t => {
  const html = documentWith(`<p style="margin:0;font:20px/30px Arial">Before picture</p>
    <img src="${image}" alt="" style="display:block;width:400px;height:400px;margin-top:10px">`);
  await expectPaginationFailure(await fixture(t, html), 'PAGINATION_OVERSIZE');
});

test('a verified overlapping picture and paragraph move together to the next slide', async t => {
  const html = documentWith(`<p style="width:150px;margin:0;font:20px/30px Arial">${words.slice(0, 16).join(' ')}</p>
    <p style="position:relative;width:150px;margin:0;font:20px/30px Arial">${words.slice(16, 22).join(' ')}</p>
    <img src="${image}" alt="" style="display:block;position:relative;top:-28px;width:80px;height:80px">`);
  const {input, output} = await fixture(t, html);
  const result = await convert(input, output, pageFlags);
  assert.equal(result.exitCode, 0, JSON.stringify(result.report));
  const ppt = await inspectPptxIndependently(await fs.readFile(output));
  assert.equal(ppt.slides.length, 2);
  const [first, second] = ppt.slides;
  assert.ok(Math.abs(result.report.pagination.pages[0].toY - 256) <= 0.01,
    'the safe cut precedes the overlapping objects, earlier than the 360px page cap');
  assert.deepEqual(first.objects.flatMap(object => object.text.match(/\S+/g) ?? []), words.slice(0, 16));
  assert.deepEqual(second.objects.flatMap(object => object.text.match(/\S+/g) ?? []), words.slice(16, 22));
  assert.equal(ppt.pictureCount, 1);
  const picture = second.objects.find(object => object.kind === 'image');
  const lastLine = second.objects.find(object => object.kind === 'text' && object.text.includes('Item11'));
  assert.ok(picture && lastLine);
  assert.ok(picture.order > lastLine.order, 'the image must remain above the text in painting order');
  assert.ok(picture.boxEmu.y < lastLine.boxEmu.y + lastLine.boxEmu.h && lastLine.boxEmu.y < picture.boxEmu.y + picture.boxEmu.h);
});

test('individually fitting picture and text objects stop if their overlapping paint cannot fit together', async t => {
  const html = documentWith(`<p style="position:relative;width:150px;margin:0;font:20px/30px Arial">${words.slice(0, 8).join(' ')}</p>
    <img src="${image}" alt="" style="display:block;position:relative;top:-28px;width:300px;height:300px">`);
  await expectPaginationFailure(await fixture(t, html), 'PAGINATION_OVERSIZE');
});

test('a multi-column document stops instead of changing reading order', async t => {
  const html = longDocument.replace('width:640px', 'width:640px;height:950px;column-count:2;column-fill:auto');
  await expectPaginationFailure(await fixture(t, html));
});
