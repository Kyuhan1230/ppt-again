import { adaptMappedPagination } from './html-pagination-adapter.mjs';
import { observePaginationDocument } from './html-pagination-observer.mjs';
import { diagnostic, validateModel } from './model.mjs';
import { publicDiagnostic } from './public-diagnostics.mjs';
import { normalizeSlideSizes } from './slide-size-policy.mjs';

const defaultTarget = { targetWidthEmu: 12192000, targetHeightEmu: 6858000 };

function paginationError(message, reasonKey = 'PAGINATION_LAYOUT_UNSUPPORTED') {
  return { ...diagnostic('PAGINATION_UNSUPPORTED', message), reasonKey };
}

function coalesceRuns(tokens) {
  const runs = [];
  for (const token of tokens) {
    if (!['copy', 'collapsed-space'].includes(token.kind) || !token.text) continue;
    const style = token.style;
    const last = runs.at(-1);
    if (last && JSON.stringify(last.style) === JSON.stringify(style)) last.text += token.text;
    else runs.push({ text: token.text, style });
  }
  return runs;
}

function pageModel(page, width) {
  const id = `slide-${page.index + 1}`;
  const nodes = [], objects = [];
  for (const part of page.parts) {
    if (part.kind === 'table' || part.kind === 'image') {
      nodes.push({
        id: part.fragmentId, kind: part.kind, boxCssPx: part.boxCssPx, zOrder: nodes.length, parentId: null,
        style: part.originalStyle,
        sourceRef: { slideId: id, path: part.objectId, sourceRanges: part.sourceRanges, ...(part.sourceCells ? { sourceCells: part.sourceCells } : {}) },
        content: part.content ?? part.originalContent
      });
      continue;
    }
    if (part.kind !== 'text') throw new Error('분할된 객체를 PPT에 배치할 수 없습니다.');
    const byLine = new Map();
    for (const token of part.tokens) {
      if (!byLine.has(token.lineId)) byLine.set(token.lineId, []);
      byLine.get(token.lineId).push(token);
    }
    for (const line of part.lines) {
      const runs = coalesceRuns(byLine.get(line.lineId) ?? []);
      if (!runs.length) continue;
      const boxCssPx = { ...line.boxCssPx };
      if (boxCssPx.x < 0 || boxCssPx.y < 0 || boxCssPx.x + boxCssPx.w > width + 1e-6 || boxCssPx.y + boxCssPx.h > page.pageHeight + 1e-6) {
        throw new Error('문단의 줄 위치가 출력 장표를 벗어납니다.');
      }
      const nodeId = `${part.fragmentId}:${line.lineId}`;
      // Explicitly position each measured line. PowerPoint can edit its runs
      // without rewrapping across a page boundary or accumulating line-height drift.
      const style = { ...runs[0].style, lineHeight: 'normal', whiteSpace: 'nowrap', valign: 'top', align: 'left' };
      nodes.push({
        id: nodeId, kind: 'text', boxCssPx, zOrder: nodes.length, parentId: null, style,
        sourceRef: { slideId: id, path: part.objectId, lineId: line.lineId, sourceRanges: byLine.get(line.lineId).flatMap(token => token.sources) },
        content: { text: runs.map(run => run.text).join(''), runs }
      });
      objects.push({ nodeId, kind: 'text', boxCssPx, lines: [{ ...boxCssPx }], baseline: line.baseline });
    }
  }
  return {
    slide: { id, widthCssPx: width, heightCssPx: page.pageHeight, background: { color: '#ffffff' }, nodes },
    measurement: { slideId: id, objects, assetsReady: true, fontsReady: true, layoutStable: true }
  };
}

export async function measurePaginatedInput(input, runtime, { viewport } = {}) {
  const result = { model: null, measurements: [], environment: { ...runtime.environment }, diagnostics: [...input.diagnostics], previews: [] };
  const reject = (message, reasonKey) => {
    result.diagnostics.push(paginationError(message, reasonKey));
    return result;
  };
  if (result.diagnostics.some(item => item.severity === 'error')) return result;
  if (input.profile !== 'static-slides-v1' || input.embeddedSlides || input.source.slideSelector !== null) {
    return reject('문서 전체를 나누는 기능은 일반 HTML 문서에 사용할 수 있습니다.');
  }
  const target = { mode: 'contain', ...defaultTarget, ...input.sizePolicy };
  const session = await runtime.render(input, { viewport });
  try {
    // An explicit document mode selects body. Internal DOM and pixel stability
    // checks still run in this same sandboxed session before any model is built.
    const observed = await observePaginationDocument({
      session, selector: 'body', viewport: session.page.viewportSize(),
      assets: input.assets, assetRefs: input.assetRefs
    });
    if (observed.status !== 'observed') {
      for (const item of observed.diagnostics) result.diagnostics.push(paginationError(item.message, item.reasonKey));
      return result;
    }
    if (observed.coverage?.outsideSelection !== 'verified-no-visible-content') return reject('문서 바깥의 표시 내용과 가림을 확인할 수 없습니다.');
    const { documentWidth, documentHeight } = observed;
    const pageHeight = documentWidth * target.targetHeightEmu / target.targetWidthEmu;
    const planned = adaptMappedPagination({ ...observed, mode: 'paginate', pageHeight });
    if (planned.diagnostics.length) {
      for (const item of planned.diagnostics) result.diagnostics.push(paginationError(item.message, item.reasonKey));
      return result;
    }
    const pages = planned.pages.map(page => pageModel(page, documentWidth));
    const assetIds = new Set(pages.flatMap(page => page.slide.nodes.filter(node => node.kind === 'image').map(node => node.content.assetId)));
    const assets = input.assets.filter(asset => assetIds.has(asset.id));
    const slides = pages.map(page => page.slide);
    const hasTables = slides.some(slide => slide.nodes.some(node => node.kind === 'table'));
    const model = { version: hasTables ? 2 : 1, sourceDigest: input.source.sha256, slides, assets };
    const normalized = normalizeSlideSizes(model, pages.map(page => page.measurement), target);
    if (!validateModel(normalized.model).valid) return reject('분할 장표의 PPT 모델을 확인할 수 없습니다.');
    const families = new Set(model.slides.flatMap(slide => slide.nodes.flatMap(node => {
      const text = node.kind === 'table' ? node.content.cells : node.kind === 'text' ? [node] : [];
      return text.flatMap(item => [item.style, ...(item.runs ?? item.content.runs).map(run => run.style)]).map(style => style.fontFamily);
    })));
    Object.assign(result, normalized, {
      pagination: {
        mode: 'paginate', sourceWidth: documentWidth, sourceHeight: documentHeight, pageHeight,
        pages: planned.pages.map(page => ({ fromY: page.sourceStart, toY: page.sourceEnd })),
        oversizePolicy: 'stop', textEditable: true, textObjectGranularity: 'line',
        sizeSource: input.sizePolicy?.targetWidthEmu === undefined ? 'default-16:9' : 'configured',
        editableTableCount: model.slides.reduce((count, slide) => count + slide.nodes.filter(node => node.kind === 'table').length, 0),
        imageCount: model.slides.reduce((count, slide) => count + slide.nodes.filter(node => node.kind === 'image').length, 0)
      }
    });
    Object.assign(result.environment, { inputProfile: input.profile, viewport: observed.viewport, fonts: [...families].map(family => ({ family, available: true, custom: false })) });
    return result;
  } catch (error) {
    return reject(error.message, error.reasonKey);
  } finally {
    try {
      await session.context.close();
    } catch (cause) {
      // The conversion worker still closes the browser before replying. Keep
      // any established conversion failure when reporting the cleanup error.
      throw Object.assign(new Error('JOB_CLEANUP_FAILED', { cause }), {
        code: 'JOB_CLEANUP_FAILED',
        publicDiagnostics: result.diagnostics.map(item => publicDiagnostic(item))
      });
    }
  }
}
