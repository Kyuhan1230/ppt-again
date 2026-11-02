import {canvasDiagnostics} from './canvas.mjs';
import {validateTable} from './tables.mjs';
import {validateChart} from './charts.mjs';
const record = v => v !== null && typeof v === 'object' && !Array.isArray(v);
const nonempty = v => typeof v === 'string' && v.trim().length > 0;
const digest = v => typeof v === 'string' && /^[a-f0-9]{64}$/.test(v);
const kinds = new Set(['text', 'shape', 'line', 'image', 'group', 'table', 'chart']);
const codes = new Set(['INPUT_INVALID', 'SLIDE_ROOT_REQUIRED', 'ASSET_MISSING', 'FONT_UNAVAILABLE', 'FONT_SUBSTITUTED', 'FONT_PREVIEW_SUBSTITUTED',
  'FONT_INVALID','FONT_FORMAT_UNSUPPORTED','FONT_VARIABLE_UNSUPPORTED','FONT_STYLE_UNSUPPORTED','FONT_STYLE_MISMATCH','FONT_GLYPH_MISSING',
  'FONT_SOURCE_UNSUPPORTED','FONT_SUBSET_UNSUPPORTED','FONT_ASSET_LIMIT','FONT_EMBEDDING_RESTRICTED','FONT_EMBEDDING_INVALID','FONT_EMBEDDING_UNSUPPORTED',
  'FONT_PREVIEW_ONLY','FONT_BITMAP_ONLY','FONT_SUBSETTING_RESTRICTED','FONT_NOT_PORTABLE',
  'CHART_DATA_REQUIRED','CHART_DATA_MISMATCH','LIMIT_EXCEEDED',
  'TABLE_LAYOUT_APPROXIMATED', 'PARAGRAPH_SPACING_APPROXIMATED', 'SVG_TEXT_APPROXIMATED', 'SCRIPT_REQUIRED', 'UNSUPPORTED_OBJECT', 'SVG_TEXT_NOT_EDITABLE', 'RENDER_INCOMPLETE', 'EXPORT_FAILED']);

export function diagnostic(code, message, location = {}, severity = 'error') {
  return {code, severity, slideId: location.slideId ?? null, nodeId: location.nodeId ?? null,
    sourceRef: location.sourceRef ?? null, message};
}

// This boundary checks data, never repairs it or substitutes values silently.
export function validateModel(value) {
  const errors = [];
  const check = (condition, path, message) => { if (!condition) errors.push({path, message}); };
  const array = (v, path) => { check(Array.isArray(v), path, 'Expected an array'); return Array.isArray(v) ? v : []; };
  const diagnostics = (items, prefix = 'diagnostics') => {
    for (const [i, d] of array(items, prefix).entries()) {
      const p = `${prefix}[${i}]`;
      check(record(d), p, 'Expected a diagnostic'); if (!record(d)) continue;
      check(codes.has(d.code), `${p}.code`, 'Unknown diagnostic code');
      check(['error', 'warning'].includes(d.severity), `${p}.severity`, 'Expected error or warning');
      check(nonempty(d.message), `${p}.message`, 'Expected a message');
      for (const field of ['slideId', 'nodeId']) check(d[field] === null || nonempty(d[field]), `${p}.${field}`, 'Expected ID or explicit null');
      check(d.sourceRef === null || record(d.sourceRef), `${p}.sourceRef`, 'Expected source reference or explicit null');
    }
  };
  const assets = (items, prefix) => {
    const seen = new Set();
    for (const [i, a] of array(items, prefix).entries()) {
      const p = `${prefix}[${i}]`;
      if (!record(a)) { check(false, p, 'Expected an asset'); continue; }
      check(nonempty(a.id) && !seen.has(a.id), `${p}.id`, 'Asset ID is missing or duplicated'); seen.add(a.id);
      check(nonempty(a.mime), `${p}.mime`, 'Expected MIME type');
      check(digest(a.sha256), `${p}.sha256`, 'Expected SHA-256');
      check(a.bytes instanceof Uint8Array && a.bytes.length > 0, `${p}.bytes`, 'Expected asset bytes');
    }
    return seen;
  };
  const deck = (d, prefix = '') => {
    const p = key => prefix + key;
    if (!record(d)) { check(false, prefix || 'model', 'Expected a deck'); return; }
    check([1,2].includes(d.version), p('version'), 'Unsupported version');
    check(digest(d.sourceDigest), p('sourceDigest'), 'Expected SHA-256');
    const assetIds = assets(d.assets, p('assets')), slideIds = new Set();
    if(d.fonts!==undefined){
      check(d.version===2,p('fonts'),'Fonts require model version 2');
      const fonts=array(d.fonts,p('fonts')),seen=new Set(),usedAssets=new Set();let bytes=0;
      check(fonts.length<=32,p('fonts'),'Too many font faces');
      for(const [i,font] of fonts.entries()){
        const fp=p(`fonts[${i}]`);if(!record(font)){check(false,fp,'Expected font face');continue;}
        check(nonempty(font.id)&&!seen.has(font.id),`${fp}.id`,'Font ID missing or duplicated');seen.add(font.id);
        for(const key of ['cssFamily','family'])check(nonempty(font[key]),`${fp}.${key}`,'Expected font family');
        check([400,700].includes(font.weight),`${fp}.weight`,'Unsupported font weight');check(['normal','italic'].includes(font.style),`${fp}.style`,'Unsupported font style');
        check(['supplied','pptx'].includes(font.origin),`${fp}.origin`,'Expected font origin');
        const asset=(Array.isArray(d.assets)?d.assets:[]).find(a=>a?.id===font.assetId);check(assetIds.has(font.assetId)&&typeof asset?.mime==='string'&&asset.mime.startsWith('font/'),`${fp}.assetId`,'Unknown font asset');
        if(asset&&!usedAssets.has(asset.id)){usedAssets.add(asset.id);bytes+=asset.bytes?.length??0;}
      }
      check(bytes<=32*1024*1024,p('fonts'),'Font bytes exceed limit');
    }
    const slides = array(d.slides, p('slides'));let tableCells=0,chartPoints=0;
    check(slides.length > 0, p('slides'), 'Expected at least one slide');
    if(slides.every(record))for(const message of canvasDiagnostics({...d,slides}))check(false,p('pptxSizeEmu'),message);
    for (const [i, slide] of slides.entries()) {
      const sp = p(`slides[${i}]`);
      if (!record(slide)) { check(false, sp, 'Expected slide'); continue; }
      check(nonempty(slide.id) && !slideIds.has(slide.id), `${sp}.id`, 'Slide ID missing or duplicated'); slideIds.add(slide.id);
      for (const k of ['widthCssPx', 'heightCssPx']) check(Number.isFinite(slide[k]) && slide[k] > 0, `${sp}.${k}`, 'Expected positive finite dimension');
      const nodes = array(slide.nodes, `${sp}.nodes`), byId = new Map();
      for (const [j, n] of nodes.entries()) {
        const np = `${sp}.nodes[${j}]`;
        if (!record(n)) { check(false, np, 'Expected node'); continue; }
        check(nonempty(n.id) && !byId.has(n.id), `${np}.id`, 'Object ID missing or duplicated'); byId.set(n.id, n);
        check(kinds.has(n.kind), `${np}.kind`, 'Unsupported node kind');
        for (const k of ['x', 'y', 'w', 'h']) {
          const v = n.boxCssPx?.[k];
          check(Number.isFinite(v) && (!['w', 'h'].includes(k) || (n.kind === 'line' ? v >= 0 : v > 0)), `${np}.boxCssPx.${k}`, 'Invalid coordinate or dimension');
        }
        check(Number.isInteger(n.zOrder) && n.zOrder >= 0, `${np}.zOrder`, 'Invalid stacking order');
        check(record(n.content), `${np}.content`, 'Expected content');
        check(record(n.style), `${np}.style`, 'Expected style');
        check(record(n.sourceRef) && n.sourceRef.slideId === slide.id && (nonempty(n.sourceRef.elementId) || nonempty(n.sourceRef.path)), `${np}.sourceRef`, 'Expected source location on this slide');
        check(n.parentId === null || nonempty(n.parentId), `${np}.parentId`, 'Expected parent or null');
        if (n.kind === 'text') check(typeof n.content?.text === 'string', `${np}.content.text`, 'Expected text');
        if (n.kind === 'image') {
          check(assetIds.has(n.content?.assetId), `${np}.content.assetId`, 'Unknown asset reference');
          if(n.content?.crop!==undefined){
            const crop=n.content.crop;
            check(record(crop)&&['left','top','right','bottom'].every(key=>Number.isInteger(crop[key])&&crop[key]>=0&&crop[key]<100000)&&crop.left+crop.right<100000&&crop.top+crop.bottom<100000,`${np}.content.crop`,'Invalid image crop');
          }
        }
        if(n.kind==='table'){
          check(d.version===2,`${np}.kind`,'Tables require model version 2');
          const table=validateTable(n.content);for(const error of table.errors)check(false,`${np}.content.${error.path}`,error.message);
          if(table.valid){
            tableCells+=n.content.rows*n.content.columns;
            check(Math.abs(n.content.columnWidthsCssPx.reduce((a,b)=>a+b,0)-n.boxCssPx?.w)<=.05&&Math.abs(n.content.rowHeightsCssPx.reduce((a,b)=>a+b,0)-n.boxCssPx?.h)<=.05,`${np}.boxCssPx`,'Table grid dimensions do not match its box');
          }
        }
        if(n.kind==='chart'){
          check(d.version===2,`${np}.kind`,'Charts require model version 2');
          const chart=validateChart(n.content);for(const error of chart.errors)check(false,`${np}.content.${error.path}`,error.message);
          if(chart.valid)chartPoints+=n.content.categories.length*n.content.series.length;
          const s=n.style??{};
          check(nonempty(s.fontFamily)&&Number.isFinite(s.fontSize)&&s.fontSize>0&&s.fontWeight===400&&s.fontStyle==='normal'&&s.color==='#223845'&&(s.opacity??1)===1,`${np}.style`,'Chart text uses one explicit positive font size and regular style');
          check(Object.keys(s).every(key=>['fontFamily','fontSize','fontWeight','fontStyle','color','opacity','cssZIndex'].includes(key)),`${np}.style`,'Unsupported chart display effect');
        }
      }
      for (const [j, n] of nodes.entries()) {
        if (!record(n) || n.parentId === null) continue;
        const visited = new Set([n.id]); let parent = n.parentId;
        while (parent !== null) {
          const target = byId.get(parent);
          if (!target || target.kind !== 'group' || visited.has(parent)) {
            check(false, `${sp}.nodes[${j}].parentId`, 'Missing, non-group or cyclic parent'); break;
          }
          visited.add(parent); parent = target.parentId;
        }
      }
    }
    check(tableCells<=10000,p('slides'),'Document exceeds 10000 table cells');
    check(chartPoints<=10000,p('slides'),'Document exceeds 10000 chart points');
  };
  if (!record(value)) check(false, '$', 'Expected a contract object');
  else if ('schemaVersion' in value) {
    check(value.schemaVersion === 1, 'schemaVersion', 'Unsupported version');
    check(['project-data-v1', 'static-slides-v1', 'svg-slides-v1', 'editable-slides-v1', 'editable-slides-v2'].includes(value.profile), 'profile', 'Unsupported input profile');
    check(typeof value.source?.html === 'string', 'source.html', 'Expected HTML');
    check(digest(value.source?.sha256), 'source.sha256', 'Expected SHA-256');
    check(value.source?.slideSelector === null || nonempty(value.source?.slideSelector), 'source.slideSelector', 'Expected selector or null');
    check(value.profile === 'project-data-v1' ? record(value.project) : value.project === null, 'project', 'Project does not match profile');
    assets(value.assets, 'assets'); diagnostics(value.diagnostics);
  } else if ('model' in value) {
    deck(value.model, 'model.'); array(value.measurements, 'measurements');
    check(record(value.environment), 'environment', 'Expected explicit environment'); diagnostics(value.diagnostics);
  } else if ('status' in value) {
    check(['generated', 'failed'].includes(value.status), 'status', 'Expected generated or failed');
    check(value.status === 'generated' ? value.pptxBytes instanceof Uint8Array && value.pptxBytes.length > 0 : value.pptxBytes === null, 'pptxBytes', 'Bytes do not match export status');
    array(value.objectMap, 'objectMap'); check(record(value.profile), 'profile', 'Expected export profile'); diagnostics(value.diagnostics);
    if (value.status === 'generated') check(!value.diagnostics?.some(d => d.severity === 'error'), 'diagnostics', 'Generated output cannot have blocking errors');
  } else if ('version' in value) deck(value);
  else check(false, '$', 'Unknown contract');
  return {valid: errors.length === 0, errors};
}
