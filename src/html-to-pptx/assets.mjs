import {createHash} from 'node:crypto';
import path from 'node:path';
export const sha256 = bytes => createHash('sha256').update(bytes).digest('hex');
const mimeByExtension = {'.png':'image/png','.jpg':'image/jpeg','.jpeg':'image/jpeg','.svg':'image/svg+xml','.woff':'font/woff','.woff2':'font/woff2','.ttf':'font/ttf','.otf':'font/otf','.css':'text/css'};
export function assetPath(value) {
  if (typeof value !== 'string') throw new Error('제공 자산 경로가 필요합니다.');
  let decoded; try { decoded = decodeURIComponent(value); } catch { throw new Error('잘못된 자산 경로입니다.'); }
  if (!decoded || /[\\\x00-\x1f?#]/.test(decoded) || decoded.startsWith('/') || /^[a-z][a-z\d+.-]*:/i.test(decoded)) throw new Error('제공된 상대 자산 경로만 허용합니다.');
  const normalized = path.posix.normalize(decoded);
  if (normalized === '..' || normalized.startsWith('../') || normalized.startsWith('/') || normalized === '.') throw new Error('제공 범위 밖 자산 경로입니다.');
  return normalized;
}
function checkBytes(bytes, mime) {
  if (!(bytes instanceof Uint8Array) || bytes.length === 0 || bytes.length > 16 * 1024 * 1024) throw new Error('자산 바이트가 없거나 16MiB를 초과합니다.');
  const b = Buffer.from(bytes), text = () => b.toString('utf8');
  const valid = mime === 'image/png' ? b.subarray(0,8).equals(Buffer.from([137,80,78,71,13,10,26,10]))
    : mime === 'image/jpeg' ? b[0]===255 && b[1]===216 && b[2]===255
    : mime === 'image/svg+xml' ? /<svg(?:\s|>)/i.test(text())
    : mime === 'font/woff2' ? b.subarray(0,4).toString()==='wOF2'
    : mime === 'font/woff' ? b.subarray(0,4).toString()==='wOFF'
    : mime === 'font/ttf' ? b.subarray(0,4).equals(Buffer.from([0,1,0,0]))
    : mime === 'font/otf' ? b.subarray(0,4).toString()==='OTTO'
    : mime === 'text/css';
  if (!valid) throw new Error('지원 자산의 MIME과 바이트가 일치하지 않습니다.');
  if (mime === 'image/svg+xml') {
    const svg=text();
    const urls=[...svg.matchAll(/url\(\s*(?:"([^"]*)"|'([^']*)'|([^\s)]+))\s*\)/gi)].map(m=>m[1]??m[2]??m[3]);
    const refs=[...svg.matchAll(/(?:^|\s)(?:xlink:)?href\s*=\s*["']([^"']*)["']/gi)].map(m=>m[1]);
    if (/<(?:script|foreignObject|iframe)\b|<!ENTITY/i.test(svg) || /\bon\w+\s*=/i.test(svg) || [...urls,...refs].some(u=>!u.startsWith('#'))) throw new Error('SVG의 스크립트·외부 참조는 지원하지 않습니다.');
  }
  return b;
}
export function resolveAsset(url, provided = [], {basePath = null} = {}) {
  let bytes, mime, sourcePath = null;
  if (typeof url !== 'string') throw new Error('자산 URL이 없습니다.');
  if (url.startsWith('data:')) {
    const match = /^data:([^;,]+)(;base64)?,([\s\S]*)$/i.exec(url);
    if (!match) throw new Error('잘못된 내장 자산입니다.');
    mime = match[1].toLowerCase();
    if (match[2]) {
      if (!/^(?:[A-Za-z0-9+/]{4})*(?:[A-Za-z0-9+/]{2}==|[A-Za-z0-9+/]{3}=)?$/.test(match[3])) throw new Error('잘못된 base64 자산입니다.');
      bytes = Buffer.from(match[3], 'base64');
    } else { try { bytes = Buffer.from(decodeURIComponent(match[3])); } catch { throw new Error('잘못된 내장 자산 인코딩입니다.'); } }
  } else {
    if(basePath) {
      if (/^(?:\/|[a-z][a-z\d+.-]*:)|\\/i.test(url)) throw new Error('제공된 상대 자산 경로만 허용합니다.');
      sourcePath=assetPath(path.posix.join(path.posix.dirname(assetPath(basePath)),decodeURIComponent(url)));
    } else sourcePath = assetPath(url);
    const found = provided.find(a => assetPath(a.path) === sourcePath);
    if (!found) throw new Error(`제공 자산이 없습니다: ${sourcePath}`);
    mime = found.mime ?? mimeByExtension[path.posix.extname(sourcePath).toLowerCase()]; bytes = found.bytes;
  }
  bytes = checkBytes(bytes,mime);
  const hash = sha256(bytes);
  return {id:`asset-${hash}`,path:sourcePath,mime,sha256:hash,bytes:new Uint8Array(bytes)};
}
export function assetDataUrl(asset) { return `data:${asset.mime};base64,${Buffer.from(asset.bytes).toString('base64')}`; }
