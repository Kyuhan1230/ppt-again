import {parse} from 'parse5';
import * as cssTree from 'css-tree';
import {assetPath, resolveAsset, sha256} from './assets.mjs';
import {diagnostic} from './model.mjs';

export const attribute = (node, name) => node.attrs?.find(a => a.name === name)?.value;
export function* walkHtml(node) { yield node; for (const child of node.childNodes ?? []) yield* walkHtml(child); }
const nodeText = node => (node.childNodes ?? []).filter(n=>n.nodeName==='#text').map(n=>n.value).join('');
const location = node => ({sourceRef:{path:`${node.tagName ?? node.nodeName}@${node.sourceCodeLocation?.startOffset ?? 0}`}});

export async function readInput({html, assets: provided = [], slideSelector = null} = {}) {
  const original = typeof html === 'string' ? html : '';
  const input = {schemaVersion:1,profile:'static-slides-v1',source:{html:original,sha256:sha256(original),slideSelector},
    assets:[],project:null,diagnostics:[],styles:[],fontFaces:[],assetRefs:[]};
  const fail = (code,message,loc,severity) => input.diagnostics.push(diagnostic(code,message,loc,severity));
  if (!original.trim() || Buffer.byteLength(original)>16*1024*1024) { fail('INPUT_INVALID','HTML 문자열이 없거나 16MiB를 초과합니다.'); return input; }
  if (slideSelector !== null && (typeof slideSelector!=='string'||!slideSelector.trim())) fail('INPUT_INVALID','장표 선택자는 비어 있지 않은 문자열이어야 합니다.');
  if (!Array.isArray(provided)) { fail('INPUT_INVALID','제공 자산 목록은 배열이어야 합니다.'); provided=[]; }
  const paths = new Set(); let total = 0;
  for (const a of provided) {
    try {
      const p=assetPath(a.path);if(paths.has(p))throw new Error(`중복 자산 경로: ${p}`);paths.add(p);
      if (!(a.bytes instanceof Uint8Array)) throw new Error(`자산 바이트가 없습니다: ${p}`);
      total+=a.bytes.length;
    } catch(e) { fail('INPUT_INVALID',e.message,{sourceRef:{path:'assets'}}); }
  }
  if(total>64*1024*1024) { fail('INPUT_INVALID','제공 자산의 총 크기가 64MiB를 초과합니다.'); return input; }
  if(input.diagnostics.some(d=>d.severity==='error'))return input;
  const byId = new Map();
  const addAsset = (url,loc,usage='image',basePath=null) => {
    try {
      const a=resolveAsset(url,provided,{basePath});
      if ((usage==='image' && !a.mime.startsWith('image/')) || (usage==='font'&&!a.mime.startsWith('font/')) || (usage==='stylesheet'&&a.mime!=='text/css')) throw new Error('자산의 용도와 MIME이 다릅니다.');
      if(!byId.has(a.id)){byId.set(a.id,a);input.assets.push(a);}
      input.assetRefs.push({...loc,url,basePath,assetId:a.id,usage}); return a;
    } catch(e) { fail(usage==='font'?'FONT_UNAVAILABLE':'ASSET_MISSING',e.message,loc); return null; }
  };
  const scanCss = (css,loc,basePath=null) => {
    input.styles.push({css,...loc,basePath});
    try {
      const ast=cssTree.parse(css);
      cssTree.walk(ast,function(n){
        if(n.type==='Atrule'&&n.name.toLowerCase()==='import')fail('UNSUPPORTED_OBJECT','CSS @import는 제공된 단일 스타일시트로 합쳐 주세요.',loc);
        if(n.type==='Atrule'&&n.name.toLowerCase()==='font-face') {
          const declarations={};n.block?.children.forEach(d=>{if(d.type==='Declaration')declarations[d.property]=cssTree.generate(d.value);});
          input.fontFaces.push({family:(declarations['font-family']??'').replace(/^["']|["']$/g,''),weight:declarations['font-weight']??'normal',style:declarations['font-style']??'normal',css:cssTree.generate(n),...loc});
        }
        if(n.type==='Url'&&!n.value.startsWith('#')) addAsset(n.value,loc,this.atrule?.name.toLowerCase()==='font-face'?'font':'image',basePath);
      });
    }catch(e){fail('INPUT_INVALID',`CSS를 해석할 수 없습니다: ${e.message}`,loc);}
  };
  const tree = parse(original,{sourceCodeLocationInfo:true});
  const nodes=[...walkHtml(tree)];
  const dataNodes=nodes.filter(n=>n.tagName==='script'&&attribute(n,'id')==='project-data');
  if(dataNodes.length) {
    input.profile='project-data-v1';
    if(dataNodes.length!==1)fail('INPUT_INVALID','project-data 스크립트는 한 개여야 합니다.',location(dataNodes[0]));
    else {
      try { input.project=JSON.parse(nodeText(dataNodes[0])); }
      catch { fail('INPUT_INVALID','project-data의 JSON이 올바르지 않습니다.',location(dataNodes[0])); }
    }
  }
  // A presentation-sized outer SVG already declares its slide coordinate system.
  // Keep explicit HTML roots authoritative; small toolbar icons are not slides.
  if(input.profile==='static-slides-v1'&&!slideSelector&&!nodes.some(n=>attribute(n,'data-slide')!==undefined)) {
    const roots=nodes.filter(n=>{
      if(n.tagName!=='svg')return false;
      for(let p=n.parentNode;p;p=p.parentNode)if(p.tagName==='svg')return false;
      const v=(attribute(n,'viewBox')??'').trim().split(/[ ,]+/).map(Number);
      return v.length===4&&v.every(Number.isFinite)&&v[2]>=300&&v[3]>=200;
    });
    if(roots.length){input.profile='svg-slides-v1';input.svgRootPaths=roots.map(n=>location(n).sourceRef.path);}
  }
  for(const n of nodes) {
    if(n.tagName==='style')scanCss(nodeText(n),location(n));
    if(n.tagName==='link'&&attribute(n,'rel')?.toLowerCase().split(/\s+/).includes('stylesheet')) {
      const a=addAsset(attribute(n,'href'),location(n),'stylesheet');if(a)scanCss(Buffer.from(a.bytes).toString('utf8'),location(n),a.path);
    }
    if(attribute(n,'style'))scanCss(`x{${attribute(n,'style')}}`,location(n));
    if(input.profile!=='project-data-v1'&&n.tagName==='img')addAsset(attribute(n,'src'),location(n));
  }
  const p=input.project;
  if(input.profile==='project-data-v1'&&p!==null) {
    if(typeof p!=='object'||Array.isArray(p)||p.version!==1||!Array.isArray(p.slides)||!p.slides.length) {
      fail('INPUT_INVALID','project-data v1의 장표 목록이 필요합니다.',{sourceRef:{path:'project-data'}});input.project=null;return input;
    }
    for(const k of ['width','height'])if(!Number.isFinite(p[k])||p[k]<=0)fail('INPUT_INVALID',`장표 ${k}는 양수여야 합니다.`,{sourceRef:{path:`project-data.${k}`}});
    const slideIds=new Set();
    for(const [i,s] of p.slides.entries()) {
      if(!s||typeof s!=='object'){fail('INPUT_INVALID','잘못된 장표입니다.',{sourceRef:{path:`project-data.slides[${i}]`}});continue;}
      const loc={slideId:s.id??null,sourceRef:{slideId:s.id??null,path:`project-data.slides[${i}]`}};
      if(typeof s.id!=='string'||!s.id||slideIds.has(s.id))fail('INPUT_INVALID','장표 ID가 없거나 중복됩니다.',loc);slideIds.add(s.id);
      if(!Array.isArray(s.nodes)){fail('INPUT_INVALID','장표 객체 목록이 없습니다.',loc);continue;}
      const ids=new Set();
      for(const [j,n] of s.nodes.entries()) {
        if(!n||typeof n!=='object'){fail('INPUT_INVALID','잘못된 객체입니다.',{...loc,sourceRef:{...loc.sourceRef,path:`${loc.sourceRef.path}.nodes[${j}]`}});continue;}
        const at={slideId:s.id??null,nodeId:n.id??null,sourceRef:{slideId:s.id??null,elementId:n.id??null,path:`${loc.sourceRef.path}.nodes[${j}]`}};
        if(typeof n.id!=='string'||!n.id||ids.has(n.id))fail('INPUT_INVALID','객체 ID가 없거나 중복됩니다.',at);ids.add(n.id);
        for(const k of ['x','y','w','h'])if(!Number.isFinite(n[k])||(['w','h'].includes(k)&&(n.type==='line'?n[k]<0:n[k]<=0)))fail('INPUT_INVALID',`객체 ${k} 좌표/크기가 올바르지 않습니다.`,at);
        if(!['text','shape','line','svg','image'].includes(n.type))fail('UNSUPPORTED_OBJECT',`지원하지 않는 객체 종류: ${n.type}`,at);
        if(n.rotation)fail('UNSUPPORTED_OBJECT','회전 객체는 아직 지원하지 않습니다.',at);
        if(n.type==='text') {
          if(typeof n.text!=='string'||!Number.isFinite(n.fontSize)||n.fontSize<=0)fail('INPUT_INVALID','텍스트 문구 또는 글자 크기가 올바르지 않습니다.',at);
        }
        if(['svg','image'].includes(n.type))addAsset(n.src,at);
      }
    }
  }
  return input;
}
