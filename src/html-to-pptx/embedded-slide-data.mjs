import {parseFragment,serialize} from 'parse5';
import * as cssTree from 'css-tree';
import {resolveAsset} from './assets.mjs';

const attribute=(node,name)=>node.attrs?.find(a=>a.name===name)?.value;
const record=value=>value!==null&&typeof value==='object'&&!Array.isArray(value);
const text=node=>(node.childNodes??[]).map(n=>n.value??'').join('');
const walk=function*(node){yield node;for(const child of node.childNodes??[])yield*walk(child);};
const escape=value=>String(value).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const reject=(message,code='INPUT_INVALID')=>{throw Object.assign(new Error(message),{code});};

// This lexer extracts one literal declaration; it never evaluates the rest of
// a source program. Regex literals and ambiguous slash syntax are unsupported.
function styleDeclarations(code){
  const tokens=[];let i=0;
  while(i<code.length){
    const c=code[i];if(/\s/.test(c)){i++;continue;}
    if(code.startsWith('//',i)){i=code.indexOf('\n',i+2);if(i<0)break;continue;}
    if(code.startsWith('/*',i)){const end=code.indexOf('*/',i+2);if(end<0)reject('스타일 선언 주석이 닫히지 않았습니다.');i=end+2;continue;}
    if(c==='"'||c==="'"){
      const quote=c;let closed=false;i++;
      while(i<code.length){if(code[i]==='\\'){i+=2;continue;}if(code[i++]===quote){closed=true;break;}}
      if(!closed)reject('스타일 선언의 문자열이 닫히지 않았습니다.');tokens.push({type:'string'});continue;
    }
    if(c==='`'){
      let value='',closed=false;const slots=[];i++;
      while(i<code.length){
        if(code[i]==='`'){i++;closed=true;break;}
        if(code[i]==='\\'){
          const next=code[++i],escapes={'n':'\n','r':'\r','t':'\t','\\':'\\','`':'`','$':'$'};
          if(next==='\n'){i++;continue;}if(!Object.hasOwn(escapes,next))reject('지원하지 않는 스타일 문자열 이스케이프입니다.','UNSUPPORTED_OBJECT');value+=escapes[next];i++;continue;
        }
        if(code.startsWith('${',i)){
          const end=code.indexOf('}',i+2),name=end<0?'':code.slice(i+2,end).trim();
          if(!/^[A-Za-z_$][\w$]*$/.test(name))reject('스타일 안의 계산·함수·속성 보간은 실행하지 않습니다.','UNSUPPORTED_OBJECT');
          slots.push(name);value+='${'+name+'}';i=end+1;continue;
        }
        value+=code[i++];
      }
      if(!closed)reject('스타일 template가 닫히지 않았습니다.');tokens.push({type:'template',value,slots});continue;
    }
    if(/[A-Za-z_$]/.test(c)){const start=i++;while(i<code.length&&/[\w$]/.test(code[i]))i++;tokens.push({type:'identifier',value:code.slice(start,i)});continue;}
    if(/\d/.test(c)){const start=i++;while(i<code.length&&/[\d.]/.test(code[i]))i++;tokens.push({type:'number',value:code.slice(start,i)});continue;}
    if(c==='/'&&(!['identifier','number'].includes(tokens.at(-1)?.type)||['return','throw','yield','await'].includes(tokens.at(-1)?.value)||!/^\s*\d/.test(code.slice(i+1))))reject('정규식 또는 모호한 원본 코드에서 스타일을 추출하지 않습니다.','UNSUPPORTED_OBJECT');
    tokens.push({type:c,value:c});i++;
  }
  const found=[];
  for(let j=0;j<tokens.length;j++){
    const token=tokens[j];if(token.type!=='template'||!/^\s*<style(?:\s|>)/i.test(token.value))continue;
    if(!['const','let'].includes(tokens[j-3]?.value)||tokens[j-2]?.type!=='identifier'||tokens[j-1]?.type!=='='||tokens[j+1]?.type!==';')reject('스타일은 계산이나 호출 없는 단일 literal 선언이어야 합니다.','UNSUPPORTED_OBJECT');
    found.push(token);
  }
  return found;
}

function decodeStyle(template,font){
  if(!template)return {html:'',frame:null,fontUsed:false};
  const fragment=parseFragment(template.value),nodes=fragment.childNodes.filter(n=>n.nodeName!=='#text'||n.value.trim());
  if(!nodes.length||nodes.some(n=>n.tagName!=='style'))reject('추출한 template에는 style 요소만 있어야 합니다.','UNSUPPORTED_OBJECT');
  if(new Set(template.slots).size>1)reject('스타일에는 내장 글꼴 URL 슬롯 한 종류만 사용할 수 있습니다.','UNSUPPORTED_OBJECT');
  let replaced=0;const frames=new Map();
  for(const node of nodes){
    const ast=cssTree.parse(text(node));
    cssTree.walk(ast,function(n){
      if(n.type==='Url'&&n.value.includes('${')){
        if(!font||n.value!=='${'+template.slots[0]+'}'||this.atrule?.name.toLowerCase()!=='font-face'||this.declaration?.property.toLowerCase()!=='src')reject('내장 글꼴 슬롯은 font-face src의 URL 전체에만 사용할 수 있습니다.','UNSUPPORTED_OBJECT');
        n.value=font.path;replaced++;
      }
      if(n.type==='Rule'&&!this.atrule&&n.prelude?.type==='SelectorList')for(const selector of n.prelude.children){
        const parts=selector.children.toArray();if(parts.length!==1||parts[0].type!=='IdSelector')continue;
        const id=parts[0].name,frame=frames.get(id)??{id};
        for(const declaration of n.block.children){
          if(declaration.type!=='Declaration'||!['width','height'].includes(declaration.property))continue;
          const values=declaration.value.children.toArray();if(values.length===1&&values[0].type==='Dimension'&&values[0].unit==='px')frame[declaration.property]=Number(values[0].value);
        }
        frames.set(id,frame);
      }
    });
    node.childNodes=[{nodeName:'#text',value:cssTree.generate(ast),parentNode:node}];
  }
  if(replaced!==template.slots.length||serialize(fragment).includes('${'))reject('글꼴 URL 외의 스타일 보간은 지원하지 않습니다.','UNSUPPORTED_OBJECT');
  const candidates=[...frames.values()].filter(f=>/^[A-Za-z_][\w-]*$/.test(f.id)&&[f.width,f.height].every(v=>Number.isFinite(v)&&v>0));
  if(candidates.length!==1)reject('정적 스타일에서 장표의 고정 크기 영역을 한 개로 확인할 수 없습니다.','SLIDE_ROOT_REQUIRED');
  return {html:serialize(fragment),frame:candidates[0],fontUsed:replaced>0};
}

// The pages[].html strings are the content authority. Source styles/defs and an
// optional static style literal supplement them; source templates are not data.
export function readEmbeddedSlideData(tree){
  const result={matched:false,documents:[],diagnostics:[]};
  try{
    const scripts=[...walk(tree)].filter(n=>n.tagName==='script'),candidates=[];
    for(const script of scripts.filter(n=>attribute(n,'type')?.trim().toLowerCase()==='application/json')){
      let value;try{value=JSON.parse(text(script));}catch{continue;}
      if(record(value)&&Object.hasOwn(value,'pages')&&Object.hasOwn(value,'sources'))candidates.push(value);
    }
    if(!candidates.length)return result;result.matched=true;
    if(candidates.length!==1)reject('장표 데이터 선언은 한 개여야 합니다.');const data=candidates[0];
    if(!Array.isArray(data.pages)||!data.pages.length||data.pages.length>200||!record(data.sources))reject('장표 데이터에는 1~200장의 pages와 sources 객체가 필요합니다.');
    let font;
    if(data.font!==undefined){
      if(!record(data.font)||typeof data.font.mime!=='string'||!data.font.mime.startsWith('font/')||typeof data.font.base64!=='string')reject('내장 글꼴 데이터 형식이 올바르지 않습니다.','FONT_INVALID');
      try{const asset=resolveAsset(`data:${data.font.mime};base64,${data.font.base64}`);font={path:`embedded-font-${asset.sha256}.${asset.mime.slice(5)}`,mime:asset.mime,bytes:asset.bytes};}catch{reject('내장 글꼴의 형식과 바이트를 확인할 수 없습니다.','FONT_INVALID');}
    }
    const declarations=scripts.filter(n=>attribute(n,'type')?.trim().toLowerCase()!=='application/json').flatMap(n=>styleDeclarations(text(n)));
    if(declarations.length>1)reject('정적 스타일 template를 한 개로 확정할 수 없습니다.');
    const style=decodeStyle(declarations[0],font);
    if(font&&!style.fontUsed)reject('내장 글꼴과 실제 font-face URL 선언의 연결을 확인할 수 없습니다.','FONT_SOURCE_UNSUPPORTED');
    let renderedBytes=0,characters=0,objects=0;
    for(const [index,page] of data.pages.entries()){
      if(!record(page)||typeof page.html!=='string'||!page.html.trim()||typeof page.source!=='string'||!Object.hasOwn(data.sources,page.source))reject('현재 장표 HTML 또는 참조 source가 없습니다.');
      const source=data.sources[page.source];if(!record(source)||typeof source.styles!=='string'||source.defs!==undefined&&typeof source.defs!=='string')reject('source의 styles와 defs는 정적 HTML 문자열이어야 합니다.');
      const head=parseFragment(source.styles);if(head.childNodes.some(n=>n.nodeName==='#text'?n.value.trim():!['style','link','#comment'].includes(n.tagName??n.nodeName)))reject('source styles에는 스타일 선언만 사용할 수 있습니다.','UNSUPPORTED_OBJECT');
      const defs=parseFragment(source.defs??''),defsNodes=[...walk(defs)],definitionRoots=defs.childNodes.filter(n=>n.nodeName!=='#comment'&&(n.nodeName!=='#text'||n.value.trim()));
      if(definitionRoots.some(n=>n.tagName!=='svg'||n.namespaceURI!=='http://www.w3.org/2000/svg'||n.childNodes.some(c=>c.nodeName==='#text'?c.value.trim():c.nodeName!=='#comment'&&c.tagName!=='defs'))||defsNodes.some(n=>n.tagName&&(n.namespaceURI!=='http://www.w3.org/2000/svg'||['script','foreignObject','style','a','image','set','discard'].includes(n.tagName)||n.tagName.startsWith('animate')||n.attrs.some(a=>a.name.startsWith('on')||['data-slide','data-slidehandoff'].includes(a.name)))))reject('defs에는 실행·본문·장표가 없는 정적 SVG 정의만 사용할 수 있습니다.','UNSUPPORTED_OBJECT');
      const fragment=parseFragment(page.html),roots=fragment.childNodes.filter(n=>n.nodeName!=='#comment'&&(n.nodeName!=='#text'||n.value.trim())&&!['style','script'].includes(n.tagName));
      const pageNodes=[...walk(fragment)],pageObjects=pageNodes.filter(n=>n.tagName).length+defsNodes.filter(n=>n.tagName).length+(style.frame?1:0);
      if(pageObjects>10000||(objects+=pageObjects)>100000)reject('내장 장표의 객체 수 한도를 초과했습니다.','LIMIT_EXCEEDED');
      for(const node of [...pageNodes,...defsNodes])if(node.nodeName==='#text'&&!['style','script'].includes(node.parentNode?.tagName))if(node.value.length>100000||(characters+=node.value.length)>2000000)reject('내장 장표의 문구가 글자 수 한도를 초과했습니다.','LIMIT_EXCEEDED');
      if(roots.length!==1||!roots[0].tagName)reject('각 장표 HTML에는 표시 루트 한 개가 필요합니다.','SLIDE_ROOT_REQUIRED');
      if(pageNodes.some(n=>attribute(n,'data-slidehandoff')!==undefined||attribute(n,'data-slide')!==undefined&&(n!==roots[0]||style.frame)||n.tagName==='script'&&(attribute(n,'type')?.trim().toLowerCase()!=='application/json'||!attribute(n,'data-chart-for')||attribute(n,'id')==='project-data')))reject('장표 안에는 중첩 문서 데이터·스크립트·추가 장표를 넣을 수 없습니다. 차트 JSON만 지원합니다.','UNSUPPORTED_OBJECT');
      let body;
      if(style.frame){
        if([...walk(fragment)].some(n=>attribute(n,'id')===style.frame.id))reject('현재 장표와 보완 스타일의 루트 ID가 중복됩니다.');
        body=`<div id="${escape(style.frame.id)}" data-slide style="position:relative!important;left:0!important;top:0!important;transform:none!important;transform-origin:0 0!important;margin:0!important">${serialize(fragment)}</div>`;
      }else{
        if(roots[0].tagName!=='svg')roots[0].attrs=[...(roots[0].attrs??[]).filter(a=>a.name!=='data-slide'),{name:'data-slide',value:''}];body=serialize(fragment);
      }
      const html=`<!doctype html><html><head><meta charset="utf-8">${source.styles}${style.html}</head><body>${source.defs??''}${body}</body></html>`;
      if((renderedBytes+=Buffer.byteLength(html))>16*1024*1024)reject('정적 장표 문서 합계가 입력 크기 한도를 초과했습니다.','LIMIT_EXCEEDED');
      result.documents.push({html,assets:font?[font]:[]});
    }
    return result;
  }catch(error){result.diagnostics.push({code:error.code??'INPUT_INVALID',severity:'error',slideId:null,nodeId:null,sourceRef:null,message:error.message});return result;}
}
