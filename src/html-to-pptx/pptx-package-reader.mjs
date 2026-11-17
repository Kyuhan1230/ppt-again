import path from 'node:path';
import {strFromU8} from 'fflate';
import {DOMParser} from '@xmldom/xmldom';
import {createPackageBudget,unzipBudgeted} from './package-budget.mjs';
import {validatePngStructure} from './png-duotone.mjs';

const elements=n=>[...(n?.childNodes??[])].filter(c=>c.nodeType===1);
const attr=(n,k,fallback=null)=>n?.hasAttribute(k)?n.getAttribute(k):fallback;
const REL='http://schemas.openxmlformats.org/package/2006/relationships',CT='http://schemas.openxmlformats.org/package/2006/content-types';

// Validate the whole package before any slide is visited. The returned budget
// is shared with every linked workbook during subsequent chart parsing.
export function openPptxPackage(bytes){
  // Every connected XLSX consumes the same package-wide expansion budget.
  const budget=createPackageBudget(),parts=unzipBudgeted(bytes,{budget});
  const validatedMathPngs=new Set();
  const validateMathPng=target=>{if(!validatedMathPngs.has(target)){validatePngStructure(parts[target]);validatedMathPngs.add(target);}};
  const xml=file=>{
    if(!parts[file])throw new Error('PPTX 필수 부분이 없습니다.');
    const s=strFromU8(parts[file]);if(/<!DOCTYPE|<!ENTITY/i.test(s))throw new Error('DTD·엔터티는 지원하지 않습니다.');
    try{return new DOMParser({onError:()=>{throw new Error('잘못된 PPTX XML입니다.');}}).parseFromString(s,'application/xml');}catch{throw new Error('잘못된 PPTX XML입니다.');}
  };
  const relationCache=new Map();
  const relationRecords=file=>{
    if(relationCache.has(file))return relationCache.get(file);
    const rfile=path.posix.join(path.posix.dirname(file),'_rels',path.posix.basename(file)+'.rels');
    const map=new Map();relationCache.set(file,map);if(!parts[rfile])return map;
    const doc=xml(rfile);if(doc.documentElement.namespaceURI!==REL||doc.documentElement.localName!=='Relationships')throw new Error('PPTX 관계 목록이 잘못되었습니다.');
    for(const r of elements(doc.documentElement)){
      const id=attr(r,'Id');if(r.namespaceURI!==REL||!id||map.has(id))throw new Error('PPTX 관계 ID 또는 네임스페이스가 잘못되었습니다.');
      if(r.localName!=='Relationship'||!['Internal',null].includes(attr(r,'TargetMode')))throw new Error('외부 참조가 있는 PPTX는 지원하지 않습니다.');
      const target=attr(r,'Target','');
      if(!target||/^\/\/|^[a-z][a-z\d+.-]*:|\\|[\u0000-\u001f?#]|%(?:2f|5c|3a|2e|00)/i.test(target))throw new Error('외부 PPTX 경로입니다.');
      const resolved=target.startsWith('/')?target.slice(1):path.posix.normalize(path.posix.join(path.posix.dirname(file),target));
      if(resolved.startsWith('../'))throw new Error('PPTX 경로가 범위를 벗어났습니다.');
      map.set(id,{target:resolved,type:attr(r,'Type')});
    }
    return map;
  };
  const relationships=file=>new Map([...relationRecords(file)].map(([id,r])=>[id,r.target]));
  const types=xml('[Content_Types].xml');if(types.documentElement.namespaceURI!==CT||types.documentElement.localName!=='Types')throw new Error('PPTX ContentType 목록이 잘못되었습니다.');
  const defaults=new Map(),overrides=new Map();
  for(const node of elements(types.documentElement)){
    if(node.namespaceURI!==CT||!['Default','Override'].includes(node.localName))throw new Error('PPTX ContentType 선언이 잘못되었습니다.');
    const isDefault=node.localName==='Default',key=isDefault?attr(node,'Extension','').toLowerCase():attr(node,'PartName',''),map=isDefault?defaults:overrides;
    if(!key||map.has(key))throw new Error('PPTX ContentType 선언이 중복되거나 없습니다.');map.set(key,attr(node,'ContentType'));
  }
  const contentTypeFor=part=>overrides.get('/'+part)??defaults.get(path.posix.extname(part).slice(1).toLowerCase());
  // Inspect every relationship, including unused external/OLE/script packages.
  for(const name of Object.keys(parts)){
    if(!name.endsWith('/')&&/vbaProject|activeX/i.test(name)||/macroEnabled|vbaProject|activeX|oleObject/i.test(contentTypeFor(name)??''))throw new Error('매크로·ActiveX·OLE 문서는 지원하지 않습니다.');
    if(name.endsWith('.rels')){
      if(name!=='_rels/.rels'&&path.posix.basename(path.posix.dirname(name))!=='_rels')throw new Error('PPTX 관계 부분의 경로가 잘못되었습니다.');
      const owner=name==='_rels/.rels'?'':path.posix.join(path.posix.dirname(path.posix.dirname(name)),path.posix.basename(name,'.rels'));
      for(const r of relationRecords(owner).values())if(/\/(?:oleObject|vbaProject|activeX|control)$/i.test(r.type??''))throw new Error('실행 가능한 내장 객체는 지원하지 않습니다.');
    }
  }
  return {budget,parts,xml,relationCache,relationRecords,relationships,contentTypeFor,validateMathPng};
}
