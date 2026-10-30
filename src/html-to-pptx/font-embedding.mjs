import {unzipSync,zipSync,strFromU8,strToU8} from 'fflate';
import {DOMParser,XMLSerializer} from '@xmldom/xmldom';
import {sha256} from './assets.mjs';
import {diagnostic} from './model.mjs';
import {inspectFont,decideFontEmbedding} from './font-policy.mjs';

const P='http://schemas.openxmlformats.org/presentationml/2006/main',R='http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const REL='http://schemas.openxmlformats.org/package/2006/relationships',CT='http://schemas.openxmlformats.org/package/2006/content-types';
const FONT_LIMIT=32*1024*1024,PPT_LIMIT=64*1024*1024;
const fail=(code,message)=>{throw Object.assign(new Error(message),{code});};
const children=node=>[...(node?.childNodes??[])].filter(n=>n.nodeType===1);
const slot=face=>face.style==='italic'?(face.weight===700?'boldItalic':'italic'):(face.weight===700?'bold':'regular');
const faceKey=face=>JSON.stringify([face.family.toLowerCase(),slot(face)]);

function checkedFace(face,asset){
  if(!asset||!(asset.bytes instanceof Uint8Array))fail('FONT_SOURCE_UNSUPPORTED','포함할 글꼴 자산이 없습니다.');
  const meta=inspectFont(asset.bytes),decision=decideFontEmbedding(meta,{forEditing:true,subset:false});
  if(!decision.allowed)fail(decision.reason,'이 글꼴은 편집용으로 포함할 수 없습니다.');
  if(face.family!==meta.family||face.weight!==meta.weight||face.style!==(meta.italic?'italic':'normal'))fail('FONT_STYLE_MISMATCH','선언된 family·스타일과 글꼴 파일이 다릅니다.');
  const hash=sha256(asset.bytes);if(hash!==asset.sha256)fail('FONT_INVALID','글꼴 자산의 지문이 일치하지 않습니다.');
  return {...face,bytes:asset.bytes,sha256:hash,meta};
}

export function planFontEmbedding(model){
  const faces=[],diagnostics=[],report=[];
  const issue=(code,message)=>diagnostics.push(diagnostic(code,message));
  const declared=model?.fonts??[];
  if(!Array.isArray(declared)||declared.length>32){issue('FONT_ASSET_LIMIT','글꼴은 최대 32개 face만 포함할 수 있습니다.');return {faces,diagnostics,report};}
  const assets=new Map(),referenced=new Set(declared.map(f=>f?.assetId));let total=0;
  for(const asset of model?.assets??[]){
    if(!referenced.has(asset?.id))continue;
    if(assets.has(asset.id)){issue('FONT_INVALID','글꼴 자산 ID가 중복됩니다.');return {faces,diagnostics,report};}
    assets.set(asset.id,asset);total+=asset.bytes?.length??0;
  }
  if(total>FONT_LIMIT){issue('FONT_ASSET_LIMIT','글꼴 합계가 32 MiB를 초과했습니다.');return {faces,diagnostics,report};}
  const slots=new Map();
  for(const face of declared){
    const entry={fontId:face?.id??null,family:face?.family??null,weight:face?.weight??null,style:face?.style??null,sha256:null,status:'planned',fullFont:true};
    try{
      if(!face||typeof face.id!=='string'||!face.id)fail('FONT_INVALID','글꼴 ID가 필요합니다.');
      const prepared=checkedFace(face,assets.get(face.assetId)),key=faceKey(prepared),previous=slots.get(key);
      entry.sha256=prepared.sha256;
      if(previous&&previous.sha256!==prepared.sha256)fail('FONT_STYLE_MISMATCH','동일 family·스타일에 다른 글꼴 파일이 있습니다.');
      if(!previous){slots.set(key,prepared);faces.push(prepared);}
    }catch(error){
      const code=error.code??'FONT_INVALID';issue(code,error.message);
      entry.status=code==='FONT_SOURCE_UNSUPPORTED'?'missing':/RESTRICTED|PREVIEW_ONLY|BITMAP_ONLY/.test(code)?'restricted':'unsupported';entry.fullFont=false;
    }
    report.push(entry);
  }
  return {faces:diagnostics.length?[]:faces,diagnostics,report};
}

// Uncompressed EOT v2.01 wrapper, described at https://www.w3.org/submissions/EOT/.
// The original sfnt, names and permission bits are retained; this is not a font converter.
function eotPayload(face){
  const bytes=Buffer.from(face.bytes.buffer,face.bytes.byteOffset,face.bytes.byteLength),tables=new Map();
  for(let p=12;p<12+bytes.readUInt16BE(4)*16;p+=16){const offset=bytes.readUInt32BE(p+8);tables.set(bytes.toString('ascii',p,p+4),bytes.subarray(offset,offset+bytes.readUInt32BE(p+12)));}
  const os2=tables.get('OS/2'),head=tables.get('head'),name=tables.get('name'),english=new Map(),storage=name.readUInt16BE(4),decoder=new TextDecoder('utf-16be',{fatal:true});
  for(let p=6;p<6+name.readUInt16BE(2)*12;p+=12){
    const id=name.readUInt16BE(p+6);
    if(name.readUInt16BE(p)!==3||![1,10].includes(name.readUInt16BE(p+2))||name.readUInt16BE(p+4)!==0x409||![1,2,4,5].includes(id))continue;
    const offset=storage+name.readUInt16BE(p+10),length=name.readUInt16BE(p+8);
    if(length%2)fail('FONT_INVALID','EOT 이름의 UTF-16 길이가 잘못되었습니다.');
    let value;try{value=decoder.decode(name.subarray(offset,offset+length));}catch{fail('FONT_INVALID','EOT 이름의 UTF-16 데이터가 잘못되었습니다.');}
    if(english.has(id)&&english.get(id)!==value)fail('FONT_INVALID','EOT 영문 이름이 일치하지 않습니다.');english.set(id,value);
  }
  if([1,2,4,5].some(id=>!english.get(id))||english.get(1)!==face.family||english.get(2)!==face.meta.subfamily)fail('FONT_FORMAT_UNSUPPORTED','EOT 시험에는 일치하는 영문 family·style·version·full name이 필요합니다.');
  const header=Buffer.alloc(82);header.writeUInt32LE(bytes.length,4);header.writeUInt32LE(0x20001,8);os2.copy(header,16,32,42);header[26]=1;header[27]=face.meta.italic?1:0;header.writeUInt32LE(face.weight,28);header.writeUInt16LE(face.meta.fsType,32);header.writeUInt16LE(0x504C,34);
  for(let i=0;i<4;i++)header.writeUInt32LE(os2.readUInt32BE(42+i*4),36+i*4);
  if(face.meta.os2Version>0)for(let i=0;i<2;i++)header.writeUInt32LE(os2.readUInt32BE(78+i*4),52+i*4);
  header.writeUInt32LE(head.readUInt32BE(8),60);
  const strings=[1,2,5,4].flatMap((id,i)=>{const value=Buffer.from(english.get(id),'utf16le'),prefix=Buffer.alloc(i?4:2);prefix.writeUInt16LE(value.length,i?2:0);return [prefix,value];});
  const payload=Buffer.concat([header,...strings,Buffer.alloc(4),bytes]);payload.writeUInt32LE(payload.length,0);return payload;
}

export function unwrapPptxFont(input){
  if(!(input instanceof Uint8Array)||input.length<12||input.length>16*1024*1024)fail('FONT_INVALID','내장 글꼴 크기가 잘못되었습니다.');
  const bytes=Buffer.from(input.buffer,input.byteOffset,input.byteLength);
  if(bytes.readUInt32BE(0)===0x10000)return {bytes:new Uint8Array(bytes),meta:inspectFont(bytes),container:'sfnt'};
  if(bytes.length<82||bytes.readUInt32LE(0)!==bytes.length||bytes.readUInt16LE(34)!==0x504C)fail('FONT_INVALID','EOT 글꼴 헤더가 잘못되었습니다.');
  // ponytail: support only the uncompressed v1/v2.01 containers; compressed native EOT remains unsupported.
  const version=bytes.readUInt32LE(8);
  if(![0x10000,0x20001].includes(version)||bytes.readUInt32LE(12)!==0)fail('FONT_FORMAT_UNSUPPORTED','이 EOT 버전·압축 방식은 지원하지 않습니다.');
  if(bytes.subarray(64,82).some(b=>b!==0))fail('FONT_INVALID','EOT 예약 영역이 잘못되었습니다.');
  let offset=82;const names=[],decoder=new TextDecoder('utf-16le',{fatal:true});
  for(let i=0;i<4;i++){
    if(i){if(offset+2>bytes.length||bytes.readUInt16LE(offset)!==0)fail('FONT_INVALID','EOT 이름 구분이 잘못되었습니다.');offset+=2;}
    if(offset+2>bytes.length)fail('FONT_INVALID','EOT 이름 길이가 없습니다.');
    const length=bytes.readUInt16LE(offset);offset+=2;
    if(!length||length%2||offset+length>bytes.length)fail('FONT_INVALID','EOT 이름 범위가 잘못되었습니다.');
    try{names.push(decoder.decode(bytes.subarray(offset,offset+length)));}catch{fail('FONT_INVALID','EOT 이름의 UTF-16 데이터가 잘못되었습니다.');}offset+=length;
  }
  if(version===0x20001){
    if(offset+4>bytes.length||bytes.readUInt16LE(offset)!==0)fail('FONT_INVALID','EOT 이름 뒤 패딩이 잘못되었습니다.');
    if(bytes.readUInt16LE(offset+2)!==0)fail('FONT_FORMAT_UNSUPPORTED','제한된 RootString이 있는 EOT는 지원하지 않습니다.');
    offset+=4;
  }
  const size=bytes.readUInt32LE(4);if(size<12||offset+size!==bytes.length)fail('FONT_INVALID','EOT 글꼴 데이터 크기가 일치하지 않습니다.');
  const sfnt=new Uint8Array(bytes.subarray(offset)),meta=inspectFont(sfnt);
  if(names[0]!==meta.family||names[1]!==meta.subfamily||bytes.readUInt32LE(28)!==meta.weight||bytes[27]!==Number(meta.italic)||bytes.readUInt16LE(32)!==meta.fsType)fail('FONT_INVALID','EOT 헤더와 글꼴 메타데이터가 다릅니다.');
  return {bytes:sfnt,meta,container:'eot'};
}

export function embedPptxFonts(bytes,faces,{container}={}){
  if(!['sfnt','eot'].includes(container))fail('FONT_EMBEDDING_UNSUPPORTED','검증할 글꼴 컨테이너를 명시해야 합니다.');
  if(!(bytes instanceof Uint8Array)||bytes.length<22||bytes.length>PPT_LIMIT)fail('EXPORT_FAILED','PPTX는 64 MiB 이하의 bytes여야 합니다.');
  if(!Array.isArray(faces)||faces.length>32)fail('FONT_ASSET_LIMIT','글꼴 face 제한을 초과했습니다.');
  if(!faces.length)return bytes;
  const slots=new Map();let fontBytes=0;
  for(const face of faces){
    const prepared=checkedFace(face,face),key=faceKey(prepared),previous=slots.get(key);
    if(previous){if(previous.sha256!==prepared.sha256)fail('FONT_STYLE_MISMATCH','같은 글꼴 스타일에 다른 파일이 있습니다.');continue;}
    const payload=container==='sfnt'?prepared.bytes:eotPayload(prepared);
    if(payload.length>16*1024*1024||(fontBytes+=payload.length)>FONT_LIMIT)fail('FONT_ASSET_LIMIT','포함 글꼴의 단일 또는 합계 크기 제한을 초과했습니다.');
    slots.set(key,{...prepared,payload});
  }
  let count=0,total=0;const names=new Set();
  const parts=unzipSync(bytes,{filter:part=>{
    if(part.compression===0&&part.size!==part.originalSize)fail('EXPORT_FAILED','ZIP 저장 크기가 일치하지 않습니다.');
    if(![0,8].includes(part.compression)||++count>10000||Math.max(part.size,part.originalSize)>16*1024*1024||(total+=Math.max(part.size,part.originalSize))>128*1024*1024)fail('EXPORT_FAILED','PPTX 압축 해제 한도를 초과했습니다.');
    if(!part.name||part.name.startsWith('/')||part.name.includes('\\')||part.name.split('/').some(p=>p==='..')||names.has(part.name))fail('EXPORT_FAILED','PPTX 내부 경로가 잘못되거나 중복됩니다.');
    names.add(part.name);return true;
  }});
  if(total+fontBytes>128*1024*1024)fail('EXPORT_FAILED','글꼴을 포함한 PPTX의 압축 해제 한도를 초과했습니다.');
  const xml=(file,namespace,localName)=>{
    if(!parts[file])fail('EXPORT_FAILED','PPTX 필수 XML이 없습니다.');
    const source=strFromU8(parts[file]);if(/<!DOCTYPE|<!ENTITY/i.test(source))fail('EXPORT_FAILED','DTD·엔터티는 허용하지 않습니다.');
    const doc=new DOMParser({onError:()=>fail('EXPORT_FAILED','PPTX XML이 잘못되었습니다.')}).parseFromString(source,'application/xml');
    if(doc.documentElement.namespaceURI!==namespace||doc.documentElement.localName!==localName)fail('EXPORT_FAILED','PPTX XML 루트가 잘못되었습니다.');return doc;
  };
  const presentation=xml('ppt/presentation.xml',P,'presentation'),rels=xml('ppt/_rels/presentation.xml.rels',REL,'Relationships'),types=xml('[Content_Types].xml',CT,'Types'),root=presentation.documentElement;
  const make=(doc,namespace,name,attributes={})=>{const node=doc.createElementNS(namespace,name);for(const [key,value] of Object.entries(attributes))node.setAttribute(key,String(value));return node;};
  const lists=children(root).filter(n=>n.namespaceURI===P&&n.localName==='embeddedFontLst');
  // ponytail: this writer handles newly generated PPTX; an existing font list needs a separate merge contract.
  if(lists.length>1||lists.some(n=>children(n).length))fail('FONT_EMBEDDING_UNSUPPORTED','이미 포함된 글꼴 목록은 자동 병합하지 않습니다.');
  const list=lists[0]??make(presentation,P,'p:embeddedFontLst');
  if(!lists.length){const later=['custShowLst','photoAlbum','custDataLst','kinsoku','defaultTextStyle','modifyVerifier','extLst'];root.insertBefore(list,children(root).find(n=>n.namespaceURI===P&&later.includes(n.localName))??null);}
  root.setAttribute('embedTrueTypeFonts','1');root.setAttribute('saveSubsetFonts','0');
  const defaults=children(types.documentElement).filter(n=>n.localName==='Default'&&n.getAttribute('Extension').toLowerCase()==='fntdata');
  if(defaults.length>1||defaults.some(n=>n.getAttribute('ContentType')!=='application/x-fontdata'))fail('EXPORT_FAILED','fntdata ContentType이 일치하지 않습니다.');
  if(!defaults.length)types.documentElement.insertBefore(make(types,CT,'Default',{Extension:'fntdata',ContentType:'application/x-fontdata'}),types.documentElement.firstChild);
  const ids=new Set();for(const relation of children(rels.documentElement)){const id=relation.getAttribute('Id');if(!id||ids.has(id))fail('EXPORT_FAILED','PPTX 관계 ID가 중복되거나 없습니다.');ids.add(id);}
  const reserved=new Set(children(types.documentElement).filter(n=>n.localName==='Override').map(n=>n.getAttribute('PartName'))),families=new Map();let nextId=1,nextFile=1;
  for(const face of slots.values()){if(!families.has(face.family))families.set(face.family,[]);families.get(face.family).push(face);}
  const order=['regular','bold','italic','boldItalic'];
  for(const [family,variants] of families){
    const embedded=make(presentation,P,'p:embeddedFont');embedded.appendChild(make(presentation,P,'p:font',{typeface:family}));
    for(const face of variants.sort((a,b)=>order.indexOf(slot(a))-order.indexOf(slot(b)))){
      while(ids.has(`rId${nextId}`))nextId++;const rid=`rId${nextId++}`;ids.add(rid);
      while(parts[`ppt/fonts/font${nextFile}.fntdata`]||reserved.has(`/ppt/fonts/font${nextFile}.fntdata`))nextFile++;
      const target=`fonts/font${nextFile++}.fntdata`;parts[`ppt/${target}`]=face.payload;
      rels.documentElement.appendChild(make(rels,REL,'Relationship',{Id:rid,Target:target,Type:`${R}/font`}));
      const variant=make(presentation,P,`p:${slot(face)}`);variant.setAttributeNS(R,'r:id',rid);embedded.appendChild(variant);
    }
    list.appendChild(embedded);
  }
  const serializer=new XMLSerializer();
  for(const [file,doc] of [['ppt/presentation.xml',presentation],['ppt/_rels/presentation.xml.rels',rels],['[Content_Types].xml',types]])parts[file]=strToU8(serializer.serializeToString(doc));
  const output=zipSync(parts,{level:6});if(output.length>PPT_LIMIT)fail('EXPORT_FAILED','생성 PPTX가 64 MiB를 초과했습니다.');return output;
}
