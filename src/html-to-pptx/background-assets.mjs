import {assetDataUrl,resolveAsset} from './assets.mjs';
import {validatePngStructure} from './png-duotone.mjs';
import {DOMParser} from '@xmldom/xmldom';
import * as cssTree from 'css-tree';

const imageMimes=new Set(['image/png','image/jpeg','image/svg+xml']);
// Check marker and segment boundaries without claiming to decode JPEG pixels.
function validateJpegStructure(bytes){
  const data=Buffer.from(bytes),invalid=()=>{throw new Error('배경 JPEG 구조가 잘못되었습니다.');};
  if(data.length<4||data[0]!==0xff||data[1]!==0xd8)invalid();
  let offset=2,frame=false,scan=false,scanBytes=0,scans=0,markers=0;
  const frames=new Set([0xc0,0xc1,0xc2,0xc3,0xc5,0xc6,0xc7,0xc9,0xca,0xcb,0xcd,0xce,0xcf]);
  while(offset<data.length){
    if(scan){
      const next=data.indexOf(0xff,offset);
      if(next<0)invalid();scanBytes+=next-offset;offset=next;
      while(offset<data.length&&data[offset]===0xff)offset++;
      if(offset>=data.length)invalid();
      const marker=data[offset];
      if(marker===0x00||marker>=0xd0&&marker<=0xd7){scanBytes++;offset++;continue;}
      if(!scanBytes)invalid();scan=false;offset=next;
    }
    if(data[offset++]!==0xff)invalid();
    while(offset<data.length&&data[offset]===0xff)offset++;
    if(offset>=data.length||++markers>10000)invalid();
    const marker=data[offset++];
    if(marker===0xd9){if(offset!==data.length||!frame||!scans)invalid();return;}
    if(marker===0x00||marker===0xd8||marker>=0xd0&&marker<=0xd7)invalid();
    if(marker===0x01)continue;
    if(offset+2>data.length)invalid();
    const length=data.readUInt16BE(offset);
    if(length<2||offset+length>data.length)invalid();
    if(frames.has(marker)){
      if(length<11)invalid();
      const height=data.readUInt16BE(offset+3),width=data.readUInt16BE(offset+5),components=data[offset+7];
      if(!width||!height||width*height>32000000||components<1||components>4||length!==8+3*components)invalid();
      frame=true;
    }
    if(marker===0xda){
      const components=data[offset+2];
      if(!frame||components<1||components>4||length!==6+2*components)invalid();
      scan=true;scanBytes=0;scans++;
    }
    offset+=length;
  }
  invalid();
}
function validateSvgStructure(bytes){
  const invalid=()=>{throw new Error('배경 SVG 구조가 잘못되었습니다.');};
  let source;try{source=new TextDecoder('utf-8',{fatal:true}).decode(bytes);}catch{invalid();}
  if(/<!DOCTYPE|<!ENTITY/i.test(source))invalid();
  let doc;try{doc=new DOMParser({onError:invalid}).parseFromString(source,'image/svg+xml');}catch{invalid();}
  if(doc.documentElement?.localName!=='svg'||doc.documentElement.namespaceURI!=='http://www.w3.org/2000/svg')invalid();
  const inspectCss=(value,context)=>{
    let tree;try{tree=cssTree.parse(value,{context});}catch{invalid();}
    try{cssTree.walk(tree,node=>{
      if(node.type==='Raw')invalid();
      if(node.type==='Atrule'&&cssTree.ident.decode(node.name).toLowerCase()==='import')invalid();
      if(node.type==='Url'&&!/^#[^\s]*$/.test(node.value))invalid();
      if(node.type==='Function'&&['image-set','-webkit-image-set'].includes(cssTree.ident.decode(node.name).toLowerCase()))invalid();
    });}catch{invalid();}
  };
  const pending=[doc.documentElement];let count=0;
  while(pending.length){
    const element=pending.pop();if(++count>100000)invalid();
    const name=element.localName?.toLowerCase();
    if(['script','foreignobject','iframe','animate','animatetransform','animatemotion','set'].includes(name))invalid();
    if(name==='style')inspectCss(element.textContent??'','stylesheet');
    for(let i=0;i<element.attributes.length;i++){
      const attr=element.attributes.item(i),local=attr.localName?.toLowerCase()??'',value=attr.value??'';
      if(local.startsWith('on')||local==='base'&&attr.namespaceURI==='http://www.w3.org/XML/1998/namespace')invalid();
      if((local==='href'||local==='src')&&!/^#[^\s]*$/.test(value.trim()))invalid();
      if(local==='style')inspectCss(value,'declarationList');
      else if(local!=='href'&&local!=='src'&&/url|image-set|\\/i.test(value))inspectCss(value,'value');
    }
    for(let child=element.firstChild;child;child=child.nextSibling)if(child.nodeType===1)pending.push(child);
  }
}
// Decorative assets come from the input model; validate MIME, bytes and hash
// before embedding the same data in either output format.
export function verifiedBackgroundDataUrl(asset){
  if(!asset||!imageMimes.has(asset.mime)||!(asset.bytes instanceof Uint8Array))throw new Error('배경 그림 자산이 없거나 지원하지 않는 형식입니다.');
  const verified=resolveAsset(`data:${asset.mime};base64,${Buffer.from(asset.bytes).toString('base64')}`);
  if(verified.sha256!==asset.sha256)throw new Error('배경 그림 자산의 무결성을 확인할 수 없습니다.');
  if(verified.mime==='image/png')validatePngStructure(verified.bytes);
  if(verified.mime==='image/jpeg')validateJpegStructure(verified.bytes);
  if(verified.mime==='image/svg+xml')validateSvgStructure(verified.bytes);
  return assetDataUrl(verified);
}

export function validateBackgroundCapture(capture,asset,bounds){
  if(capture===undefined)return;
  const bytes=asset?.bytes;
  if(capture?.content!=='computed-background-only'||capture.sourceChildrenCopied!==0||
    !Number.isSafeInteger(capture.width)||!Number.isSafeInteger(capture.height)||capture.width<=0||capture.height<=0||
    !Number.isFinite(capture.renderScale)||capture.renderScale<=0||asset?.mime!=='image/png'||
    !(bytes instanceof Uint8Array)||bytes.length<24||Buffer.from(bytes).readUInt32BE(16)!==capture.width||Buffer.from(bytes).readUInt32BE(20)!==capture.height||
    !Number.isFinite(bounds?.w)||!Number.isFinite(bounds?.h)||bounds.w<=0||bounds.h<=0||
    Math.abs(capture.width-bounds.w*capture.renderScale)>1||Math.abs(capture.height-bounds.h*capture.renderScale)>1)
    throw new Error('배경 합성 정보가 잘못되었습니다.');
}
