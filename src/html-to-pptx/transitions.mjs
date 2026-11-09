import {unzipSync,zipSync,strFromU8,strToU8} from 'fflate';
import {DOMParser,XMLSerializer} from '@xmldom/xmldom';
const P='http://schemas.openxmlformats.org/presentationml/2006/main',P14='http://schemas.microsoft.com/office/powerpoint/2010/main';
const types=new Set(['none','fade','push','wipe','split','cover','uncover','dissolve','random']);
const dirs={push:['l','r','u','d'],wipe:['l','r','u','d'],cover:['l','r','u','d','lu','ld','ru','rd'],uncover:['l','r','u','d','lu','ld','ru','rd'],split:['in','out']};
export function validTransition(t){
 if(!t||typeof t!=='object'||Array.isArray(t)||!types.has(t.type)||Object.keys(t).some(k=>!['type','duration','speed','advanceClick','advanceTime','direction','orientation','throughBlack'].includes(k)))return false;
 if(t.duration!==undefined&&(!Number.isInteger(t.duration)||t.duration<0||t.duration>600000))return false;
 if(t.speed!==undefined&&!['slow','med','fast'].includes(t.speed))return false;
 if(t.advanceClick!==undefined&&typeof t.advanceClick!=='boolean')return false;
 if(t.advanceTime!==undefined&&(!Number.isInteger(t.advanceTime)||t.advanceTime<0||t.advanceTime>2147483647))return false;
 if(t.direction!==undefined&&!dirs[t.type]?.includes(t.direction))return false;
 if(t.orientation!==undefined&&(t.type!=='split'||!['horz','vert'].includes(t.orientation)))return false;
 if(t.throughBlack!==undefined&&(t.type!=='fade'||typeof t.throughBlack!=='boolean'))return false;
 return true;
}
export function readTransition(node){
 if(!node||node.namespaceURI!==P)return null;
 const t={type:'none'},attr=(name)=>node.hasAttribute(name)?node.getAttribute(name):null;
 for(const a of [...node.attributes])if(!((!a.namespaceURI&&['spd','advClick','advTm'].includes(a.name))||(a.namespaceURI===P14&&a.localName==='dur')||a.name==='xmlns:p14'))return null;
 const duration=node.getAttributeNS(P14,'dur');if(duration!=='')t.duration=Number(duration);
 if(attr('spd')!==null)t.speed=attr('spd');
 if(attr('advClick')!==null){if(!['0','1','true','false'].includes(attr('advClick')))return null;t.advanceClick=['1','true'].includes(attr('advClick'));}
 if(attr('advTm')!==null)t.advanceTime=Number(attr('advTm'));
 const children=[...node.childNodes].filter(n=>n.nodeType===1);if(children.length>1)return null;
 if(children.length){const c=children[0];if(c.namespaceURI!==P||!types.has(c.localName)||c.localName==='none')return null;t.type=c.localName;
  for(const a of [...c.attributes])if(a.namespaceURI||!['dir','orient','thruBlk'].includes(a.name))return null;
  if(c.hasAttribute('dir'))t.direction=c.getAttribute('dir');if(c.hasAttribute('orient'))t.orientation=c.getAttribute('orient');
  if(c.hasAttribute('thruBlk')){const v=c.getAttribute('thruBlk');if(!['0','1','true','false'].includes(v))return null;t.throughBlack=['1','true'].includes(v);}
  if([...c.childNodes].some(n=>n.nodeType===1))return null;
 }
 return validTransition(t)?t:null;
}
export function applyTransitions(bytes,model){
 if(!model.slides.some(s=>s.transition))return bytes;
 const parts=unzipSync(bytes),ser=new XMLSerializer();
 for(const [i,slide] of model.slides.entries()){
  if(!slide.transition)continue;const t=slide.transition;if(!validTransition(t))throw new Error('Invalid slide transition');
  const path=`ppt/slides/slide${i+1}.xml`,doc=new DOMParser().parseFromString(strFromU8(parts[path]),'application/xml'),root=doc.documentElement;
  const e=doc.createElementNS(P,'p:transition');if(t.duration!==undefined)e.setAttributeNS(P14,'p14:dur',String(t.duration));
  if(t.speed!==undefined)e.setAttribute('spd',t.speed);if(t.advanceClick!==undefined)e.setAttribute('advClick',t.advanceClick?'1':'0');if(t.advanceTime!==undefined)e.setAttribute('advTm',String(t.advanceTime));
  if(t.type!=='none'){const child=doc.createElementNS(P,`p:${t.type}`);if(t.direction!==undefined)child.setAttribute('dir',t.direction);if(t.orientation!==undefined)child.setAttribute('orient',t.orientation);if(t.throughBlack!==undefined)child.setAttribute('thruBlk',t.throughBlack?'1':'0');e.appendChild(child);}
  const before=[...root.childNodes].find(n=>n.nodeType===1&&['timing','extLst'].includes(n.localName));root.insertBefore(e,before??null);parts[path]=strToU8(ser.serializeToString(doc));
 }
 return zipSync(parts,{level:6});
}
