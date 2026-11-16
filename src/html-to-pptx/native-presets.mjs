import {unzipSync,zipSync,strFromU8,strToU8} from 'fflate';
import {DOMParser,XMLSerializer} from '@xmldom/xmldom';
const A='http://schemas.openxmlformats.org/drawingml/2006/main';
const all=(n,name)=>[...n.getElementsByTagNameNS('*',name)];
export function applyNativePresetShapes(bytes,model,objectMap){
 if(!model.slides.some(s=>s.nodes.some(n=>n.content?.editableVector)))return bytes;
 const parts=unzipSync(bytes),ser=new XMLSerializer();
 for(const [i,slide] of model.slides.entries()){
  const nodes=slide.nodes.filter(n=>n.content?.editableVector);if(!nodes.length)continue;
  const path=`ppt/slides/slide${i+1}.xml`,doc=new DOMParser().parseFromString(strFromU8(parts[path]),'application/xml');
  const create=(name,attrs={})=>{const e=doc.createElementNS(A,`a:${name}`);for(const [k,v] of Object.entries(attrs))e.setAttribute(k,String(v));return e;};
  for(const n of nodes){
   const v=n.content.editableVector,entry=objectMap.find(e=>e.pptSlideNumber===i+1&&e.sourceNodeId===n.id),nv=all(doc,'cNvPr').find(e=>Number(e.getAttribute('id'))===entry?.pptObjectId);
   if(!nv||nv.parentNode?.localName!=='nvSpPr')throw new Error('Native preset correspondence missing');
   const shape=nv.parentNode.parentNode,pr=[...shape.childNodes].find(x=>x.localName==='spPr'),geom=[...pr.childNodes].find(x=>x.localName==='prstGeom');
   if(!geom)throw new Error('Native preset geometry missing');
   if(['downArrow','rightArrow'].includes(v.shape)){
    const av=all(geom,'avLst')[0];while(av.firstChild)av.removeChild(av.firstChild);
    av.appendChild(create('gd',{name:'adj1',fmla:`val ${v.adj1}`}));av.appendChild(create('gd',{name:'adj2',fmla:`val ${v.adj2}`}));
   }
   if(v.gradient){
    const gradient=create('gradFill',{flip:'none',rotWithShape:'1'}),list=create('gsLst');
    for(const [pos,color] of [[0,v.gradient.from],[100000,v.gradient.to]]){const stop=create('gs',{pos});stop.appendChild(create('srgbClr',{val:color.slice(1).toUpperCase()}));list.appendChild(stop);}
    gradient.appendChild(list);gradient.appendChild(create('lin',{ang:'0',scaled:'1'}));gradient.appendChild(create('tileRect'));
    const old=[...pr.childNodes].find(x=>['solidFill','noFill','gradFill'].includes(x.localName));if(old)pr.replaceChild(gradient,old);else pr.appendChild(gradient);
   }
  }
  parts[path]=strToU8(ser.serializeToString(doc));
 }
 return zipSync(parts,{level:6});
}
