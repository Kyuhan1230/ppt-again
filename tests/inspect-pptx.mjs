// Independent OOXML reader. Never imports exporter or its coordinate helpers.
import {unzipSync,strFromU8} from 'fflate';
import {DOMParser} from '@xmldom/xmldom';
const all=(node,name)=>[...node.getElementsByTagNameNS('*',name)];
const child=(node,name)=>[...node.childNodes].find(n=>n.localName===name);
const parse=bytes=>new DOMParser().parseFromString(strFromU8(bytes),'application/xml');
export async function inspectPptxIndependently(bytes,expectedSource=[]){
  const parts=unzipSync(bytes),presentation=parse(parts['ppt/presentation.xml']);
  const size=all(presentation,'sldSz')[0],widthEmu=Number(size.getAttribute('cx')),heightEmu=Number(size.getAttribute('cy'));
  const slides=[];
  for(const file of Object.keys(parts).filter(n=>/^ppt\/slides\/slide\d+\.xml$/.test(n)).sort((a,b)=>Number(a.match(/(\d+)\.xml$/)[1])-Number(b.match(/(\d+)\.xml$/)[1]))){
    const doc=parse(parts[file]),tree=all(doc,'spTree')[0],objects=[];
    const visit=(parent,parentId=null)=>{
      for(const node of [...parent.childNodes].filter(n=>['sp','cxnSp','pic','grpSp','graphicFrame'].includes(n.localName))){
        const nv=all(node,'cNvPr')[0],properties=child(node,node.localName==='grpSp'?'grpSpPr':'spPr'),xfrm=properties&&all(properties,'xfrm')[0],off=xfrm&&child(xfrm,'off'),ext=xfrm&&child(xfrm,'ext');
        const paragraphs=child(node,'txBody')?[...child(node,'txBody').childNodes].filter(n=>n.localName==='p'):[];
        const text=paragraphs.map(p=>[...p.childNodes].map(c=>c.localName==='br'?'\n':c.localName==='r'?all(c,'t').map(t=>t.textContent).join(''):'').join('')).join('\n');
        const runs=all(node,'r').map(r=>{const pr=child(r,'rPr');return {text:all(r,'t').map(t=>t.textContent).join(''),font:pr&&all(pr,'latin')[0]?.getAttribute('typeface'),fontSizePt:Number(pr?.getAttribute('sz'))/100,bold:pr?.getAttribute('b')==='1',italic:pr?.getAttribute('i')==='1',color:pr&&all(pr,'srgbClr')[0]?.getAttribute('val')};});
        const shape=properties&&all(properties,'prstGeom')[0]?.getAttribute('prst'),line=properties&&child(properties,'ln');
        const kind=node.localName==='grpSp'?'group':node.localName==='pic'?'image':node.localName==='cxnSp'||shape==='line'?'line':paragraphs.length&&text.length?'text':'shape';
        const object={id:Number(nv.getAttribute('id')),name:nv.getAttribute('name'),kind,tag:node.localName,parentId,order:objects.length,shape,text,runs,boxEmu:off&&ext?{x:Number(off.getAttribute('x')),y:Number(off.getAttribute('y')),w:Number(ext.getAttribute('cx')),h:Number(ext.getAttribute('cy'))}:null,lineWidthEmu:line?Number(line.getAttribute('w')):0,bodyPr:paragraphs.length?Object.fromEntries([...all(node,'bodyPr')[0].attributes].map(a=>[a.name,a.value])):null};
        object.lineOpacity=line&&!child(line,'noFill')?Number(all(line,'alpha')[0]?.getAttribute('val')??100000):0;
        const fill=properties&&child(properties,'solidFill'),fillColor=fill&&child(fill,'srgbClr');
        object.fillColor=fillColor?.getAttribute('val')??null;
        object.fillOpacity=fillColor?Number(child(fillColor,'alpha')?.getAttribute('val')??100000):0;
        object.pictureOpacity=kind==='image'?Number(all(node,'alphaModFix')[0]?.getAttribute('amt')??100000):null;
        objects.push(object);if(kind==='group')visit(node,object.id);
      }
    };visit(tree);slides.push({file,number:slides.length+1,objects});
  }
  const objects=slides.flatMap(s=>s.objects);
  return {widthEmu,heightEmu,slides,nativeTextCount:objects.filter(n=>n.kind==='text').length,nativeShapeCount:objects.filter(n=>n.kind==='shape').length,nativeLineCount:objects.filter(n=>n.kind==='line').length,pictureCount:objects.filter(n=>n.kind==='image').length,groupCount:objects.filter(n=>n.kind==='group').length,missingSourceIds:expectedSource.filter(e=>!slides[e.slideNumber-1]?.objects.some(o=>o.name===e.id)).map(e=>e.id),fullSlideRasterReplacementCount:slides.filter(s=>s.objects.length===1&&s.objects[0].kind==='image'&&Math.abs(s.objects[0].boxEmu.w-widthEmu)<2&&Math.abs(s.objects[0].boxEmu.h-heightEmu)<2).length};
}
