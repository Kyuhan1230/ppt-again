import {unzipSync,zipSync,strFromU8,strToU8} from 'fflate';
import {DOMParser,XMLSerializer} from '@xmldom/xmldom';
const P='http://schemas.openxmlformats.org/presentationml/2006/main',A='http://schemas.openxmlformats.org/drawingml/2006/main';
const all=(n,name)=>[...n.getElementsByTagNameNS('*',name)];
export function groupDiagnostics(slide){
  const errors=[],leaves=[...slide.nodes].filter(n=>n.kind!=='group').sort((a,b)=>a.zOrder-b.zOrder);
  for(const group of slide.nodes.filter(n=>n.kind==='group')){
    const children=slide.nodes.filter(n=>n.parentId===group.id),indices=children.map(n=>leaves.indexOf(n)).sort((a,b)=>a-b);
    if(group.parentId||group.style.rotation||children.some(n=>n.kind==='group'||n.style.rotation))errors.push({group,message:'중첩·회전 그룹은 지원하지 않습니다.'});
    else if(!children.length||indices.some((n,i)=>n!==indices[0]+i))errors.push({group,message:'그룹 구성원이 겹침 순서에서 연속되지 않습니다.'});
  }
  return errors;
}
export function applyFlatGroups(bytes,model,objectMap,pointScale){
  const parts=unzipSync(bytes),serializer=new XMLSerializer();
  for(const [i,slide] of model.slides.entries()){
    const groups=slide.nodes.filter(n=>n.kind==='group'),bars=slide.nodes.filter(n=>n.kind==='line'&&n.style.linePaint?.kind==='bar'),tables=slide.nodes.filter(n=>n.kind==='table');if(!groups.length&&!bars.length&&!tables.length)continue;
    const file=`ppt/slides/slide${i+1}.xml`,doc=new DOMParser().parseFromString(strFromU8(parts[file]),'application/xml'),tree=all(doc,'spTree')[0];
    // PptxGenJS numbers tables separately, which can collide with ordinary shapes.
    for(const table of tables){
      const entry=objectMap.find(e=>e.pptSlideNumber===i+1&&e.sourceNodeId===table.id),nv=all(doc,'cNvPr').find(n=>n.getAttribute('name')===table.id&&n.parentNode.localName==='nvGraphicFramePr');
      if(!entry||!nv)throw new Error('Table output correspondence missing');nv.setAttribute('id',String(entry.pptObjectId));
    }
    let nextId=Math.max(...all(doc,'cNvPr').map(n=>Number(n.getAttribute('id'))))+1;
    const create=(ns,name,attrs={})=>{const e=doc.createElementNS(ns,name);for(const [k,v] of Object.entries(attrs))e.setAttribute(k,String(v));return e;};
    // Office paints thin filled bars differently when a transparent stroke is
    // present. Declare no outline, matching the HTML background-only paint.
    for(const bar of bars){
      const entry=objectMap.find(e=>e.pptSlideNumber===i+1&&e.sourceNodeId===bar.id),nv=all(doc,'cNvPr').find(n=>Number(n.getAttribute('id'))===entry?.pptObjectId);
      if(!nv)throw new Error('Bar output correspondence missing');
      const properties=[...nv.parentNode.parentNode.childNodes].find(n=>n.localName==='spPr'),line=[...properties.childNodes].find(n=>n.localName==='ln'),empty=create(A,'a:ln');
      empty.appendChild(create(A,'a:noFill'));if(line)properties.replaceChild(empty,line);else properties.appendChild(empty);
    }
    for(const group of groups){
      const childIds=new Set(slide.nodes.filter(n=>n.parentId===group.id).map(n=>n.id)),entries=objectMap.filter(e=>e.pptSlideNumber===i+1&&(childIds.has(e.sourceNodeId)||childIds.has(e.derivedFrom))),pptIds=new Set(entries.map(e=>e.pptObjectId));
      const members=[...tree.childNodes].filter(n=>n.nodeType===1&&pptIds.has(Number(all(n,'cNvPr')[0]?.getAttribute('id'))));
      if(members.length!==entries.length)throw new Error('Group output correspondence missing');
      const g=create(P,'p:grpSp'),nv=create(P,'p:nvGrpSpPr'),id=nextId++;
      nv.appendChild(create(P,'p:cNvPr',{id,name:group.id}));nv.appendChild(create(P,'p:cNvGrpSpPr'));nv.appendChild(create(P,'p:nvPr'));g.appendChild(nv);
      const pr=create(P,'p:grpSpPr'),xfrm=create(A,'a:xfrm'),b=group.boxCssPx,emu=v=>Math.round(v*pointScale*12700);
      for(const [name,attrs] of [['off',{x:emu(b.x),y:emu(b.y)}],['ext',{cx:emu(b.w),cy:emu(b.h)}],['chOff',{x:emu(b.x),y:emu(b.y)}],['chExt',{cx:emu(b.w),cy:emu(b.h)}]])xfrm.appendChild(create(A,`a:${name}`,attrs));
      pr.appendChild(xfrm);g.appendChild(pr);tree.insertBefore(g,members[0]);for(const member of members)g.appendChild(member);
      for(const entry of entries)entry.pptParentId=id;
      objectMap.push({sourceSlideId:slide.id,sourceNodeId:group.id,sourceRef:group.sourceRef,pptSlideNumber:i+1,pptObjectId:id,pptObjectName:group.id,support:'native-group',parentId:null});
    }
    parts[file]=strToU8(serializer.serializeToString(doc));
  }
  return zipSync(parts,{level:6});
}
