import {unzipSync,zipSync,strFromU8,strToU8} from 'fflate';
import {DOMParser,XMLSerializer} from '@xmldom/xmldom';

const A='http://schemas.openxmlformats.org/drawingml/2006/main';
const elements=(node,name)=>[...(node?.childNodes??[])].filter(n=>n.nodeType===1&&n.localName===name);
const child=(node,name)=>elements(node,name)[0];
const all=(node,name)=>[...node.getElementsByTagNameNS('*',name)];

export function paragraphTextRuns(paragraph,makeRun,fallbackStyle){
  const items=[];
  for(const run of paragraph.runs.filter(r=>!r.marker)){
    const lines=run.text.split('\n');
    for(const [index,text] of lines.entries()){
      if(index>0&&!items.length)items.push(makeRun({text:'',style:run.style}));
      if(text||index>0||!items.length){const item=makeRun({text,style:run.style});if(index>0)item.options.softBreakBefore=true;items.push(item);}
    }
  }
  return items.length?items:[makeRun({text:'',style:paragraph.runs[0]?.style??fallbackStyle})];
}

// PptxGenJS cannot independently set a paragraph's bullet and hanging indent.
export function applyTextParagraphs(bytes,model,pointScale,paint,resolveFont){
  if(!model.slides.some(slide=>slide.nodes.some(node=>node.content?.paragraphs||node.content?.cells?.some(cell=>cell.paragraphs))))return bytes;
  const parts=unzipSync(bytes),serializer=new XMLSerializer();
  for(const [slideIndex,slide] of model.slides.entries()){
    const targets=slide.nodes.filter(node=>node.content?.paragraphs||node.content?.cells?.some(cell=>cell.paragraphs));
    if(!targets.length)continue;
    const path=`ppt/slides/slide${slideIndex+1}.xml`,doc=new DOMParser().parseFromString(strFromU8(parts[path]),'application/xml');
    const create=(name,attributes={})=>{const element=doc.createElementNS(A,`a:${name}`);for(const [key,value] of Object.entries(attributes))element.setAttribute(key,String(value));return element;};
    const write=(body,paragraphs)=>{
      const output=elements(body,'p');
      if(output.length!==paragraphs.length)throw new Error('문단 수가 변환 과정에서 달라졌습니다.');
      if(paragraphs[0].beforePx>0||paragraphs.at(-1).afterPx>0)child(body,'bodyPr')?.setAttribute('spcFirstLastPara','1');
      const numberGroups=[];
      for(const [index,paragraph] of paragraphs.entries()){
        const p=output[index];let properties=child(p,'pPr');
        if(!properties){properties=create('pPr');p.insertBefore(properties,p.firstChild);}
        let end=child(p,'endParaRPr');
        if(paragraph.sourceEnd?.mode==='implicit'){if(end)p.removeChild(end);}
        else if(paragraph.sourceEnd?.mode==='explicit'){
          if(!end){end=create('endParaRPr');p.appendChild(end);}
          for(const key of ['spc','sz','kern','b','i']){const value=paragraph.sourceEnd.attrs[key];if(value===undefined)end.removeAttribute(key);else end.setAttribute(key,value);}
        }
        for(const [key,value] of [['marL',paragraph.leftPx],['marR',paragraph.rightPx],['indent',paragraph.indentPx]])properties.setAttribute(key,String(Math.round(value*pointScale*12700)));
        if(paragraph.bullet)properties.setAttribute('lvl',String(paragraph.bullet.level??0));
        for(const item of [...properties.childNodes])if(item.nodeType===1&&['lnSpc','spcBef','spcAft','buNone','buChar','buAutoNum','buBlip','buClr','buClrTx','buSzTx','buSzPct','buSzPts','buFont','buFontTx'].includes(item.localName))properties.removeChild(item);
        const anchor=[...properties.childNodes].find(item=>item.nodeType===1&&['tabLst','defRPr','extLst'].includes(item.localName))??null;
        const insert=item=>properties.insertBefore(item,anchor);
        const lineValue=parseFloat(paragraph.lineHeight)*(paragraph.lineHeight.endsWith('px')?1:(paragraph.runs.find(r=>!r.marker)?.style.fontSize??paragraph.runs[0]?.style.fontSize));
        for(const [name,value] of [['lnSpc',lineValue],['spcBef',paragraph.beforePx],['spcAft',paragraph.afterPx]]){
          const spacing=create(name),original=paragraph.sourceSpacing?.[name];
          if(original?.kind==='default')continue;
          spacing.appendChild(original?create(original.kind,{val:original.value}):create('spcPts',{val:Math.round(value*pointScale*100)}));insert(spacing);
        }
        if(paragraph.bullet){
          const marker=paragraph.runs.find(run=>run.marker),style=marker?.style??paragraph.runs[0].style;
          if(paragraph.bullet.followTextColor)insert(create('buClrTx'));
          else{const color=paint(style.color),bulletColor=create('buClr'),rgb=create('srgbClr',{val:color.color});if(color.transparency)rgb.appendChild(create('alpha',{val:Math.round((100-color.transparency)*1000)}));bulletColor.appendChild(rgb);insert(bulletColor);}
          const sourceSize=paragraph.bullet.sourceSize;
          insert(sourceSize?create(sourceSize.kind,sourceSize.kind==='buSzTx'?{}:{val:sourceSize.value}):create('buSzPts',{val:Math.round(style.fontSize*pointScale*100)}));
          insert(paragraph.bullet.followTextFont?create('buFontTx'):create('buFont',{typeface:resolveFont(style,slide).fontFace}));
          if(paragraph.bullet.type==='char'){
            insert(create('buChar',{char:paragraph.bullet.char}));numberGroups.length=Math.min(numberGroups.length,paragraph.bullet.level??0);
          }else{
            const level=paragraph.bullet.level??0,previous=numberGroups[level],restart=paragraph.bullet.restart!==false||!previous;
            const groupStart=restart?paragraph.bullet.start:previous.groupStart,count=restart?paragraph.bullet.start:previous.count+1;
            if(restart&&previous&&previous.groupStart===groupStart)throw new Error('같은 시작 번호로 즉시 다시 시작하는 목록은 아직 지원하지 않습니다.');
            if(count>32767)throw new Error('목록 번호가 지원 범위를 벗어났습니다.');
            numberGroups[level]={groupStart,count};numberGroups.length=level+1;
            insert(create('buAutoNum',{type:paragraph.bullet.format,startAt:groupStart}));
          }
        }else{insert(create('buNone'));numberGroups.length=0;}
      }
    };
    for(const node of targets){
      const nv=all(doc,'cNvPr').find(item=>item.getAttribute('name')===node.id);
      if(!nv)throw new Error('문단을 적용할 PPT 객체를 찾지 못했습니다.');
      const object=nv.parentNode.parentNode;
      if(node.kind==='text')write(child(object,'txBody'),node.content.paragraphs);
      else if(node.kind==='table'){
        const rows=elements(all(object,'tbl')[0],'tr');
        for(const cell of node.content.cells.filter(item=>item.paragraphs)){
          const nativeCell=elements(rows[cell.row],'tc')[cell.col];
          if(!nativeCell)throw new Error('문단을 적용할 PPT 표 셀을 찾지 못했습니다.');
          write(child(nativeCell,'txBody'),cell.paragraphs);
        }
      }
    }
    parts[path]=strToU8(serializer.serializeToString(doc));
  }
  return zipSync(parts,{level:6});
}
