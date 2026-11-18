import {sha256} from './assets.mjs';
import {measurePage} from './html-object-readers.mjs';
import {validateDisplayMapping} from './html-display-mapping.mjs';
import {captureDocumentTiles} from './html-pagination-tiles.mjs';
import {readPaginationFontEvidence} from './html-pagination-fonts.mjs';

const unsupported=message=>({status:'unsupported',diagnostics:[{severity:'error',code:'PAGINATION_UNSUPPORTED',message}]});
const timeoutFailure=()=>({status:'unsupported',diagnostics:[{severity:'error',code:'PAGINATION_UNSUPPORTED',reasonKey:'PAGINATION_TIMEOUT',message:'문서 관측 시간이 초과되었습니다.'}],cleanupRequired:true});
const validViewport=v=>v&&Number.isSafeInteger(v.width)&&Number.isSafeInteger(v.height)&&v.width>0&&v.height>0&&v.width<=8192&&v.height<=8192;

// Read-only, deliberately narrow first observer. It does not validate a user's
// selection token or create a PPT-ready model. Unsupported content is rejected
// rather than converted from incomplete DOM observations.
async function readSnapshot({selector}){
  await document.fonts.ready;
  let roots;try{roots=[...document.querySelectorAll(selector)];}catch{return {error:'선택 영역을 확인할 수 없습니다.'};}
  if(roots.length!==1)return {error:'선택한 문서 영역은 하나여야 합니다.'};
  const root=roots[0],rootRect=root.getBoundingClientRect(),viewport={width:innerWidth,height:innerHeight,dpr:devicePixelRatio,scrollX,scrollY};
  if(!rootRect.width||!rootRect.height||rootRect.width>8192||rootRect.height>32768||rootRect.width*rootRect.height>32000000)return {error:'선택 영역 크기를 안전하게 관측할 수 없습니다.'};
  const transparent=color=>color==='transparent'||/^rgba\([^)]*[, /]0(?:\.0+)?\)$/.test(color);
  const visiblePseudo=element=>['::before','::after','::marker'].some(pseudo=>!['none','normal'].includes(getComputedStyle(element,pseudo).content));
  const pseudoProperties=['font-family','font-size','font-weight','font-style','font-kerning','font-feature-settings','font-variant-caps','font-variant-numeric','font-variant-east-asian','color','-webkit-text-fill-color','-webkit-text-stroke-width','-webkit-text-stroke-color','-webkit-text-security','text-shadow','text-emphasis-style','text-emphasis-color','text-transform','text-decoration-line','letter-spacing','word-spacing','background-color','background-image'];
  const textPseudoEffect=element=>{const style=getComputedStyle(element);return ['::first-letter','::first-line'].some(pseudo=>pseudoProperties.some(name=>getComputedStyle(element,pseudo).getPropertyValue(name)!==style.getPropertyValue(name)));};
  const inspectStyle=(element,{replacedImage=false}={})=>{
    const s=getComputedStyle(element),duration=value=>value.split(',').some(part=>parseFloat(part)!==0);
    const unsupportedOverflow=value=>['hidden','scroll','auto'].includes(value)||value==='clip'&&!replacedImage;
    return s.display==='none'||s.display.split(/\s+/).includes('list-item')||s.visibility!=='visible'||s.contentVisibility!=='visible'||s.contain!=='none'||s.willChange!=='auto'||s.isolation!=='auto'||Number(s.opacity)!==1||s.transform!=='none'||s.scale!=='none'||s.rotate!=='none'||s.translate!=='none'||Number(s.zoom)!==1||s.filter!=='none'||s.clipPath!=='none'||s.clip!=='auto'||s.maskImage!=='none'||s.mixBlendMode!=='normal'||s.boxShadow!=='none'||s.textShadow!=='none'||s.animationName!=='none'||duration(s.transitionDuration)||['fixed','sticky'].includes(s.position)||unsupportedOverflow(s.overflowX)||unsupportedOverflow(s.overflowY)||visiblePseudo(element);
  };
  const supportedFlow=style=>style.columnCount==='auto'&&style.columnWidth==='auto'&&!/\b(?:flex|grid)\b/.test(style.display);
  for(let element=root;element;element=element.parentElement){
    if(inspectStyle(element)||textPseudoEffect(element))return {error:'숨김·스크롤·변환 또는 합성 효과가 있는 문서는 관측할 수 없습니다.'};
    const ancestorStyle=getComputedStyle(element),border=['Top','Right','Bottom','Left'].some(side=>parseFloat(ancestorStyle[`border${side}Width`])>0&&ancestorStyle[`border${side}Style`]!=='none'&&!transparent(ancestorStyle[`border${side}Color`]));
    if(!supportedFlow(ancestorStyle))return {error:'다단 또는 별도 읽기 순서가 있는 문서는 분할할 수 없습니다.'};
    if(!transparent(ancestorStyle.backgroundColor)||ancestorStyle.backgroundImage!=='none'||border||ancestorStyle.textDecorationLine!=='none'||ancestorStyle.textEmphasisStyle!=='none'||ancestorStyle.outlineStyle!=='none'&&parseFloat(ancestorStyle.outlineWidth)>0&&!transparent(ancestorStyle.outlineColor))return {error:'문서 또는 부모 배경·테두리·글자 장식을 확인할 수 없습니다.'};
    // A screenshot may contain an unrelated overlay without revealing its
    // ownership. For this profile accept no rendered sibling subtree at any
    // ancestor level. Hidden-by-display metadata such as HEAD cannot paint.
    for(const sibling of element.parentNode?.childNodes??[]){
      if(sibling===element)continue;
      if(sibling.nodeType===Node.ELEMENT_NODE&&getComputedStyle(sibling).display!=='none'||
          sibling.nodeType===Node.TEXT_NODE&&sibling.data.trim()){
        return {error:'선택 문서의 바깥 콘텐츠가 있거나 겹칠 수 있습니다.'};
      }
    }
  }
  if([...root.querySelectorAll('*')].length>10000)return {error:'문서의 객체 수가 관측 한도를 초과합니다.'};
  const children=[...root.children].filter(element=>!['LINK','STYLE'].includes(element.tagName)||getComputedStyle(element).display!=='none');
  if(!children.length||children.some(element=>!['P','H1','H2','H3','H4','H5','H6','TABLE','IMG'].includes(element.tagName))||[...root.childNodes].some(node=>node.nodeType===Node.TEXT_NODE&&node.data.trim()))return {error:'현재 관측기는 직접 배치된 일반 문단·제목·표·그림을 지원합니다.'};
  const rootStyle=getComputedStyle(root);
  if(!transparent(rootStyle.backgroundColor)||rootStyle.backgroundImage!=='none')return {error:'문서 배경의 분할 정책을 확인할 수 없습니다.'};
  const fonts=[...document.fonts].map(face=>({family:face.family,status:face.status,weight:face.weight,style:face.style}));
  if(fonts.some(face=>face.status!=='loaded'||['arial','malgun gothic'].includes(face.family.replace(/^['"]|['"]$/g,'').toLowerCase())))return {error:'문서 글꼴의 실제 원본을 확인할 수 없습니다.'};
  const paragraphs=[],tables=[],images=[];let characters=0,imageBytes=0;
  const opaqueRgb=color=>{
    const rgb=/^rgb\(\s*(\d{1,3})\s*,\s*(\d{1,3})\s*,\s*(\d{1,3})\s*\)$/.exec(color);
    return Boolean(rgb&&rgb.slice(1).every(value=>Number(value)<=255));
  };
  const sides=['Top','Right','Bottom','Left'];
  const boxOf=element=>{
    const box=element.getBoundingClientRect();
    if(box.width<=0||box.height<=0||box.x<rootRect.x-.000001||box.y<rootRect.y-.000001||box.right>rootRect.right+.000001||box.bottom>rootRect.bottom+.000001)return null;
    return {x:box.x-rootRect.x,y:box.y-rootRect.y,w:box.width,h:box.height};
  };
  const hasOutline=style=>style.outlineStyle!=='none'&&parseFloat(style.outlineWidth)>0&&!transparent(style.outlineColor);
  const textFamily=style=>style.fontFamily.split(',')[0].trim().replace(/^['"]|['"]$/g,'');
  for(const [index,p] of children.entries()){
    const zOrder=index;
    const objectStyle=getComputedStyle(p);
    if(!supportedFlow(objectStyle)||!['static','relative'].includes(objectStyle.position)||objectStyle.zIndex!=='auto'||objectStyle.float!=='none'||objectStyle.perspective!=='none'||objectStyle.backdropFilter!=='none')return {error:'문서 객체의 겹치는 순서와 표시 영역을 확인할 수 없습니다.'};
    if(p.tagName==='IMG'){
      const box=boxOf(p),src=p.currentSrc,assetId=p.dataset.assetId,sourcePath=p.dataset.sourcePath;
      if(inspectStyle(p,{replacedImage:true})||!box||!sourcePath||!/^asset-[a-f0-9]{64}$/.test(assetId??'')||objectStyle.objectFit!=='fill'||!transparent(objectStyle.backgroundColor)||objectStyle.backgroundImage!=='none'||hasOutline(objectStyle)||sides.some(side=>parseFloat(objectStyle[`border${side}Width`])!==0||parseFloat(objectStyle[`padding${side}`])!==0)||[objectStyle.borderTopLeftRadius,objectStyle.borderTopRightRadius,objectStyle.borderBottomLeftRadius,objectStyle.borderBottomRightRadius].some(value=>parseFloat(value)!==0))return {error:'그림의 표시 범위나 자르기·장식을 확인할 수 없습니다.'};
      const match=/^data:(image\/(?:png|jpeg));base64,([a-z0-9+/]+={0,2})$/i.exec(src);
      if(!match||(imageBytes+=match[2].length)>23000000)return {error:'그림은 검증된 PNG 또는 JPEG 자산이어야 합니다.'};
      try{
        await p.decode();
        if(!p.complete||!p.naturalWidth||!p.naturalHeight||p.currentSrc!==src)return {error:'그림의 실제 자산을 확인할 수 없습니다.'};
        const bytes=Uint8Array.from(atob(match[2]),value=>value.charCodeAt(0));
        const digest=Array.from(new Uint8Array(await crypto.subtle.digest('SHA-256',bytes)),value=>value.toString(16).padStart(2,'0')).join('');
        if(assetId!==`asset-${digest}`)return {error:'그림의 실제 자산이 입력과 다릅니다.'};
        images.push({id:`image-${index+1}`,zOrder,position:objectStyle.position,sourcePath,box,assetId,mime:match[1]});
      }catch{return {error:'그림을 해독하고 실제 자산을 확인할 수 없습니다.'};}
      continue;
    }
    if(p.tagName==='TABLE'){
      const box=boxOf(p),sourcePath=p.dataset.sourcePath,rows=[...p.rows],cells=[],rowPaintBoxes=[];
      if(!box||!sourcePath||rows.length<1||rows.length>100||p.querySelectorAll('td,th').length>2000||!['table','inline-table'].includes(objectStyle.display)||objectStyle.emptyCells!=='show'||objectStyle.borderCollapse!=='collapse'&&objectStyle.borderSpacing!=='0px'||[p,...p.querySelectorAll('*')].some(element=>{
        const s=getComputedStyle(element);
        return !['TABLE','THEAD','TFOOT','TBODY','TR','TD','TH'].includes(element.tagName)||inspectStyle(element)||textPseudoEffect(element)||s.perspective!=='none'||s.backdropFilter!=='none'||s.position!=='static'&&element!==p||s.zIndex!=='auto'||hasOutline(s)||s.backgroundImage!=='none'||s.backgroundClip!=='border-box'||s.emptyCells!=='show'||!transparent(s.backgroundColor)&&!opaqueRgb(s.backgroundColor)||sides.some(side=>parseFloat(s[`border${side}Width`])!==0)||[s.borderTopLeftRadius,s.borderTopRightRadius,s.borderBottomLeftRadius,s.borderBottomRightRadius].some(value=>parseFloat(value)!==0);
      }))return {error:'표의 행 경계·셀 장식 또는 테두리를 확인할 수 없습니다.'};
      const grid=Array.from({length:rows.length},()=>Array(50).fill(false));
      for(const [row,tr] of rows.entries()){
        const rowBox=boxOf(tr);
        if(!rowBox||Math.abs(rowBox.x-box.x)>.05||Math.abs(rowBox.w-box.w)>.05||Math.abs(rowBox.y-(rowPaintBoxes.length?rowPaintBoxes.at(-1).y+rowPaintBoxes.at(-1).h:box.y))>.05)return {error:'표 행이 하나의 연속된 격자를 이루지 않습니다.'};
        rowPaintBoxes.push(rowBox);
        let col=0;
        for(const cell of tr.cells){
          while(grid[row][col])col++;
          const rowSpan=cell.rowSpan,colSpan=cell.colSpan;
          if(rowSpan<1||row+rowSpan>rows.length||colSpan<1||col+colSpan>50)return {error:'병합 셀의 범위가 표 격자를 벗어났습니다.'};
          for(let y=row;y<row+rowSpan;y++)for(let x=col;x<col+colSpan;x++){
            if(grid[y][x])return {error:'두 병합 셀의 영역이 겹칩니다.'};
            grid[y][x]=true;
          }
          const text=cell.textContent,style=getComputedStyle(cell),family=textFamily(style),cellBox=cell.getBoundingClientRect();
          if(cell.children.length||!/^([\x20-\x7e\uac00-\ud7a3]*)$/.test(text)||/^ | $| {2}/.test(text)||(characters+=text.length)>30000||!['arial','malgun gothic'].includes(family.toLowerCase())||!['400','700'].includes(style.fontWeight)||style.fontStyle!=='normal'||style.fontFeatureSettings!=='normal'||style.fontVariationSettings!=='normal'||style.fontStretch!=='100%'||style.letterSpacing!=='normal'||style.wordSpacing!=='0px'||!['normal','nowrap'].includes(style.whiteSpace)||style.textTransform!=='none'||style.textDecorationLine!=='none'||style.textEmphasisStyle!=='none'||style.writingMode!=='horizontal-tb'||style.direction!=='ltr'||!['normal','isolate'].includes(style.unicodeBidi)||!opaqueRgb(style.color)||style.webkitTextFillColor!==style.color||parseFloat(style.webkitTextStrokeWidth)!==0||style.webkitTextSecurity!=='none')return {error:'표 셀은 표시를 확인할 수 있는 일반 한 줄 문구여야 합니다.'};
          if(text){
            const range=document.createRange();range.selectNodeContents(cell);
            const rects=[...range.getClientRects()].filter(rect=>rect.width>0&&rect.height>0);
            const line=rects[0],canvas=document.createElement('canvas').getContext('2d');
            canvas.font=`${style.fontWeight} ${style.fontSize} "${family}"`;
            const metrics=canvas.measureText(text);
            // Small typographic side bearings can cross an internal column
            // edge without crossing a page cut. The actual line advance
            // must still fit its cell, and ink must stay inside the table.
            const tableLeft=rootRect.x+box.x,tableRight=tableLeft+box.w;
            if(rects.length!==1||Math.abs(line.width-metrics.width)>.5||line.left<cellBox.left-.05||line.right>cellBox.right+.05||line.left-metrics.actualBoundingBoxLeft<tableLeft-.05||line.left+metrics.actualBoundingBoxRight>tableRight+.05||line.top-1<cellBox.top||line.bottom+1>cellBox.bottom)return {error:'표 셀의 글자가 셀 또는 병합 행 범위를 벗어날 수 있습니다.'};
          }
          const cellPath=cell.dataset.sourcePath;
          if(!cellPath)return {error:'표 셀의 독립 원본 위치가 없습니다.'};
          cells.push({row,col,rowSpan,colSpan,text,family,cellPath,sourceId:`table-${index+1}-cell-${row+1}-${col+1}`});
          col+=colSpan;
        }
      }
      if(Math.abs(rowPaintBoxes.at(-1).y+rowPaintBoxes.at(-1).h-box.y-box.h)>.05)return {error:'표의 첫 행과 마지막 행이 실제 표 표시 범위와 다릅니다.'};
      tables.push({id:`table-${index+1}`,zOrder,position:objectStyle.position,sourcePath,box,rowPaintBoxes,cells});
      continue;
    }
    const style=getComputedStyle(p),walker=document.createTreeWalker(p,NodeFilter.SHOW_TEXT|NodeFilter.SHOW_ELEMENT),sources=[];
    while(walker.nextNode()){const node=walker.currentNode;if(node.nodeType===Node.TEXT_NODE)sources.push({node,kind:'text',text:node.data});else if(node.tagName==='BR')sources.push({node,kind:'br',text:'\n'});}
    const inlineProperties=['fontFamily','fontSize','fontWeight','fontStyle','fontStretch','fontSizeAdjust','fontKerning','fontVariantLigatures','fontVariantCaps','fontVariantNumeric','fontVariantEastAsian','fontVariantPosition','fontVariantAlternates','fontFeatureSettings','fontVariationSettings','webkitTextSecurity','letterSpacing','wordSpacing','textTransform','textDecorationLine','textEmphasisStyle','writingMode','direction','webkitTextStrokeWidth','whiteSpace','lineHeight','verticalAlign'];
    for(const element of p.querySelectorAll('*')){
      const inline=getComputedStyle(element);
      if(!['SPAN','BR'].includes(element.tagName)||inline.display!=='inline'||inline.position!=='static'||inspectStyle(element)||textPseudoEffect(element)||!['normal','isolate'].includes(inline.unicodeBidi)||inlineProperties.some(name=>inline[name]!==style[name])||!opaqueRgb(inline.color)||inline.webkitTextFillColor!==inline.color||!transparent(inline.backgroundColor)||inline.backgroundImage!=='none'||inline.outlineStyle!=='none'&&parseFloat(inline.outlineWidth)>0&&!transparent(inline.outlineColor)||['Top','Right','Bottom','Left'].some(side=>parseFloat(inline[`border${side}Width`])>0||parseFloat(inline[`padding${side}`])!==0||parseFloat(inline[`margin${side}`])!==0))return {error:'인라인 요소의 독립 서식 또는 장식을 아직 관측할 수 없습니다.'};
    }
    if(inspectStyle(p)||!sources.length||sources.some(source=>!source.text))return {error:'문단의 텍스트 노드를 관측할 수 없습니다.'};
    if(textPseudoEffect(p))return {error:'첫 글자 또는 첫 줄의 별도 서식을 확인할 수 없습니다.'};
    const text=sources.map(source=>source.text).join('');
    if(!text||text.length>30000||(characters+=text.length)>30000||!(/^[\n\x20-\x7e\uac00-\ud7a3]+$/.test(text))||/^ | $| {2}|^\n|\n$|\n\n| \n|\n /.test(text)||sources.some(source=>source.kind==='text'&&/[\r\n]/.test(source.text)))return {error:'공백 축약 또는 문자 배치를 안전하게 관측할 수 없습니다.'};
    const family=style.fontFamily.split(',')[0].trim().replace(/^['"]|['"]$/g,''),fontSize=parseFloat(style.fontSize);
    const opaqueColor=opaqueRgb(style.color);
    if(!['arial','malgun gothic'].includes(family.toLowerCase())||!['400','700'].includes(style.fontWeight)||style.fontStyle!=='normal'||style.fontStretch!=='100%'||style.fontSizeAdjust!=='none'||style.fontKerning!=='auto'||['fontVariantLigatures','fontVariantCaps','fontVariantNumeric','fontVariantEastAsian','fontVariantPosition','fontVariantAlternates'].some(name=>style[name]!=='normal')||style.fontFeatureSettings!=='normal'||style.fontVariationSettings!=='normal'||style.webkitTextSecurity!=='none'||!Number.isFinite(fontSize)||fontSize<=0||style.letterSpacing!=='normal'||style.wordSpacing!=='0px'||style.textTransform!=='none'||style.textDecorationLine!=='none'||style.textEmphasisStyle!=='none'||style.writingMode!=='horizontal-tb'||style.direction!=='ltr'||!['normal','isolate'].includes(style.unicodeBidi)||!['left','start'].includes(style.textAlign)||parseFloat(style.textIndent)!==0||!opaqueColor||style.webkitTextFillColor!==style.color||parseFloat(style.webkitTextStrokeWidth)!==0||!['normal','nowrap'].includes(style.whiteSpace)||!transparent(style.backgroundColor)||style.backgroundImage!=='none'||style.outlineStyle!=='none'&&parseFloat(style.outlineWidth)>0&&!transparent(style.outlineColor)||['Top','Right','Bottom','Left'].some(side=>parseFloat(style[`border${side}Width`])>0||parseFloat(style[`padding${side}`])!==0))return {error:'문단의 글꼴·공백·가시성 또는 장식은 아직 독립 관측할 수 없습니다.'};
    try{await new FontFace('__pagination_font_probe',`local("${family}")`).load();}catch{return {error:'문단 글꼴을 확인할 수 없습니다.'};}
    const range=document.createRange(),nodeRects=[];
    for(const source of sources){if(source.kind==='text'){range.selectNodeContents(source.node);nodeRects.push(...[...range.getClientRects()].filter(rect=>rect.width>0&&rect.height>0));}}
    const rects=[];
    for(const rect of nodeRects){let line=rects.at(-1);if(line&&rect.top<line.top-.1)return {error:'문단의 실제 줄 순서를 확인할 수 없습니다.'};if(!line||rect.top>line.top+.1){line={x:rect.x,y:rect.y,left:rect.left,top:rect.top,right:rect.right,width:rect.width,height:rect.height};rects.push(line);}else{line.left=Math.min(line.left,rect.left);line.x=line.left;line.right=Math.max(line.right,rect.right);line.width=line.right-line.left;line.height=Math.max(line.height,rect.height);}}
    if(!rects.length||rects.length>5000||rects.some((rect,i)=>i>0&&rect.top<=rects[i-1].top+.1))return {error:'문단의 실제 줄 순서를 확인할 수 없습니다.'};
    const lineAt=top=>{let low=0,high=rects.length;while(low<high){const mid=(low+high)>>1;if(rects[mid].top<top-.1)low=mid+1;else high=mid;}return low<rects.length&&Math.abs(rects[low].top-top)<.1?low:-1;};
    const box=p.getBoundingClientRect();
    if(box.width<=0||box.height<=0||box.x<rootRect.x-.000001||box.y<rootRect.y-.000001||box.right>rootRect.right+.000001||box.bottom>rootRect.bottom+.000001)return {error:'문단의 실제 위치가 선택 영역 밖으로 벗어났습니다.'};
    const context=document.createElement('canvas').getContext('2d');context.font=`${style.fontWeight} ${fontSize}px "${family}"`;
    const lineTexts=rects.map(()=>''),parts=[];let lastLine=-1;
    let nodeIndex=0,nodeStart=0;
    for(const {index:offset,segment} of new Intl.Segmenter('und',{granularity:'grapheme'}).segment(text)){
      while(offset>=nodeStart+sources[nodeIndex].text.length){nodeStart+=sources[nodeIndex].text.length;nodeIndex++;}
      const source=sources[nodeIndex];
      if(offset+segment.length>nodeStart+source.text.length)return {error:'하나의 글자 조합이 여러 원문 노드에 걸쳐 있습니다.'};
      if(source.kind==='br'){
        const observed=[...source.node.getClientRects()].filter(rect=>rect.height>0),rect=observed[0],lineIndex=rect&&lineAt(rect.top);
        if(observed.length!==1||rect.width>.000001||lineIndex!==lastLine||lineIndex<0||lineIndex+1>=rects.length||Math.abs(rect.x-rects[lineIndex].right)>.2)return {error:'BR의 실제 줄 경계를 확인할 수 없습니다.'};
        parts.push({nodeIndex,start:0,end:1,lineIndex,hardBreak:true});continue;
      }
      range.setStart(source.node,offset-nodeStart);range.setEnd(source.node,offset-nodeStart+segment.length);const observed=[...range.getClientRects()].filter(rect=>rect.height>0);
      if(observed.length!==1&&observed.length!==2)return {error:'문자의 실제 줄 위치를 확인할 수 없습니다.'};
      const charRect=observed[0],lineIndex=lineAt(charRect.top);
      if(observed.length===2){
        const next=observed[1],nextIndex=lineAt(next.top);
        if(text[offset]!==' '||charRect.width>.000001||next.width>.000001||lineIndex<0||nextIndex!==lineIndex+1||Math.abs(charRect.x-rects[lineIndex].right)>.2||Math.abs(next.x-rects[nextIndex].left)>.2)return {error:'줄 경계 공백의 실제 위치를 확인할 수 없습니다.'};
      }
      if(lineIndex<0||lineIndex<lastLine||lineIndex>lastLine+1||text[offset]!==' '&&charRect.width<=0)return {error:'문자의 실제 줄 순서를 확인할 수 없습니다.'};
      const suppressed=text[offset]===' '&&charRect.width<=.000001;
      if(!suppressed)lineTexts[lineIndex]+=segment;
      parts.push({nodeIndex,start:offset-nodeStart,end:offset-nodeStart+segment.length,lineIndex,suppressed});lastLine=lineIndex;
    }
    const lines=rects.map((rect,lineIndex)=>{
      const rendered=lineTexts[lineIndex],metric=context.measureText(rendered);
      if(!rendered||![metric.width,metric.actualBoundingBoxLeft,metric.actualBoundingBoxRight,metric.actualBoundingBoxAscent,metric.actualBoundingBoxDescent].every(Number.isFinite)||Math.abs(metric.width-rect.width)>.5)return null;
      return {box:{x:rect.x-rootRect.x,y:rect.y-rootRect.y,w:rect.width,h:rect.height},text:rendered,metric:{width:metric.width,left:metric.actualBoundingBoxLeft,right:metric.actualBoundingBoxRight,ascent:metric.actualBoundingBoxAscent,descent:metric.actualBoundingBoxDescent}};
    });
    if(lines.some(line=>!line))return {error:'문단의 실제 글꼴 표시 범위를 확인할 수 없습니다.'};
    const modelStyle={fontFamily:family,fontSize,fontWeight:Number(style.fontWeight),fontStyle:'normal',color:style.color,textDecoration:'none',letterSpacing:'normal'};
    paragraphs.push({id:`paragraph-${index+1}`,zOrder,position:objectStyle.position,sources:sources.map((source,nodeIndex)=>({sourceId:`source-${index+1}-${nodeIndex+1}`,kind:source.kind,text:source.text,style:{...modelStyle,color:getComputedStyle(source.kind==='br'?source.node:source.node.parentElement).color}})),text,style:modelStyle,lineHeight:style.lineHeight,whiteSpace:style.whiteSpace,box:{x:box.x-rootRect.x,y:box.y-rootRect.y,w:box.width,h:box.height},lines,parts});
  }
  return {viewport,root:{x:rootRect.x,y:rootRect.y,width:rootRect.width,height:rootRect.height},fonts,paragraphs,tables,images,characters,documentSize:{width:document.documentElement.scrollWidth,height:document.documentElement.scrollHeight}};
}

async function probeBaselines(context,paragraphs){
  const page=await context.newPage();
  try{
    await page.setContent('<!doctype html><html><body style="margin:0"></body></html>');
    return await page.evaluate(items=>items.map(item=>{
      const host=document.createElement('span');host.style.cssText=`display:inline-block;font-size:${item.style.fontSize}px;font-weight:${item.style.fontWeight};font-style:normal;line-height:${item.lineHeight};white-space:pre`;host.style.fontFamily=`"${item.style.fontFamily}"`;
      const text=document.createTextNode(item.lines[0].text),marker=document.createElement('span');marker.style.cssText='display:inline-block;width:0;height:0;vertical-align:baseline;padding:0;border:0;margin:0';host.append(text,marker);document.body.append(host);
      const range=document.createRange();range.selectNode(text);const offset=marker.getBoundingClientRect().top-range.getBoundingClientRect().top;host.remove();return offset;
    }),paragraphs);
  }finally{await page.close();}
}

export async function observePaginationDocument({session,selector,viewport,timeoutMs=30000,assets=[],assetRefs=[]}={}){
  if(!session?.page||!session?.context||typeof selector!=='string'||!selector.trim()||!validViewport(viewport)||!Number.isSafeInteger(timeoutMs)||timeoutMs<1||timeoutMs>30000||!Array.isArray(session.blocked)||session.sourceHadScripts!==false)return unsupported('스크립트가 없는 선택 문서와 화면 크기가 필요합니다.');
  if(session.blocked.length)return unsupported('차단된 외부 자산이 있어 문서를 끝까지 관측할 수 없습니다.');
  const current=session.page.viewportSize();if(!current||current.width!==viewport.width||current.height!==viewport.height)return unsupported('선택한 화면 크기가 현재 화면과 다릅니다.');
  const deadline=Date.now()+timeoutMs;let timedOut=false;
  const beforeDeadline=operation=>new Promise((resolve,reject)=>{
    const remaining=deadline-Date.now(),expire=()=>{timedOut=true;try{Promise.resolve(session.context.close()).catch(()=>{});}catch{}reject(new Error('관측 시간 초과'));};
    if(remaining<=0){expire();return;}const timer=setTimeout(expire,remaining);
    Promise.resolve().then(operation).then(value=>{clearTimeout(timer);resolve(value);},error=>{clearTimeout(timer);reject(error);});
  });
  try{
    const first=await beforeDeadline(()=>session.page.evaluate(readSnapshot,{selector}));if(first.error)return unsupported(first.error);
    if(first.viewport.width!==viewport.width||first.viewport.height!==viewport.height||first.viewport.dpr!==1)return unsupported('선택 화면 또는 문서 크기를 안전하게 확인할 수 없습니다.');
    const readFonts=snapshot=>readPaginationFontEvidence({context:session.context,page:session.page,selector,paragraphs:snapshot.paragraphs,tables:snapshot.tables});
    const fontEvidence=await beforeDeadline(()=>readFonts(first));
    const characterCount=first.characters,documentY=first.root.y+first.viewport.scrollY,pixelHeight=Math.ceil(documentY+first.root.height)-Math.floor(documentY),tileCount=Math.ceil(pixelHeight/1024);
    if(characterCount*(tileCount*4+3)>2000000)return unsupported('문서 글자와 타일의 관측 작업량이 지원 한도를 초과합니다.');
    const offsets=await beforeDeadline(()=>probeBaselines(session.context,first.paragraphs));
    await beforeDeadline(()=>new Promise(resolve=>setTimeout(resolve,60)));
    const firstState=JSON.stringify(first);
    const readVerifiedState=async()=>{
      const current=await session.page.evaluate(readSnapshot,{selector});
      if(JSON.stringify(current)!==firstState)throw new Error('문서 타일 관측 중 원본 범위 또는 표시 상태가 바뀌었습니다.');
      return current;
    };
    const tiles=await beforeDeadline(()=>captureDocumentTiles({page:session.page,rect:{x:first.root.x+first.viewport.scrollX,y:first.root.y+first.viewport.scrollY,width:first.root.width,height:first.root.height},viewport,readState:readVerifiedState,timeoutMs:Math.max(1,deadline-Date.now())}));
    const nativeTables=first.tables.length?await beforeDeadline(()=>session.page.evaluate(measurePage,{selector,profile:'static-slides-v1',timeoutMs:Math.max(1,deadline-Date.now()),sourceHadScripts:false,tablesOnly:true})):null;
    const after=await beforeDeadline(()=>session.page.evaluate(readSnapshot,{selector}));
    if(after.error)return unsupported(after.error);
    const finalFonts=await beforeDeadline(()=>readFonts(after));
    if(JSON.stringify(after)!==firstState||JSON.stringify(fontEvidence)!==JSON.stringify(finalFonts)||session.blocked.length)return unsupported('관측 중 화면·글꼴·자산이 바뀌었습니다.');
    const sourceLedger=[],tableSourceLedger=[],observations=[],tokens=[],lineIds=[],lineOwners=[],objects=[],lineMeasurements=[];
    if(nativeTables){
      const tableNodes=nativeTables.slides?.[0]?.nodes;
      if(nativeTables.diagnostics.length||nativeTables.slides.length!==1||tableNodes?.length!==first.tables.length)return unsupported('편집 가능한 표의 격자·서식·위치를 원본과 일치시킬 수 없습니다.');
      for(const table of first.tables){
        const node=tableNodes.find(candidate=>candidate.kind==='table'&&candidate.sourceRef.path===table.sourcePath);
        if(!node||!['x','y','w','h'].every(key=>Math.abs(table.box[key]-node.boxCssPx[key])<=.05)||node.content.rows!==table.rowPaintBoxes.length||node.content.cells.length!==table.cells.length||node.content.rowHeightsCssPx.some((height,i)=>Math.abs(height-table.rowPaintBoxes[i].h)>.05))return unsupported('표의 원문 위치와 편집 가능한 표 격자가 다릅니다.');
        const content=structuredClone(node.content),sourceRanges=[];
        for(const [i,cell] of table.cells.entries()){
          const target=content.cells[i];
          if(!['row','col','rowSpan','colSpan','text'].every(key=>target[key]===cell[key])||cell.text&&target.runs.length!==1||target.runs.map(run=>run.text).join('')!==cell.text)return unsupported('표의 셀·병합 또는 원문 순서가 다릅니다.');
          if(cell.text){
            const range={sourceId:cell.sourceId,start:0,end:cell.text.length};
            Object.assign(target.runs[0],range);
            sourceRanges.push(range);
            tableSourceLedger.push({objectId:table.id,sourceId:cell.sourceId,row:cell.row,col:cell.col,text:cell.text,ranges:[{start:0,end:cell.text.length}]});
          }
        }
        objects.push({...node,id:table.id,zOrder:table.zOrder,parentId:null,boxCssPx:table.box,paintBoxCssPx:table.box,splitSafe:true,rowPaintBoxes:table.rowPaintBoxes,sourceRanges,content});
      }
    }
    for(const image of first.images){
      const asset=assets.find(candidate=>candidate.id===image.assetId);
      const reference=assetRefs.find(candidate=>candidate.usage==='image'&&candidate.sourceRef?.path===image.sourcePath);
      if(!asset||!reference||reference.assetId!==image.assetId||image.mime!==asset.mime||sha256(asset.bytes)!==asset.sha256||image.assetId!==`asset-${asset.sha256}`)return unsupported('그림의 실제 자산을 검증된 입력과 연결할 수 없습니다.');
      objects.push({id:image.id,zOrder:image.zOrder,kind:'image',parentId:null,splitSafe:false,boxCssPx:image.box,paintBoxCssPx:image.box,style:{opacity:1},content:{assetId:image.assetId},sourceRef:{slideId:'document',path:image.sourcePath}});
    }
    for(const [index,p] of first.paragraphs.entries()){
      const sources=p.sources.map(item=>({...item,objectId:p.id,ranges:[{start:0,end:item.text.length}],whiteSpace:p.whiteSpace})),box=p.box,lines=[];
      let left=box.x,top=box.y,right=box.x+box.w,bottom=box.y+box.h;
      for(const [lineIndex,item] of p.lines.entries()){
        const baseline=item.box.y+offsets[index];if(!Number.isFinite(baseline)||baseline<item.box.y-.05||baseline>item.box.y+item.box.h+.05)return unsupported('문단의 기준선을 확인할 수 없습니다.');
        const l=Math.min(item.box.x,item.box.x-item.metric.left)-1,t=Math.min(item.box.y,baseline-item.metric.ascent)-1,r=Math.max(item.box.x+item.box.w,item.box.x+item.metric.right)+1,b=Math.max(item.box.y+item.box.h,baseline+item.metric.descent)+1;
        if(l<0||t<0||r>first.root.width||b>first.root.height)return unsupported('글자의 실제 표시 범위가 선택 문서 밖으로 벗어났습니다.');
        left=Math.min(left,l);top=Math.min(top,t);right=Math.max(right,r);bottom=Math.max(bottom,b);
        const lineId=`line-${index+1}-${lineIndex+1}`,line={lineId,boxCssPx:item.box,paintBoxCssPx:{x:l,y:t,w:r-l,h:b-t},baseline};lines.push(line);lineIds.push(lineId);lineOwners.push({lineId,objectId:p.id});lineMeasurements.push({...line,objectId:p.id});
      }
      let previousLine=0,previousStyle=p.style;
      for(const part of p.parts){
        if(part.lineIndex!==previousLine){tokens.push({kind:'layout-break',objectId:p.id,lineId:lines[previousLine].lineId,nextLineId:lines[part.lineIndex].lineId,sources:[],text:'\n',style:previousStyle});previousLine=part.lineIndex;}
        const source=sources[part.nodeIndex],text=source.text.slice(part.start,part.end),lineId=lines[part.lineIndex].lineId,kind=part.suppressed?'suppressed':'visible';
        previousStyle=source.style;
        if(part.hardBreak){
          observations.push({sourceId:source.sourceId,objectId:p.id,start:0,end:1,lineId,kind:'hard-break',text:'\n',style:source.style});
          tokens.push({kind:'hard-break',objectId:p.id,lineId,nextLineId:lines[part.lineIndex+1].lineId,sources:[{sourceId:source.sourceId,start:0,end:1}],text:'\n',style:source.style});
          previousLine=part.lineIndex+1;
          continue;
        }
        observations.push({sourceId:source.sourceId,objectId:p.id,start:part.start,end:part.end,lineId,kind,text:part.suppressed?'':text,style:source.style});
        tokens.push({kind:part.suppressed?'suppressed-space':'copy',objectId:p.id,lineId,sources:[{sourceId:source.sourceId,start:part.start,end:part.end}],text:part.suppressed?'':text,style:source.style});
      }
      const paint={x:left,y:top,w:right-left,h:bottom-top};sourceLedger.push(...sources);
      objects.push({id:p.id,zOrder:p.zOrder,kind:'text',parentId:null,splitSafe:true,style:{...p.style,lineHeight:p.lineHeight,whiteSpace:p.whiteSpace,align:'left',valign:'top',padding:{top:0,right:0,bottom:0,left:0},fill:'transparent',strokeWidth:0,backgroundImage:'none',opacity:1},boxCssPx:box,paintBoxCssPx:paint,content:{text:p.text,runs:sources.map(source=>({text:source.text,style:source.style}))},lines});
    }
    objects.sort((a,b)=>a.zOrder-b.zOrder);
    // With static descendants and no additional stacking contexts, relative
    // siblings at z-index:auto paint in DOM order. Other overlapping flow
    // objects can paint in separate phases even without a declared z-index.
    const byId=new Map([...first.paragraphs,...first.tables,...first.images].map(item=>[item.id,item.position]));
    const orderedPaint=objects.map(object=>({box:object.paintBoxCssPx,position:byId.get(object.id)})).sort((a,b)=>a.box.y-b.box.y);
    let comparisons=0;
    for(let i=0;i<orderedPaint.length;i++)for(let j=i+1;j<orderedPaint.length;j++){
      const a=orderedPaint[i].box,b=orderedPaint[j].box;
      if(b.y>=a.y+a.h-.000001)break;
      if(++comparisons>100000)return unsupported('겹친 문서 객체의 수가 지원 한도를 초과합니다.');
      const overlaps=a.x<b.x+b.w-.000001&&b.x<a.x+a.w-.000001;
      if(overlaps&&(orderedPaint[i].position!=='relative'||orderedPaint[j].position!=='relative'))return unsupported('서로 겹친 객체의 실제 출력 순서를 확인할 수 없습니다.');
    }
    const sourceObjectIds=objects.map(object=>object.id);
    const displayMapping={sourceObjectIds,objectIds:[...sourceObjectIds],sourceBreakIds:sourceLedger.filter(source=>source.kind==='br').map(source=>source.sourceId),lineIds,lineOwners,sourceLedger,observations,tokens};
    if(!validateDisplayMapping(displayMapping).valid)return unsupported('원문과 표시 글자의 대응을 확인할 수 없습니다.');
    if(Date.now()>=deadline){timedOut=true;try{Promise.resolve(session.context.close()).catch(()=>{});}catch{}return timeoutFailure();}
    return {status:'observed',documentWidth:first.root.width,documentHeight:first.root.height,sourceObjectIds,sourceLedger,tableSourceLedger,observations,objects,lineMeasurements,displayMapping,fontEvidence,viewport,fingerprint:sha256(JSON.stringify({first,fontEvidence,tiles:tiles.fingerprint})),tileCoverage:{bounds:tiles.bounds,count:tiles.tiles.length},selectionVerified:false,exportReady:false,coverage:{selectedRootOnly:true,outsideSelection:'verified-no-visible-content'},remainingChecks:['verified-selection','worker-deadline-supervisor','model-and-ppt-output'],diagnostics:[]};
  }catch(error){
    if(error?.reasonKey==='PAGINATION_TIMEOUT'&&!timedOut){timedOut=true;try{Promise.resolve(session.context.close()).catch(()=>{});}catch{}}
    return timedOut?timeoutFailure():unsupported('문서의 화면·글꼴·기준선을 안전하게 관측할 수 없습니다.');
  }
}
