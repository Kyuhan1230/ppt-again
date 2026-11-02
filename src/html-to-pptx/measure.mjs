import {diagnostic,validateModel} from './model.mjs';
import {resolveAsset} from './assets.mjs';
import {identifyFontMappings,installedNativeAliases} from './font-profiles.mjs';
import {captureDecorations} from './decorations.mjs';
import {measureSvgInput,measureSvgFragments} from './svg-slides.mjs';
import {collectFontAssets} from './font-assets.mjs';
import {normalizeWoff2Fonts} from './woff2-fonts.mjs';
import {applyFallbackFont} from './fallback-font.mjs';
import {renderChartSvg} from './charts.mjs';

export function scopeEmbeddedModelErrors(errors,slides){
  return errors.map(error=>{const match=/^slides\[(\d+)\]\./.exec(error.path),slide=match?slides[Number(match[1])]:null;return diagnostic('INPUT_INVALID',`${error.path}: ${error.message}`,slide?{slideId:slide.id}:{});});
}

// Runs inside a script-disabled browser. This is repository-owned measurement code.
async function measurePage({selector,profile,timeoutMs,sourceHadScripts,allowTableOverlay=false,nativeAliases=[],charts=[]}) {
  const editable=['editable-slides-v1','editable-slides-v2'].includes(profile),declared=editable||profile==='project-data-v1';
  const diagnostics=[];
  const issue=(code,message,slideId=null,nodeId=null,sourceRef=null)=>diagnostics.push({code,severity:'error',slideId,nodeId,sourceRef,message});
  const timeout=(promise)=>Promise.race([promise,new Promise((_,reject)=>setTimeout(()=>reject(new Error('loading timeout')),timeoutMs))]);
  try { await timeout(document.fonts.ready); }catch {issue('RENDER_INCOMPLETE','글꼴 로딩이 끝나지 않았습니다.');}
  try { await timeout(Promise.all([...document.images].map(img=>img.decode()))); }catch {issue('RENDER_INCOMPLETE','이미지 디코딩이 완료되지 않았습니다.');}
  const backgroundUrls=new Set();
  for(const el of document.querySelectorAll('*'))for(const match of getComputedStyle(el).backgroundImage.matchAll(/url\((?:"([^"]*)"|'([^']*)'|([^)]*))\)/g))backgroundUrls.add(match[1]??match[2]??match[3]);
  for(const url of backgroundUrls){try{const img=new Image();img.src=url;await timeout(img.decode());}catch{issue('RENDER_INCOMPLETE','CSS 배경 이미지 디코딩이 완료되지 않았습니다.');}}
  let roots;
  try{roots=[...document.querySelectorAll(selector)];}catch{issue('SLIDE_ROOT_REQUIRED','올바른 장표 선택자가 필요합니다.');return {slides:[],measurements:[],diagnostics};}
  if(!roots.length){issue(sourceHadScripts&&!document.body.innerText.trim()?'SCRIPT_REQUIRED':'SLIDE_ROOT_REQUIRED','장표 영역을 자동으로 찾지 못했습니다. 고정 크기의 HTML 장표 또는 SVG 장표인지 확인해 주세요.');return {slides:[],measurements:[],diagnostics};}
  const round=v=>Math.round(v*10000)/10000;
  const visibleColor=c=>c!=='transparent'&&!/^rgba\(.*[, /]0\)$/.test(c);
  const fontAvailability=new Map(),customFonts=new Set();
  const familyName=s=>s.split(',')[0].trim().replace(/^["']|["']$/g,'');
  const fontAvailable=async(family)=>{
    if(fontAvailability.has(family))return fontAvailability.get(family);
    let ok=['serif','sans-serif','monospace','system-ui','cursive','fantasy'].includes(family);
    if(!ok){
      const faces=[...document.fonts].filter(f=>f.family.replace(/^["']|["']$/g,'')===family);
      if(faces.length){customFonts.add(family);try{await timeout(document.fonts.load(`20px "${family.replaceAll('"','\\"')}"`));ok=faces.every(f=>f.status==='loaded');}catch{ok=false;}}
      else{
        for(const localName of [family,...nativeAliases.filter(a=>a.family===family).map(a=>a.localName)]){
          try{const face=new FontFace('__availability_probe',`local(${JSON.stringify(localName)})`);await timeout(face.load());ok=true;break;}catch{ok=false;}
        }
      }
    }
    fontAvailability.set(family,ok);return ok;
  };
  const slides=[],measurements=[];let tableCells=0;
  for(const [slideIndex,root] of roots.entries()) {
    const id=(declared?root.dataset.slideId:null)||root.id||`slide-${slideIndex+1}`;
    const rootStyle=getComputedStyle(root),before=root.getBoundingClientRect();
    const width=parseFloat(rootStyle.width),height=parseFloat(rootStyle.height);
    const widthValue=root.computedStyleMap?.().get('width')?.toString(),heightValue=root.computedStyleMap?.().get('height')?.toString();
    if(!width||!height||(widthValue==='auto'||heightValue==='auto')&&rootStyle.aspectRatio==='auto') {issue('SLIDE_ROOT_REQUIRED','장표의 고정 너비와 높이를 확인할 수 없습니다.',id,null,{slideId:id,path:root.dataset.sourcePath??selector});continue;}
    if(roots.some(r=>r!==root&&r.contains(root))){issue('SLIDE_ROOT_REQUIRED','장표 루트가 중첩됩니다.',id);continue;}
    if(rootStyle.transform!=='none') {const m=new DOMMatrix(rootStyle.transform);if(m.b!==0||m.c!==0||m.a<=0||Math.abs(m.a-m.d)>.0001){issue('UNSUPPORTED_OBJECT','장표 루트의 비균일/회전 변환은 지원하지 않습니다.',id);continue;}}
    const sourceTransformScale=round(before.width/width);
    if(slides.length&&(Math.abs(slides[0].widthCssPx-width)>.01||Math.abs(slides[0].heightCssPx-height)>.01))issue('UNSUPPORTED_OBJECT','한 PPT 안의 서로 다른 장표 크기는 지원하지 않습니다.',id);
    const renderScale=1280/Math.max(width,height),referenceSize={width:Math.round(width*renderScale),height:Math.round(height*renderScale)};
    // The responsive frame is presentation chrome, not a clipping mask on the
    // logical slide. Its dimensions must not crop a normalized measurement.
    if(editable&&root.parentElement?.classList.contains('slide-view')){
      root.parentElement.style.width=`${referenceSize.width}px`;root.parentElement.style.maxWidth='none';
      root.parentElement.style.height=`${referenceSize.height}px`;root.parentElement.style.overflow='visible';
    }
    // Render at the declared output size from the start; never rescale captures.
    root.style.setProperty('transform',`scale(${renderScale})`,'important');root.style.setProperty('transform-origin','0 0','important');
    const rr=root.getBoundingClientRect();
    const boxOf=el=>{const r=el.getBoundingClientRect();return {x:round((r.x-rr.x)/renderScale),y:round((r.y-rr.y)/renderScale),w:round(r.width/renderScale),h:round(r.height/renderScale)};};
    const nodes=[],objects=[],seen=new Set(),textOwners=new Set(),domOrder=new WeakMap(),orderKeys=new Map();let order=0,domIndex=0;
    const domWalker=document.createTreeWalker(root,NodeFilter.SHOW_ELEMENT|NodeFilter.SHOW_TEXT);domOrder.set(root,domIndex++);
    while(domWalker.nextNode())domOrder.set(domWalker.currentNode,domIndex++);
    const elements=declared?[...root.querySelectorAll('[data-source-id]')]:[root,...root.querySelectorAll('*')];
    if(editable&&[...root.children].some(e=>!e.hasAttribute('data-source-id')))issue('UNSUPPORTED_OBJECT','객체 표식이 없는 장표 콘텐츠가 있습니다.',id);
    const hiddenByOpacity=el=>{for(let p=el;p;p=p.parentElement)if(Number(getComputedStyle(p).opacity)===0)return true;return false;};
    const plainStyle=st=>({fontFamily:familyName(st.fontFamily),fontSize:parseFloat(st.fontSize),fontWeight:Number(st.fontWeight)||400,fontStyle:st.fontStyle,color:st.color,align:st.textAlign,lineHeight:st.lineHeight,whiteSpace:st.whiteSpace,
      fill:st.backgroundColor,backgroundImage:st.backgroundImage,stroke:st.borderTopColor,strokeWidth:parseFloat(st.borderTopWidth)||0,borderStyle:st.borderTopStyle,radius:parseFloat(st.borderTopLeftRadius)||0,opacity:Number(st.opacity),padding:{top:parseFloat(st.paddingTop),right:parseFloat(st.paddingRight),bottom:parseFloat(st.paddingBottom),left:parseFloat(st.paddingLeft)}});
    const textRuns=el=>{
      const runs=[],walker=document.createTreeWalker(el,NodeFilter.SHOW_TEXT|NodeFilter.SHOW_ELEMENT);let child;
      while((child=walker.nextNode())){
        if(child.nodeType!==3&&child.tagName!=='BR')continue;
        const parent=child.nodeType===3?child.parentElement:child;
        if(parent.closest('svg'))continue;
        const cs=getComputedStyle(parent);
        if(hiddenByOpacity(parent)||cs.visibility==='hidden'||!parent.getClientRects().length)continue;
        let text=child.nodeType===3?child.data:'\n';
        if(child.nodeType===3&&!['pre','pre-wrap','break-spaces'].includes(cs.whiteSpace))text=text.replace(cs.whiteSpace==='pre-line'?/[\t\r\f ]+/g:/[\t\r\n\f ]+/g,' ');
        const style={fontFamily:familyName(cs.fontFamily),fontSize:parseFloat(cs.fontSize),fontWeight:Number(cs.fontWeight)||400,fontStyle:cs.fontStyle,color:cs.color,textDecoration:cs.textDecorationLine,letterSpacing:cs.letterSpacing};
        if(parseFloat(cs.webkitTextStrokeWidth)>0)style.outline={color:cs.webkitTextStrokeColor,width:parseFloat(cs.webkitTextStrokeWidth)};
        if(text)runs.push({text,style});
      }
      if(profile!=='project-data-v1'&&!['pre','pre-wrap','break-spaces'].includes(getComputedStyle(el).whiteSpace)){
        for(let i=1;i<runs.length;i++)if(runs[i-1].text.endsWith(' '))runs[i].text=runs[i].text.replace(/^ +/,'');
        if(runs.length){runs[0].text=runs[0].text.replace(/^ +/,'');runs.at(-1).text=runs.at(-1).text.replace(/ +$/,'');}
      }
      return runs.filter(r=>r.text.length);
    };
    const measureTable=async(table,box,at)=>{
      const rows=[...table.rows],tableStyle=getComputedStyle(table);
      if(!rows.length||rows.length>100||rows.reduce((n,row)=>n+row.cells.length,0)>2000){at('INPUT_INVALID','표는 100행·50열·2000셀, 문서는 10000셀 이내여야 합니다.');return null;}
      if(table.querySelector('table')||[...table.querySelectorAll('*')].some(e=>!['THEAD','TBODY','TFOOT','TR','TH','TD','COLGROUP','COL','DIV','SPAN','BR','B','I','STRONG','EM','U','SMALL'].includes(e.tagName))){at('UNSUPPORTED_OBJECT','중첩 표·그림·입력 요소·캡션은 일반 표에 넣을 수 없습니다.');return null;}
      if(!['table','inline-table'].includes(tableStyle.display)||tableStyle.direction!=='ltr'||tableStyle.emptyCells!=='show'||tableStyle.borderCollapse!=='collapse'&&tableStyle.borderSpacing.split(' ').some(v=>parseFloat(v)!==0)){at('UNSUPPORTED_OBJECT','표의 방향·빈 셀 숨김·셀 사이 간격은 아직 지원하지 않습니다.');return null;}
      const sides=['Top','Right','Bottom','Left'];
      const ruledRows=rows.filter(row=>sides.some(side=>parseFloat(getComputedStyle(row)[`border${side}Width`])>0));
      const simpleRowRules=tableStyle.borderCollapse==='collapse'&&ruledRows.length===1&&ruledRows[0]===rows[0]&&[table,...table.querySelectorAll('thead,tbody,tfoot,colgroup,col')].every(e=>sides.every(side=>parseFloat(getComputedStyle(e)[`border${side}Width`])===0))&&[...rows[0].cells].every(e=>e.rowSpan===1&&sides.every(side=>parseFloat(getComputedStyle(e)[`border${side}Width`])===0))&&(!rows[1]||[...rows[1].cells].every(e=>parseFloat(getComputedStyle(e).borderTopWidth)===0))&&['Left','Right'].every(side=>parseFloat(getComputedStyle(rows[0])[`border${side}Width`])===0)&&['Top','Bottom'].every(side=>['none','solid'].includes(getComputedStyle(rows[0])[`border${side}Style`]));
      const simpleCellBlocks=cell=>{
        if(!cell||!['TD','TH'].includes(cell.tagName)||cell.children.length<2||[...cell.childNodes].some(n=>n.nodeType===3&&n.textContent.trim()))return false;
        const base=getComputedStyle(cell);
        return [...cell.children].every(e=>{
          const cs=getComputedStyle(e),range=document.createRange();range.selectNodeContents(e);
          const lines=new Set([...range.getClientRects()].filter(r=>r.width>0&&r.height>0).map(r=>Math.round(r.top*2)));
          const bounds=e.getBoundingClientRect(),ink=range.getBoundingClientRect();
          const inline=[...e.querySelectorAll('*')].every(child=>{const s=getComputedStyle(child);return ['SPAN','B','I','STRONG','EM','U'].includes(child.tagName)&&s.display==='inline'&&s.fontSize===cs.fontSize&&s.lineHeight===cs.lineHeight&&!visibleColor(s.backgroundColor)&&['Top','Right','Bottom','Left'].every(side=>parseFloat(s[`margin${side}`])===0&&parseFloat(s[`padding${side}`])===0&&parseFloat(s[`border${side}Width`])===0);});
          return e.tagName==='SPAN'&&inline&&cs.display==='block'&&!visibleColor(cs.backgroundColor)&&['Top','Right','Bottom','Left'].every(side=>parseFloat(cs[`margin${side}`])===0&&parseFloat(cs[`padding${side}`])===0&&parseFloat(cs[`border${side}Width`])===0)&&cs.textAlign===base.textAlign&&cs.lineHeight===base.lineHeight&&(cs.whiteSpace===base.whiteSpace||cs.whiteSpace==='nowrap'&&base.whiteSpace==='normal'&&ink.left>=bounds.left-.1&&ink.right<=bounds.right+.1)&&Number.isFinite(parseFloat(cs.lineHeight))&&Math.abs(parseFloat(cs.height)-parseFloat(cs.lineHeight))<.1&&lines.size===1;
        });
      };
      for(const element of [table,...table.querySelectorAll('*')]){
        const cs=getComputedStyle(element);
        const relativeCell=['TD','TH'].includes(element.tagName)&&cs.position==='relative'&&['top','right','bottom','left'].every(side=>cs[side]==='auto'||parseFloat(cs[side])===0);
        if(cs.display==='none'||cs.visibility!=='visible'||hiddenByOpacity(element)||cs.transform!=='none'||cs.filter!=='none'||cs.clipPath!=='none'||cs.perspective!=='none'||cs.animationName!=='none'||cs.writingMode!=='horizontal-tb'||cs.direction!=='ltr'||Number(cs.opacity)!==1||cs.backgroundImage!=='none'||cs.boxShadow!=='none'||cs.textShadow!=='none'||['hidden','clip','scroll','auto'].includes(cs.overflowX)||['hidden','clip','scroll','auto'].includes(cs.overflowY)||element!==table&&cs.position!=='static'&&!relativeCell)at('UNSUPPORTED_OBJECT','표 안의 숨김·변환·이미지 배경·효과는 아직 지원하지 않습니다.');
        if(element.tagName==='DIV'){
          const cell=element.parentElement,cellStyle=getComputedStyle(cell);
          const range=document.createRange();range.selectNodeContents(element);
          const lines=new Set([...range.getClientRects()].filter(rect=>rect.width>0&&rect.height>0).map(rect=>Math.round(rect.top*2)));
          if(!['TD','TH'].includes(cell.tagName)||cell.children.length!==1||[...cell.childNodes].some(n=>n.nodeType===3&&n.textContent.trim())||cs.display!=='block'||visibleColor(cs.backgroundColor)||sides.some(side=>parseFloat(cs[`margin${side}`])!==0||parseFloat(cs[`padding${side}`])!==0)||cs.textAlign!==cellStyle.textAlign||cs.lineHeight!==cellStyle.lineHeight||cs.whiteSpace!==cellStyle.whiteSpace||lines.size!==1||!Number.isFinite(parseFloat(cs.lineHeight))||Math.abs(parseFloat(cs.height)-parseFloat(cs.lineHeight))>.1||Math.abs(element.getBoundingClientRect().width/renderScale-(cell.clientWidth-parseFloat(cellStyle.paddingLeft)-parseFloat(cellStyle.paddingRight)))>1.1)at('UNSUPPORTED_OBJECT','셀 안 블록은 여백이나 별도 배치가 없는 한 줄 문구만 지원합니다.');
        }
        if(['TD','TH'].includes(element.tagName)&&tableStyle.borderCollapse!=='collapse'&&sides.some(side=>parseFloat(cs[`border${side}Width`])>0)){at('UNSUPPORTED_OBJECT','셀 테두리가 있는 표는 border-collapse:collapse로 지정해 주세요.');return null;}
        if(!['TD','TH'].includes(element.tagName)&&sides.some(side=>parseFloat(cs[`border${side}Width`])>0)&&!(simpleRowRules&&element===ruledRows[0]))at('UNSUPPORTED_OBJECT','표 테두리는 각 셀에 지정해 주세요.');
        if(['SPAN','B','I','STRONG','EM','U','SMALL'].includes(element.tagName)&&(visibleColor(cs.backgroundColor)||sides.some(side=>parseFloat(cs[`padding${side}`])>0||parseFloat(cs[`margin${side}`])!==0)||cs.display!=='inline'&&!simpleCellBlocks(element.parentElement)))at('UNSUPPORTED_OBJECT','셀 안의 별도 배경·여백·블록 배치는 아직 지원하지 않습니다.');
        if(['COL','COLGROUP'].includes(element.tagName)&&visibleColor(cs.backgroundColor))at('UNSUPPORTED_OBJECT','열 배경은 셀 배경으로 지정해 주세요.');
        for(const pseudo of ['::before','::after']){const ps=getComputedStyle(element,pseudo);if(['none','normal','""'].includes(ps.content))continue;let glyph='';try{glyph=JSON.parse(ps.content);}catch{}const arrow=['TD','TH'].includes(element.tagName)&&/^[→←↑↓❯❮]$/u.test(glyph)&&ps.position==='absolute'&&ps.display==='block'&&ps.transform==='none'&&ps.filter==='none'&&ps.clipPath==='none'&&ps.textShadow==='none'&&Number(ps.opacity)===1&&!visibleColor(ps.backgroundColor)&&ps.backgroundImage==='none'&&sides.every(side=>parseFloat(ps[`padding${side}`])===0&&parseFloat(ps[`border${side}Width`])===0);at('UNSUPPORTED_OBJECT',arrow?'표의 장식 화살표는 배경 그림에 보존합니다.':'표의 생성 콘텐츠를 명시적인 셀 문구로 넣어 주세요.');}
      }
      const occupied=Array.from({length:rows.length},()=>Array(50).fill(false)),starts=[];let columns=0;
      for(const [row,tr] of rows.entries()){
        if(getComputedStyle(tr).display!=='table-row'){at('UNSUPPORTED_OBJECT','표의 각 행은 실제 table-row여야 합니다.');return null;}
        let col=0;
        for(const cell of tr.cells){
          while(occupied[row][col])col++;
          const rowSpan=cell.rowSpan,colSpan=cell.colSpan;
          if(rowSpan<1||row+rowSpan>rows.length||colSpan<1||col+colSpan>50){at('UNSUPPORTED_OBJECT','병합 범위가 표 밖이거나 rowspan=0입니다.');return null;}
          for(let r=row;r<row+rowSpan;r++)for(let c=col;c<col+colSpan;c++){
            if(occupied[r][c]){at('UNSUPPORTED_OBJECT','표의 병합 범위가 겹칩니다.');return null;}occupied[r][c]=true;
          }
          starts.push({row,col,rowSpan,colSpan,cell});col+=colSpan;columns=Math.max(columns,col);
        }
      }
      if(!columns||rows.length*columns>2000||(tableCells+=rows.length*columns)>10000){at('INPUT_INVALID','표는 100행·50열·2000셀, 문서는 10000셀 이내여야 합니다.');return null;}
      if(occupied.some(row=>row.slice(0,columns).some(value=>!value))){at('UNSUPPORTED_OBJECT','시작 셀과 병합 범위가 전체 표 격자를 덮어야 합니다.');return null;}
      const rowBoxes=rows.map(boxOf),edges=new Map(),cells=[];let consistent=true;
      const boundary=(index,x)=>{if(index<=0||index>=columns)return;if(edges.has(index)&&Math.abs(edges.get(index)-x)>.05)consistent=false;else edges.set(index,x);};
      for(const {col,colSpan,cell} of starts){const bounds=boxOf(cell);boundary(col,bounds.x);boundary(col+colSpan,bounds.x+bounds.w);}
      let colIndex=0;
      for(const group of table.querySelectorAll('colgroup'))for(const col of group.children.length?[...group.children]:[group]){
        const bounds=boxOf(col);boundary(colIndex,bounds.x);colIndex+=col.span;boundary(colIndex,bounds.x+bounds.w);
      }
      if(colIndex>columns){at('UNSUPPORTED_OBJECT','명시 열에 대응하는 셀 또는 병합 범위가 없습니다.');return null;}
      const xCuts=Array.from({length:columns+1},(_,i)=>i===0?box.x:i===columns?box.x+box.w:edges.get(i)),yCuts=[box.y,...rowBoxes.slice(1).map(b=>b.y),box.y+box.h];
      if(!consistent||xCuts.some(x=>!Number.isFinite(x))){at('UNSUPPORTED_OBJECT','모든 행과 명시 열에서 표의 열 경계를 확정할 수 없습니다.');return null;}
      for(const {row,col,rowSpan,colSpan,cell} of starts){
          const cs=getComputedStyle(cell),bounds=boxOf(cell);
          const end=rowBoxes[row+rowSpan-1];
          if(cs.display!=='table-cell'||Math.abs(bounds.y-rowBoxes[row].y)>.05||Math.abs(bounds.y+bounds.h-end.y-end.h)>.05){at('UNSUPPORTED_OBJECT','표의 셀 경계를 하나의 격자로 확정할 수 없습니다.');return null;}
          const style=plainStyle(cs),runs=simpleCellBlocks(cell)?[...cell.children].flatMap((e,i)=>{const parts=textRuns(e);return i&&parts.length?[{text:'\n',style:parts[0].style},...parts]:parts;}):textRuns(cell),text=cell.innerText;
          if(['normal','nowrap'].includes(cs.whiteSpace)&&[...cell.querySelectorAll('*')].every(e=>['normal','nowrap'].includes(getComputedStyle(e).whiteSpace))){
            for(let i=0;i<runs.length;i++)if(runs[i].text==='\n'){
              if(i>0)runs[i-1].text=runs[i-1].text.replace(/ +$/,'');
              if(i+1<runs.length)runs[i+1].text=runs[i+1].text.replace(/^ +/,'');
            }
          }
          style.align=({start:'left',end:'right'})[style.align]??style.align;style.valign=({top:'top',middle:'mid',bottom:'bottom'})[cs.verticalAlign];
          if(!style.valign)at('UNSUPPORTED_OBJECT','셀 세로 정렬은 top·middle·bottom만 지원합니다.');
          if(runs.map(r=>r.text).join('')!==text)at('UNSUPPORTED_OBJECT','현재 셀 문구와 텍스트 조각의 공백·줄바꿈을 일치시킬 수 없습니다.');
          let translucent=false;
          for(let ancestor=cell;ancestor&&table.contains(ancestor);ancestor=ancestor.parentElement){
            const color=getComputedStyle(ancestor).backgroundColor;if(!visibleColor(color))continue;
            if(translucent){at('UNSUPPORTED_OBJECT','겹친 표 배경의 반투명 합성은 아직 지원하지 않습니다.');break;}
            style.fill=color;const rgb=color.match(/^rgba?\(([^)]+)\)$/),alpha=rgb?Number(rgb[1].split(/[, /]+/)[3]??1):1;
            if(alpha>=1)break;translucent=true;
          }
          style.borders={};for(const [name,side] of [['top','Top'],['right','Right'],['bottom','Bottom'],['left','Left']]){
            const borderSource=simpleRowRules&&cell.parentElement===ruledRows[0]&&['Top','Bottom'].includes(side)?getComputedStyle(ruledRows[0]):cs;
            const width=parseFloat(borderSource[`border${side}Width`])||0,borderStyle=borderSource[`border${side}Style`];
            if(!['none','solid'].includes(borderStyle))at('UNSUPPORTED_OBJECT','셀 테두리는 단색 실선만 지원합니다.');
            style.borders[name]={color:borderSource[`border${side}Color`],width,style:width?'solid':'none'};
          }
          for(const family of new Set([style.fontFamily,...runs.map(r=>r.style.fontFamily)]))if(!await fontAvailable(family))at('FONT_UNAVAILABLE',`표 글꼴을 사용할 수 없습니다: ${family}`);
          cells.push({row,col,rowSpan,colSpan,text,runs,style});
      }
      return {rows:rows.length,columns,columnWidthsCssPx:xCuts.slice(1).map((v,i)=>round(v-xCuts[i])),rowHeightsCssPx:yCuts.slice(1).map((v,i)=>round(v-yCuts[i])),cells};
    };
    const initialLayout=elements.map(boxOf);
    const lineMetrics=el=>{
      const walker=document.createTreeWalker(el,NodeFilter.SHOW_TEXT);const chars=[];let tn;
      while((tn=walker.nextNode())) {
        if(tn.parentElement.closest('svg'))continue;
        for(let i=0;i<tn.length;i++){
          const range=document.createRange();range.setStart(tn,i);range.setEnd(tn,i+1);
          const r=range.getBoundingClientRect(),char=tn.data[i];
          if(r.height>0&&r.width>0&&char!=='\n'&&char!=='\r')chars.push({text:char,x:round((r.x-rr.x)/renderScale),y:round((r.y-rr.y)/renderScale),w:round(r.width/renderScale),h:round(r.height/renderScale)});
        }
      }
      const lines=[];
      for(const c of chars){let line=lines.find(l=>Math.abs(l.y-c.y)<.5);if(!line){line={text:'',x:c.x,y:c.y,w:0,h:c.h};lines.push(line);}line.text+=c.text;line.w=round(Math.max(line.x+line.w,c.x+c.w)-Math.min(line.x,c.x));line.x=Math.min(line.x,c.x);line.h=Math.max(line.h,c.h);}
      const marker=document.createElement('span');marker.style.cssText='display:inline-block;width:0;height:0;vertical-align:baseline;padding:0;border:0;margin:0';el.prepend(marker);
      const baseline=round((marker.getBoundingClientRect().top-rr.y)/renderScale);marker.remove();
      return {lines,baseline,baselineMethod:'zero-size inline baseline marker',lineAdvance:lines.length>1?round(lines[1].y-lines[0].y):null};
    };
    const directBaseline=(value,st,rect)=>{
      const probe=document.createElement('span');probe.style.cssText='all:initial!important;position:fixed!important;top:0!important;left:0!important;visibility:hidden!important;display:inline-block!important;white-space:pre!important';
      for(const key of ['font-family','font-size','font-weight','font-style','font-stretch','line-height','letter-spacing','word-spacing'])probe.style.setProperty(key,st.getPropertyValue(key),'important');
      const text=document.createTextNode(value),marker=document.createElement('span');marker.style.cssText='all:initial!important;display:inline-block!important;width:0!important;height:0!important;vertical-align:baseline!important';probe.append(text,marker);document.body.append(probe);
      const range=document.createRange();range.selectNode(text);const result=round((rect.top-rr.top)/renderScale+marker.getBoundingClientRect().top-range.getBoundingClientRect().top);probe.remove();return result;
    };
    for(const [index,el] of elements.entries()) {
      if(['STYLE','SCRIPT','LINK','META','BR'].includes(el.tagName)||el.parentElement?.closest('svg')||el.closest('svg')&&el.tagName.toLowerCase()!=='svg')continue;
      if(el.closest('table')&&el.closest('table')!==el)continue;
      if(el.closest('figure[data-chart-id]')&&el.closest('figure[data-chart-id]')!==el)continue;
      const st=getComputedStyle(el);if(st.display==='none'||st.visibility==='hidden'||hiddenByOpacity(el))continue;
      const sourceId=declared?el.dataset.sourceId:null;
      const nodeId=sourceId||el.id||`node-${index+1}`,sourceRef={slideId:id,...(sourceId?{elementId:sourceId}:{path:el.dataset.sourcePath??`node[${index}]`})};
      const at=(code,msg)=>issue(code,msg,id,nodeId,sourceRef);
      if(['CANVAS','IFRAME','VIDEO','AUDIO','OBJECT','EMBED'].includes(el.tagName)){at('UNSUPPORTED_OBJECT',`${el.tagName} 객체는 지원하지 않습니다.`);continue;}
      if(profile!=='project-data-v1'&&((el!==root&&st.transform!=='none')||st.filter!=='none'||st.perspective!=='none'||st.clipPath!=='none'||st.animationName!=='none'))at('UNSUPPORTED_OBJECT','객체의 CSS 변환·필터·클리핑·애니메이션을 지원하지 않습니다.');
      if(Number(st.opacity)<1&&el.children.length)at('UNSUPPORTED_OBJECT','여러 자식을 합성하는 컨테이너의 투명도는 지원하지 않습니다.');
      if(st.backgroundClip==='text')at('UNSUPPORTED_OBJECT','글자 윤곽으로 자르는 배경은 지원하지 않습니다.');
      for(const pseudo of ['::before','::after']){
        const ps=getComputedStyle(el,pseudo);if(['none','normal','""'].includes(ps.content))continue;
        let glyph='';try{glyph=JSON.parse(ps.content);}catch{}
        const supported=!declared&&!el.closest('table')&&typeof glyph==='string'&&/^[→←↑↓❯❮]$/u.test(glyph)&&ps.position==='absolute'&&ps.display==='block'&&ps.transform==='none'&&ps.filter==='none'&&ps.clipPath==='none'&&ps.mixBlendMode==='normal'&&ps.textShadow==='none'&&Number(ps.opacity)===1&&!visibleColor(ps.backgroundColor)&&ps.backgroundImage==='none'&&['Top','Right','Bottom','Left'].every(side=>parseFloat(ps[`padding${side}`])===0&&parseFloat(ps[`border${side}Width`])===0&&parseFloat(ps[`margin${side}`])===0);
        if(!supported){at('UNSUPPORTED_OBJECT',`CSS ${pseudo}의 생성 콘텐츠는 명시적인 HTML 요소로 넣어 주세요.`);continue;}
        const probe=document.createElement('span');probe.textContent=glyph;probe.style.cssText='all:initial!important;position:absolute!important;display:block!important;pointer-events:none!important';
        for(const prop of ['top','left','width','height','font-family','font-size','font-weight','font-style','font-stretch','line-height','letter-spacing','word-spacing','white-space','text-align','color'])probe.style.setProperty(prop,ps.getPropertyValue(prop),'important');
        el.append(probe);const rect=probe.getBoundingClientRect(),box={x:round((rect.x-rr.x)/renderScale),y:round((rect.y-rr.y)/renderScale),w:round(rect.width/renderScale),h:round(rect.height/renderScale)},baseline=directBaseline(glyph,ps,rect);probe.remove();
        if(box.w<=0||box.h<=0){at('UNSUPPORTED_OBJECT',`CSS ${pseudo}의 표시 위치를 확인할 수 없습니다.`);continue;}
        const pseudoId=`${nodeId}::generated-${pseudo.slice(2)}`,fontStyle={fontFamily:familyName(ps.fontFamily),fontSize:parseFloat(ps.fontSize),fontWeight:Number(ps.fontWeight)||400,fontStyle:ps.fontStyle,color:ps.color,textDecoration:ps.textDecorationLine,letterSpacing:ps.letterSpacing};
        if(seen.has(pseudoId)){at('INPUT_INVALID','중복 객체 ID입니다.');continue;}seen.add(pseudoId);if(!await fontAvailable(fontStyle.fontFamily))at('FONT_UNAVAILABLE',`글꼴을 사용할 수 없습니다: ${fontStyle.fontFamily}`);
        const pseudoStyle={...fontStyle,align:'left',lineHeight:ps.lineHeight,whiteSpace:ps.whiteSpace,opacity:1,padding:{top:0,right:0,bottom:0,left:0},cssZIndex:0};
        const afterOrder=Math.max(domOrder.get(el),...[...el.querySelectorAll('*')].map(child=>domOrder.get(child)??0));orderKeys.set(pseudoId,afterOrder+.5);
        nodes.push({id:pseudoId,kind:'text',boxCssPx:box,zOrder:order++,content:{text:glyph,runs:[{text:glyph,style:fontStyle}]},style:pseudoStyle,sourceRef,parentId:null});objects.push({nodeId:pseudoId,kind:'text',boxCssPx:box,lines:[{text:glyph,...box}],baseline,baselineMethod:'font-matched zero-size baseline probe',lineAdvance:null});
      }
      let kind=declared?el.dataset.kind:null;
      const chart=el.tagName==='FIGURE'?charts.find(c=>c.id===el.dataset.chartId):null;
      if(chart&&!declared)kind='chart';
      let box=boxOf(el);
      if(editable){
        if(!['text','shape','line','image','group','table','chart'].includes(kind)||['table','chart'].includes(kind)&&profile!=='editable-slides-v2'){at('UNSUPPORTED_OBJECT','지원하지 않는 편집 객체입니다.');continue;}
        if(!['text','table','chart'].includes(kind)&&(el.textContent.trim()||kind!=='line'&&el.children.length))at('UNSUPPORTED_OBJECT','이 객체 안에 추가된 콘텐츠를 편집 가능한 객체로 읽을 수 없습니다.');
        if(kind==='text'&&[...el.querySelectorAll('*')].some(c=>!['SPAN','BR','B','I','STRONG','EM','U'].includes(c.tagName)))at('UNSUPPORTED_OBJECT','글자 객체 안의 비텍스트 콘텐츠는 지원하지 않습니다.');
        const fields={x:'left',y:'top',w:'width',h:'height'},authored={};
        // CSSOM stringifies numeric values to fewer significant digits. Read
        // current inline px literals, and independently compare visible bounds.
        const literals=new Map([...el.getAttribute('style').matchAll(/(?:^|;)\s*(left|top|width|height)\s*:\s*(-?(?:\d+\.?\d*|\.\d+)px)\s*(?:!important\s*)?(?=;|$)/gi)].map(m=>[m[1].toLowerCase(),m[2]]));
        for(const [key,prop] of Object.entries(fields)){
          const value=literals.get(prop)??el.style.getPropertyValue(prop);
          if(!/^-?(?:\d+\.?\d*|\.\d+)px$/.test(value)){at('INPUT_INVALID','편집 객체의 논리 좌표는 px 값이어야 합니다.');continue;}
          authored[key]=parseFloat(value);
          if(Math.abs(authored[key]-box[key])>.05)at('UNSUPPORTED_OBJECT','현재 표시 배치와 객체의 논리 좌표가 다릅니다.');
        }
        if(st.position!=='absolute'||el.parentElement!==root)at('UNSUPPORTED_OBJECT','편집 객체는 장표에 직접 배치되어야 합니다.');
        box=authored;
      }
      if((box.w<=0||box.h<=0)&&!(editable&&kind==='line'))continue;
      const isSvg=el.tagName.toLowerCase()==='svg';
      const children=[...el.children].filter(c=>!['BR','STYLE'].includes(c.tagName));
      const containsBlock=children.some(c=>!['inline','inline-block','contents'].includes(getComputedStyle(c).display));
      const paragraphLines=!declared&&el.tagName==='P'&&children.length>1&&children.every(c=>c.tagName==='SPAN'&&getComputedStyle(c).display==='block'&&getComputedStyle(c).lineHeight===st.lineHeight&&getComputedStyle(c).fontSize===st.fontSize&&!visibleColor(getComputedStyle(c).backgroundColor)&&getComputedStyle(c).position==='static'&&!c.querySelector('br'))&&![...el.childNodes].some(n=>n.nodeType===3&&n.textContent.trim());
      const text=declared?el.textContent:el.querySelector('svg')?textRuns(el).map(run=>run.text).join(''):el.innerText;
      const hasText=typeof text==='string'&&text.trim().length>0;
      const inlineParent=[...textOwners].some(owner=>owner.contains(el));
      const direct=[];
      if(hasText&&containsBlock&&!inlineParent&&[...el.childNodes].some(n=>n.nodeType===3&&n.textContent.trim())){
        let visible;
        for(let clip=el;clip;clip=clip.parentElement){
          const cs=getComputedStyle(clip),axes=[cs.overflowX,cs.overflowY],restricted=axes.map(v=>['hidden','clip','scroll','auto'].includes(v));
          if(restricted.some(Boolean)){
            visible??=[el,...el.querySelectorAll('*')].filter(e=>{const s=getComputedStyle(e);return s.display!=='none'&&s.visibility==='visible'&&!hiddenByOpacity(e);});
            const bounds=clip.getBoundingClientRect(),limits={left:bounds.left+parseFloat(cs.borderLeftWidth)*renderScale,right:bounds.right-parseFloat(cs.borderRightWidth)*renderScale,top:bounds.top+parseFloat(cs.borderTopWidth)*renderScale,bottom:bounds.bottom-parseFloat(cs.borderBottomWidth)*renderScale};
            const outside=r=>restricted[0]&&(r.left<limits.left-.1||r.right>limits.right+.1)||restricted[1]&&(r.top<limits.top-.1||r.bottom>limits.bottom+.1);
            if(visible.some(e=>e!==clip&&outside(e.getBoundingClientRect()))||visible.some(e=>[...e.childNodes].some(n=>{if(n.nodeType!==3||!n.data.trim())return false;const r=document.createRange();r.selectNode(n);return [...r.getClientRects()].some(outside);})))at('UNSUPPORTED_OBJECT','잘린 직접 글자 또는 하위 콘텐츠는 지원하지 않습니다.');
          }
          if(clip===root)break;
        }
        // Adapted from dom-to-pptx's measured text-node handling (MIT):
        // keep unsupported whitespace and multiple lines as explicit failures.
        for(const [part,child] of [...el.childNodes].entries()){
          if(child.nodeType!==3||!child.data.trim())continue;
          if(editable||declared||!['normal','nowrap'].includes(st.whiteSpace)||st.textTransform!=='none'||st.direction!=='ltr'||st.writingMode!=='horizontal-tb'){
            at('UNSUPPORTED_OBJECT','블록과 섞인 직접 글자의 공백·방향·변환은 지원하지 않습니다.');continue;
          }
          const first=child.data.search(/\S/u),last=child.data.search(/\s*$/u),range=document.createRange();
          range.setStart(child,first);range.setEnd(child,last);
          const rects=[...range.getClientRects()].filter(r=>r.width>0&&r.height>0),r=range.getBoundingClientRect();
          if(rects.length!==1||!r.width||!r.height){at('UNSUPPORTED_OBJECT','블록과 섞인 여러 줄 직접 글자는 지원하지 않습니다.');continue;}
          const value=child.data.slice(first,last).replace(/[\t\r\n\f ]+/g,' '),fontStyle={fontFamily:familyName(st.fontFamily),fontSize:parseFloat(st.fontSize),fontWeight:Number(st.fontWeight)||400,fontStyle:st.fontStyle,color:st.color,textDecoration:st.textDecorationLine,letterSpacing:st.letterSpacing};
          if(parseFloat(st.webkitTextStrokeWidth)>0)fontStyle.outline={color:st.webkitTextStrokeColor,width:parseFloat(st.webkitTextStrokeWidth)};
          if(!await fontAvailable(fontStyle.fontFamily))at('FONT_UNAVAILABLE',`글꼴을 사용할 수 없습니다: ${fontStyle.fontFamily}`);
          const box={x:round((r.x-rr.x)/renderScale),y:round((r.y-rr.y)/renderScale),w:round(r.width/renderScale),h:round(r.height/renderScale)},id=`${nodeId}::direct-text-${part}`;
          const stacking=[];for(let parent=el;parent&&parent!==root;parent=parent.parentElement){const ps=getComputedStyle(parent),container=parent.parentElement&&getComputedStyle(parent.parentElement);if(ps.zIndex!=='auto'&&(ps.position!=='static'||/^(inline-)?(grid|flex)$/.test(container?.display)))stacking.push(Number(ps.zIndex));}
          if(stacking.length>1)at('UNSUPPORTED_OBJECT','중첩된 CSS 겹침 문맥은 지원하지 않습니다.');
          const style={...plainStyle(st),...fontStyle,align:'left',valign:'top',padding:{top:0,right:0,bottom:0,left:0},cssZIndex:stacking.at(-1)??0};
          const baseline=directBaseline(value,st,r);
          direct.push({node:{id,kind:'text',boxCssPx:box,zOrder:0,content:{text:value,runs:[{text:value,style:fontStyle}]},style,sourceRef,parentId:null},object:{nodeId:id,kind:'text',boxCssPx:box,lines:[{text:value,...box}],baseline,baselineMethod:'font-matched zero-size baseline probe',lineAdvance:null},domOrder:domOrder.get(child)});
        }
      }
      const addDirect=()=>{for(const item of direct){if(seen.has(item.node.id)){at('INPUT_INVALID','중복 객체 ID입니다.');continue;}seen.add(item.node.id);item.node.zOrder=order++;orderKeys.set(item.node.id,item.domOrder);nodes.push(item.node);objects.push(item.object);}};
      if(!kind){
        if(el.tagName==='TABLE')kind='table';
        else if(el.tagName==='IMG'||isSvg)kind='image';
        else if(el.tagName==='HR')kind='line';
        else if(hasText&&(!containsBlock||paragraphLines)&&!inlineParent)kind='text';
        else if(el!==root&&!inlineParent&&(visibleColor(st.backgroundColor)||parseFloat(st.borderTopWidth)>0||st.backgroundImage!=='none'))kind='shape';
        else {addDirect();continue;}
      }
      if(kind==='text'&&el.querySelector('svg'))at('UNSUPPORTED_OBJECT','HTML 글자와 SVG가 같은 줄에 섞인 콘텐츠는 별도 요소로 나누어 주세요.');
      if(seen.has(nodeId)){at('INPUT_INVALID','중복 객체 ID입니다.');continue;}seen.add(nodeId);
      const style=kind==='chart'?{fontFamily:familyName(st.fontFamily),fontSize:parseFloat(st.fontSize),fontWeight:400,fontStyle:'normal',color:'#223845',opacity:1}:plainStyle(st);
      if(editable&&kind!=='chart'){
        style.valign=st.alignContent==='center'?'mid':st.alignContent==='end'?'bottom':'top';
        if(kind==='shape'){style.svgGeometry=el.dataset.geometry==='centerline';style.shape=st.borderTopLeftRadius==='50%'?'ellipse':'rect';}
        if(kind==='line'){
          const line=el.querySelector('svg > line');
          if(!line||el.querySelectorAll('svg > *').length!==1){at('UNSUPPORTED_OBJECT','편집 선은 SVG line 한 개여야 합니다.');continue;}
          const ls=getComputedStyle(line);Object.assign(style,{svgGeometry:true,stroke:ls.stroke,strokeWidth:parseFloat(ls.strokeWidth),endpoints:{start:{x:box.x+line.x1.baseVal.value,y:box.y+line.y1.baseVal.value},end:{x:box.x+line.x2.baseVal.value,y:box.y+line.y2.baseVal.value}}});
          if(ls.strokeDasharray!=='none'||ls.markerStart!=='none'||ls.markerEnd!=='none')at('UNSUPPORTED_OBJECT','점선·화살표 선은 아직 지원하지 않습니다.');
        }
      }
      // A background-painted rule has rectangular filled ends. Retain that
      // paint primitive before authored JSON fields restore their semantic style.
      if(kind==='line'&&visibleColor(st.backgroundColor)&&st.backgroundImage==='none'&&['Top','Right','Bottom','Left'].every(side=>parseFloat(st[`border${side}Width`])===0)){
        style.linePaint={kind:'bar',color:st.backgroundColor,thickness:Math.min(box.w,box.h),horizontal:box.w>=box.h};
      }
      const stacking=[];for(let parent=el;parent&&parent!==root;parent=parent.parentElement){const ps=getComputedStyle(parent),container=parent.parentElement&&getComputedStyle(parent.parentElement);if(ps.zIndex!=='auto'&&(ps.position!=='static'||/^(inline-)?(grid|flex)$/.test(container?.display)))stacking.push(Number(ps.zIndex));if(ps.order!=='0')at('UNSUPPORTED_OBJECT','CSS order를 지정한 항목은 원본 DOM 순서로 정리해 주세요.');}
      if(stacking.length>1)at('UNSUPPORTED_OBJECT','중첩된 CSS 겹침 문맥은 지원하지 않습니다.');
      style.cssZIndex=stacking.at(-1)??0;
      let runs;
      if(kind==='text') {
        textOwners.add(el);runs=paragraphLines?children.flatMap((child,i)=>{const parts=textRuns(child);return i&&parts.length?[{text:'\n',style:parts[0].style},...parts]:parts;}):textRuns(el);
        for(const family of new Set([style.fontFamily,...runs.map(r=>r.style.fontFamily)]))if(!await fontAvailable(family))at('FONT_UNAVAILABLE',`글꼴을 사용할 수 없습니다: ${family}`);
        if(runs.map(r=>r.text).join('')!==text)at('UNSUPPORTED_OBJECT','화면 문구와 텍스트 조각의 공백·줄바꿈을 일치시킬 수 없습니다.');
        if(el!==root&&(style.backgroundImage!=='none'||visibleColor(style.fill)||style.strokeWidth>0)) {
          orderKeys.set(`${nodeId}::background`,domOrder.get(el));
          nodes.push({id:`${nodeId}::background`,kind:'shape',boxCssPx:box,zOrder:order++,content:{},style:{...style},sourceRef,parentId:null});
        }
      }
      let content=kind==='text'?{text,runs,...(paragraphLines?{tableCellGroup:true}:{})}:kind==='image'?{assetId:el.dataset.assetId??null}:{};
      if(kind==='image'&&el.tagName==='IMG'){
        const view=st.getPropertyValue('object-view-box');
        if(view&&view!=='none'){
          const match=/^inset\(\s*((?:\d+(?:\.\d+)?%\s*){1,4})\)$/i.exec(view),values=match?.[1].trim().split(/\s+/).map(v=>Math.round(parseFloat(v)*1000));
          if(!values||st.objectFit!=='fill')at('UNSUPPORTED_OBJECT','그림 자르기는 백분율 inset과 object-fit:fill만 지원합니다.');
          else{
            const [top,right=top,bottom=top,left=right]=values,crop={left,top,right,bottom};
            if(Object.values(crop).some(v=>v<0||v>=100000)||left+right>=100000||top+bottom>=100000)at('UNSUPPORTED_OBJECT','그림의 자르기 범위가 올바르지 않습니다.');
            else if(Object.values(crop).some(Boolean))content.crop=crop;
          }
        }
      }
      if(kind==='chart'){
        if(!chart){at('CHART_DATA_REQUIRED','차트 객체에 연결된 현재 JSON 데이터가 없습니다.');continue;}
        if(Number(st.opacity)!==1||st.boxShadow!=='none'||st.textShadow!=='none'||st.direction!=='ltr'||st.writingMode!=='horizontal-tb'||['Top','Right','Bottom','Left'].some(side=>parseFloat(st[`padding${side}`])!==0||parseFloat(st[`border${side}Width`])!==0)||el!==root&&(visibleColor(st.backgroundColor)||st.backgroundImage!=='none'))at('UNSUPPORTED_OBJECT','차트 영역의 배경·테두리·여백·투명도·효과는 별도 객체로 지정해 주세요.');
        if(!await fontAvailable(style.fontFamily))at('FONT_UNAVAILABLE',`차트 글꼴을 사용할 수 없습니다: ${style.fontFamily}`);
        content=chart.content;
      }
      if(kind==='table'){
        if(el.tagName!=='TABLE'){at('UNSUPPORTED_OBJECT','표 객체는 실제 HTML table이어야 합니다.');continue;}
        const tableIssues=[];
        content=await measureTable(el,box,(code,message)=>tableIssues.push({code,message}));if(!content){for(const d of tableIssues)at(d.code,d.message);continue;}
        const allowed=new Set(['셀 안 블록은 여백이나 별도 배치가 없는 한 줄 문구만 지원합니다.','셀 안의 별도 배경·여백·블록 배치는 아직 지원하지 않습니다.','현재 셀 문구와 텍스트 조각의 공백·줄바꿈을 일치시킬 수 없습니다.','표의 장식 화살표는 배경 그림에 보존합니다.']);
        if(allowTableOverlay&&tableIssues.length&&tableIssues.every(d=>d.code==='UNSUPPORTED_OBJECT'&&allowed.has(d.message))&&!declared){
          const pieces=[],tableCssZIndex=style.cssZIndex;let safe=true,pieceSequence=0;
          for(const cell of el.querySelectorAll('td,th')){
            const start=pieces.length;
            const cellRect=cell.getBoundingClientRect(),walker=document.createTreeWalker(cell,NodeFilter.SHOW_TEXT);let child;
            while((child=walker.nextNode())){
              if(!child.data.trim())continue;const parent=child.parentElement,cs=getComputedStyle(parent);if(cs.visibility!=='visible'||!parent.getClientRects().length||hiddenByOpacity(parent))continue;
              const range=document.createRange(),groups=[];let current=null;
              for(let i=0;i<child.data.length;i++){
                range.setStart(child,i);range.setEnd(child,i+1);const rect=[...range.getClientRects()].find(r=>r.width>0&&r.height>0);if(!rect)continue;
                const ch=child.data[i];if(!current||Math.abs(current.top-rect.top)>.5){current={text:'',left:rect.left,top:rect.top,right:rect.right,bottom:rect.bottom};groups.push(current);}
                current.text+=ch;current.right=Math.max(current.right,rect.right);current.bottom=Math.max(current.bottom,rect.bottom);
              }
              for(const group of groups){
                group.text=group.text.replace(/[\t\r\n\f ]+/g,' ');if(!group.text.trim())continue;
                if(group.left<cellRect.left-1||group.right>cellRect.right+1||group.top<cellRect.top-1||group.bottom>cellRect.bottom+1){safe=false;break;}
                const fontStyle={fontFamily:familyName(cs.fontFamily),fontSize:parseFloat(cs.fontSize),fontWeight:Number(cs.fontWeight)||400,fontStyle:cs.fontStyle,color:cs.color,textDecoration:cs.textDecorationLine,letterSpacing:cs.letterSpacing};
                if(parseFloat(cs.webkitTextStrokeWidth)>0){safe=false;break;}
                if(!await fontAvailable(fontStyle.fontFamily)){safe=false;break;}
                const rect={top:group.top},b={x:round((group.left-rr.x)/renderScale),y:round((group.top-rr.y)/renderScale),w:round((Math.min(cellRect.right,group.right+3*renderScale)-group.left)/renderScale),h:round((group.bottom-group.top)/renderScale)};
                const partId=nodeId+'::table-text-'+pieceSequence++,style={...plainStyle(cs),...fontStyle,align:'left',valign:'top',padding:{top:0,right:0,bottom:0,left:0},cssZIndex:tableCssZIndex};
                pieces.push({element:parent,node:{id:partId,kind:'text',boxCssPx:b,zOrder:0,content:{text:group.text,runs:[{text:group.text,style:fontStyle}]},style,sourceRef,parentId:null},object:{nodeId:partId,kind:'text',boxCssPx:b,lines:[{text:group.text,...b}],baseline:directBaseline(group.text,cs,rect),baselineMethod:'font-matched zero-size baseline probe',lineAdvance:null}});
              }
            }
            // Join aligned block lines in one cell, keeping each inline run's style.
            const fragments=pieces.slice(start),blocks=[...cell.children],base=getComputedStyle(cell);
            const grouped=blocks.length>1&&blocks.every(e=>{
              const cs=getComputedStyle(e),r=e.getBoundingClientRect(),range=document.createRange();range.selectNodeContents(e);
              const lines=new Set([...range.getClientRects()].filter(x=>x.width>0&&x.height>0).map(x=>Math.round(x.top*2)));
              return e.tagName==='SPAN'&&cs.display==='block'&&cs.lineHeight===base.lineHeight&&cs.fontSize===base.fontSize&&cs.textAlign===base.textAlign&&!visibleColor(cs.backgroundColor)&&['Top','Right','Bottom','Left'].every(side=>parseFloat(cs['margin'+side])===0&&parseFloat(cs['padding'+side])===0&&parseFloat(cs['border'+side+'Width'])===0)&&lines.size===1&&Math.abs(r.height/renderScale-parseFloat(cs.lineHeight))<.2;
            })&&fragments.length&&fragments.every(p=>p.node.style.fontSize===parseFloat(base.fontSize));
            if(grouped){
              const target=cell.innerText,chars=[];for(const piece of fragments)for(const ch of piece.node.content.text)if(!/\s/u.test(ch))chars.push({ch,style:piece.node.content.runs[0].style});
              if(!/\n\s*\n/.test(target)&&chars.map(x=>x.ch).join('')===target.replace(/\s/g,'')){
                let cursor=0;const runs=[];for(const ch of target){const nonspace=!/\s/u.test(ch),font=nonspace?chars[cursor++].style:(chars[Math.min(cursor,chars.length-1)]?.style??fragments[0].node.content.runs[0].style),key=JSON.stringify(font),last=runs.at(-1);if(last&&last.key===key)last.text+=ch;else runs.push({text:ch,style:font,key});}
                const first=fragments[0],last=fragments.at(-1),firstBox=first.node.boxCssPx,lastBox=last.node.boxCssPx;
                const left=Math.min(...fragments.map(x=>x.node.boxCssPx.x)),right=round((cellRect.right-rr.x-parseFloat(base.paddingRight)*renderScale)/renderScale);
                const lines=blocks.map(block=>{const r=block.getBoundingClientRect();return {text:block.innerText,x:round((r.x-rr.x)/renderScale),y:round((r.y-rr.y)/renderScale),w:round(r.width/renderScale),h:round(r.height/renderScale)};});
                const mergedBox={x:left,y:firstBox.y,w:Math.max(.01,round(right-left)),h:round(lastBox.y+lastBox.h-firstBox.y)};
                first.node.boxCssPx=mergedBox;first.node.style.lineHeight=base.lineHeight;first.node.style.fontWeight=Number(base.fontWeight)||400;first.node.content={text:target,runs:runs.map(({key,...r})=>r),tableCellGroup:true};
                first.object.boxCssPx=mergedBox;first.object.lines=lines;first.object.lineAdvance=round(parseFloat(base.lineHeight));pieces.splice(start,fragments.length,first);
              }
            }
            // A cell can contain distinct labeled items or a multiline heading.
            // Join each simple block, never the neighboring blocks by proximity alone.
            for(const block of [...blocks].reverse()){
              const bs=getComputedStyle(block);
              if(!['SPAN','DIV'].includes(block.tagName)||bs.display!=='block'||[...block.querySelectorAll('*')].some(e=>!['BR','B','STRONG','I','EM','U','SPAN'].includes(e.tagName)||!['inline','inline-block'].includes(getComputedStyle(e).display))||visibleColor(bs.backgroundColor))continue;
              const indexes=[];for(let i=start;i<pieces.length;i++)if(block.contains(pieces[i].element))indexes.push(i);
              if(indexes.length<2||indexes.some((x,i)=>i&&x!==indexes[i-1]+1))continue;
              const parts=indexes.map(i=>pieces[i]),target=block.innerText,chars=[];
              for(const part of parts)for(const ch of part.node.content.text)if(!/\s/u.test(ch))chars.push({ch,part,style:part.node.content.runs[0].style});
              if(!target.trim()||/\n\s*\n/.test(target)||chars.map(x=>x.ch).join('')!==target.replace(/\s/g,''))continue;
              let cursor=0,previous=null,output='',runs=[];
              const append=(ch,font)=>{output+=ch;const key=JSON.stringify(font),last=runs.at(-1);if(last&&last.key===key)last.text+=ch;else runs.push({text:ch,style:font,key});};
              for(const ch of target){const nonspace=!/\s/u.test(ch),current=chars[Math.min(cursor,chars.length-1)],font=current?.style??parts[0].node.content.runs[0].style;
                if(nonspace&&previous&&previous.part!==current.part&&!/\s/u.test(output.at(-1)??'')&&Math.abs(previous.part.node.boxCssPx.y-current.part.node.boxCssPx.y)<font.fontSize*.7){const gap=current.part.node.boxCssPx.x-(previous.part.node.boxCssPx.x+previous.part.node.boxCssPx.w);if(gap>2){const spaces=Math.min(10,Math.round(gap/(font.fontSize*.27)));for(let j=0;j<spaces;j++)append(' ',font);}}
                append(ch,font);if(nonspace){previous=current;cursor++;}}
              const first=parts[0],left=Math.min(...parts.map(p=>p.node.boxCssPx.x)),top=Math.min(...parts.map(p=>p.node.boxCssPx.y)),bottom=Math.max(...parts.map(p=>p.node.boxCssPx.y+p.node.boxCssPx.h)),r=block.getBoundingClientRect(),right=Math.min(cellRect.right,r.right),box={x:left,y:top,w:Math.max(.01,round((right-rr.x)/renderScale-left)),h:round(bottom-top)};
              first.node.boxCssPx=box;first.node.style.lineHeight=bs.lineHeight;first.node.style.fontWeight=Number(bs.fontWeight)||400;first.node.content={text:output,runs:runs.map(({key,...r})=>r),tableCellGroup:true};first.object.boxCssPx=box;first.object.lines=parts.map(p=>p.object.lines[0]);first.object.lineAdvance=parseFloat(bs.lineHeight)||null;pieces.splice(indexes[0],indexes.length,first);
            }
          }
          const expected=el.innerText.replace(/\s+/g,''),actual=pieces.map(p=>p.node.content.text).join('').replace(/\s+/g,'');
          if(expected!==actual||!pieces.length)safe=false;
          if(safe){
            kind='image';content={assetId:'pending-table-background',tableBackground:true};
            diagnostics.push({code:'TABLE_LAYOUT_APPROXIMATED',severity:'warning',slideId:id,nodeId,sourceRef,message:'복잡한 표의 배경은 그림으로, 문구는 별도 편집 가능한 글자 상자로 변환합니다. 표 전체 이동 시 글자도 함께 옮겨 주세요.'});
            orderKeys.set(nodeId,domOrder.get(el));nodes.push({id:nodeId,kind,boxCssPx:box,zOrder:order++,content,style,sourceRef,parentId:null});objects.push({nodeId,kind,boxCssPx:box});
            for(const piece of pieces){seen.add(piece.node.id);piece.node.zOrder=order++;orderKeys.set(piece.node.id,domOrder.get(el));nodes.push(piece.node);objects.push(piece.object);}
            continue;
          }
        }
        for(const d of tableIssues)at(d.code,d.message);
      }
      if(isSvg)content.inlineSvg=new XMLSerializer().serializeToString(el);
      orderKeys.set(nodeId,domOrder.get(el));
      nodes.push({id:nodeId,kind,boxCssPx:box,zOrder:order++,content,style,sourceRef,parentId:editable?(el.dataset.parentId??null):null});
      objects.push({nodeId,kind,boxCssPx:box,...(kind==='text'?lineMetrics(el):{})});
      addDirect();
    }
    const layoutStable=JSON.stringify(initialLayout)===JSON.stringify(elements.map(boxOf));
    if(profile!=='project-data-v1'&&!nodes.length&&sourceHadScripts)issue('SCRIPT_REQUIRED','스크립트 없이 표시되는 장표 콘텐츠가 없습니다. 렌더된 정적 HTML을 제공해 주세요.',id);
    if(!layoutStable)issue('RENDER_INCOMPLETE','측정 중 레이아웃 경계가 바뀌었습니다.',id);
    nodes.sort((a,b)=>a.style.cssZIndex-b.style.cssZIndex||(orderKeys.get(a.id)??0)-(orderKeys.get(b.id)??0)||a.zOrder-b.zOrder);nodes.forEach((n,i)=>{n.zOrder=i;});
    slides.push({id,widthCssPx:width,heightCssPx:height,background:{color:rootStyle.backgroundColor,image:rootStyle.backgroundImage},nodes});
    measurements.push({slideId:id,referenceSize,renderScale,sourceTransformScale,layoutStable,objects,assetsReady:diagnostics.every(d=>d.code!=='RENDER_INCOMPLETE'),fontsReady:diagnostics.every(d=>d.code!=='FONT_UNAVAILABLE')});
  }
  const layoutSnapshot=[...document.querySelectorAll(`${selector},${selector} *`)].map(e=>{const r=e.getBoundingClientRect();return [r.x,r.y,r.width,r.height];});
  return {slides,measurements,diagnostics,layoutSnapshot,fonts:[...fontAvailability].map(([family,available])=>({family,available,custom:customFonts.has(family)}))};
}

export async function measureInput(input,runtime) {
  const result={model:null,measurements:[],environment:{...runtime.environment},diagnostics:[...input.diagnostics],previews:[]};
  let fallbackFonts=[];
  const attachFonts=async()=>{
    const normalized=await normalizeWoff2Fonts(input);
    result.model.assets=[...result.model.assets,...normalized.assets.filter(a=>!result.model.assets.some(existing=>existing.id===a.id))];
    const fonts=collectFontAssets(normalized,result.model);result.diagnostics.push(...fonts.diagnostics);
    if(fonts.faces.length){result.model.version=2;result.model.fonts=fonts.faces;result.environment.fontUsages=fonts.usages;}
    if(fallbackFonts.length){
      const used=new Set(),usages=[];
      for(const slide of result.model.slides){let substituted=false;for(const node of slide.nodes){
        const styles=node.kind==='table'?node.content.cells.flatMap(c=>[c.style,...c.runs.map(r=>r.style)]):node.kind==='text'?[node.style,...node.content.runs.map(r=>r.style)]:node.kind==='chart'?[node.style]:[];
        for(const style of styles)if(style?.fontFamily==='PPTAgainFallback'){const weight=Number(style.fontWeight)>=600?700:400;used.add(weight);substituted=true;usages.push({fontId:`fallback-pretendard-${weight}`,slideId:slide.id,nodeId:node.id,text:''});}
      }if(substituted)result.diagnostics.push(diagnostic('FONT_SUBSTITUTED','원래 글꼴 대신 Pretendard를 사용합니다.',{slideId:slide.id},'warning'));}
      if(used.size){result.model.version=2;result.model.fonts??=[];result.environment.fontUsages??=[];for(const f of fallbackFonts.filter(f=>used.has(f.weight))){if(!result.model.assets.some(a=>a.id===f.id))result.model.assets.push({id:f.id,sha256:f.sha256,mime:f.mime,bytes:f.bytes});result.model.fonts.push({id:`fallback-pretendard-${f.weight}`,cssFamily:'PPTAgainFallback',family:'Pretendard',weight:f.weight,style:'normal',assetId:f.id,origin:'supplied'});}result.environment.fontUsages.push(...usages);}
    }
    for(const m of result.measurements)if(fonts.diagnostics.some(d=>d.severity==='error'&&(!d.slideId||d.slideId===m.slideId)))m.fontsReady=false;
  };
  if(result.diagnostics.some(d=>d.severity==='error'))return result;
  if(input.embeddedSlides){
    const slides=[],assets=new Map(),fonts=new Map(),mapped=new Map(),available=new Map(),usages=[];let version=1,assetBytes=0;
    let batch=[];
    for(const [index,child] of input.embeddedSlides.entries()){
      if(index%2===0){const settled=await Promise.allSettled(input.embeddedSlides.slice(index,index+2).map(page=>measureInput(page,runtime)));const failed=settled.find(item=>item.status==='rejected');if(failed)throw failed.reason;batch=settled.map(item=>item.value);}
      const measured=batch[index%2],ids=new Map();
      if(measured.model&&measured.model.slides.length!==1)result.diagnostics.push(diagnostic('SLIDE_ROOT_REQUIRED','내장 데이터의 각 페이지는 정확히 한 장표여야 합니다.',{slideId:`slide-${index+1}`}));
      for(const slide of measured.model?.slides??[]){
        const id=`slide-${slides.length+1}`;if(slides.length&&(Math.abs(slides[0].widthCssPx-slide.widthCssPx)>.01||Math.abs(slides[0].heightCssPx-slide.heightCssPx)>.01))result.diagnostics.push(diagnostic('UNSUPPORTED_OBJECT','한 PPT 안의 서로 다른 장표 크기는 지원하지 않습니다.',{slideId:id}));
        ids.set(slide.id,id);slide.id=id;for(const node of slide.nodes)node.sourceRef.slideId=id;slides.push(slide);
      }
      result.diagnostics.push(...measured.diagnostics.map(d=>({...d,slideId:ids.get(d.slideId)??`slide-${index+1}`})));
      result.measurements.push(...measured.measurements.map(m=>({...m,slideId:ids.get(m.slideId)??m.slideId})));
      result.previews.push(...measured.previews.map(p=>({...p,slideId:ids.get(p.slideId)??p.slideId})));
      for(const asset of measured.model?.assets??[])if(!assets.has(asset.id)){
        if((assetBytes+=asset.bytes.length)>64*1024*1024){result.diagnostics.push(diagnostic('LIMIT_EXCEEDED','내장 장표 자산의 합계 크기 한도를 초과했습니다.'));break;}
        assets.set(asset.id,asset);
      }
      if(assetBytes>64*1024*1024)break;
      for(const font of measured.model?.fonts??[]){
        const other=[...fonts.values()].find(f=>f.cssFamily===font.cssFamily&&f.weight===font.weight&&f.style===font.style&&f.assetId!==font.assetId);
        if(other)result.diagnostics.push(diagnostic('FONT_STYLE_MISMATCH','장표마다 같은 글꼴 이름에 서로 다른 파일이 연결되어 있습니다.'));
        fonts.set(font.id,font);
      }
      for(const mapping of measured.environment.fontMappings??[])for(const slideId of ids.values()){
        const scoped={...mapping,slideId};mapped.set(JSON.stringify(scoped),scoped);
      }
      for(const font of measured.environment.fonts??[])available.set(font.family,font);
      usages.push(...(measured.environment.fontUsages??[]).map(u=>({...u,slideId:ids.get(u.slideId)??u.slideId})));
      version=Math.max(version,measured.model?.version??1);
    }
    result.model={version,sourceDigest:input.source.sha256,slides,assets:[...assets.values()],...(fonts.size?{fonts:[...fonts.values()]}:{})};
    Object.assign(result.environment,{inputProfile:'embedded-static-documents-v1',sourceScriptsExecuted:false,fontMappings:[...mapped.values()],fonts:[...available.values()],...(usages.length?{fontUsages:usages}:{})});
    result.diagnostics.push(...scopeEmbeddedModelErrors(validateModel(result.model).errors,slides));
    return result;
  }
  let session;
  try {
    session=await runtime.render(input);
    if(input.profile==='static-slides-v1')fallbackFonts=await applyFallbackFont(session.page,installedNativeAliases());
    if(input.profile==='svg-slides-v1'){await measureSvgInput(input,runtime,session,result);await attachFonts();return result;}
    let selector=input.profile==='project-data-v1'?'[data-slide]':input.source.slideSelector??'[data-slide]';
    // Auto-detect only an unambiguous visible fixed-size container; never guess among competing containers.
    if(input.profile==='static-slides-v1'&&!input.source.slideSelector){
      const detected=await session.page.evaluate(()=>{
        if(document.querySelector('[data-slide]'))return null;
        const candidates=[...document.querySelectorAll('section,article,main,div')].filter(el=>{
          const style=getComputedStyle(el),width=el.computedStyleMap().get('width')?.toString(),height=el.computedStyleMap().get('height')?.toString();
          return el.getClientRects().length>0&&style.display!=='none'&&style.visibility==='visible'&&/^\d+(?:\.\d+)?px$/.test(width??'')&&/^\d+(?:\.\d+)?px$/.test(height??'')&&parseFloat(width)>=300&&parseFloat(height)>=200;
        });
        if(candidates.length!==1)return null;
        let el=candidates[0],parts=[];while(el&&el!==document.body){const siblings=[...el.parentElement.children].filter(child=>child.tagName===el.tagName);parts.unshift(el.tagName.toLowerCase()+':nth-of-type('+(siblings.indexOf(el)+1)+')');el=el.parentElement;}
        return el?'body > '+parts.join(' > '):null;
      });
      if(detected)selector=detected;
    }
    const dimensions=selector=>{try{return [...document.querySelectorAll(selector)].map(r=>{const st=getComputedStyle(r);return [parseFloat(st.width),parseFloat(st.height)];});}catch{return [];}};
    const initialDimensions=await session.page.evaluate(dimensions,selector);
    await session.page.setViewportSize({width:1151,height:877});
    const alternateDimensions=await session.page.evaluate(dimensions,selector);
    await session.page.setViewportSize({width:1280,height:960});
    if(initialDimensions.some((size,i)=>size.some((value,j)=>Math.abs(value-(alternateDimensions[i]?.[j]??NaN))>.01))){
      result.diagnostics.push(diagnostic('SLIDE_ROOT_REQUIRED','화면 크기에 따라 장표 크기가 달라집니다. 고정 너비와 높이를 지정해 주세요.'));return result;
    }
    const size=await session.page.evaluate(selector=>{try{const r=document.querySelector(selector);if(!r)return null;const st=getComputedStyle(r),w=parseFloat(st.width),h=parseFloat(st.height);return w>0&&h>0?{width:Math.round(w*1280/Math.max(w,h)),height:Math.round(h*1280/Math.max(w,h))}:null;}catch{return null;}},selector);
    if(size&&size.width>0&&size.height>0)await session.page.setViewportSize(size);
    result.environment.viewport=session.page.viewportSize();
    if(input.charts?.length){
      const frames=await session.page.evaluate(({selector,ids})=>{
        let roots;try{roots=[...document.querySelectorAll(selector)];}catch{return [];}
        return [...document.querySelectorAll('figure[data-chart-id]')].filter(el=>ids.includes(el.dataset.chartId)).map(el=>{
          const st=getComputedStyle(el),width=parseFloat(st.width),height=parseFloat(st.height);
          return {id:el.dataset.chartId,width,height,fontFamily:st.fontFamily.split(',')[0].trim().replace(/^["']|["']$/g,''),fontSize:parseFloat(st.fontSize),fixed:el.computedStyleMap?.().get('width')?.toString()!=='auto'&&el.computedStyleMap?.().get('height')?.toString()!=='auto',roots:roots.filter(root=>root===el||root.contains(el)).length};
        });
      },{selector,ids:input.charts.map(c=>c.id)});
      const displays=[];
      for(const chart of input.charts){
        const frame=frames.find(f=>f.id===chart.id);
        if(!frame||frame.roots!==1||!frame.fixed||![frame.width,frame.height].every(v=>Number.isFinite(v)&&v>0))result.diagnostics.push(diagnostic('INPUT_INVALID','차트는 한 장표 안에서 너비와 높이를 지정해 주세요.',chart));
        else displays.push({...frame,svg:renderChartSvg(chart.content,frame)});
      }
      if(result.diagnostics.some(d=>d.severity==='error'))return result;
      await session.page.evaluate(displays=>{
        for(const display of displays){
          const figure=[...document.querySelectorAll('figure[data-chart-id]')].find(el=>el.dataset.chartId===display.id),host=document.createElement('div');
          // Author styles cannot change the generated marks. The shadow tree
          // still shares the document's FontFaceSet; no font bytes are copied.
          const shadow=host.attachShadow({mode:'closed'});
          shadow.innerHTML=`<style>:host{all:initial!important;display:block!important;width:${display.width}px!important;height:${display.height}px!important;direction:ltr!important;unicode-bidi:normal!important}:host::before,:host::after{content:none!important}svg{display:block;overflow:hidden}</style>${display.svg}`;
          figure.replaceChildren(host);
        }
      },displays);
    }
    let timer;
    const measured=await Promise.race([
      session.page.evaluate(measurePage,{selector,profile:input.profile,timeoutMs:runtime.timeoutMs,sourceHadScripts:session.sourceHadScripts,allowTableOverlay:input.allowTableOverlay===true,nativeAliases:installedNativeAliases(),charts:input.charts??[]}),
      new Promise((_,reject)=>{timer=setTimeout(()=>reject(new Error('측정 제한 시간 초과')),Math.max(15000,runtime.timeoutMs));})
    ]).finally(()=>clearTimeout(timer));
    if(measured.layoutSnapshot){
      await new Promise(resolve=>setTimeout(resolve,60));
      const after=await session.page.evaluate(selector=>[...document.querySelectorAll(`${selector},${selector} *`)].map(e=>{const r=e.getBoundingClientRect();return [r.x,r.y,r.width,r.height];}),selector);
      if(JSON.stringify(after)!==JSON.stringify(measured.layoutSnapshot)){measured.diagnostics.push(diagnostic('RENDER_INCOMPLETE','시간 간격을 둔 검사에서 배치가 바뀌었습니다.'));for(const m of measured.measurements)m.layoutStable=false;}
    }
    const slots=measured.slides.flatMap(slide=>slide.nodes.filter(n=>n.kind==='image'&&n.content.inlineSvg&&n.sourceRef.path).map(n=>({slideId:slide.id,path:n.sourceRef.path,renderScale:measured.measurements.find(m=>m.slideId===slide.id)?.renderScale})));
    if(slots.length){
      const svg=await measureSvgFragments(session,{selector,timeoutMs:runtime.timeoutMs,slots});
      measured.diagnostics.push(...svg.diagnostics);
      for(const font of svg.fonts){
        const prior=measured.fonts.find(f=>f.family===font.family);
        if(prior){prior.available=prior.available&&font.available;prior.custom=prior.custom||font.custom;}else measured.fonts.push(font);
        if(!font.available)measured.diagnostics.push(diagnostic('FONT_UNAVAILABLE',`글꼴을 사용할 수 없습니다: ${font.family}`));
      }
      for(const slide of measured.slides){
        const measurement=measured.measurements.find(m=>m.slideId===slide.id),used=new Set(slide.nodes.map(n=>n.id)),merged=[];
        for(const slot of slide.nodes){
          const part=svg.fragments.find(f=>f.slideId===slide.id&&f.svgPath===slot.sourceRef.path);
          if(!part){merged.push(slot);continue;}
          const ids=new Map();
          for(const node of part.nodes){let id=`${slot.id}::${node.id}`,i=1;while(used.has(id))id=`${slot.id}::${node.id}::${++i}`;used.add(id);ids.set(node.id,id);}
          for(const node of [...part.nodes].sort((a,b)=>a.zOrder-b.zOrder))merged.push({...node,id:ids.get(node.id),parentId:node.parentId?ids.get(node.parentId):null});
          measurement.objects=measurement.objects.filter(o=>o.nodeId!==slot.id);
          measurement.objects.push(...part.objects.map(o=>({...o,nodeId:ids.get(o.nodeId)})));
        }
        slide.nodes=merged;slide.nodes.forEach((node,i)=>{node.zOrder=i;});
        if(svg.fonts.some(font=>!font.available))measurement.fontsReady=false;
      }
    }
    result.diagnostics.push(...measured.diagnostics);result.measurements=measured.measurements;result.environment.fonts=measured.fonts??[];
    const mappings=identifyFontMappings(input);
    const availableNative=await session.page.evaluate(async names=>{const available=[];for(const name of names){try{await new FontFace('__native_probe',`local("${name}")`).load();available.push(name);}catch{}}return available;},[...new Set(mappings.map(m=>m.nativeLocalName))]);
    result.environment.fontMappings=mappings.filter(m=>availableNative.includes(m.nativeLocalName));
    result.model={version:measured.slides.some(s=>s.nodes.some(n=>['table','chart'].includes(n.kind)))?2:1,sourceDigest:input.source.sha256,slides:measured.slides,assets:[...input.assets]};
    if(['editable-slides-v1','editable-slides-v2'].includes(input.profile))result.model.pptxSizeEmu=input.pptxSizeEmu;
    for(const slide of result.model.slides) {
      const source=input.project?.slides.find(s=>s.id===slide.id);
      for(const n of slide.nodes) {
        const original=source?.nodes.find(o=>o.id===n.id);
        if(original) {
          n.sourceRef={slideId:slide.id,elementId:n.id};
          n.boxCssPx={x:original.x,y:original.y,w:original.w,h:original.h};
          // Preserve source semantics; browser measurements remain independent evidence.
          Object.assign(n.style,{dash:original.dash??false,stroke:original.stroke??n.style.stroke,strokeWidth:original.strokeWidth??n.style.strokeWidth,fill:original.fill??n.style.fill,radius:original.radius??n.style.radius});
          if(n.style.linePaint?.kind==='bar')Object.assign(n.style.linePaint,{thickness:Math.max(1,original.strokeWidth??1),horizontal:original.w>original.h});
          if(n.kind==='image')n.content.assetId=input.assetRefs.find(r=>r.slideId===slide.id&&r.nodeId===n.id)?.assetId??null;
        }
        if(n.content.inlineSvg) {
          try{const a=resolveAsset(`data:image/svg+xml;base64,${Buffer.from(n.content.inlineSvg).toString('base64')}`);if(!result.model.assets.some(x=>x.id===a.id))result.model.assets.push(a);n.content={assetId:a.id,...(n.content.artworkOnly?{artworkOnly:true}:{})};}
          catch(e){result.diagnostics.push(diagnostic('UNSUPPORTED_OBJECT',e.message,{slideId:slide.id,nodeId:n.id,sourceRef:n.sourceRef}));}
        }
        if(n.content.inlinePng){try{const a=resolveAsset(n.content.inlinePng);if(!result.model.assets.some(x=>x.id===a.id))result.model.assets.push(a);n.content={assetId:a.id,artworkOnly:true};}catch(e){result.diagnostics.push(diagnostic('UNSUPPORTED_OBJECT',e.message,{slideId:slide.id,nodeId:n.id,sourceRef:n.sourceRef}));}}
      }
      if(source)for(const groupId of new Set(source.nodes.map(n=>n.group).filter(Boolean))) {
        const children=slide.nodes.filter(n=>source.nodes.find(o=>o.id===n.id)?.group===groupId);if(!children.length)continue;
        const x=Math.min(...children.map(n=>n.boxCssPx.x)),y=Math.min(...children.map(n=>n.boxCssPx.y));
        const id=`group-${groupId}`;for(const n of children)n.parentId=id;
        slide.nodes.push({id,kind:'group',boxCssPx:{x,y,w:Math.max(...children.map(n=>n.boxCssPx.x+n.boxCssPx.w))-x,h:Math.max(...children.map(n=>n.boxCssPx.y+n.boxCssPx.h))-y},zOrder:Math.min(...children.map(n=>n.zOrder)),content:{},style:{},sourceRef:{slideId:slide.id,elementId:groupId},parentId:null});
      }
    }
    await attachFonts();
    if(!result.diagnostics.some(d=>d.severity==='error')) {
      for(const slide of result.model.slides)for(const node of slide.nodes.filter(n=>n.kind==='image'&&n.content.tableBackground)){
        const target=session.page.locator('[data-source-path='+JSON.stringify(node.sourceRef.path)+']');
        const originalRootStyle=await session.page.evaluate(({selector,path})=>{const el=[...document.querySelectorAll('[data-source-path]')].find(e=>e.dataset.sourcePath===path),root=[...document.querySelectorAll(selector)].find(r=>r.contains(el));if(!root)throw new Error('Table slide not found');const old=root.getAttribute('style'),marks=[],rules=[];for(const cell of el.querySelectorAll('td,th'))for(const pseudo of ['::before','::after']){const ps=getComputedStyle(cell,pseudo);let glyph='';try{glyph=JSON.parse(ps.content);}catch{}if(!/^[→←↑↓❯❮]$/u.test(glyph))continue;let mark=marks.find(m=>m.el===cell);if(!mark){mark={el:cell,old:cell.getAttribute('data-pptagain-pseudo'),id:String(marks.length)};marks.push(mark);cell.setAttribute('data-pptagain-pseudo',mark.id);}rules.push(`[data-pptagain-mask] [data-pptagain-pseudo="${mark.id}"]${pseudo}{color:${ps.color}!important;-webkit-text-fill-color:${ps.color}!important}`);}root.style.setProperty('position','fixed','important');root.style.setProperty('left','0','important');root.style.setProperty('top','0','important');root.style.setProperty('margin','0','important');el.setAttribute('data-pptagain-mask','');const style=document.createElement('style');style.setAttribute('data-pptagain-mask-style','');style.textContent='[data-pptagain-mask], [data-pptagain-mask] * { color: transparent !important; -webkit-text-fill-color: transparent !important; text-shadow: none !important; }'+rules.join('');document.head.append(style);return {old,marks:marks.map(m=>({path:m.el.dataset.sourcePath,old:m.old}))};},{selector,path:node.sourceRef.path});
        let bytes;try{const clip=await target.boundingBox();if(!clip||clip.x<0||clip.y<0||clip.width<=0||clip.height<=0)throw new Error('Table outside slide');bytes=await session.page.screenshot({clip,animations:'disabled',scale:'css',omitBackground:true});}finally{await session.page.evaluate(({selector,path,state})=>{const el=[...document.querySelectorAll('[data-source-path]')].find(e=>e.dataset.sourcePath===path),root=[...document.querySelectorAll(selector)].find(r=>r.contains(el));el.removeAttribute('data-pptagain-mask');for(const m of state.marks){const cell=[...el.querySelectorAll('[data-source-path]')].find(e=>e.dataset.sourcePath===m.path);if(cell){if(m.old===null)cell.removeAttribute('data-pptagain-pseudo');else cell.setAttribute('data-pptagain-pseudo',m.old);}}document.querySelector('[data-pptagain-mask-style]')?.remove();if(state.old===null)root.removeAttribute('style');else root.setAttribute('style',state.old);},{selector,path:node.sourceRef.path,state:originalRootStyle});}
        const asset=resolveAsset('data:image/png;base64,'+bytes.toString('base64'));if(!result.model.assets.some(a=>a.id===asset.id))result.model.assets.push(asset);node.content={assetId:asset.id};
      }
      await captureDecorations(session,result.model,result.measurements,selector);
      const validation=validateModel(result.model);
      for(const e of validation.errors)result.diagnostics.push(diagnostic('INPUT_INVALID',`${e.path}: ${e.message}`));
    }
    if(runtime.capture&&!result.diagnostics.some(d=>d.code==='SLIDE_ROOT_REQUIRED')) {
      for(const [index,m] of result.measurements.entries()) {
        await session.page.setViewportSize(m.referenceSize);
        const root=session.page.locator(selector).nth(index);
        const oldStyle=await root.evaluate(el=>{const old=el.getAttribute('style');el.style.setProperty('position','fixed','important');el.style.setProperty('left','0','important');el.style.setProperty('top','0','important');el.style.setProperty('margin','0','important');return old;});
        let bytes;try{bytes=await root.screenshot({animations:'disabled',scale:'css'});}finally{await root.evaluate((el,old)=>{if(old===null)el.removeAttribute('style');else el.setAttribute('style',old);},oldStyle);}
        if(bytes.readUInt32BE(16)!==m.referenceSize.width||bytes.readUInt32BE(20)!==m.referenceSize.height){result.diagnostics.push(diagnostic('RENDER_INCOMPLETE','캡처 크기가 선언한 기준 화면 크기와 다릅니다.',{slideId:m.slideId}));continue;}
        result.previews.push({slideId:m.slideId,bytes:new Uint8Array(bytes),...m.referenceSize});
      }
    }
    return result;
  } catch(e) {
    result.diagnostics.push(diagnostic('RENDER_INCOMPLETE',`브라우저 측정이 완료되지 않았습니다: ${e.message}`));return result;
  } finally {await session?.context.close();}
}
