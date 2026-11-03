const A='http://schemas.openxmlformats.org/drawingml/2006/main';
const P='http://schemas.openxmlformats.org/presentationml/2006/main';
const R='http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const elements=n=>[...(n?.childNodes??[])].filter(c=>c.nodeType===1);
const child=(n,name)=>elements(n).find(c=>c.namespaceURI===A&&c.localName===name);
const readInt=(node,name,required=false)=>{
  if(!node?.hasAttribute(name)&&!required)return 0;
  const value=node?.getAttribute(name)?.trim();
  if(!/^[+-]?\d+$/.test(value??'')||!Number.isSafeInteger(Number(value)))throw new Error('수식 표의 좌표 또는 이미지 채우기 값이 잘못되었습니다.');
  return Number(value);
};

export function readMathTransform(transform){
  const off=child(transform,'off'),ext=child(transform,'ext');
  if(!off||!ext)throw new Error('수식 그림의 좌표 또는 크기가 없습니다.');
  const box={x:readInt(off,'x',true),y:readInt(off,'y',true),w:readInt(ext,'cx',true),h:readInt(ext,'cy',true)};
  if(box.w<=0||box.h<=0)throw new Error('수식 그림의 크기가 잘못되었습니다.');
  const flip=name=>{
    const value=transform.hasAttribute(name)?transform.getAttribute(name).trim():'0';
    if(!['0','1','false','true'].includes(value))throw new Error('수식 그림의 뒤집기 값이 잘못되었습니다.');
    return ['1','true'].includes(value);
  };
  return {box,rotation:readInt(transform,'rot'),flipH:flip('flipH'),flipV:flip('flipV')};
}

// A stretched fill may extend past the shape or occupy only part of it.
// Intersect the stretched image with the rectangular shape and describe the
// visible portion as a normal image box plus percentages cropped from the PNG.
export function stretchImageFragment(box,fillRect){
  if(fillRect){
    const allowed=new Set(['l','t','r','b']);
    for(const attribute of [...fillRect.attributes])if(attribute.namespaceURI!=='http://www.w3.org/2000/xmlns/'&&(attribute.namespaceURI||!allowed.has(attribute.name)))throw new Error('수식 그림의 채우기 속성이 잘못되었습니다.');
    if(elements(fillRect).length)throw new Error('수식 그림의 채우기 구조가 잘못되었습니다.');
  }
  const raw=Object.fromEntries(['l','t','r','b'].map(k=>[k,readInt(fillRect,k)]));
  if(Object.values(raw).some(v=>Math.abs(v)>1000000))throw new Error('수식 그림의 채우기 값이 잘못되었습니다.');
  const imageLeft=box.w*raw.l/100000,imageTop=box.h*raw.t/100000;
  const imageRight=box.w*(1-raw.r/100000),imageBottom=box.h*(1-raw.b/100000);
  const imageWidth=imageRight-imageLeft,imageHeight=imageBottom-imageTop;
  if(imageWidth<=0||imageHeight<=0)throw new Error('수식 그림의 채우기 크기가 잘못되었습니다.');
  const left=Math.max(0,imageLeft),top=Math.max(0,imageTop),right=Math.min(box.w,imageRight),bottom=Math.min(box.h,imageBottom);
  if(right<=left||bottom<=top)return null;
  const crop={left:Math.round((left-imageLeft)/imageWidth*100000),top:Math.round((top-imageTop)/imageHeight*100000),right:Math.round((imageRight-right)/imageWidth*100000),bottom:Math.round((imageBottom-bottom)/imageHeight*100000)};
  if(Object.values(crop).some(v=>v<0||v>=100000)||crop.left+crop.right>=100000||crop.top+crop.bottom>=100000)throw new Error('수식 그림의 자르기 값이 잘못되었습니다.');
  return {box:{x:box.x+left,y:box.y+top,w:right-left,h:bottom-top},crop};
}

// A narrowly scoped Office Math fallback: a one-column table whose otherwise
// empty cells each show a clipped portion of one embedded PNG. Unsupported
// configurations return null so callers retain the existing placeholder.
export function readMathTableFallback(frame){
  if(frame?.namespaceURI!==P||frame.localName!=='graphicFrame')return null;
  const graphic=child(frame,'graphic'),data=child(graphic,'graphicData');
  if(data?.getAttribute('uri')!=='http://schemas.openxmlformats.org/drawingml/2006/table')return null;
  const table=child(data,'tbl'),grid=child(table,'tblGrid');
  if(!table||elements(data).length!==1)return null;
  if(elements(table).some(n=>n.namespaceURI!==A||!['tblPr','tblGrid','tr','extLst'].includes(n.localName)))return null;
  const cols=elements(grid).filter(n=>n.namespaceURI===A&&n.localName==='gridCol');
  const rows=elements(table).filter(n=>n.namespaceURI===A&&n.localName==='tr');
  if(cols.length!==1||elements(grid).length!==1||!rows.length||rows.length>100)return null;
  const transform=elements(frame).find(n=>n.namespaceURI===P&&n.localName==='xfrm'),off=child(transform,'off'),ext=child(transform,'ext');
  if(!off||!ext||elements(transform).length!==2||elements(transform).some(n=>n.namespaceURI!==A||!['off','ext'].includes(n.localName)))return null;
  const parsedTransform=readMathTransform(transform);if(parsedTransform.rotation||parsedTransform.flipH||parsedTransform.flipV)return null;
  const {x,y,w,h}=parsedTransform.box,columnWidth=readInt(cols[0],'w',true);
  if(columnWidth<=0)throw new Error('수식 표의 좌표 또는 이미지 채우기 값이 잘못되었습니다.');
  const heights=rows.map(row=>readInt(row,'h',true));
  if(heights.some(value=>value<=0))throw new Error('수식 표의 좌표 또는 이미지 채우기 값이 잘못되었습니다.');
  if(Math.abs(columnWidth-w)>1||Math.abs(heights.reduce((a,b)=>a+b,0)-h)>1)return null;
  const fragments=[];let currentY=0,rid=null,omittedBorder=false;
  for(let i=0;i<rows.length;i++){
    const cells=elements(rows[i]).filter(n=>n.namespaceURI===A&&n.localName==='tc');
    if(cells.length!==1||elements(rows[i]).some(n=>n.namespaceURI!==A||!['tc','extLst'].includes(n.localName)))return null;
    const cell=cells[0],props=child(cell,'tcPr'),body=child(cell,'txBody');
    if(!props||!body||['rowSpan','gridSpan','hMerge','vMerge'].some(k=>cell.hasAttribute(k))||[...body.getElementsByTagNameNS('*','t')].some(t=>t.textContent.trim()))return null;
    const fills=elements(props).filter(n=>n.namespaceURI===A&&['blipFill','solidFill','noFill','gradFill','pattFill','grpFill'].includes(n.localName));
    if(elements(props).some(n=>n.namespaceURI!==A||!['blipFill','lnL','lnR','lnT','lnB','lnTlToBr','lnBlToTr','extLst'].includes(n.localName)))return null;
    if(elements(props).some(n=>['lnL','lnR','lnT','lnB','lnTlToBr','lnBlToTr'].includes(n.localName)&&!child(n,'noFill')))omittedBorder=true;
    if(fills.length!==1||fills[0].localName!=='blipFill')return null;
    const fill=fills[0],blip=child(fill,'blip'),stretch=child(fill,'stretch'),fillRect=child(stretch,'fillRect');
    if(!blip||!stretch||elements(fill).length!==2||elements(blip).length||elements(stretch).length>1||elements(stretch).some(n=>n.namespaceURI!==A||n.localName!=='fillRect'))return null;
    if(blip.hasAttributeNS(R,'link'))throw new Error('수식 그림 관계 또는 이미지가 잘못되었습니다.');
    const cellRid=blip.getAttributeNS(R,'embed');if(!cellRid)throw new Error('수식 그림 관계 또는 이미지가 잘못되었습니다.');
    if(rid!==null&&cellRid!==rid)return null;rid=cellRid;
    const cellHeight=heights[i],fragment=stretchImageFragment({x,y:y+currentY,w:columnWidth,h:cellHeight},fillRect);
    if(fragment)fragments.push(fragment);
    currentY+=cellHeight;
  }
  return fragments.length?{rid,fragments,omittedBorder}:null;
}
