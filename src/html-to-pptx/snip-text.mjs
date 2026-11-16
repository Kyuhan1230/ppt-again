const A='http://schemas.openxmlformats.org/drawingml/2006/main';
const elements=n=>[...(n?.childNodes??[])].filter(c=>c.nodeType===1);
const child=(n,name)=>elements(n).find(c=>c.localName===name);
const attr=(n,name,otherwise=null)=>n?.hasAttribute(name)?n.getAttribute(name):otherwise;

// A constrained preset shape: editable text is kept separately from its
// clipped-corner background. Unknown guides or transforms stay unsupported.
export function snippedTextGeometry(geometry,xfrm,raw,map){
 if(geometry?.namespaceURI!==A||attr(geometry,'prst')!=='snip1Rect'||xfrm?.namespaceURI!==A)return null;
 const rotation=attr(xfrm,'rot','0'),flipH=attr(xfrm,'flipH','0'),flipV=attr(xfrm,'flipV','0');
 if(!/^0+$/.test(rotation)||!['0','1'].includes(flipH)||!['0','1'].includes(flipV))return null;
 const lists=elements(geometry);if(lists.length!==1||lists[0].namespaceURI!==A||lists[0].localName!=='avLst')return null;
 const guides=elements(lists[0]);if(guides.length>1)return null;let adjustment=16667;
 if(guides.length){
  const guide=guides[0],match=/^val\s+(-?\d+)$/.exec(attr(guide,'fmla',''));
  if(guide.namespaceURI!==A||guide.localName!=='gd'||attr(guide,'name')!=='adj'||elements(guide).length||!match)return null;
  adjustment=Number(match[1]);if(!Number.isSafeInteger(adjustment))return null;adjustment=Math.max(0,Math.min(50000,adjustment));
 }
 if(![raw.w,raw.h,map.sx,map.sy].every(n=>Number.isFinite(n)&&n>0))return null;
 const cut=Math.min(raw.w,raw.h)*adjustment/100000/9525;
 return {cutX:cut*map.sx,cutY:cut*map.sy,flipH:flipH==='1',flipV:flipV==='1'};
}

export function snippedTextSvg(box,geometry,fill){
 if(![box.w,box.h,geometry.cutX,geometry.cutY].every(n=>Number.isFinite(n)&&n>=0&&n<=100000)||!/^#[a-f\d]{6}$/i.test(fill))return null;
 const points=[[0,0],[box.w-geometry.cutX,0],[box.w,geometry.cutY],[box.w,box.h],[0,box.h]].map(([x,y])=>[geometry.flipH?box.w-x:x,geometry.flipV?box.h-y:y].map(n=>Number(n.toFixed(6))).join(',')).join(' ');
 return `<svg xmlns="http://www.w3.org/2000/svg" width="${box.w}" height="${box.h}" viewBox="0 0 ${box.w} ${box.h}"><polygon points="${points}" fill="${fill}"/></svg>`;
}
