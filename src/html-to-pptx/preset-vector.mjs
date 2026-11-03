import {DOMParser,XMLSerializer} from '@xmldom/xmldom';
import {parseFragment} from 'parse5';

// These are structured, bounded primitive shapes, never arbitrary SVG markup.
// The same generator serves the asset and the editable inline HTML so that
// visual output and reimport cannot diverge through hidden metadata.
export function validPresetVector(v){
  if(!v||typeof v!=='object'||Array.isArray(v))return false;
  const keys=['shape','width','height','angle','fill','stroke','strokeWidth','adj1','adj2','gradient','flipV'];
  if(Object.keys(v).some(k=>!keys.includes(k)))return false;
  if(!['ellipse','downArrow','rightArrow','bentArrow','triangle'].includes(v.shape))return false;
  if(v.flipV!==undefined&&(v.flipV!==true||v.shape!=='bentArrow'))return false;
  if(![v.width,v.height].every(x=>Number.isFinite(x)&&x>=0.00001&&x<=100000))return false;
  if(!Number.isFinite(v.angle)||Math.abs(v.angle)>=360||!Number.isFinite(v.strokeWidth)||v.strokeWidth<0||v.strokeWidth>1024)return false;
  if(typeof v.stroke!=='string'||!/^(?:none|#[0-9a-f]{6})$/i.test(v.stroke))return false;
  if(v.fill==='gradient'){
    const g=v.gradient;if(!g||typeof g!=='object'||Array.isArray(g)||Object.keys(g).some(k=>!['id','from','to'].includes(k))||!/^g-[a-f0-9]{20}$/.test(g.id)||![g.from,g.to].every(c=>typeof c==='string'&&/^#[0-9a-f]{6}$/i.test(c)))return false;
  }else if(typeof v.fill!=='string'||!/^(?:none|#[0-9a-f]{6})$/i.test(v.fill)||v.gradient!==undefined)return false;
  if(!Number.isSafeInteger(v.adj1)||v.adj1<0||v.adj1>100000||!Number.isSafeInteger(v.adj2)||v.adj2<0||v.adj2>2147483647)return false;
  if(v.shape==='bentArrow'&&(v.adj1!==25000||v.adj2!==25000))return false;
  if(v.adj2>100000*(v.shape==='rightArrow'?v.width:v.height)/Math.min(v.width,v.height))return false;
  const a=v.angle*Math.PI/180,c=Math.abs(Math.cos(a)),s=Math.abs(Math.sin(a));
  return [v.width*c+v.height*s+v.strokeWidth,v.width*s+v.height*c+v.strokeWidth].every(x=>Number.isFinite(x)&&x>0&&x<=100000);
}

export function presetVectorMarkup(v){
  if(!validPresetVector(v))throw new Error('SVG 도형의 숫자 또는 색상이 올바르지 않습니다.');
  const a=v.angle*Math.PI/180,c=Math.abs(Math.cos(a)),s=Math.abs(Math.sin(a));
  const width=v.width*c+v.height*s+v.strokeWidth,height=v.width*s+v.height*c+v.strokeWidth;
  const n=x=>String(Number(x.toFixed(10)));
  const shape=v.shape==='ellipse'?`<ellipse cx="0" cy="0" rx="${n(v.width/2)}" ry="${n(v.height/2)}"/>`:v.shape==='triangle'?`<path d="M ${n(v.width/2)} 0 L ${n(-v.width/2)} ${n(-v.height/2)} L ${n(-v.width/2)} ${n(v.height/2)} Z"/>`:v.shape==='bentArrow'?(()=>{
    // ECMA preset defaults: adj1=adj2=adj3=25000; adj4=43750.
    const w=v.width,h=v.height,ss=Math.min(w,h),th=ss*.25,aw=ss*.25,dh=aw-th/2,ah=ss*.25;
    const bd=Math.min(ss*.4375,w-ah,h-dh),inner=Math.max(bd-th,0),x3=th+inner,x4=w-ah,y3=dh+th,y4=y3+dh,y5=dh+bd,y6=y3+inner;
    const x=q=>n(q-w/2),y=q=>n(q-h/2);
    return `<path d="M ${x(0)} ${y(h)} L ${x(0)} ${y(y5)} A ${n(bd)} ${n(bd)} 0 0 1 ${x(bd)} ${y(dh)} L ${x(x4)} ${y(dh)} L ${x(x4)} ${y(0)} L ${x(w)} ${y(aw)} L ${x(x4)} ${y(y4)} L ${x(x4)} ${y(y3)} L ${x(x3)} ${y(y3)} A ${n(inner)} ${n(inner)} 0 0 0 ${x(th)} ${y(y6)} L ${x(th)} ${y(h)} Z"/>`;
  })():v.shape==='rightArrow'?(()=>{
    const x=v.width/2,y=v.height/2,shaft=v.height*v.adj1/200000,base=x-Math.min(v.width,v.height)*v.adj2/100000;
    return `<path d="M ${n(-x)} ${n(-shaft)} L ${n(base)} ${n(-shaft)} L ${n(base)} ${n(-y)} L ${n(x)} 0 L ${n(base)} ${n(y)} L ${n(base)} ${n(shaft)} L ${n(-x)} ${n(shaft)} Z"/>`;
  })():(()=>{
    const x=v.width/2,y=v.height/2,shaft=v.width*v.adj1/200000,base=y-Math.min(v.width,v.height)*v.adj2/100000;
    return `<path d="M ${n(-shaft)} ${n(-y)} L ${n(shaft)} ${n(-y)} L ${n(shaft)} ${n(base)} L ${n(x)} ${n(base)} L 0 ${n(y)} L ${n(-x)} ${n(base)} L ${n(-shaft)} ${n(base)} Z"/>`;
  })();
  const defs=v.gradient?`<defs><linearGradient id="${v.gradient.id}" gradientUnits="userSpaceOnUse" color-interpolation="linearRGB" x1="${n(-v.width/2)}" y1="0" x2="${n(v.width/2)}" y2="0"><stop offset="0" stop-color="${v.gradient.from}"/><stop offset="1" stop-color="${v.gradient.to}"/></linearGradient></defs>`:'';
  const fill=v.gradient?`url(#${v.gradient.id})`:v.fill;
  const body=`${defs}<g transform="translate(${n(width/2)} ${n(height/2)}) rotate(${n(v.angle)})${v.flipV?' scale(1 -1)':''}" fill="${fill}" stroke="${v.stroke}" stroke-width="${n(v.strokeWidth)}">${shape}</g>`;
  return {width,height,body,svg:`<svg xmlns="http://www.w3.org/2000/svg" width="${n(width)}" height="${n(height)}" viewBox="0 0 ${n(width)} ${n(height)}">${body}</svg>`};
}

// Office renders SVG gradients in sRGB even when linearRGB was requested.
// Approximate the validated two-stop preset in sRGB with sampled linear light.
export function officeCompatiblePresetGradient(svg){
  if(!svg.includes('data-preset-vector')||!svg.includes('linearRGB'))return svg;
  const doc=new DOMParser().parseFromString(svg,'image/svg+xml');
  const gradient=doc.getElementsByTagNameNS('http://www.w3.org/2000/svg','linearGradient')[0];
  if(!gradient||gradient.getAttribute('color-interpolation')!=='linearRGB')return svg;
  const stops=[...gradient.getElementsByTagNameNS('http://www.w3.org/2000/svg','stop')];
  if(stops.length!==2||stops[0].getAttribute('offset')!=='0'||stops[1].getAttribute('offset')!=='1')return svg;
  const colors=stops.map(s=>s.getAttribute('stop-color'));
  if(!colors.every(c=>/^#[0-9a-f]{6}$/i.test(c)))return svg;
  const linear=c=>{const v=c/255;return v<=0.04045?v/12.92:((v+0.055)/1.055)**2.4;};
  const srgb=v=>Math.round(255*(v<=0.0031308?v*12.92:1.055*v**(1/2.4)-0.055));
  const channels=colors.map(c=>[1,3,5].map(i=>linear(parseInt(c.slice(i,i+2),16))));
  while(gradient.firstChild)gradient.removeChild(gradient.firstChild);
  gradient.setAttribute('color-interpolation','sRGB');
  for(let i=0;i<=16;i++){
    const t=i/16,color=channels[0].map((x,j)=>srgb(x*(1-t)+channels[1][j]*t).toString(16).padStart(2,'0')).join('');
    const stop=doc.createElementNS('http://www.w3.org/2000/svg','stop');stop.setAttribute('offset',String(t));stop.setAttribute('stop-color',`#${color}`);gradient.appendChild(stop);
  }
  return new XMLSerializer().serializeToString(doc);
}

// A generated SVG can be edited in HTML, but its input remains restricted to
// elementary inline geometry: no scripts, resources, references or CSS paints.
export function validInlinePresetVector(root){
  const ns='http://www.w3.org/2000/svg';
  const attr=(node,name)=>node.attrs?.find(a=>a.name===name)?.value;
  const allowed=(node,names)=>node.attrs?.every(a=>names.includes(a.name)&&(!a.namespace||a.name==='xmlns'&&a.namespace==='http://www.w3.org/2000/xmlns/'&&a.value===ns));
  const elements=node=>(node.childNodes??[]).filter(n=>n.nodeName!=='#text'||n.value.trim());
  const decimal='[-+]?(?:\\d+(?:\\.\\d*)?|\\.\\d+)(?:[eE][-+]?\\d+)?';
  const scalar=value=>typeof value==='string'&&new RegExp(`^${decimal}$`).test(value)&&Number.isFinite(Number(value))&&Math.abs(Number(value))<=100000;
  if(root.tagName!=='svg'||root.namespaceURI!==ns||!allowed(root,['xmlns','data-source-id','data-parent-id','data-kind','data-preset-vector','data-pptx-vector','viewBox','style']))return false;
  const view=(attr(root,'viewBox')??'').trim().split(/[ ,]+/);if(view.length!==4||!view.every(scalar)||Number(view[2])<=0||Number(view[3])<=0)return false;
  const groups=elements(root);let gradientId=null;
  if(groups.length===2){
    const defs=groups.shift();if(defs.tagName!=='defs'||defs.namespaceURI!==ns||!allowed(defs,[]))return false;
    const gradients=elements(defs);if(gradients.length!==1||gradients[0].tagName!=='linearGradient'||gradients[0].namespaceURI!==ns)return false;
    const gradient=gradients[0];if(!allowed(gradient,['id','gradientUnits','color-interpolation','x1','y1','x2','y2'])||attr(gradient,'gradientUnits')!=='userSpaceOnUse'||attr(gradient,'color-interpolation')!=='linearRGB'||!/^g-[a-f0-9]{20}$/.test(attr(gradient,'id')??'')||!['x1','y1','x2','y2'].every(k=>scalar(attr(gradient,k))))return false;
    gradientId=attr(gradient,'id');const stops=elements(gradient);
    if(stops.length!==2||stops.some((stop,i)=>stop.tagName!=='stop'||stop.namespaceURI!==ns||!allowed(stop,['offset','stop-color'])||elements(stop).length||attr(stop,'offset')!==String(i)||!/^#[0-9a-f]{6}$/i.test(attr(stop,'stop-color')??'')))return false;
  }
  if(groups.length!==1||groups[0].tagName!=='g'||groups[0].namespaceURI!==ns)return false;
  const group=groups[0];if(!allowed(group,['transform','fill','stroke','stroke-width']))return false;
  if(!/^(?:none|#[0-9a-f]{6})$/i.test(attr(group,'stroke')??'')||(gradientId?attr(group,'fill')!==`url(#${gradientId})`:!/^(?:none|#[0-9a-f]{6})$/i.test(attr(group,'fill')??'')))return false;
  if(!scalar(attr(group,'stroke-width'))||Number(attr(group,'stroke-width'))<0||Number(attr(group,'stroke-width'))>1024)return false;
  const transform=new RegExp(`^translate\\((${decimal})[ ,]+(${decimal})\\)\\s+rotate\\((${decimal})\\)(?:\\s+scale\\(1 -1\\))?$`).exec(attr(group,'transform')??'');
  if(!transform||!transform.slice(1).every(scalar)||Math.abs(Number(transform[3]))>=360)return false;
  const shapes=elements(group);if(shapes.length!==1||shapes[0].namespaceURI!==ns)return false;
  const shape=shapes[0];if(elements(shape).length)return false;
  if(shape.tagName==='ellipse')return allowed(shape,['cx','cy','rx','ry'])&&['cx','cy','rx','ry'].every(k=>scalar(attr(shape,k)))&&Number(attr(shape,'rx'))>0&&Number(attr(shape,'ry'))>0;
  if(shape.tagName!=='path'||!allowed(shape,['d']))return false;
  const d=attr(shape,'d')??'',tokens=d.trim().split(/[ ,]+/);
  if(tokens.length<10||tokens.length>194||tokens[0]!=='M'||tokens.at(-1)!=='Z')return false;
  if(!scalar(tokens[1])||!scalar(tokens[2]))return false;
  for(let i=3;i<tokens.length-1;){
    if(tokens[i]==='L'){if(!scalar(tokens[i+1])||!scalar(tokens[i+2]))return false;i+=3;}
    else if(tokens[i]==='A'){
      if(![tokens[i+1],tokens[i+2],tokens[i+3],tokens[i+6],tokens[i+7]].every(scalar)||Number(tokens[i+1])<=0||Number(tokens[i+2])<=0||!['0','1'].includes(tokens[i+4])||!['0','1'].includes(tokens[i+5]))return false;i+=8;
    }else return false;
    if(i>tokens.length-1)return false;
  }
  return true;
}

// Metadata only identifies a shape when its actual editable SVG geometry still
// matches. Edited colors are read from the SVG; changed paths fall back to art.
export function recoverPresetVector(root){
  const attr=(n,k)=>n.attrs?.find(a=>a.name===k)?.value;
  const children=n=>(n.childNodes??[]).filter(x=>x.tagName);
  let v;try{v=JSON.parse(attr(root,'data-pptx-vector'));}catch{return null;}
  if(!validPresetVector(v)||!validInlinePresetVector(root))return null;
  const elems=children(root),g=elems.find(x=>x.tagName==='g'),defs=elems.find(x=>x.tagName==='defs');
  v={...v,stroke:attr(g,'stroke'),strokeWidth:Number(attr(g,'stroke-width'))};
  if(defs){const grad=children(defs)[0],stops=children(grad);v.fill='gradient';v.gradient={id:attr(grad,'id'),from:attr(stops[0],'stop-color'),to:attr(stops[1],'stop-color')};}
  else{v.fill=attr(g,'fill');delete v.gradient;}
  if(!validPresetVector(v))return null;
  const canonical=children(parseFragment(presetVectorMarkup(v).svg))[0],cg=children(canonical).find(x=>x.tagName==='g');
  const a=children(g)[0],b=children(cg)[0];
  const equivalent=(x,y)=>{const tokenize=s=>String(s).match(/[A-Za-z]+|[-+]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][-+]?\d+)?/g)??[],left=tokenize(x),right=tokenize(y);return left.length===right.length&&left.every((v,i)=>v===right[i]||Number.isFinite(Number(v))&&Number.isFinite(Number(right[i]))&&Math.abs(Number(v)-Number(right[i]))<1e-7);};
  if(!equivalent(attr(root,'viewBox'),attr(canonical,'viewBox'))||!equivalent(attr(g,'transform'),attr(cg,'transform'))||a.tagName!==b.tagName)return null;
  const keys=a.tagName==='ellipse'?['cx','cy','rx','ry']:['d'];if(keys.some(k=>!equivalent(attr(a,k),attr(b,k))))return null;
  if(defs){const grad=children(defs)[0],expected=children(children(canonical).find(x=>x.tagName==='defs'))[0];if(['x1','x2','y1','y2','gradientUnits','color-interpolation'].some(k=>attr(grad,k)!==attr(expected,k)))return null;}
  return v;
}
