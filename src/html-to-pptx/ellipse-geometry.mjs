const A='http://schemas.openxmlformats.org/drawingml/2006/main';
const children=node=>[...(node?.childNodes??[])].filter(n=>n.nodeType===1);
const child=(node,name)=>children(node).find(n=>n.namespaceURI===A&&n.localName===name);

// Only approximate closed, explicit paths that follow one ellipse closely.
// Formulas, arcs and multiple paths remain unsupported rather than guessed.
export function nearlyEllipticalGeometry(geometry){
 if(geometry?.namespaceURI!==A||geometry.localName!=='custGeom')return false;
 const paths=children(child(geometry,'pathLst'));if(paths.length!==1||paths[0].namespaceURI!==A||paths[0].localName!=='path')return false;
 const path=paths[0],dimensions=['w','h'].map(k=>path.getAttribute(k));
 if(path.hasAttribute('fill')&&path.getAttribute('fill')!=='norm')return false;
 if(dimensions.some(v=>!/^\d+$/.test(v)||!Number.isSafeInteger(Number(v))||Number(v)<=0||Number(v)>1e9))return false;
 const [w,h]=dimensions.map(Number),commands=children(path);if(commands.length<3||commands.length>64)return false;
 if(commands[0].localName!=='moveTo'||commands.at(-1).localName!=='close'||commands.some(c=>c.namespaceURI!==A))return false;
 const points=node=>children(node).map(p=>{
  const x=p.getAttribute('x'),y=p.getAttribute('y');
  return p.namespaceURI===A&&p.localName==='pt'&&/^\d+$/.test(x)&&/^\d+$/.test(y)&&Number(x)<=w&&Number(y)<=h?{x:Number(x)/w,y:Number(y)/h}:null;
 });
 let current=null,start=null;const samples=[];
 const add=p=>{samples.push(p);current=p;};
 for(const [index,command] of commands.entries()){
  const p=points(command);if(p.some(v=>!v))return false;
  if(index===0&&command.localName==='moveTo'&&p.length===1){start=p[0];add(start);continue;}
  if(command.localName==='close'&&index===commands.length-1&&p.length===0){
   const from=current;for(let i=1;i<=12;i++){const t=i/12;add({x:from.x+(start.x-from.x)*t,y:from.y+(start.y-from.y)*t});}continue;
  }
  if(command.localName==='lnTo'&&p.length===1){const from=current;for(let i=1;i<=12;i++){const t=i/12;add({x:from.x+(p[0].x-from.x)*t,y:from.y+(p[0].y-from.y)*t});}continue;}
  if(command.localName==='cubicBezTo'&&p.length===3){const from=current;for(let i=1;i<=16;i++){const t=i/16,u=1-t;add({x:u**3*from.x+3*u*u*t*p[0].x+3*u*t*t*p[1].x+t**3*p[2].x,y:u**3*from.y+3*u*u*t*p[0].y+3*u*t*t*p[1].y+t**3*p[2].y});}continue;}
  return false;
 }
 if(samples.some(p=>Math.abs(Math.hypot(p.x*2-1,p.y*2-1)-1)>.03))return false;
 let winding=0,direction=0;
 for(let i=1;i<samples.length;i++){
  const a=samples[i-1],b=samples[i];let delta=Math.atan2(b.y-.5,b.x-.5)-Math.atan2(a.y-.5,a.x-.5);
  if(delta>Math.PI)delta-=2*Math.PI;if(delta< -Math.PI)delta+=2*Math.PI;
  if(Math.abs(delta)>1e-6){if(!direction)direction=Math.sign(delta);if(delta*direction< -1e-4)return false;winding+=delta;}
 }
 return Math.abs(Math.abs(winding)-2*Math.PI)<.05;
}
