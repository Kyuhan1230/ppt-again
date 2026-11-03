import {EMU_PER_CSS_PX} from './canvas.mjs';
const A='http://schemas.openxmlformats.org/drawingml/2006/main';
const elements=n=>[...(n?.childNodes??[])].filter(c=>c.nodeType===1);
const attr=(n,k,fallback=null)=>n?.hasAttribute(k)?n.getAttribute(k):fallback;

// Static, unscaled solid outer shadows can be approximated by CSS box-shadow.
export function simpleOuterShadow(node,scale){
 if(node?.namespaceURI!==A||node.localName!=='outerShdw'||!Number.isFinite(scale)||scale<=0)return null;
 const allowed=new Set(['blurRad','dist','dir','algn','rotWithShape','sx','sy','kx','ky']);
 if([...(node.attributes??[])].some(a=>a.namespaceURI||!allowed.has(a.name)))return null;
 if(!['tl','t','tr','l','ctr','r','bl','b','br'].includes(attr(node,'algn','ctr'))||!['0','1','true','false'].includes(attr(node,'rotWithShape','1')))return null;
 if(attr(node,'sx','100000')!=='100000'||attr(node,'sy','100000')!=='100000'||attr(node,'kx','0')!=='0'||attr(node,'ky','0')!=='0')return null;
 const vals=['blurRad','dist','dir'].map(k=>attr(node,k,'0'));if(vals.some(v=>!/^\d+$/.test(v)))return null;
 const [blurRaw,distRaw,direction]=vals.map(Number);if(![blurRaw,distRaw,direction].every(Number.isSafeInteger)||direction>21600000)return null;
 const blur=blurRaw/EMU_PER_CSS_PX*scale,dist=distRaw/EMU_PER_CSS_PX*scale;if(blur>2000||dist>2000)return null;
 const colors=elements(node),c=colors[0];if(colors.length!==1||c.namespaceURI!==A||!['srgbClr','prstClr'].includes(c.localName))return null;
 const hex=c.localName==='srgbClr'?attr(c,'val',''):({black:'000000',white:'ffffff'})[attr(c,'val','')];if(!/^[a-f\d]{6}$/i.test(hex??''))return null;
 const effects=elements(c);if(effects.length>1||effects.some(e=>e.namespaceURI!==A||e.localName!=='alpha'||elements(e).length))return null;
 const alphaRaw=effects.length?attr(effects[0],'val',''):'100000';if(!/^\d+$/.test(alphaRaw)||Number(alphaRaw)>100000)return null;
 const radians=direction/60000*Math.PI/180;
 return {x:Math.cos(radians)*dist,y:Math.sin(radians)*dist,blur,r:parseInt(hex.slice(0,2),16),g:parseInt(hex.slice(2,4),16),b:parseInt(hex.slice(4,6),16),alpha:Number(alphaRaw)/100000};
}
