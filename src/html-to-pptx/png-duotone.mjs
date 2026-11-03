import {deflateSync,inflateSync} from 'node:zlib';

const A='http://schemas.openxmlformats.org/drawingml/2006/main';
const SIGNATURE=Buffer.from([137,80,78,71,13,10,26,10]);
const MAX_PIXELS=4_000_000,MAX_BYTES=16*1024*1024;
const children=n=>[...(n?.childNodes??[])].filter(c=>c.nodeType===1);
const clamp=(n,min,max)=>Math.min(max,Math.max(min,n));
const srgbToLinear=c=>c<=.04045?c/12.92:((c+.055)/1.055)**2.4;
const linearToSrgb=c=>c<=.0031308?12.92*c:1.055*c**(1/2.4)-.055;
const invalid=()=>{throw new Error('PNG 이미지 구조 또는 색상 효과가 잘못되었습니다.');};
const limit=()=>{throw Object.assign(new Error('이미지 픽셀 또는 압축 해제 한도를 초과했습니다.'),{code:'LIMIT_EXCEEDED'});};

function saturation(rgb,multiplier){
 const max=Math.max(...rgb),min=Math.min(...rgb),light=(max+min)/2,delta=max-min;
 if(!delta)return rgb;
 const saturation=clamp(delta/(1-Math.abs(2*light-1))*multiplier,0,1);
 let hue=max===rgb[0]?((rgb[1]-rgb[2])/delta)%6:max===rgb[1]?(rgb[2]-rgb[0])/delta+2:(rgb[0]-rgb[1])/delta+4;
 hue=(hue*60+360)%360;
 const chroma=(1-Math.abs(2*light-1))*saturation,x=chroma*(1-Math.abs((hue/60)%2-1)),m=light-chroma/2;
 return (hue<60?[chroma,x,0]:hue<120?[x,chroma,0]:hue<180?[0,chroma,x]:hue<240?[0,x,chroma]:hue<300?[x,0,chroma]:[chroma,0,x]).map(c=>c+m);
}

// Only opaque, deterministic RGB endpoints with supported ordered transforms
// are accepted. Unknown DrawingML effects remain unsupported by the caller.
export function duotoneColors(duotone,lookupScheme){
 if(duotone?.namespaceURI!==A||duotone.localName!=='duotone')return null;
 const stops=children(duotone);if(stops.length!==2)return null;
 const parse=stop=>{
  if(stop.namespaceURI!==A)return null;
  let base=stop;
  if(stop.localName==='schemeClr')base=lookupScheme(stop.getAttribute('val'));
  if(!base||base.namespaceURI!==A)return null;
  let hex=base.localName==='srgbClr'?base.getAttribute('val'):base.localName==='sysClr'?base.getAttribute('lastClr'):base.localName==='prstClr'?({black:'000000',white:'FFFFFF'})[base.getAttribute('val')]:null;
  if(!/^[a-f\d]{6}$/i.test(hex??''))return null;
  let rgb=[0,2,4].map(i=>parseInt(hex.slice(i,i+2),16)/255);
  const effects=[...children(base),...(base===stop?[]:children(stop))];
  for(const effect of effects){
   if(effect.namespaceURI!==A||!/^\d+$/.test(effect.getAttribute('val')??''))return null;
   const value=Number(effect.getAttribute('val'));if(!Number.isSafeInteger(value))return null;
   if(effect.localName==='shade'&&value<=100000)rgb=rgb.map(c=>linearToSrgb(srgbToLinear(c)*value/100000));
   else if(effect.localName==='satMod'&&value<=1000000)rgb=saturation(rgb,value/100000);
   else if(effect.localName==='alpha'&&value===100000)continue;
   else return null;
  }
  return rgb.map(c=>Math.round(clamp(c,0,1)*255));
 };
 const result=stops.map(parse);return result.every(Boolean)?result:null;
}

const crcTable=Uint32Array.from({length:256},(_,n)=>{let x=n;for(let i=0;i<8;i++)x=x&1?0xedb88320^(x>>>1):x>>>1;return x>>>0;});
function crc32(bytes){let value=0xffffffff;for(const byte of bytes)value=crcTable[(value^byte)&255]^(value>>>8);return (value^0xffffffff)>>>0;}

// Validate a PNG used as a math fallback before promising a preview. This
// checks the complete chunk stream and bounds decompression for all standard
// color types, bit depths, and Adam7 interlace; it does not transform pixels.
export function validatePngStructure(bytes){
 const input=Buffer.from(bytes),maxInflated=128*1024*1024,maxPixels=32_000_000;
 if(input.length<33||input.length>MAX_BYTES||!input.subarray(0,8).equals(SIGNATURE))invalid();
 if(input.readUInt32BE(8)!==13||input.toString('ascii',12,16)!=='IHDR')invalid();
 const width=input.readUInt32BE(16),height=input.readUInt32BE(20),depth=input[24],typeCode=input[25];
 const depths={0:[1,2,4,8,16],2:[8,16],3:[1,2,4,8],4:[8,16],6:[8,16]},channels={0:1,2:3,3:1,4:2,6:4};
 if(!width||!height||!depths[typeCode]?.includes(depth)||input[26]!==0||input[27]!==0||![0,1].includes(input[28]))invalid();
 if(width>maxPixels||height>maxPixels||width*height>maxPixels)limit();
 let offset=8,count=0,seenHeader=false,seenData=false,endedData=false,seenEnd=false,seenPalette=false,paletteCount=0;const data=[];
 while(offset<input.length){
  if(++count>10000)limit();
  if(offset+12>input.length)invalid();const size=input.readUInt32BE(offset),end=offset+12+size;
  if(size>MAX_BYTES||end>input.length)invalid();const type=input.toString('ascii',offset+4,offset+8);
  if(!/^[A-Za-z]{4}$/.test(type)||type[2]!==type[2].toUpperCase()||crc32(input.subarray(offset+4,offset+8+size))!==input.readUInt32BE(offset+8+size))invalid();
  if(!seenHeader){if(type!=='IHDR'||size!==13)invalid();seenHeader=true;}
  else if(type==='IHDR')invalid();
  else if(type==='IDAT'){if(endedData||seenEnd)invalid();seenData=true;data.push(input.subarray(offset+8,offset+8+size));}
  else if(type==='IEND'){if(!seenData||size!==0||end!==input.length)invalid();seenEnd=true;}
  else if(type==='PLTE'){if(seenPalette||seenData||!size||size>768||size%3||[0,4].includes(typeCode))invalid();seenPalette=true;paletteCount=size/3;if(typeCode===3&&paletteCount>2**depth)invalid();}
  else {if(seenData)endedData=true;if(type[0]===type[0].toUpperCase())invalid();}
  offset=end;if(seenEnd)break;
 }
 if(!seenEnd||!data.length||typeCode===3&&!seenPalette)invalid();
 const passes=input[28]===0?[[0,0,1,1]]:[[0,0,8,8],[4,0,8,8],[0,4,4,8],[2,0,4,4],[0,2,2,4],[1,0,2,2],[0,1,1,2]];
 const rows=[];let expected=0;
 for(const [x,y,dx,dy] of passes){
  const w=Math.max(0,Math.ceil((width-x)/dx)),h=Math.max(0,Math.ceil((height-y)/dy));if(!w||!h)continue;
  const rowBytes=Math.ceil(w*channels[typeCode]*depth/8),size=(rowBytes+1)*h;
  if(!Number.isSafeInteger(size)||expected+size>maxInflated)limit();expected+=size;rows.push({w,rowBytes,h});
 }
 const compressed=Buffer.concat(data);let inflated;
 try{const result=inflateSync(compressed,{maxOutputLength:expected+1,info:true});if(result.engine.bytesWritten<compressed.length-data.at(-1).length)invalid();inflated=result.buffer;}catch{invalid();}
 if(inflated.length!==expected)invalid();let position=0;
 for(const pass of rows){
  let previous=typeCode===3?Buffer.alloc(pass.rowBytes):null,current=typeCode===3?Buffer.alloc(pass.rowBytes):null;
  for(let y=0;y<pass.h;y++){
   const filter=inflated[position++];if(filter>4)invalid();
   if(typeCode!==3){position+=pass.rowBytes;continue;}
   for(let x=0;x<pass.rowBytes;x++){
    const left=x?current[x-1]:0,up=previous[x],upperLeft=x?previous[x-1]:0;
    const prediction=filter===0?0:filter===1?left:filter===2?up:filter===3?Math.floor((left+up)/2):paeth(left,up,upperLeft);
    current[x]=(inflated[position++]+prediction)&255;
   }
   for(let x=0;x<pass.w;x++){const bit=x*depth,index=(current[bit>>3]>>(8-depth-(bit%8)))&((1<<depth)-1);if(index>=paletteCount)invalid();}
   [previous,current]=[current,previous];
  }
 }
}

function chunk(type,data){const out=Buffer.alloc(12+data.length);out.writeUInt32BE(data.length,0);out.write(type,4,4,'ascii');data.copy(out,8);out.writeUInt32BE(crc32(out.subarray(4,8+data.length)),8+data.length);return out;}
function paeth(a,b,c){const p=a+b-c,pa=Math.abs(p-a),pb=Math.abs(p-b),pc=Math.abs(p-c);return pa<=pb&&pa<=pc?a:pb<=pc?b:c;}

// Return null for a valid but unsupported PNG encoding. On a supported PNG,
// check chunks/CRC and cap dimensions and zlib output before decoding pixels.
export function applyPngDuotone(bytes,colors){
 const input=Buffer.from(bytes);if(input.length<33||input.length>MAX_BYTES||!input.subarray(0,8).equals(SIGNATURE))invalid();
 if(input.readUInt32BE(8)!==13||input.toString('ascii',12,16)!=='IHDR')invalid();
 const width=input.readUInt32BE(16),height=input.readUInt32BE(20);
 if(!width||!height)invalid();if(width>MAX_PIXELS||height>MAX_PIXELS||width*height>MAX_PIXELS)limit();
 const depth=input[24],typeCode=input[25],depths={0:[1,2,4,8,16],2:[8,16],3:[1,2,4,8],4:[8,16],6:[8,16]};
 if(!depths[typeCode]?.includes(depth)||input[26]!==0||input[27]!==0||![0,1].includes(input[28]))invalid();
 if(!Array.isArray(colors)||colors.length!==2||!colors.every(rgb=>Array.isArray(rgb)&&rgb.length===3&&rgb.every(v=>Number.isInteger(v)&&v>=0&&v<=255)))invalid();
 let offset=8,seenHeader=false,seenData=false,endedData=false,seenEnd=false,seenPalette=false,unsupportedProfile=false,chunkCount=0;const data=[],singletons=new Set();
 while(offset<input.length){
  if(++chunkCount>10000)limit();
  if(offset+12>input.length)invalid();const size=input.readUInt32BE(offset),end=offset+12+size;
  if(size>MAX_BYTES||end>input.length)invalid();const type=input.toString('ascii',offset+4,offset+8);
  if(!/^[A-Za-z]{4}$/.test(type)||type[2]!==type[2].toUpperCase()||crc32(input.subarray(offset+4,offset+8+size))!==input.readUInt32BE(offset+8+size))invalid();
  if(!seenHeader){if(type!=='IHDR'||size!==13)invalid();seenHeader=true;}
  else if(type==='IHDR')invalid();
  else if(type==='IDAT'){if(endedData||seenEnd)invalid();seenData=true;data.push(input.subarray(offset+8,offset+8+size));}
  else if(type==='IEND'){if(!seenData||size!==0||end!==input.length)invalid();seenEnd=true;}
  else if(type==='PLTE'){if(seenPalette||seenData||!size||size>768||size%3)invalid();seenPalette=true;}
  else {
   if(seenData)endedData=true;
   if(type[0]===type[0].toUpperCase())invalid();
   if(['sRGB','gAMA','cHRM','iCCP','pHYs','sBIT','tRNS','bKGD','hIST','tIME','eXIf'].includes(type)){
    if(singletons.has(type))invalid();singletons.add(type);
    if(type!=='tIME'&&seenData)invalid();
    const payload=input.subarray(offset+8,offset+8+size);
    if(type==='sRGB'&&(size!==1||payload[0]>3))invalid();
    if(type==='gAMA'){if(size!==4||payload.readUInt32BE(0)===0)invalid();if(payload.readUInt32BE(0)!==45455)unsupportedProfile=true;}
    if(type==='pHYs'&&(size!==9||payload[8]>1))invalid();
    if(type==='cHRM'){if(size!==32)invalid();unsupportedProfile=true;}
    if(type==='iCCP'){if(size<3||payload.indexOf(0)<1||payload.indexOf(0)>79||payload[payload.indexOf(0)+1]!==0)invalid();unsupportedProfile=true;}
    if(type==='sBIT')unsupportedProfile=true;
    if(type==='tRNS'&&typeCode===6)invalid();
   }
  }
  offset=end;if(seenEnd)break;
 }
 if(!seenEnd||!data.length)invalid();
 if(unsupportedProfile||depth!==8||typeCode!==6||input[28]!==0)return null;
 const stride=width*4,expected=(stride+1)*height;
 const compressed=Buffer.concat(data);let inflated;try{const result=inflateSync(compressed,{maxOutputLength:expected,info:true});if(result.engine.bytesWritten!==compressed.length)invalid();inflated=result.buffer;}catch{invalid();}
 if(inflated.length!==expected)invalid();
 const pixels=Buffer.alloc(stride*height);let src=0;
 for(let y=0;y<height;y++){
  const filter=inflated[src++];if(filter>4)invalid();const base=y*stride;
  for(let x=0;x<stride;x++){
   const a=x>=4?pixels[base+x-4]:0,b=y?pixels[base+x-stride]:0,c=y&&x>=4?pixels[base+x-stride-4]:0;
   const prediction=filter===0?0:filter===1?a:filter===2?b:filter===3?Math.floor((a+b)/2):paeth(a,b,c);
   pixels[base+x]=(inflated[src++]+prediction)&255;
  }
 }
 const rows=Buffer.alloc(expected);let target=0;
 for(let y=0;y<height;y++){
  rows[target++]=0;const base=y*stride;
  for(let x=0;x<stride;x+=4){
   const i=base+x,luminance=(.2126*pixels[i]+.7152*pixels[i+1]+.0722*pixels[i+2])/255;
   for(let c=0;c<3;c++)rows[target++]=Math.round(colors[0][c]+(colors[1][c]-colors[0][c])*luminance);
   rows[target++]=pixels[i+3];
  }
 }
 const header=Buffer.alloc(13);header.writeUInt32BE(width,0);header.writeUInt32BE(height,4);header[8]=8;header[9]=6;
 const output=Buffer.concat([SIGNATURE,chunk('IHDR',header),chunk('IDAT',deflateSync(rows)),chunk('IEND',Buffer.alloc(0))]);
 if(output.length>MAX_BYTES)limit();return output;
}
