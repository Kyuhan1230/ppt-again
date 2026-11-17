import {inflateSync} from 'node:zlib';
import {sha256} from './assets.mjs';
import {validatePngStructure} from './png-duotone.mjs';

const MAX_WIDTH=8192,MAX_HEIGHT=32768,MAX_PIXELS=32000000,MAX_TILES=64,TILE_HEIGHT=1024,OVERLAP=2,MAX_MS=30000;
const fail=message=>Object.assign(new Error(message),{code:'PAGINATION_UNSUPPORTED'});
// close() is requested, not awaited. The owning worker must confirm context
// termination and escalate to its process supervisor if cleanup does not end.
const timeout=()=>Object.assign(fail('문서 타일 관측 시간이 초과되었습니다.'),{reasonKey:'PAGINATION_TIMEOUT',cleanupRequired:true});
const validViewport=v=>v&&Number.isSafeInteger(v.width)&&Number.isSafeInteger(v.height)&&v.width>0&&v.height>0&&v.width<=8192&&v.height<=8192;

// Chromium screenshots are non-interlaced 8-bit RGB/RGBA PNGs. Decode rows
// to compare actual pixels at tile seams, not the compressed PNG bytes.
function pngEdges(bytes,width,height){
  const data=Buffer.from(bytes);validatePngStructure(data);
  if(data.readUInt32BE(16)!==width||data.readUInt32BE(20)!==height||data[24]!==8||![2,6].includes(data[25])||data[28]!==0)throw fail('캡처한 타일의 픽셀 크기 또는 형식을 확인할 수 없습니다.');
  const channels=data[25]===6?4:3,rowBytes=width*channels,compressed=[];
  for(let offset=8;offset<data.length;){const size=data.readUInt32BE(offset),type=data.toString('ascii',offset+4,offset+8);if(type==='IDAT')compressed.push(data.subarray(offset+8,offset+8+size));offset+=size+12;}
  const expected=(rowBytes+1)*height,raw=inflateSync(Buffer.concat(compressed),{maxOutputLength:expected});if(raw.length!==expected)throw fail('캡처한 타일의 픽셀을 확인할 수 없습니다.');
  let previous=Buffer.alloc(rowBytes),first=[],last=[];
  for(let y=0;y<height;y++){
    const start=y*(rowBytes+1),filter=raw[start],row=Buffer.from(raw.subarray(start+1,start+1+rowBytes));if(filter>4)throw fail('캡처한 타일의 픽셀을 확인할 수 없습니다.');
    for(let i=0;i<rowBytes;i++){
      const left=i>=channels?row[i-channels]:0,up=previous[i],upperLeft=i>=channels?previous[i-channels]:0;
      let prediction=0;if(filter===1)prediction=left;else if(filter===2)prediction=up;else if(filter===3)prediction=Math.floor((left+up)/2);else if(filter===4){const p=left+up-upperLeft,a=Math.abs(p-left),b=Math.abs(p-up),c=Math.abs(p-upperLeft);prediction=a<=b&&a<=c?left:b<=c?up:upperLeft;}
      row[i]=(row[i]+prediction)&255;
    }
    if(y<OVERLAP)first.push(row);last.push(row);if(last.length>OVERLAP)last.shift();previous=row;
  }
  return {first,last};
}

export async function captureDocumentTiles({page,rect,viewport,readState,timeoutMs=MAX_MS}={}){
  try{return await captureChecked({page,rect,viewport,readState,timeoutMs});}catch(error){if(error?.code==='PAGINATION_UNSUPPORTED')throw error;throw fail('문서 타일을 안전하게 확인할 수 없습니다.');}
}
async function captureChecked({page,rect,viewport,readState,timeoutMs}){
  if(!page||typeof page.screenshot!=='function'||typeof page.viewportSize!=='function'||typeof page.close!=='function'||typeof readState!=='function'||!Number.isSafeInteger(timeoutMs)||timeoutMs<1||timeoutMs>MAX_MS||!validViewport(viewport)||!rect||!['x','y','width','height'].every(key=>Number.isFinite(rect[key]))||rect.x<0||rect.y<0||rect.width<=0||rect.height<=0)throw fail('문서 타일 범위가 올바르지 않습니다.');
  const x=Math.floor(rect.x),y=Math.floor(rect.y),right=Math.ceil(rect.x+rect.width),bottom=Math.ceil(rect.y+rect.height),width=right-x,height=bottom-y;
  if(width>MAX_WIDTH||height>MAX_HEIGHT||width*height>MAX_PIXELS||Math.ceil(height/TILE_HEIGHT)>MAX_TILES)throw fail('문서 타일 크기가 지원 한도를 초과합니다.');
  const bounds={x,y,width,height},tiles=[];
  for(let offset=0;offset<height;offset+=TILE_HEIGHT){const ownedHeight=Math.min(TILE_HEIGHT,height-offset),overlap=Math.min(OVERLAP,height-offset-ownedHeight);tiles.push({x,y:y+offset,width,height:ownedHeight+overlap,owned:{x,y:y+offset,width,height:ownedHeight}});}
  const deadline=Date.now()+timeoutMs;
  const abortPage=()=>{try{Promise.resolve(page.close({runBeforeUnload:false})).catch(()=>{});}catch{}};
  const beforeDeadline=operation=>new Promise((resolve,reject)=>{
    const remaining=deadline-Date.now();
    const expire=()=>{abortPage();reject(timeout());};
    if(remaining<=0){expire();return;}
    const timer=setTimeout(expire,remaining);
    Promise.resolve().then(operation).then(value=>{clearTimeout(timer);resolve(value);},error=>{clearTimeout(timer);reject(error);});
  });
  const check=async expected=>{
    if(Date.now()>=deadline){abortPage();throw timeout();}
    const actualViewport=page.viewportSize(),state=await beforeDeadline(readState);
    if(!actualViewport||actualViewport.width!==viewport.width||actualViewport.height!==viewport.height||!state?.viewport||state.viewport.width!==viewport.width||state.viewport.height!==viewport.height||state.viewport.dpr!==1||!Number.isFinite(state.viewport.scrollX)||!Number.isFinite(state.viewport.scrollY)||!Number.isFinite(state.documentSize?.width)||!Number.isFinite(state.documentSize?.height)||right>state.documentSize.width||bottom>state.documentSize.height)throw fail('문서 화면 또는 범위가 바뀌었습니다.');
    const digest=sha256(JSON.stringify(state));if(expected!==undefined&&digest!==expected)throw fail('문서 타일 관측 중 화면이 바뀌었습니다.');return digest;
  };
  const stateHash=await check(),hashes=[],edges=[];
  const capture=async tile=>{
    await check(stateHash);
    const bytes=await beforeDeadline(()=>page.screenshot({fullPage:true,clip:{x:tile.x,y:tile.y,width:tile.width,height:tile.height},animations:'disabled',scale:'css',timeout:Math.max(1,Math.min(5000,deadline-Date.now()))}));
    if(!(bytes instanceof Uint8Array)||!bytes.length||bytes.length>16*1024*1024)throw fail('문서 타일이 누락되었거나 너무 큽니다.');
    const pixels=pngEdges(bytes,tile.width,tile.height);await check(stateHash);return {hash:sha256(bytes),pixels};
  };
  for(const [index,tile] of tiles.entries()){
    const result=await capture(tile);hashes.push(result.hash);edges.push(result.pixels);
    if(index>0){const previous=tiles[index-1],overlap=previous.y+previous.height-tile.y;if(overlap<=0||overlap>OVERLAP||edges[index-1].last.slice(-overlap).some((row,i)=>!row.equals(result.pixels.first[i])))throw fail('문서 타일의 겹친 픽셀이 일치하지 않습니다.');}
  }
  for(let index=tiles.length-1;index>=0;index--){const result=await capture(tiles[index]);if(result.hash!==hashes[index])throw fail('다시 관측한 문서 타일이 달라졌습니다.');}
  const output=tiles.map((tile,index)=>({...tile,sha256:hashes[index]}));
  return {bounds,tiles:output,fingerprint:sha256(JSON.stringify({version:1,bounds,viewport,stateHash,tiles:output}))};
}
