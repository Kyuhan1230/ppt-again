const MAX_FONT_BYTES=16*1024*1024;
const utf16=new TextDecoder('utf-16be',{fatal:true});
const fail=(message,code='FONT_INVALID')=>{throw Object.assign(new Error(message),{code});};
const requireRange=(bytes,offset,length)=>{
  if(!Number.isInteger(offset)||!Number.isInteger(length)||offset<0||length<0||offset+length>bytes.length)fail('글꼴 데이터의 범위가 잘못되었습니다.');
};

// OS/2 embedding bits: https://learn.microsoft.com/en-us/typography/opentype/spec/os2#fstype
export function decideFontEmbedding(meta,{forEditing=true,subset=false}={}){
  const deny=reason=>({allowed:false,reason});
  if(!Number.isInteger(meta?.os2Version)||meta.os2Version<0||meta.os2Version>5)return deny('FONT_EMBEDDING_UNSUPPORTED');
  if(!Number.isInteger(meta.fsType)||meta.fsType<0||meta.fsType>0xFFFF)return deny('FONT_EMBEDDING_INVALID');
  const type=meta.fsType,permission=type&0xF;
  // ponytail: legacy high-bit semantics stay unsupported; add a version adapter only for a real input.
  if(meta.os2Version<2&&(type&0xFFF0))return deny('FONT_EMBEDDING_UNSUPPORTED');
  if(type&~0x30E)return deny('FONT_EMBEDDING_INVALID');
  if(![0,2,4,8].includes(permission))return deny(meta.os2Version<3?'FONT_EMBEDDING_UNSUPPORTED':'FONT_EMBEDDING_INVALID');
  if(permission===2)return deny('FONT_EMBEDDING_RESTRICTED');
  if(type&0x200)return deny('FONT_BITMAP_ONLY');
  if(forEditing&&permission===4)return deny('FONT_PREVIEW_ONLY');
  if(subset&&(type&0x100))return deny('FONT_SUBSETTING_RESTRICTED');
  return {allowed:true,reason:null};
}

function names(table){
  requireRange(table,0,6);
  const version=table.readUInt16BE(0),count=table.readUInt16BE(2),storage=table.readUInt16BE(4),recordsEnd=6+count*12;
  if(version>1)fail('지원하지 않는 글꼴 이름 표입니다.','FONT_FORMAT_UNSUPPORTED');
  requireRange(table,6,count*12);
  let headerEnd=recordsEnd,languageCount=0;
  if(version===1){
    requireRange(table,recordsEnd,2);languageCount=table.readUInt16BE(recordsEnd);headerEnd+=2+languageCount*4;
    requireRange(table,recordsEnd+2,languageCount*4);
    for(let i=0;i<languageCount;i++){const p=recordsEnd+2+i*4;requireRange(table,storage+table.readUInt16BE(p+2),table.readUInt16BE(p));}
  }
  if(storage<headerEnd||storage>table.length)fail('글꼴 이름 저장 위치가 잘못되었습니다.');
  const pairs=new Map(),keys=new Map();let decodedBytes=0;
  for(let i=0;i<count;i++){
    const p=6+i*12,platform=table.readUInt16BE(p),encoding=table.readUInt16BE(p+2),language=table.readUInt16BE(p+4),id=table.readUInt16BE(p+6),length=table.readUInt16BE(p+8),offset=storage+table.readUInt16BE(p+10);
    requireRange(table,offset,length);
    const recordKey=`${platform}/${encoding}/${language}/${id}`;
    const prior=keys.get(recordKey);
    if(prior){
      // Some shipped fonts repeat font-specific names while pointing to exactly the same bytes.
      if(id<256||prior.offset!==offset||prior.length!==length)fail('글꼴 이름 키가 중복됩니다.');
      continue;
    }
    keys.set(recordKey,{offset,length});
    if(language>=0x8000&&(version!==1||language-0x8000>=languageCount))fail('글꼴 이름의 언어 참조가 잘못되었습니다.');
    if(![1,2].includes(id)||!(platform===0||(platform===3&&[1,10].includes(encoding))))continue;
    if((decodedBytes+=length)>1024*1024)fail('글꼴 이름 metadata의 처리량 제한을 초과했습니다.');
    if(length%2)fail('글꼴 이름의 UTF-16 길이가 잘못되었습니다.');
    let value;try{value=utf16.decode(table.subarray(offset,offset+length));}catch{fail('글꼴 이름의 UTF-16 데이터가 잘못되었습니다.');}
    if(!value.trim()||/[\u0000-\u001F\u007F]/u.test(value))fail('글꼴 이름에 잘못된 문자가 있습니다.');
    const key=`${platform}/${encoding}/${language}`,pair=pairs.get(key)??{rank:(platform===3?0:2)+(language===0x409?0:1)};
    if(pair[id]&&pair[id]!==value)fail('같은 글꼴 이름 키에 다른 이름이 있습니다.');
    pair[id]=value;pairs.set(key,pair);
  }
  const selected=[...pairs.values()].filter(p=>p[1]&&p[2]).sort((a,b)=>a.rank-b.rank)[0];
  if(!selected)fail('Unicode 글꼴 family와 subfamily가 필요합니다.','FONT_FORMAT_UNSUPPORTED');
  return {family:selected[1],subfamily:selected[2]};
}

function pushRange(ranges,start,end){
  if(start>end)return;
  if(start<=0xDFFF&&end>=0xD800)fail('글꼴 cmap에 surrogate 코드포인트가 있습니다.');
  const previous=ranges.at(-1);
  if(previous&&previous[1]+1===start)previous[1]=end;
  else ranges.push([start,end]);
}
function coverage4(table,glyphCount){
  requireRange(table,0,16);
  const doubled=table.readUInt16BE(6),count=doubled/2,arrayEnd=16+count*8;
  if(!doubled||doubled%2)fail('cmap 4의 segment 수가 잘못되었습니다.');
  requireRange(table,0,arrayEnd);
  if(table.readUInt16BE(14+count*2)!==0)fail('cmap 4의 예약 필드가 잘못되었습니다.');
  const ranges=[];let last=-1;
  for(let i=0;i<count;i++){
    const end=table.readUInt16BE(14+i*2),start=table.readUInt16BE(16+count*2+i*2),delta=table.readInt16BE(16+count*4+i*2),pointer=16+count*6+i*2,offset=table.readUInt16BE(pointer);
    if(start>end||start<=last||offset%2)fail('cmap 4의 segment 범위가 잘못되었습니다.');
    if(i===count-1&&(start!==0xFFFF||end!==0xFFFF))fail('cmap 4의 마지막 segment가 없습니다.');
    if(offset){if(pointer+offset<arrayEnd)fail('cmap 4의 glyph 참조가 배열 앞을 가리킵니다.');requireRange(table,pointer+offset,(end-start+1)*2);}
    for(let cp=start;cp<=end;cp++){
      let glyph=offset?table.readUInt16BE(pointer+offset+(cp-start)*2):cp;
      if(!offset||glyph!==0)glyph=(glyph+delta)&0xFFFF;
      if(glyph>=glyphCount)fail('cmap이 존재하지 않는 glyph를 참조합니다.');
      if(glyph)pushRange(ranges,cp,cp);
    }
    last=end;
  }
  return ranges;
}
function coverage12(table,glyphCount){
  requireRange(table,0,16);
  if(table.readUInt16BE(2)!==0)fail('cmap 12의 예약 필드가 잘못되었습니다.');
  const count=table.readUInt32BE(12);requireRange(table,16,count*12);
  const ranges=[];let last=-1;
  for(let i=0;i<count;i++){
    const p=16+i*12,start=table.readUInt32BE(p),end=table.readUInt32BE(p+4),glyph=table.readUInt32BE(p+8);
    if(start>end||start<=last||end>0x10FFFF||glyph+end-start>=glyphCount)fail('cmap 12의 문자 또는 glyph 범위가 잘못되었습니다.');
    pushRange(ranges,start+(glyph===0?1:0),end);last=end;
  }
  return ranges;
}
function glyphCoverage(table,glyphCount){
  requireRange(table,0,4);
  if(table.readUInt16BE(0)!==0)fail('cmap 버전이 잘못되었습니다.');
  const count=table.readUInt16BE(2),headerEnd=4+count*8;requireRange(table,4,count*8);
  let selected=null;const seen=new Set();
  for(let i=0;i<count;i++){
    const p=4+i*8,platform=table.readUInt16BE(p),encoding=table.readUInt16BE(p+2),offset=table.readUInt32BE(p+4),key=`${platform}/${encoding}`;
    if(seen.has(key)||offset<headerEnd)fail('cmap encoding 또는 offset이 잘못되었습니다.');seen.add(key);requireRange(table,offset,2);
    const format=table.readUInt16BE(offset);
    if(!(platform===0||(platform===3&&[1,10].includes(encoding)))||![4,12].includes(format))continue;
    if(format===12?!(platform===0&&encoding===4||platform===3&&encoding===10):!(platform===0&&encoding<=3||platform===3&&encoding===1))fail('cmap의 encoding과 format이 일치하지 않습니다.');
    requireRange(table,offset,format===4?4:8);
    const length=format===4?table.readUInt16BE(offset+2):table.readUInt32BE(offset+4);requireRange(table,offset,length);
    const data=table.subarray(offset,offset+length),ranges=format===4?coverage4(data,glyphCount):coverage12(data,glyphCount),rank=(format===12?10:0)+(platform===3?1:0);
    if(!selected||rank>selected.rank)selected={rank,ranges};
  }
  if(!selected)fail('Unicode cmap 4 또는 12가 필요합니다.','FONT_FORMAT_UNSUPPORTED');
  return selected.ranges;
}

// Read only the identity, style, rights and repertoire needed for embedding decisions.
// Outline rendering and PowerPoint portability remain separate verification gates.
export function inspectFont(bytes){
  if(!(bytes instanceof Uint8Array)||bytes.length>MAX_FONT_BYTES)fail('글꼴은 16 MiB 이하의 bytes여야 합니다.');
  const data=Buffer.from(bytes.buffer,bytes.byteOffset,bytes.byteLength);requireRange(data,0,12);
  if(data.readUInt32BE(0)!==0x10000)fail('정적 TrueType sfnt만 지원합니다.','FONT_FORMAT_UNSUPPORTED');
  const count=data.readUInt16BE(4),directoryEnd=12+count*16;requireRange(data,12,count*16);
  if(!count)fail('글꼴 table이 없습니다.');
  const tables=new Map(),ranges=[];
  for(let i=0;i<count;i++){
    const p=12+i*16,tag=data.toString('ascii',p,p+4),offset=data.readUInt32BE(p+8),length=data.readUInt32BE(p+12);
    if(tables.has(tag)||offset<directoryEnd||offset%4)fail('글꼴 table 디렉터리가 잘못되었습니다.');
    requireRange(data,offset,length);tables.set(tag,data.subarray(offset,offset+length));
    if(length)ranges.push([offset,offset+length]);
  }
  ranges.sort((a,b)=>a[0]-b[0]);
  for(let i=1;i<ranges.length;i++)if(ranges[i][0]<ranges[i-1][1])fail('글꼴 table 범위가 겹칩니다.');
  if(tables.has('fvar'))fail('가변 글꼴은 아직 포함하지 않습니다.','FONT_VARIABLE_UNSUPPORTED');
  if(tables.has('CFF ')||tables.has('CFF2'))fail('CFF outline은 아직 지원하지 않습니다.','FONT_FORMAT_UNSUPPORTED');
  for(const tag of ['name','OS/2','cmap','head','maxp','loca','glyf'])if(!tables.has(tag))fail(`글꼴 필수 table이 없습니다: ${tag}`);
  const os2=tables.get('OS/2'),head=tables.get('head'),maxp=tables.get('maxp');
  requireRange(os2,0,2);const os2Version=os2.readUInt16BE(0);
  if(os2Version>5)fail('지원하지 않는 OS/2 버전입니다.','FONT_FORMAT_UNSUPPORTED');
  requireRange(os2,0,[68,86,96,96,96,100][os2Version]);requireRange(head,0,54);requireRange(maxp,0,32);
  if(head.readUInt32BE(0)!==0x10000||head.readUInt32BE(12)!==0x5F0F3CF5||maxp.readUInt32BE(0)!==0x10000)fail('head 또는 maxp 버전이 잘못되었습니다.');
  const weight=os2.readUInt16BE(4),fsType=os2.readUInt16BE(8),selection=os2.readUInt16BE(62),italic=Boolean(selection&1),bold=Boolean(selection&32),style=head.readUInt16BE(44),glyphCount=maxp.readUInt16BE(4),locaFormat=head.readInt16BE(50);
  if(![400,700].includes(weight)||(selection&512))fail('정적 400/700 normal/italic 스타일만 지원합니다.','FONT_STYLE_UNSUPPORTED');
  if(bold!==(weight===700)||bold!==Boolean(style&1)||italic!==Boolean(style&2)||((selection&64)&&(bold||italic)))fail('글꼴 스타일 metadata가 일치하지 않습니다.');
  if(!glyphCount||![0,1].includes(locaFormat))fail('glyph 수 또는 loca 형식이 잘못되었습니다.');
  requireRange(tables.get('loca'),0,(glyphCount+1)*(locaFormat===0?2:4));
  return {...names(tables.get('name')),weight,italic,os2Version,fsType,format:'ttf',variable:false,glyphCoverage:glyphCoverage(tables.get('cmap'),glyphCount)};
}
