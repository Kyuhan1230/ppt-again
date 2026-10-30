import {inflateRawSync,crc32} from 'node:zlib';

const reject=(message,code='INPUT_INVALID')=>{throw Object.assign(new Error(message),{code});};
export function createPackageBudget({maxBytes=128*1024*1024,maxEntries=10000,maxMemberBytes=16*1024*1024}={}){
  if(![maxBytes,maxEntries,maxMemberBytes].every(v=>Number.isSafeInteger(v)&&v>0))throw new TypeError('Invalid package limits');
  let bytes=0,entries=0;
  return {
    charge(amount){
      if(!amount||![amount.bytes,amount.entries].every(v=>Number.isSafeInteger(v)&&v>=0))throw new TypeError('Invalid package charge');
      if(amount.bytes>maxMemberBytes||bytes+amount.bytes>maxBytes||entries+amount.entries>maxEntries)reject('문서와 내장 데이터의 압축 해제 한도를 초과했습니다.','LIMIT_EXCEEDED');
      bytes+=amount.bytes;entries+=amount.entries;
    },
    snapshot:()=>({bytes,entries,maxBytes,maxEntries,maxMemberBytes})
  };
}

// Only ordinary, single-disk stored/deflated Office ZIPs are accepted. Both
// parent PPTX parts and nested XLSX parts charge the same budget before inflate.
export function unzipBudgeted(bytes,{budget}={}){
  if(!(bytes instanceof Uint8Array)||bytes.length<22||!budget?.charge)reject('ZIP 바이트 또는 공유 예산이 없습니다.');
  const data=Buffer.from(bytes.buffer,bytes.byteOffset,bytes.byteLength),length=data.length;
  const within=(at,size,end=length)=>Number.isSafeInteger(at)&&at>=0&&size>=0&&at+size<=end;
  let end=length-22;
  for(;end>=Math.max(0,length-65557);end--)if(data.readUInt32LE(end)===0x06054b50&&end+22+data.readUInt16LE(end+20)===length)break;
  if(end<Math.max(0,length-65557))reject('ZIP 중앙 디렉터리가 없습니다.');
  const count=data.readUInt16LE(end+10),centralSize=data.readUInt32LE(end+12),central=data.readUInt32LE(end+16);
  if(data.readUInt16LE(end+4)||data.readUInt16LE(end+6)||data.readUInt16LE(end+8)!==count||count===0xffff||central===0xffffffff||centralSize===0xffffffff||end>=20&&data.readUInt32LE(end-20)===0x07064b50)reject('분할 ZIP 또는 ZIP64는 지원하지 않습니다.','UNSUPPORTED_OBJECT');
  if(central+centralSize!==end||!within(central,centralSize,end))reject('ZIP 중앙 디렉터리 범위가 올바르지 않습니다.');
  budget.charge({bytes:0,entries:count});
  const records=[],names=new Set(),nameKeys=new Set();let cursor=central;
  const extra=(start,size)=>{let p=start;while(p<start+size){if(!within(p,4,start+size))reject('ZIP 확장 필드가 잘렸습니다.');const id=data.readUInt16LE(p),n=data.readUInt16LE(p+2);if(!within(p+4,n,start+size))reject('ZIP 확장 필드가 잘렸습니다.');if(id===1)reject('ZIP64는 지원하지 않습니다.','UNSUPPORTED_OBJECT');p+=4+n;}};
  for(let i=0;i<count;i++){
    if(!within(cursor,46,end)||data.readUInt32LE(cursor)!==0x02014b50)reject('잘못된 ZIP 디렉터리 항목입니다.');
    const flags=data.readUInt16LE(cursor+8),method=data.readUInt16LE(cursor+10),crc=data.readUInt32LE(cursor+16),packed=data.readUInt32LE(cursor+20),expanded=data.readUInt32LE(cursor+24),nameLength=data.readUInt16LE(cursor+28),extraLength=data.readUInt16LE(cursor+30),commentLength=data.readUInt16LE(cursor+32),offset=data.readUInt32LE(cursor+42);
    const next=cursor+46+nameLength+extraLength+commentLength;
    if(next>end||data.readUInt16LE(cursor+34)||[packed,expanded,offset].includes(0xffffffff))reject('잘못된 ZIP 항목 범위입니다.');
    if(flags&~0x080e||![0,8].includes(method)||(data.readUInt32LE(cursor+38)>>>16&0xf000)===0xa000)reject('암호화·링크·지원하지 않는 압축 형식입니다.','UNSUPPORTED_OBJECT');
    const nameBytes=data.subarray(cursor+46,cursor+46+nameLength);let name;
    try{if(!(flags&0x800)&&nameBytes.some(byte=>byte>127))reject('UTF-8이 아닌 ZIP 경로는 지원하지 않습니다.','UNSUPPORTED_OBJECT');name=new TextDecoder('utf-8',{fatal:true}).decode(nameBytes);}catch(error){if(error.code)throw error;reject('잘못된 ZIP 경로 문자입니다.');}
    const segments=name.endsWith('/')?name.slice(0,-1).split('/'):name.split('/');
    if(!name||/[\\\u0000-\u001f:%?#]/.test(name)||segments.some(part=>!part||part==='.'||part==='..')||names.has(name)||nameKeys.has(name.toLowerCase()))reject('ZIP 내부 경로가 잘못되었거나 중복됩니다.');
    names.add(name);nameKeys.add(name.toLowerCase());extra(cursor+46+nameLength,extraLength);
    if(method===0&&packed!==expanded||name.endsWith('/')&&expanded!==0)reject('ZIP 항목 크기가 일치하지 않습니다.');
    budget.charge({bytes:Math.max(packed,expanded),entries:0});
    if(!within(offset,30,central)||data.readUInt32LE(offset)!==0x04034b50||data.readUInt16LE(offset+6)!==flags||data.readUInt16LE(offset+8)!==method)reject('ZIP 로컬 헤더가 디렉터리와 다릅니다.');
    const localNameLength=data.readUInt16LE(offset+26),localExtraLength=data.readUInt16LE(offset+28),start=offset+30+localNameLength+localExtraLength;
    if(!within(offset+30,localNameLength+localExtraLength,central)||!data.subarray(offset+30,offset+30+localNameLength).equals(nameBytes)||!within(start,packed,central))reject('ZIP 로컬 경로 또는 바이트 범위가 다릅니다.');
    extra(offset+30+localNameLength,localExtraLength);
    for(const [field,expected] of [[14,crc],[18,packed],[22,expanded]]){const value=data.readUInt32LE(offset+field);if(value!==expected&&(!(flags&8)||value!==0))reject('ZIP 로컬 크기 또는 CRC가 다릅니다.');}
    let finish=start+packed;
    if(flags&8){
      if(!within(finish,12,central))reject('ZIP 데이터 기술자가 없습니다.');
      const at=data.readUInt32LE(finish)===0x08074b50?finish+4:finish;
      if(!within(at,12,central)||data.readUInt32LE(at)!==crc||data.readUInt32LE(at+4)!==packed||data.readUInt32LE(at+8)!==expanded)reject('ZIP 데이터 기술자가 다릅니다.');finish=at+12;
    }
    records.push({name,method,crc,packed,expanded,offset,start,finish});cursor=next;
  }
  if(cursor!==end)reject('ZIP 중앙 디렉터리에 확인되지 않은 항목이 있습니다.');
  let previous=0;for(const record of [...records].sort((a,b)=>a.offset-b.offset)){if(record.offset!==previous)reject('ZIP 항목 범위가 겹치거나 숨은 바이트가 있습니다.');previous=record.finish;}
  if(previous!==central)reject('ZIP 데이터 범위가 중앙 디렉터리와 다릅니다.');
  const parts=Object.create(null);
  for(const record of records){
    const compressed=data.subarray(record.start,record.start+record.packed);let decoded;
    if(record.method===0)decoded=Buffer.from(compressed);
    else{
      try{const inflated=inflateRawSync(compressed,{maxOutputLength:Math.max(1,record.expanded),info:true});if(inflated.engine.bytesWritten!==record.packed)reject('ZIP 압축 스트림 뒤에 추가 바이트가 있습니다.');decoded=inflated.buffer;}
      catch(error){if(error.code==='INPUT_INVALID')throw error;reject('ZIP 실제 해제 크기 또는 압축 스트림이 선언과 다릅니다.');}
    }
    if(decoded.length!==record.expanded||crc32(decoded)!==record.crc)reject('ZIP 실제 길이 또는 CRC가 선언과 다릅니다.');
    parts[record.name]=new Uint8Array(decoded.buffer,decoded.byteOffset,decoded.byteLength);
  }
  return parts;
}
