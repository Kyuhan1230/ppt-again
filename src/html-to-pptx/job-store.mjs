import {randomUUID} from 'node:crypto';

const reserved=new Set(['id','ownerId','createdAt','expiresAt']);
const failure=(code,statusCode)=>Object.assign(new Error(code),{code,statusCode});

export function createJobStore({ttlMs=1800000,maxJobs=10,maxBytes=268435456,onDispose=async()=>{}}={}){
  for(const value of [ttlMs,maxJobs,maxBytes])if(!Number.isSafeInteger(value)||value<=0)throw new TypeError('Invalid job limit');
  const jobs=new Map(),pending=new Set();let bytes=0,closed=false;
  // Counts retained document payloads; process/browser RSS is a separate C03.2 limit.
  function sizeOf(value){
    const seen=new WeakSet(),stack=[value];let size=0;
    while(stack.length&&size<=maxBytes){
      const item=stack.pop();
      if(typeof item==='string'){size+=Buffer.byteLength(item);continue;}
      if(!item||typeof item!=='object'||seen.has(item))continue;
      seen.add(item);
      if(ArrayBuffer.isView(item)||item instanceof ArrayBuffer){size+=item.byteLength;continue;}
      if(item instanceof Promise)continue;
      for(const [key,child]of Object.entries(item))if(!reserved.has(key))stack.push(child);
    }
    return size;
  }
  function dispose(record,reason){
    jobs.delete(record.job.id);clearTimeout(record.timer);
    const promise=Promise.resolve().then(()=>onDispose(record.job,reason)).finally(()=>{
      bytes-=record.bytes;
      for(const key of Object.keys(record.job))if(!reserved.has(key))delete record.job[key];
    });
    pending.add(promise);
    promise.then(()=>pending.delete(promise),()=>pending.delete(promise));
    return promise;
  }
  function find(ownerId,id){
    const record=jobs.get(id);
    if(!record)return null;
    if(record.job.expiresAt<=Date.now()){void dispose(record,'expired');return null;}
    return record.job.ownerId===ownerId?record:null;
  }
  return {
    get size(){return jobs.size;},
    get byteLength(){return bytes;},
    create(ownerId,payload){
      if(closed)throw failure('JOB_STORE_CLOSED',503);
      if(typeof ownerId!=='string'||!ownerId)throw new TypeError('Owner required');
      const count=sizeOf(payload);
      // Reject pressure instead of silently deleting another retained document.
      if(jobs.size>=maxJobs||bytes+count>maxBytes)throw failure('JOB_STORAGE_LIMIT',count>maxBytes?413:429);
      const now=Date.now(),job={...payload,id:randomUUID(),ownerId,createdAt:now,expiresAt:now+ttlMs};
      const record={job,bytes:count,timer:null};
      record.timer=setTimeout(()=>{if(jobs.get(job.id)===record)void dispose(record,'expired');},ttlMs);
      record.timer.unref();jobs.set(job.id,record);bytes+=count;return job.id;
    },
    get(ownerId,id){return find(ownerId,id)?.job??null;},
    update(ownerId,id,patch){
      const record=find(ownerId,id);if(!record)return false;
      if(Object.keys(patch).some(key=>reserved.has(key)))throw new TypeError('Immutable job metadata');
      const count=sizeOf({...record.job,...patch});
      if(bytes-record.bytes+count>maxBytes)throw failure('JOB_STORAGE_LIMIT',count>maxBytes?413:429);
      Object.assign(record.job,patch);bytes+=count-record.bytes;record.bytes=count;return true;
    },
    async delete(ownerId,id){const record=find(ownerId,id);if(!record)return false;await dispose(record,'deleted');return true;},
    async close(){closed=true;for(const record of jobs.values())void dispose(record,'shutdown');await Promise.all([...pending]);}
  };
}
