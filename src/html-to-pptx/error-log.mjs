import fs from 'node:fs/promises';
import path from 'node:path';
import {fileURLToPath} from 'node:url';
import {publicDiagnostic} from './public-diagnostics.mjs';

const defaultDirectory=fileURLToPath(new URL('../../.local/logs/',import.meta.url));
const sourceUrl=new URL('./',import.meta.url).href;
const sourcePath=fileURLToPath(sourceUrl).replaceAll('\\','/');
const sourceFiles=new Set([
  'browser-runtime.mjs','cli.mjs','conversion-worker.mjs','export-html.mjs','export.mjs',
  'input.mjs','job-runner.mjs','job-store.mjs','measure.mjs','powerpoint-chart-images.mjs',
  'read-pptx.mjs','server.mjs'
]);
const operations=new Set(['analyze-html','analyze-pptx','cli','export-html','export-pptx','http','preview-pptx','read-pptx','server-start']);
const errorCodes=new Set([
  'ASSET_INVALID','ASSET_LIMIT','EXPORT_FAILED','INPUT_INVALID','INPUT_LIMIT',
  'JOB_BUSY','JOB_CANCELLED','JOB_CLEANUP_FAILED','JOB_FAILED','JOB_MEMORY_LIMIT','JOB_OBSERVER_FAILED',
  'JOB_STORAGE_LIMIT','JOB_STORE_CLOSED','JOB_TIMEOUT','OBSERVER_ACCESS_DENIED',
  'OBSERVER_IDENTITY_UNAVAILABLE','OBSERVER_QUERY_FAILED','OBSERVER_TIMEOUT','OUTPUT_EXISTS','USAGE',
  'EACCES','EADDRINUSE','EADDRNOTAVAIL','ECONNRESET','EEXIST','EISDIR','EMFILE','ENOENT','ENOSPC',
  'ENOTDIR','EPERM','EPIPE','ERR_CHILD_PROCESS_STDIO_MAXBUFFER','ERR_INVALID_ARG_VALUE',
  'ERR_IPC_CHANNEL_CLOSED','ERR_SOCKET_BAD_PORT'
]);
const errorNames=new Set(['Error','AbortError','AggregateError','RangeError','ReferenceError','SyntaxError','TypeError','URIError']);
const signals=new Set(['SIGABRT','SIGHUP','SIGINT','SIGKILL','SIGSEGV','SIGTERM']);
const phases=new Set(['initial','active','closing']);
const sourceTypes=new Set(['html','pptx']);
const logName=/^errors-(\d{4}-\d{2}-\d{2})(?:\.1)?\.log$/;
const uuid=/^[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i;
const nonNegativeInteger=value=>Number.isSafeInteger(value)&&value>=0;

function safeLocations(error){
  const locations=new Set();
  function add(value){
    if(typeof value!=='string')return;
    const match=/^src\/html-to-pptx\/([a-z-]+\.mjs):[1-9]\d{0,5}:[1-9]\d{0,5}$/.exec(value);
    if(match&&sourceFiles.has(match[1]))locations.add(value);
  }
  if(Array.isArray(error.locations))for(const value of error.locations.slice(0,6))add(value);
  const stack=error.stack;
  if(typeof stack==='string'){
    for(const line of stack.slice(-4096).split('\n')){
      if(!line.startsWith('    at '))continue;
      const normalized=line.replaceAll('\\','/');
      for(const prefix of [sourceUrl,sourcePath]){
        const index=normalized.lastIndexOf(prefix);
        if(index>0&&[' ','('].includes(normalized[index-1])){
          add('src/html-to-pptx/'+normalized.slice(index+prefix.length).replace(/\)\s*$/,''));
        }
      }
    }
  }
  return [...locations].slice(0,6);
}

// The same summary can cross IPC and be sanitized again before persistence.
export function safeError(error,depth=0){
  if(!error||typeof error!=='object')return undefined;
  const known=typeof error.code==='string'&&(errorCodes.has(error.code)||publicDiagnostic({code:error.code}).code===error.code);
  const record={code:known?error.code:'UNKNOWN_ERROR',name:errorNames.has(error.name)?error.name:'Error',locations:safeLocations(error)};
  const exitCode=error.exitCode??error.code;
  if(nonNegativeInteger(exitCode)&&exitCode<=0xffffffff)record.exitCode=exitCode;
  if(typeof error.killed==='boolean')record.killed=error.killed;
  if(signals.has(error.signal))record.signal=error.signal;
  if(phases.has(error.observerPhase))record.observerPhase=error.observerPhase;
  if(depth<3&&error.cause)record.cause=safeError(error.cause,depth+1);
  return record;
}

function safeDiagnostic(item,slides){
  const index=slides.findIndex(slide=>item.slideId!=null&&slide.id===item.slideId);
  const diagnostic=publicDiagnostic({code:item.code,severity:item.severity,reasonKey:item.reasonKey,slideNumber:index<0?item.slideNumber:index+1});
  const error=safeError(item.error);
  return error?{...diagnostic,error}:diagnostic;
}

function errorRecord(event,time){
  const record={
    time:time.toISOString(),
    operation:operations.has(event.operation)?event.operation:'http',
    nodeVersion:process.version,
    platform:process.platform
  };
  for(const key of ['requestId','jobId'])if(typeof event[key]==='string'&&uuid.test(event[key]))record[key]=event[key];
  for(const key of ['sourceType','targetType'])if(sourceTypes.has(event[key]))record[key]=event[key];
  if(nonNegativeInteger(event.inputBytes))record.inputBytes=event.inputBytes;
  if(Number.isInteger(event.statusCode)&&event.statusCode>=100&&event.statusCode<=599)record.statusCode=event.statusCode;
  const usage={};
  for(const key of ['elapsedMs','rssSamples','peakRssBytes','terminationMs']){
    if(nonNegativeInteger(event.resourceUsage?.[key]))usage[key]=event.resourceUsage[key];
  }
  if(Object.keys(usage).length)record.resourceUsage=usage;
  const error=safeError(event.error);
  if(error)record.error=error;
  const diagnostics=Array.isArray(event.diagnostics)?event.diagnostics:[];
  const slides=Array.isArray(event.slides)?event.slides:[];
  record.diagnosticCount=diagnostics.length;
  record.diagnostics=diagnostics.slice(0,20).filter(item=>item&&typeof item==='object').map(item=>safeDiagnostic(item,slides));
  return record;
}

export function createErrorLogger({directory=defaultDirectory,maxBytes=1024*1024,retentionDays=7,now=()=>new Date()}={}){
  let queue=Promise.resolve();
  return event=>{
    queue=queue.then(async()=>{
      const time=now();
      const day=time.toISOString().slice(0,10);
      const cutoff=new Date(time.getTime()-(retentionDays-1)*86400000).toISOString().slice(0,10);
      const line=JSON.stringify(errorRecord(event,time))+'\n';
      const bytes=Buffer.byteLength(line);
      if(bytes>maxBytes)return false;

      await fs.mkdir(directory,{recursive:true,mode:0o700});
      for(const name of await fs.readdir(directory)){
        const date=logName.exec(name)?.[1];
        if(date&&date<cutoff)await fs.unlink(path.join(directory,name)).catch(error=>{if(error.code!=='ENOENT')throw error;});
      }
      const file=path.join(directory,`errors-${day}.log`);
      const stat=await fs.lstat(file).catch(error=>{if(error.code!=='ENOENT')throw error;return null;});
      if(stat&&!stat.isFile())return false;
      if(stat&&stat.size+bytes>maxBytes)await fs.rename(file,path.join(directory,`errors-${day}.1.log`));
      await fs.appendFile(file,line,{encoding:'utf8',mode:0o600});
      return true;
    }).catch(()=>false);
    return queue;
  };
}

export const writeErrorLog=createErrorLogger();
