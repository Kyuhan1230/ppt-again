import {readdir,readFile} from 'node:fs/promises';
import {pathToFileURL} from 'node:url';
import {safeError,safeLogOperation} from '../src/html-to-pptx/error-log.mjs';
import {publicDiagnostic} from '../src/html-to-pptx/public-diagnostics.mjs';

// Print only allowlisted fields from the existing sanitized logs. Never print
// raw log lines: malformed input could contain paths or document contents.
const defaultDirectory=new URL('../.local/logs/',import.meta.url);

function summarize(row) {
  const statusCode=row?.statusCode;
  const cliExitCode=row?.cliExitCode;
  const httpFailure=Number.isInteger(statusCode)&&statusCode>=400&&statusCode<=599;
  const cliFailure=cliExitCode===1||cliExitCode===2;
  if(!httpFailure&&!cliFailure)return null;

  const operation=safeLogOperation(row.operation);
  const validTime=typeof row.time==='string'&&/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{3}Z$/.test(row.time);
  const resourceUsage=Object.fromEntries(['elapsedMs','rssSamples','peakRssBytes','terminationMs']
    .filter(key=>Number.isSafeInteger(row.resourceUsage?.[key])&&row.resourceUsage[key]>=0)
    .map(key=>[key,row.resourceUsage[key]]));
  const diagnostics=(Array.isArray(row.diagnostics)?row.diagnostics:[]).slice(0,20).map(item=>{
    const value=publicDiagnostic({code:item?.code,severity:item?.severity,reasonKey:item?.reasonKey});
    const error=safeError(item?.error);
    return {code:value.code,severity:value.severity,...(value.reasonKey?{reasonKey:value.reasonKey}:{}),...(error?{error}:{})};
  });
  const error=safeError(row.error);
  return {
    ...(validTime?{time:row.time}:{}),...(operation?{operation}:{}),
    ...(httpFailure?{statusCode}:{}),...(cliFailure?{cliExitCode}:{}),
    ...(error?{error}:{}),resourceUsage,diagnostics
  };
}

export async function readCiErrorSummaries(directory=defaultDirectory) {
  let names=[];
  try{names=(await readdir(directory)).filter(name=>/^errors-\d{4}-\d{2}-\d{2}(?:\.1)?\.log$/.test(name)).sort();}catch{}
  const records=[];
  for(const name of names){
    let contents;
    try{contents=await readFile(new URL(name,directory),'utf8');}catch{continue;}
    for(const line of contents.split('\n')){
      let row;
      try{row=JSON.parse(line);}catch{continue;}
      const record=summarize(row);
      if(record)records.push(record);
    }
  }
  records.sort((a,b)=>(a.time??'').localeCompare(b.time??''));
  const critical=records.filter(row=>row.statusCode>=500||row.cliExitCode===1||row.diagnostics.some(item=>item.code==='RENDER_INCOMPLETE'));
  return {critical:critical.slice(-100),recent:records.slice(-30)};
}

if(process.argv[1]&&pathToFileURL(process.argv[1]).href===import.meta.url){
  console.log(JSON.stringify({message:'Sanitized CI error summaries',...await readCiErrorSummaries()}));
}
