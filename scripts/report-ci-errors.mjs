import {readdir,readFile} from 'node:fs/promises';
import {safeError} from '../src/html-to-pptx/error-log.mjs';
import {publicDiagnostic} from '../src/html-to-pptx/public-diagnostics.mjs';

// Print only allowlisted fields from the existing sanitized logs. Never print
// raw log lines: malformed input could contain paths or document contents.
const directory=new URL('../.local/logs/',import.meta.url);
let names=[];
try{names=(await readdir(directory)).filter(name=>/^errors-\d{4}-\d{2}-\d{2}(?:\.1)?\.log$/.test(name)).sort();}catch{}
const records=[];
for(const name of names){
  let contents;try{contents=await readFile(new URL(name,directory),'utf8');}catch{continue;}
  for(const line of contents.split('\n')){
    let row;try{row=JSON.parse(line);}catch{continue;}
    if(!Number.isInteger(row?.statusCode)||row.statusCode<400||row.statusCode>599)continue;
    const resourceUsage=Object.fromEntries(['elapsedMs','rssSamples','peakRssBytes','terminationMs'].filter(key=>Number.isSafeInteger(row.resourceUsage?.[key])&&row.resourceUsage[key]>=0).map(key=>[key,row.resourceUsage[key]]));
    const diagnostics=(Array.isArray(row.diagnostics)?row.diagnostics:[]).slice(0,20).map(item=>{
      const value=publicDiagnostic({code:item?.code,severity:item?.severity,reasonKey:item?.reasonKey});
      return {code:value.code,severity:value.severity,...(value.reasonKey?{reasonKey:value.reasonKey}:{}),...(safeError(item?.error)?{error:safeError(item.error)}:{})};
    });
    records.push({statusCode:row.statusCode,...(safeError(row.error)?{error:safeError(row.error)}:{}),resourceUsage,diagnostics});
  }
}
const critical=records.filter(row=>row.statusCode>=500||row.diagnostics.some(item=>item.code==='RENDER_INCOMPLETE'));
console.log(JSON.stringify({message:'Sanitized CI error summaries',critical:critical.slice(-100),recent:records.slice(-30)}));
