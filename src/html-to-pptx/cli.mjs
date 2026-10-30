#!/usr/bin/env node
import fs from 'node:fs/promises';
import path from 'node:path';
import {parseArgs} from 'node:util';
import {runConversion} from './job-runner.mjs';
import {publicDiagnostic} from './public-diagnostics.mjs';
import {assetPath} from './assets.mjs';

const MiB=1024*1024;
const help=`Usage: slidehandoff convert <input.html|input.pptx> --output <output.pptx|output.html> [--selector <css>] [--assets <directory>] [--json]\n       slidehandoff --help | --version`;
const fail=(code,message,exit=1)=>Object.assign(new Error(message),{code,exit});
async function readFileLimited(file,max){
  const handle=await fs.open(file,'r');
  try{const stat=await handle.stat();if(!stat.isFile()||stat.size>max)throw fail('INPUT_LIMIT','입력 파일의 크기 또는 형식을 확인해 주세요.');const chunks=[];let total=0;while(true){const chunk=Buffer.allocUnsafe(Math.min(65536,max+1-total));const {bytesRead}=await handle.read(chunk,0,chunk.length,null);if(!bytesRead)break;total+=bytesRead;if(total>max)throw fail('INPUT_LIMIT','입력 파일이 너무 큽니다.');chunks.push(chunk.subarray(0,bytesRead));}return Buffer.concat(chunks,total);}finally{await handle.close();}
}
async function readAssets(dir){
  const root=await fs.realpath(dir),rootStat=await fs.lstat(dir);if(rootStat.isSymbolicLink()||!rootStat.isDirectory())throw fail('ASSET_INVALID','자산 폴더를 확인해 주세요.');
  const assets=[],names=new Set();let total=0,entryCount=0;
  async function visit(folder,relative=''){
    const entries=[];for await(const entry of await fs.opendir(folder)){if(++entryCount>1024)throw fail('ASSET_LIMIT','자산 폴더는 최대 1024개 항목까지 사용할 수 있습니다.');entries.push(entry);}entries.sort((a,b)=>a.name.localeCompare(b.name));
    for(const entry of entries){
      const rel=relative?`${relative}/${entry.name}`:entry.name;const normalized=assetPath(rel);if(normalized!==rel||names.has(normalized))throw fail('ASSET_INVALID','자산 경로가 올바르지 않습니다.');names.add(normalized);
      const file=path.join(folder,entry.name),stat=await fs.lstat(file);if(stat.isSymbolicLink())throw fail('ASSET_INVALID','자산의 심볼릭 링크는 사용할 수 없습니다.');
      const real=await fs.realpath(file),inside=path.relative(root,real);if(inside.startsWith('..'+path.sep)||inside==='..'||path.isAbsolute(inside))throw fail('ASSET_INVALID','자산 경로가 폴더를 벗어났습니다.');
      if(stat.isDirectory()){await visit(file,rel);continue;}
      if(!stat.isFile()||stat.size>16*MiB||(total+=stat.size)>64*MiB)throw fail('ASSET_INVALID','자산 형식 또는 크기를 확인해 주세요.');
      const bytes=await readFileLimited(file,16*MiB);total+=bytes.length-stat.size;if(total>64*MiB)throw fail('ASSET_LIMIT','자산의 총 크기가 64MiB를 초과합니다.');assets.push({path:rel,bytes});
    }
  }
  await visit(root);return assets;
}
async function main(){
 let values,positionals;try{({values,positionals}=parseArgs({allowPositionals:true,options:{output:{type:'string'},selector:{type:'string'},assets:{type:'string'},json:{type:'boolean'},help:{type:'boolean'},version:{type:'boolean'}}}));}catch{throw fail('USAGE',help,2);}
 if(values.help&&positionals.length===0){console.log(help);return;}
 if(values.version&&positionals.length===0){const pkg=JSON.parse(await fs.readFile(new URL('../../package.json',import.meta.url),'utf8'));console.log(pkg.version);return;}
 if(positionals.length!==2||positionals[0]!=='convert'||!values.output||values.help||values.version)throw fail('USAGE',help,2);
 const inputPath=positionals[1],output=values.output,ext=path.extname(inputPath).toLowerCase(),outExt=path.extname(output).toLowerCase();
 const sourceType=ext==='.pptx'?'pptx':ext==='.html'||ext==='.htm'?'html':null,targetType=sourceType==='html'?'pptx':'html';
 if(!sourceType||outExt!==`.${targetType}`||(sourceType==='pptx'&&(values.selector!==undefined||values.assets!==undefined))||values.selector!==undefined&&!values.selector.trim())throw fail('USAGE',help,2);
 try{await fs.lstat(output);throw fail('OUTPUT_EXISTS','출력 파일이 이미 존재합니다. 다른 경로를 지정해 주세요.');}catch(e){if(e.code!=='ENOENT')throw e;}
 let bytes;try{bytes=await readFileLimited(inputPath,(sourceType==='html'?16:64)*MiB);}catch(e){if(e.code==='ENOENT')throw fail('USAGE','입력 파일을 찾을 수 없습니다.',2);throw e;}
 const input=sourceType==='html'?{html:bytes.toString('utf8'),assets:values.assets?await readAssets(values.assets):[],slideSelector:values.selector??null}:bytes;
 const controller=new AbortController(),abort=()=>controller.abort();process.on('SIGINT',abort);process.on('SIGTERM',abort);
 try{
  const {measured}=await runConversion({operation:sourceType==='html'?'analyze-html':'analyze-pptx',input,signal:controller.signal});
  let diagnostics=measured.diagnostics??[];
  if(diagnostics.some(d=>d.severity==='error'))throw Object.assign(fail(diagnostics.find(d=>d.severity==='error').code,'변환할 수 없는 내용이 있습니다.'),{diagnostics,slides:measured.model?.slides});
  const {previews,...data}=measured;
  const result=await runConversion({operation:targetType==='pptx'?'export-pptx':'export-html',input:targetType==='pptx'?data:data.model,signal:controller.signal});
  diagnostics=[...new Map([...diagnostics,...(result.diagnostics??[])].map(d=>[JSON.stringify(d),d])).values()];
  if(result.status!=='generated'||diagnostics.some(d=>d.severity==='error'))throw Object.assign(fail(diagnostics.find(d=>d.severity==='error')?.code??'EXPORT_FAILED','변환 파일을 만들지 못했습니다.'),{diagnostics,slides:data.model?.slides});
  if(controller.signal.aborted)throw fail('JOB_CANCELLED','변환이 취소됐습니다.');
  let handle;try{handle=await fs.open(output,'wx');await handle.writeFile(targetType==='pptx'?result.pptxBytes:result.html);await handle.close();handle=null;}catch(e){if(handle){await handle.close().catch(()=>{});await fs.unlink(output).catch(()=>{});}if(e.code==='EEXIST')throw fail('OUTPUT_EXISTS','출력 파일이 이미 존재합니다.');throw e;}
  const report={status:'generated',output,sourceType,targetType,diagnostics:diagnostics.map(d=>publicDiagnostic(d,data.model?.slides))};
  console.log(values.json?JSON.stringify(report):`변환 완료: ${output}${report.diagnostics.length?'\n'+report.diagnostics.map(d=>d.message).join('\n'):''}`);
 }finally{process.off('SIGINT',abort);process.off('SIGTERM',abort);}
}
main().catch(e=>{const report={status:'failed',output:null,code:e.code??'IO_ERROR',diagnostics:(e.diagnostics??[]).map(d=>publicDiagnostic(d,e.slides))};if(process.argv.includes('--json'))console.log(JSON.stringify(report));else console.error(e.code&&e.exit?e.message:'변환에 실패했습니다. 입력과 출력 경로를 확인해 주세요.');process.exitCode=e.exit??1;});
