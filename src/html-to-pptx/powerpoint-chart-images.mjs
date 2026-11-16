import {execFile} from 'node:child_process';
import {promisify} from 'node:util';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import {fileURLToPath} from 'node:url';

const execute=promisify(execFile),script=fileURLToPath(new URL('./powerpoint-chart-export.ps1',import.meta.url));
const failure=code=>Object.assign(new Error(code),{code});
const psLiteral=value=>`'${value.replaceAll("'","''")}'`;

// Only called with a fully checked package and chart identifiers from its XML.
// Office reads a private copy; neither the uploaded file nor the user's open
// presentation is saved or modified.
export async function renderPowerPointChartImages(bytes,charts){
 if(process.platform!=='win32')throw failure('POWERPOINT_UNAVAILABLE');
 if(!(bytes instanceof Uint8Array)||!Array.isArray(charts)||!charts.length||charts.length>100||charts.some(c=>!Number.isSafeInteger(c.slideNumber)||c.slideNumber<1||c.slideNumber>200||!Number.isSafeInteger(c.shapeId)||c.shapeId<1||c.shapeId>2147483647))throw failure('POWERPOINT_EXPORT_FAILED');
 const dir=await fs.mkdtemp(path.join(os.tmpdir(),'ppt-again-chart-'));
 const input=path.join(dir,'source.pptx'),request=path.join(dir,'request.json'),resultPath=path.join(dir,'result.json'),outputs=charts.map((_,i)=>path.join(dir,`chart-${i+1}.png`));
 try{
  await fs.writeFile(input,bytes,{flag:'wx'});await fs.writeFile(request,JSON.stringify(charts),{flag:'wx'});
  const source=await fs.readFile(script,'utf8');
  const command=`& {\n${source}\n} -InputPath ${psLiteral(input)} -RequestPath ${psLiteral(request)} -OutputDir ${psLiteral(dir)}`;
  try{await execute('powershell.exe',['-NoProfile','-NonInteractive','-EncodedCommand',Buffer.from(command,'utf16le').toString('base64')],{windowsHide:true,timeout:120000,maxBuffer:64*1024});}
  catch(error){const stderr=String(error.stderr??'');throw failure(error.code==='ENOENT'||stderr.includes('PPT_AGAIN_POWERPOINT_UNAVAILABLE')?'POWERPOINT_UNAVAILABLE':stderr.includes('PPT_AGAIN_POWERPOINT_OPEN_FAILED')?'POWERPOINT_OPEN_FAILED':stderr.includes('PPT_AGAIN_POWERPOINT_SECURITY_FAILED')?'POWERPOINT_SECURITY_FAILED':error.killed?'POWERPOINT_TIMEOUT':'POWERPOINT_EXPORT_FAILED');}
  const records=JSON.parse(await fs.readFile(resultPath,'utf8'));if(!Array.isArray(records)||records.length!==charts.length)throw failure('POWERPOINT_EXPORT_FAILED');
  return await Promise.all(charts.map(async (chart,i)=>{const record=records[i];if(record?.slideNumber!==chart.slideNumber||record?.shapeId!==chart.shapeId)throw failure('POWERPOINT_EXPORT_FAILED');if(record.errorCode)return {...chart,errorCode:record.errorCode};try{const stat=await fs.stat(outputs[i]);if(stat.size<1||stat.size>16*1024*1024)throw new Error();return {...chart,bytes:new Uint8Array(await fs.readFile(outputs[i]))};}catch{return {...chart,errorCode:'POWERPOINT_EXPORT_FAILED'};}}));
 }finally{
  for(const file of [input,request,resultPath,...outputs])await fs.unlink(file).catch(()=>{});
  await fs.rmdir(dir).catch(()=>{});
 }
}
