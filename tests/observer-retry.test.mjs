import test from 'node:test';
import assert from 'node:assert/strict';
import childProcess from 'node:child_process';
import {syncBuiltinESMExports} from 'node:module';
import {promisify} from 'node:util';

test('initial Windows observer retries a timeout but still fails closed if it remains unavailable', {skip: process.platform!=='win32'}, async()=>{
  const real=childProcess.execFile;
  let failures=1;
  childProcess.execFile=function(file,args,options,callback){
    if(file==='powershell.exe'&&args.at(-1).startsWith('Get-CimInstance Win32_Process |')&&failures>0){failures--;const error=Object.assign(new Error('simulated timeout'),{killed:true,signal:'SIGTERM'});queueMicrotask(()=>callback(error,'',''));return {};}
    return real.call(this,file,args,file==='powershell.exe'?{...options,timeout:15000}:options,callback);
  };
  childProcess.execFile[promisify.custom]=(file,args,options)=>new Promise((resolve,reject)=>childProcess.execFile(file,args,options,(error,stdout,stderr)=>error?reject(error):resolve({stdout,stderr})));
  syncBuiltinESMExports();
  try{
    const {runConversion}=await import('../src/html-to-pptx/job-runner.mjs');
    const input=new Uint8Array([0x50,0x4b]);
    const result=await runConversion({operation:'read-pptx',input});
    assert.ok(result.resourceUsage.rssSamples>0);
    failures=10;
    await assert.rejects(runConversion({operation:'read-pptx',input}),error=>error.code==='JOB_OBSERVER_FAILED'&&error.resourceUsage.browserPid===null&&error.resourceUsage.rssSamples===0);
  }finally{childProcess.execFile=real;syncBuiltinESMExports();}
});
