import {fork,execFile} from 'node:child_process';
import {promisify} from 'node:util';
import fs from 'node:fs/promises';
const execute=promisify(execFile),MiB=1024*1024;
const operations=new Set(['analyze-html','analyze-pptx','export-pptx','read-pptx','export-html']);
// The initial 1GiB observation threshold rejected the unchanged 8/84-slide
// inputs. The measured Windows worker+browser budget is 2GiB; sampling can
// briefly overshoot it and this is not an OS memory cap.
export const conversionLimits=Object.freeze({timeoutMs:120000,oldSpaceMb:512,maxRssBytes:2048*MiB,rssIntervalMs:1000,closeGraceMs:1000});
let active=0;
const failure=(code,statusCode=400)=>Object.assign(new Error(code),{code,statusCode});

// These checks run before fork. No input can select a program, handler or limits.
function validateInput(operation,input){
  if(!operations.has(operation))throw failure('INPUT_INVALID');
  if(operation==='analyze-html'){
    if(!input||typeof input.html!=='string'||!Array.isArray(input.assets??[]))throw failure('INPUT_INVALID');
    if(Buffer.byteLength(input.html)>16*MiB)throw failure('INPUT_LIMIT',413);
    let total=0;
    for(const asset of input.assets??[]){
      if(!asset||!(asset.bytes instanceof Uint8Array))throw failure('INPUT_INVALID');
      if(asset.bytes.byteLength>16*MiB||(total+=asset.bytes.byteLength)>64*MiB)throw failure('INPUT_LIMIT',413);
    }
  }else if(operation==='read-pptx'||operation==='analyze-pptx'){
    if(!(input instanceof Uint8Array))throw failure('INPUT_INVALID');
    if(input.byteLength>64*MiB)throw failure('INPUT_LIMIT',413);
    if(operation==='analyze-pptx'&&(input.byteLength<22||input[0]!==0x50||input[1]!==0x4b||input[2]!==3||input[3]!==4))throw failure('INPUT_INVALID');
  }else{
    const model=operation==='export-pptx'?input?.model:input;
    if(!model||!Array.isArray(model.slides))throw failure('INPUT_INVALID');
    if(model.slides.length>200||model.slides.some(slide=>slide.nodes?.length>10000)||model.slides.reduce((n,slide)=>n+(slide.nodes?.length??0),0)>100000)throw failure('INPUT_LIMIT',413);
    let total=0;for(const asset of model.assets??[]){if(!(asset.bytes instanceof Uint8Array))throw failure('INPUT_INVALID');if(asset.bytes.length>16*MiB||(total+=asset.bytes.length)>64*MiB)throw failure('INPUT_LIMIT',413);}
  }
}

async function processTree(owned,initialPid=null){
  let rows;
  if(process.platform==='win32'){
    // Only numeric metadata leaves this observer; document contents and command lines do not.
    const {stdout}=await execute('powershell.exe',['-NoProfile','-NonInteractive','-Command',"Get-CimInstance Win32_Process | Select-Object @{n='pid';e={[int]$_.ProcessId}},@{n='parent';e={[int]$_.ParentProcessId}},@{n='rss';e={[long]$_.WorkingSetSize}},@{n='created';e={$_.CreationDate.ToFileTimeUtc().ToString()}} | ConvertTo-Json -Compress"],{windowsHide:true,timeout:5000,maxBuffer:1024*1024});
    rows=JSON.parse(stdout);if(!Array.isArray(rows))rows=[rows];
  }else{
    const {stdout}=await execute('ps',['-eo','pid=,ppid=,rss='],{timeout:5000,maxBuffer:1024*1024});
    rows=await Promise.all(stdout.trim().split('\n').map(async line=>{const [pid,parent,rss]=line.trim().split(/\s+/).map(Number);const stat=await fs.readFile(`/proc/${pid}/stat`,'utf8').catch(()=>null);return {pid,parent,rss:rss*1024,created:stat?.slice(stat.lastIndexOf(')')+2).split(' ')[19]};}));
  }
  if(initialPid){const worker=rows.find(row=>row.pid===initialPid);if(!worker?.created)throw failure('JOB_OBSERVER_FAILED',500);owned.set(worker.pid,worker.created);}
  const ids=new Set(rows.filter(row=>row.created&&owned.get(row.pid)===row.created).map(row=>row.pid));
  const byPid=new Map(rows.map(row=>[row.pid,row]));
  let changed=true;while(changed){changed=false;for(const row of rows)if(ids.has(row.parent)&&!ids.has(row.pid)&&row.created&&BigInt(row.created)>=BigInt(byPid.get(row.parent).created)){ids.add(row.pid);changed=true;}}
  const selected=rows.filter(row=>ids.has(row.pid)&&row.created);
  for(const row of selected)owned.set(row.pid,row.created);
  return selected;
}

async function terminateOwned(child,owned){
  if(!owned.size){if(child.exitCode===null&&child.signalCode===null)child.kill('SIGKILL');return;}
  // If the observer failed, retain the previously verified identities. The
  // termination command still verifies them before it opens any PID tree.
  await processTree(owned).catch(()=>{});
  if(process.platform==='win32'){
    // Hold verified process handles through termination and the final child
    // census. Never let taskkill /T expand a numeric PID into an unverified tree.
    const identities=[...owned].map(([pid,created])=>`${pid}:${created}`).join(',');
    const script=`$ErrorActionPreference='Stop';$targets=@{}
function Add-Owned([int]$taskId,[long]$created){
  try{$p=Get-Process -Id $taskId -ErrorAction Stop;$handle=$p.Handle;$ticks=$p.StartTime.ToFileTimeUtc();$ticks=$ticks-($ticks%10)
    if($ticks -eq $created){$targets[$taskId]=@{process=$p;created=$ticks};return $true};$p.Dispose()
  }catch{$row=Get-CimInstance Win32_Process -Filter "ProcessId = $taskId" -ErrorAction SilentlyContinue;if($row -and $row.CreationDate.ToFileTimeUtc() -eq $created){throw 'Worker identity check failed'}};return $false
}
$activeAtStart=@{}
foreach($row in Get-CimInstance Win32_Process){$activeAtStart[[int]$row.ProcessId]=$true}
foreach($entry in '${identities}'.Split(',')){if($entry){$parts=$entry.Split(':');if($activeAtStart.ContainsKey([int]$parts[0])){$null=Add-Owned ([int]$parts[0]) ([long]$parts[1])}}}
try{
  for($pass=0;$pass -lt 4;$pass++){
    foreach($target in @($targets.Values)){$p=$target.process;if(!$p.HasExited){try{$p.Kill()}catch{if(!$p.HasExited){throw}};if(!$p.WaitForExit(5000)){throw 'Worker cleanup failed'}}}
    $added=$false
    $rows=@(Get-CimInstance Win32_Process);$parents=@{}
    foreach($row in $rows){$parents[[int]$row.ProcessId]=$row}
    foreach($row in $rows){
      $parent=[int]$row.ParentProcessId;$taskId=[int]$row.ProcessId
      if($targets.ContainsKey($parent) -and !$targets.ContainsKey($taskId)){
        $parentRow=$parents[$parent]
        if(!$parentRow -or $parentRow.CreationDate.ToFileTimeUtc() -ne $targets[$parent].created){throw 'Worker child ownership could not be verified'}
        $created=$row.CreationDate.ToFileTimeUtc()
        if($created -ge $targets[$parent].created){if(Add-Owned $taskId $created){$added=$true}}
      }
    }
    if(!$added){break}
  }
  if($added){throw 'Worker cleanup incomplete'}
}finally{foreach($target in $targets.Values){$target.process.Dispose()}}
exit 0`;
    await execute('powershell.exe',['-NoProfile','-NonInteractive','-Command',script],{windowsHide:true,timeout:10000,maxBuffer:65536});
  }else{
    const rows=await processTree(owned);
    for(const row of rows.reverse())try{process.kill(row.pid,'SIGKILL');}catch{}
  }
}

export async function runConversion({input,operation,signal,limits={}}={}){
  const started=Date.now(),resourceUsage={workerPid:null,browserPid:null,peakRssBytes:null,rssSamples:0,elapsedMs:0,terminationMs:0,forcedTermination:false,ownedPids:[]};
  let config;
  try{
    validateInput(operation,input);
    config={...conversionLimits,...limits};
    for(const value of Object.values(config))if(!Number.isSafeInteger(value)||value<=0)throw failure('INPUT_INVALID');
    if(signal?.aborted)throw failure('JOB_CANCELLED',409);
    if(active>=2)throw failure('JOB_BUSY',429);
  }catch(error){throw Object.assign(error,{resourceUsage});}
  active++;let retainSlotUntilClose=false;
  try{return await new Promise((resolve,reject)=>{
    const child=fork(new URL('./conversion-worker.mjs',import.meta.url),[],{serialization:'advanced',stdio:['ignore','ignore','ignore','ipc'],execArgv:[`--max-old-space-size=${config.oldSpaceMb}`],windowsHide:true,detached:process.platform!=='win32'});
    resourceUsage.workerPid=child.pid??null;
    let result=null,error=null,stoppedAt=null,grace=null,sampling=null,finishing=false,force=null,failSafe=null,observerFailed=false;
    const owned=new Map();
    const stop=(reason)=>{
      if(finishing||stoppedAt!==null)return;
      error=reason;stoppedAt=Date.now();
      if(child.connected)child.send({type:'cancel'},()=>{});
      grace=setTimeout(()=>{
        resourceUsage.forcedTermination=true;
        force=terminateOwned(child,owned).then(async()=>{
          if(child.exitCode===null&&child.signalCode===null)await new Promise(resolve=>{
            const done=()=>{clearTimeout(timer);child.off('exit',done);resolve();};
            const timer=setTimeout(done,1000);child.once('exit',done);
          });
          if(child.exitCode===null&&child.signalCode===null)throw new Error('Worker exit unconfirmed');
        }).catch(()=>{
          error=failure('JOB_CLEANUP_FAILED',500);
          // ChildProcess owns a process handle; this fallback targets only our worker.
          try{if(child.exitCode===null&&child.signalCode===null)child.kill('SIGKILL');}catch{}
        }).then(()=>{
          failSafe=setTimeout(()=>{
            if(finishing)return;
            clearTimeout(timeout);clearInterval(observer);signal?.removeEventListener('abort',abort);
            error=failure('JOB_CLEANUP_FAILED',500);
            retainSlotUntilClose=true;resourceUsage.elapsedMs=Date.now()-started;resourceUsage.terminationMs=Date.now()-stoppedAt;
            reject(Object.assign(error,{resourceUsage}));
          },2000);
        });
      },config.closeGraceMs);
    };
    const abort=()=>stop(failure('JOB_CANCELLED',409));
    const sample=(initial=false)=>{
      if(sampling||finishing)return;
      const observe=()=>processTree(owned,initial?child.pid:null);
      sampling=observe().catch(error=>{
        // Before the worker receives input, one timed-out census can be retried safely.
        if(initial&&error.killed&&stoppedAt===null&&!finishing)return observe();
        throw error;
      }).then(rows=>{
        resourceUsage.rssSamples++;
        const rss=rows.reduce((sum,row)=>sum+row.rss,0);
        if(rss>(resourceUsage.peakRssBytes??0)){resourceUsage.peakRssBytes=rss;resourceUsage.peakProcesses=rows.map(({pid,parent,rss})=>({pid,parent,rss}));}
        resourceUsage.ownedPids=[...new Set([...resourceUsage.ownedPids,...rows.map(row=>row.pid)])];
        if(resourceUsage.peakRssBytes>config.maxRssBytes)stop(failure('JOB_MEMORY_LIMIT',413));
      }).catch(()=>{observerFailed=true;stop(failure('JOB_OBSERVER_FAILED',500));}).finally(()=>{sampling=null;});
      return sampling;
    };
    const timeout=setTimeout(()=>stop(failure('JOB_TIMEOUT',408)),config.timeoutMs);
    const observer=setInterval(sample,config.rssIntervalMs);
    signal?.addEventListener('abort',abort,{once:true});
    child.on('message',message=>{
      if(message?.type==='browser'&&Number.isSafeInteger(message.pid)&&message.pid>0){resourceUsage.browserPid=message.pid;sample();}
      if(stoppedAt!==null)return;
      if(message?.type==='result')result=message.result;
      if(message?.type==='error')stop(failure(['INPUT_LIMIT','INPUT_INVALID','JOB_CLEANUP_FAILED'].includes(message.code)?message.code:'JOB_FAILED',message.code==='INPUT_LIMIT'?413:500));
    });
    child.once('error',()=>stop(failure('JOB_FAILED',500)));
    child.once('close',async code=>{
      finishing=true;clearTimeout(timeout);clearTimeout(grace);clearTimeout(failSafe);clearInterval(observer);signal?.removeEventListener('abort',abort);
      if(sampling)await sampling;
      if(force)await force;
      // Also covers worker crashes after the browser was launched.
      if(code!==0||stoppedAt!==null)try{await terminateOwned(child,owned);}catch{error=failure('JOB_CLEANUP_FAILED',500);}
      resourceUsage.elapsedMs=Date.now()-started;
      resourceUsage.terminationMs=stoppedAt===null?0:Date.now()-stoppedAt;
      if(signal?.aborted&&!error)error=failure('JOB_CANCELLED',409);
      if(observerFailed&&!error)error=failure('JOB_OBSERVER_FAILED',500);
      if(resourceUsage.peakRssBytes>config.maxRssBytes&&!error)error=failure('JOB_MEMORY_LIMIT',413);
      if(error||code!==0||!result)reject(Object.assign(error??failure('JOB_FAILED',500),{resourceUsage}));
      else resolve({...result,resourceUsage});
      if(retainSlotUntilClose){retainSlotUntilClose=false;active--;}
    });
    sample(true).then(()=>{
      if(!finishing&&stoppedAt===null)child.send({type:'run',operation,input},sendError=>{if(sendError)stop(failure('JOB_FAILED',500));});
    });
    if(signal?.aborted)abort();
  });}finally{if(!retainSlotUntilClose)active--;}
}
