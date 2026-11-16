import test,{before,after} from 'node:test';
import assert from 'node:assert/strict';
import childProcess from 'node:child_process';
import {EventEmitter} from 'node:events';
import {syncBuiltinESMExports} from 'node:module';
import {promisify} from 'node:util';

const worker={pid:241, parent:15, rss:1024, created:'110'};
const browser={pid:242, parent:241, rss:2048, created:'120'};
const unrelated={pid:99, parent:15, rss:1024, created:'90'};
const timeout=()=>Object.assign(new Error('simulated timeout'),{code:null,killed:true,signal:'SIGTERM'});
const accessDenied=()=>Object.assign(new Error('simulated permission error'),{code:5});
const realExecFile=childProcess.execFile,realFork=childProcess.fork;
const platform=Object.getOwnPropertyDescriptor(process,'platform');
let scenario,runConversion;

before(async()=>{
  Object.defineProperty(process,'platform',{...platform,value:'win32'});
  childProcess.execFile=function(){throw new Error('Expected the promisified command');};
  childProcess.execFile[promisify.custom]=async(file,args,options)=>{
    assert.equal(file,'powershell.exe');
    const script=args.at(-1);
    if(!script.includes('ConvertTo-Json')){
      scenario.cleanupCommands.push(script);
      return {stdout:''};
    }
    assert.equal(options.timeout,5000);
    const sample=scenario.samples.shift();
    assert.ok(sample,'Unexpected process sample');
    return {stdout:JSON.stringify(await sample())};
  };
  childProcess.fork=()=>scenario.child;
  syncBuiltinESMExports();
  ({runConversion}=await import('../src/html-to-pptx/job-runner.mjs'));
});

after(()=>{
  Object.defineProperty(process,'platform',platform);
  childProcess.execFile=realExecFile;
  childProcess.fork=realFork;
  syncBuiltinESMExports();
});

function run(samples,{onRun,signal,maxRssBytes}={}){
  const child=Object.assign(new EventEmitter(),{pid:worker.pid,connected:true,exitCode:null,signalCode:null});
  const events=[],cleanupCommands=[];
  function close(code){
    child.exitCode=code;
    child.connected=false;
    child.emit('close',code);
  }
  const succeed=()=>{
    child.emit('message',{type:'result',result:{status:'parsed'}});
    close(0);
  };
  child.send=(message,callback)=>{
    events.push(message.type);
    queueMicrotask(()=>{
      callback();
      if(message.type==='run')(onRun??succeed)({child,succeed,close});
      if(message.type==='cancel')close(1);
    });
  };
  child.kill=()=>close(1);
  scenario={child,samples:[...samples],cleanupCommands};
  const limits={rssIntervalMs:10000,closeGraceMs:30,...(maxRssBytes===undefined?{}:{maxRssBytes})};
  const pending=runConversion({operation:'read-pptx',input:new Uint8Array([0x50,0x4b]),limits,signal});
  return {pending,events,cleanupCommands};
}

async function expectObserverFailure(pending,phase,causeCode){
  await assert.rejects(pending,error=>{
    assert.equal(error.code,'JOB_OBSERVER_FAILED');
    assert.equal(error.observerPhase,phase);
    assert.equal(error.cause?.code,causeCode);
    return true;
  });
}

const initialSample=()=>[worker,unrelated];
const rejectedSample=error=>()=>{throw error();};

test('initial Windows observation retries a timeout exactly once',async()=>{
  const {pending,events}=run([rejectedSample(timeout),initialSample]);
  assert.equal((await pending).resourceUsage.rssSamples,1);
  assert.deepEqual(events,['run']);

  const failed=run([rejectedSample(timeout),rejectedSample(timeout)]);
  await expectObserverFailure(failed.pending,'initial','OBSERVER_TIMEOUT');
  assert.deepEqual(failed.events,['cancel']);
});

test('Windows observation distinguishes permission failure and does not retry it',async()=>{
  const {pending,events}=run([rejectedSample(accessDenied)]);
  await expectObserverFailure(pending,'initial','OBSERVER_ACCESS_DENIED');
  assert.deepEqual(events,['cancel']);
});

test('an active observation timeout cancels the conversion and cleans up verified descendants',async()=>{
  const {pending,events,cleanupCommands}=run([()=>[worker,browser],rejectedSample(timeout),()=>[browser]],{
    onRun:({child})=>child.emit('message',{type:'browser',pid:browser.pid})
  });
  await expectObserverFailure(pending,'active','OBSERVER_TIMEOUT');
  assert.deepEqual(events,['run','cancel']);
  assert.match(cleanupCommands[0],/242:120/);
});

test('a timeout during clean worker exit accepts a result only after a new census confirms exit',async()=>{
  const {pending,events,cleanupCommands}=run([()=>[worker,browser],rejectedSample(timeout),()=>[unrelated]],{
    onRun:({child,succeed})=>{child.emit('message',{type:'browser',pid:browser.pid});succeed();}
  });
  assert.equal((await pending).status,'parsed');
  assert.deepEqual(events,['run']);
  assert.deepEqual(cleanupCommands,[]);
});

test('a surviving browser after a closing timeout still fails and is cleaned up',async()=>{
  const {pending,cleanupCommands}=run([()=>[worker,browser],rejectedSample(timeout),()=>[browser],()=>[browser]],{
    onRun:({child,succeed})=>{child.emit('message',{type:'browser',pid:browser.pid});succeed();}
  });
  await expectObserverFailure(pending,'closing','OBSERVER_TIMEOUT');
  assert.match(cleanupCommands[0],/242:120/);
});

test('an unverified child of the reported browser prevents closing recovery',async()=>{
  const orphan={pid:243,parent:browser.pid,rss:2048,created:'130'};
  const {pending,cleanupCommands}=run([initialSample,rejectedSample(timeout),()=>[orphan],()=>[orphan]],{
    onRun:({child,succeed})=>{child.emit('message',{type:'browser',pid:browser.pid});succeed();}
  });
  await expectObserverFailure(pending,'closing','OBSERVER_TIMEOUT');
  assert.doesNotMatch(cleanupCommands[0],/243:130/);
});

test('a failed recheck after a closing timeout preserves the original cause',async()=>{
  const {pending,cleanupCommands}=run([initialSample,rejectedSample(timeout),rejectedSample(accessDenied),()=>[unrelated]],{
    onRun:({child,succeed})=>{child.emit('message',{type:'browser',pid:browser.pid});succeed();}
  });
  await expectObserverFailure(pending,'closing','OBSERVER_TIMEOUT');
  assert.equal(cleanupCommands.length,1);
});

test('a permission failure during close keeps its cause and requires cleanup',async()=>{
  const {pending,cleanupCommands}=run([()=>[worker,browser],rejectedSample(accessDenied),()=>[unrelated]],{
    onRun:({child,succeed})=>{child.emit('message',{type:'browser',pid:browser.pid});succeed();}
  });
  await expectObserverFailure(pending,'closing','OBSERVER_ACCESS_DENIED');
  assert.equal(cleanupCommands.length,1);
});

test('observed memory limits and cancellation still stop the job and clean up',async()=>{
  const overLimit=run([()=>[worker,browser],()=>[browser]],{maxRssBytes:worker.rss});
  await assert.rejects(overLimit.pending,{code:'JOB_MEMORY_LIMIT'});
  assert.deepEqual(overLimit.events,['cancel']);
  assert.match(overLimit.cleanupCommands[0],/242:120/);

  const controller=new AbortController();
  const cancelled=run([initialSample,()=>[unrelated]],{signal:controller.signal,onRun:()=>controller.abort()});
  await assert.rejects(cancelled.pending,{code:'JOB_CANCELLED'});
  assert.deepEqual(cancelled.events,['run','cancel']);
  assert.match(cancelled.cleanupCommands[0],/241:110/);
});
