import {readInput} from './input.mjs';
import {measureInput} from './measure.mjs';
import {exportPptx} from './export.mjs';
import {readPptx} from './read-pptx.mjs';
import {exportHtml} from './export-html.mjs';
import {createBrowserRuntime} from './browser-runtime.mjs';
import {chartText,validateChart} from './charts.mjs';
import {sha256} from './assets.mjs';

let runtime=null,cancelled=false,started=false,finished=false;
const limit=condition=>{if(condition)throw Object.assign(new Error('INPUT_LIMIT'),{code:'INPUT_LIMIT'});};
function checkModel(model){
  if(!model)return;
  const slides=model.slides??[];let objects=0,text=0;
  limit(slides.length>200);
  for(const slide of slides){
    limit(slide.nodes.length>10000||(objects+=slide.nodes.length)>100000);
    for(const node of slide.nodes){
      if(node.kind==='chart'&&!validateChart(node.content).valid)throw Object.assign(new Error('INPUT_INVALID'),{code:'INPUT_INVALID'});
      const contents=node.kind==='table'?(node.content?.cells??[]):node.kind==='chart'?chartText(node.content).map(text=>({text})):[node.content??node];
      for(const content of contents)for(const run of content.runs??[{text:content.text??''}]){const length=typeof run.text==='string'?run.text.length:0;limit(length>100000||(text+=length)>2000000);}
    }
  }
}
async function analyzeHtml(source){
  const input=await readInput(source);
  checkModel(input.project);
  limit(input.fontFaces.length>32||input.assets.filter(asset=>asset.mime.startsWith('font/')).reduce((n,asset)=>n+asset.bytes.length,0)>32*1024*1024);
  runtime=await createBrowserRuntime({checkLimits:true,onBrowserProcess:pid=>{if(process.connected)process.send({type:'browser',pid});}});
  if(cancelled)return null;
  const measured=await measureInput(input,runtime);limit(runtime.limitExceeded);checkModel(measured.model);
  return {input,measured};
}
async function cancel(){cancelled=true;await runtime?.close().catch(()=>{});}
process.on('disconnect',()=>{if(!finished)void cancel().finally(()=>process.exit(1));});
process.on('message',async message=>{
  if(message?.type==='cancel'){await cancel();if(!started&&process.connected){finished=true;process.disconnect();}return;}
  if(started||message?.type!=='run')return;
  started=true;let result,error;
  try{
    if(cancelled)return;
    if(message.operation==='analyze-html'){
      result=await analyzeHtml(message.input);
    }else if(message.operation==='analyze-pptx'){
      const bytes=message.input instanceof Uint8Array?message.input:message.input.bytes,options=message.input instanceof Uint8Array?{}:message.input.options;
      const parsed=await readPptx(bytes,options);checkModel(parsed.model);
      const input={profile:'pptx-v1',source:{sha256:sha256(bytes)},options};
      let measured={model:parsed.model,measurements:[],environment:{},diagnostics:[...parsed.diagnostics],previews:[]};
      if(parsed.status==='parsed'){
        const generated=exportHtml(parsed.model);measured.diagnostics.push(...generated.diagnostics);
        if(generated.status==='generated'){
          // Exported HTML comes from the validated current PPT model. A preview
          // renderer is optional for this direction, never for HTML→PPTX.
          try{
            const preview=await analyzeHtml({html:generated.html});
            if(cancelled)return;
            const fatal=preview.measured.diagnostics.find(d=>d.severity==='error'&&['INPUT_INVALID','INPUT_LIMIT','LIMIT_EXCEEDED'].includes(d.code));
            if(fatal)throw Object.assign(new Error('INPUT_LIMIT'),{code:'INPUT_LIMIT'});
            const previewDiagnostics=preview.measured.diagnostics.map(d=>d.severity!=='error'?d:d.code==='FONT_UNAVAILABLE'?{...d,code:'FONT_PREVIEW_SUBSTITUTED',severity:'warning'}:{code:'PPTX_PREVIEW_UNAVAILABLE',severity:'warning',reasonKey:'PREVIEW_UNAVAILABLE',slideId:d.slideId??null,message:'HTML은 만들었지만 미리보기를 확인하지 못했습니다.'});
            const diagnostics=[...measured.diagnostics,...previewDiagnostics];
            measured={...preview.measured,model:parsed.model,diagnostics:[...new Map(diagnostics.map(d=>[JSON.stringify(d),d])).values()]};
          }catch(error){
            if(cancelled||['INPUT_INVALID','INPUT_LIMIT','LIMIT_EXCEEDED','JOB_CANCELLED','JOB_MEMORY_LIMIT','JOB_TIMEOUT'].includes(error.code))throw error;
            measured.diagnostics.push({code:'PPTX_PREVIEW_UNAVAILABLE',severity:'warning',reasonKey:'PREVIEW_UNAVAILABLE',message:'HTML은 만들었지만 미리보기를 확인하지 못했습니다.'});
          }
        }
      }
      result={input,measured};
    }else if(message.operation==='export-pptx'){
      checkModel(message.input.model);result=await exportPptx(message.input,{fontEmbedding:message.input.model.fonts?.length?'auto':'off'});
      limit(result.pptxBytes?.length>64*1024*1024);
    }else if(message.operation==='read-pptx'){
      result=await readPptx(message.input instanceof Uint8Array?message.input:message.input.bytes,message.input instanceof Uint8Array?{}:message.input.options);checkModel(result.model);
    }else if(message.operation==='export-html'){
      checkModel(message.input);result=exportHtml(message.input);
    }else throw new Error('INPUT_INVALID');
  }catch(caught){error=caught;}
  finally{
    try{await runtime?.close();}catch{error=Object.assign(new Error('JOB_CLEANUP_FAILED'),{code:'JOB_CLEANUP_FAILED'});}
    runtime=null;finished=true;
    if(process.connected){
      const response=error?{type:'error',code:error.code??'JOB_FAILED'}:cancelled?{type:'error',code:'JOB_CANCELLED'}:{type:'result',result};
      process.send(response,()=>process.disconnect());
    }
  }
});
