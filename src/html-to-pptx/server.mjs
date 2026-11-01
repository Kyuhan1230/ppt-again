import http from 'node:http';
import fs from 'node:fs/promises';
import path from 'node:path';
import {fileURLToPath,pathToFileURL} from 'node:url';
import {createJobStore} from './job-store.mjs';
import {createSessionManager} from './session.mjs';
import {runConversion} from './job-runner.mjs';
import {diagnostic} from './model.mjs';
import {publicDiagnostic} from './public-diagnostics.mjs';
const root=fileURLToPath(new URL('../../',import.meta.url));
const json=(res,status,data)=>{if(res.destroyed)return;res.writeHead(status,{'content-type':'application/json; charset=utf-8','cache-control':'no-store','x-content-type-options':'nosniff'});res.end(JSON.stringify(data));};
const pptxMime='application/vnd.openxmlformats-officedocument.presentationml.presentation';
const mime={'.html':'text/html; charset=utf-8','.js':'text/javascript; charset=utf-8','.css':'text/css; charset=utf-8'};
const downloadName=job=>(job.fileName.replace(/[\\/]/g,'_').replace(/[\x00-\x1f\x7f<>:"|?*]/g,'_').replace(/\.(?:html?|pptx)$/i,'').slice(0,120).toWellFormed()||'slides')+(job.partialExport?'-partial':'')+'.'+(job.targetType??'pptx');
function partialSlides(job){
  const measured=job.measured,slides=measured?.model?.slides??[],errors=(measured?.diagnostics??[]).filter(d=>d.severity==='error'),ids=new Set(slides.map(s=>s.id));
  if(job.sourceType!=='html'||job.targetType!=='pptx'||!errors.length||errors.some(d=>!d.slideId||!ids.has(d.slideId)))return null;
  const blocked=new Set(errors.map(d=>d.slideId)),included=slides.filter(s=>!blocked.has(s.id));
  if(!included.length||included.length===slides.length)return null;
  return {included,excludedSlideNumbers:slides.flatMap((s,i)=>blocked.has(s.id)?[i+1]:[])};
}
function decodeBase64(value,maxBytes){
  const fail=statusCode=>{throw Object.assign(new Error('잘못된 파일 바이트입니다.'),{statusCode});};
  if(typeof value!=='string')fail(400);
  if(value.length>4*Math.ceil(maxBytes/3))fail(413);
  if(value.length%4||!/^[A-Za-z0-9+/]*={0,2}$/.test(value)||value&&Buffer.from(value.slice(-4),'base64').toString('base64')!==value.slice(-4))fail(400);
  const bytes=Buffer.from(value,'base64');if(bytes.length>maxBytes)fail(413);return bytes;
}
async function readBody(req,limit) {
  let size=0;const chunks=[];
  for await(const c of req){size+=c.length;if(size>limit){const e=new Error('입력 크기 제한을 초과했습니다.');e.statusCode=413;throw e;}chunks.push(c);}
  try{return JSON.parse(Buffer.concat(chunks).toString('utf8'));}catch{const e=new Error('올바른 JSON 입력이 필요합니다.');e.statusCode=400;throw e;}
}
export function summarizeJob(job,{developmentReports=false}={}) {
  const sourceType=job.sourceType??'html',targetType=job.targetType??'pptx';
  if(!job.measured)return {id:job.id,sourceType,targetType,status:'analyzing',exportAvailable:false,diagnostics:[],slides:[]};
  const measured=job.measured,failed=measured.diagnostics.some(d=>d.severity==='error');
  const diagnostics=job.exportResult?.diagnostics??measured.diagnostics,partial=partialSlides(job);
  return {id:job.id,fileName:job.fileName,sourceType,targetType,status:failed?'failed':'analyzed',profile:job.input.profile,sourceDigest:job.input.source.sha256,
    slideCount:measured.model?.slides.length??0,objectCount:measured.model?.slides.flatMap(s=>s.nodes).filter(n=>n.kind!=='group').length??0,
    diagnostics:developmentReports?diagnostics:diagnostics.map(d=>publicDiagnostic(d,measured.model?.slides)),exportAvailable:!failed&&!job.exporting,partialExportAvailable:Boolean(partial&&!job.exporting),partialSlideCount:partial?.included.length??0,excludedSlideNumbers:partial?.excludedSlideNumbers??[],downloadSlideCount:job.partialExport?partial?.included.length??0:measured.model?.slides.length??0,visualQuality:'unmeasured',
    exportStatus:job.exporting?'converting':job.exportResult?.status??'pending',
    downloadUrl:job.exportResult?.status==='generated'?`/api/jobs/${job.id}/download`:null,
    downloadFileName:downloadName(job),objectMapUrl:targetType==='pptx'&&developmentReports&&job.exportResult?.status==='generated'?`/api/jobs/${job.id}/object-map`:null,
    slides:measured.model?.slides.map((s,i)=>({id:developmentReports?s.id:`slide-${i+1}`,number:i+1,widthCssPx:s.widthCssPx,heightCssPx:s.heightCssPx,objectCount:s.nodes.filter(n=>n.kind!=='group').length,
      previewUrl:measured.previews.some(p=>p.slideId===s.id)?`/api/jobs/${job.id}/preview/${i+1}.png`:null}))??[]};
}
export async function createHtmlServer({port=42337,host='127.0.0.1',maxBodyBytes=96*1024*1024,developmentReports=false,jobTtlMs=1800000,maxJobBytes=268435456,conversionLimits={}}={}) {
  if(host!=='127.0.0.1')throw new Error('HTML 입력 서버는 로컬 루프백에서만 실행합니다.');
  const serveDev=developmentReports?(await import('./dev-reports.mjs')).serveDevelopmentRoute:null;
  const summarize=job=>summarizeJob(job,{developmentReports});
  const sessions=createSessionManager();
  const jobs=createJobStore({ttlMs:jobTtlMs,maxBytes:maxJobBytes,onDispose:async job=>{job.abortController?.abort();await job.pending?.catch(()=>{});}});
  const missing=res=>json(res,404,{error:'결과가 없거나 만료됐습니다. 다시 분석해 주세요.'});
  let active=0,url;
  const server=http.createServer(async(req,res)=>{
    try {
      if(req.headers.host!==new URL(url).host)return json(res,403,{error:'허용되지 않은 호스트입니다.'});
      if(req.headers.origin&&req.headers.origin!==url)return json(res,403,{error:'같은 로컬 화면에서 요청해 주세요.'});
      const route=new URL(req.url,url).pathname;
      if(req.method==='GET'&&route==='/api/session')return json(res,200,{csrfToken:sessions.bootstrap(req,res).csrfToken});
      const session=sessions.read(req),ownerId=session?.ownerId;
      if(['POST','DELETE'].includes(req.method)&&route.startsWith('/api/')&&!sessions.permitsWrite(req,session))return json(res,403,{error:'화면을 새로 연 뒤 다시 실행해 주세요.'});
      if(req.method==='POST'&&route==='/api/analyze') {
        if(active>=2)return json(res,429,{error:'현재 입력을 분석 중입니다. 잠시 뒤 다시 실행해 주세요.'});
        active++;
        try {
          const body=await readBody(req,maxBodyBytes);
          if(!body||typeof body!=='object'||Array.isArray(body))return json(res,400,{error:'변환할 파일이 필요합니다.'});
          const sourceType=body.sourceType===undefined?'html':body.sourceType;
          if(!['html','pptx'].includes(sourceType))return json(res,400,{error:'HTML 또는 PPTX 파일이 필요합니다.'});
          const targetType=sourceType==='html'?'pptx':'html';let input;
          if(sourceType==='pptx'){
            if(body.html!==undefined||body.assets!==undefined||body.slideSelector!==undefined||body.fileName!==undefined&&(typeof body.fileName!=='string'||!/\.pptx$/i.test(body.fileName)))return json(res,400,{error:'PPTX 파일 하나를 선택해 주세요.'});
            input=decodeBase64(body.pptxBase64,64*1024*1024);
          }else{
            if(body.pptxBase64!==undefined)return json(res,400,{error:'HTML과 PPTX를 한 번에 입력할 수 없습니다.'});
            if(body.assets!==undefined&&!Array.isArray(body.assets))return json(res,400,{error:'자산 목록은 배열이어야 합니다.'});
            const assets=(body.assets??[]).map(a=>({path:a?.path,mime:a?.mime||undefined,bytes:decodeBase64(a?.base64,16*1024*1024)}));
            input={html:body.html,assets,slideSelector:body.slideSelector??null,allowTableOverlay:true};
          }
          const abortController=new AbortController();
          const id=jobs.create(ownerId,{direction:`${sourceType}-to-${targetType}`,sourceType,targetType,fileName:typeof body.fileName==='string'?body.fileName.slice(0,200):(sourceType==='pptx'?'입력.pptx':'입력 HTML'),input,abortController});
          const disconnected=()=>{if(!res.writableEnded)abortController.abort();};res.on('close',disconnected);
          try{
            const pending=runConversion({input,operation:sourceType==='pptx'?'analyze-pptx':'analyze-html',signal:abortController.signal,limits:conversionLimits});
            if(!jobs.update(ownerId,id,{pending})){abortController.abort();await pending.catch(()=>{});return missing(res);}
            const result=await pending;
            if(abortController.signal.aborted||!jobs.update(ownerId,id,{input:result.input,measured:result.measured,resourceUsage:result.resourceUsage,pending:null}))return missing(res);
            const saved=jobs.get(ownerId,id);if(!saved)return missing(res);
            const summary=summarize(saved);return json(res,summary.status==='failed'?422:200,summary);
          }catch(error){await jobs.delete(ownerId,id);throw error;}
          finally{res.off('close',disconnected);}
        }finally{active--;}
      }
      if(req.method==='POST'&&route==='/api/export'){
        const body=await readBody(req,4096);
        if(!body||typeof body.jobId!=='string'||!body.jobId)return json(res,400,{error:'분석 작업 식별자가 필요합니다.'});
        const job=jobs.get(ownerId,body.jobId);
        if(!job)return missing(res);
        const target=body.target===undefined?(job.targetType??'pptx'):body.target;
        if(target!==(job.targetType??'pptx'))return json(res,400,{error:'HTML은 PPTX로, PPTX는 HTML로 변환할 수 있습니다.'});
        if(job.exporting||!job.measured)return json(res,409,{error:'이미 장표를 처리 중입니다.'});
        const partial=body.partial===true?partialSlides(job):null;
        if(body.partial!==undefined&&typeof body.partial!=='boolean'||body.partial===true&&!partial)return json(res,422,summarize(job));
        if(!partial&&job.measured.diagnostics.some(d=>d.severity==='error'))return json(res,422,summarize(job));
        if(job.exportResult?.status==='generated')return json(res,200,summarize(job));
        if(active>=2)return json(res,429,{error:'다른 장표를 처리 중입니다. 잠시 뒤 다시 실행해 주세요.'});
        active++;
        const abortController=new AbortController(),disconnected=()=>{if(!res.writableEnded)abortController.abort();};res.on('close',disconnected);
        try{
          if(!jobs.update(ownerId,job.id,{exporting:true,abortController}))return missing(res);
          const {previews,...whole}=job.measured;
          const ids=partial?new Set(partial.included.map(s=>s.id)):null;
          const measured=partial?{...whole,model:{...whole.model,slides:partial.included},measurements:whole.measurements.filter(m=>ids.has(m.slideId)),diagnostics:whole.diagnostics.filter(d=>d.severity!=='error'&&(!d.slideId||ids.has(d.slideId)))}:whole;
          const pending=runConversion({input:target==='html'?measured.model:measured,operation:target==='html'?'export-html':'export-pptx',signal:abortController.signal,limits:conversionLimits});
          if(!jobs.update(ownerId,job.id,{pending})){abortController.abort();await pending.catch(()=>{});return missing(res);}
          const exportResult=await pending;
          exportResult.diagnostics=[...new Map([...whole.diagnostics,...exportResult.diagnostics].map(d=>[JSON.stringify(d),d])).values()];
          if(abortController.signal.aborted||!jobs.update(ownerId,body.jobId,{exportResult,partialExport:Boolean(partial),exporting:false,pending:null}))return missing(res);
          return json(res,exportResult.status==='generated'?200:422,summarize(job));
        }catch(error){await jobs.delete(ownerId,body.jobId);throw error;}
        finally{res.off('close',disconnected);active--;}
      }
      const jobRoute=/^\/api\/jobs\/([a-f0-9-]+)$/.exec(route);
      if(req.method==='DELETE'&&jobRoute){
        if(!await jobs.delete(ownerId,jobRoute[1]))return missing(res);
        res.writeHead(204,{'cache-control':'no-store'});return res.end();
      }
      if(req.method==='GET') {
        if(serveDev&&await serveDev({route,res,root}))return;
        const output=/^\/api\/jobs\/([a-f0-9-]+)\/(download|object-map)$/.exec(route);
        if(output){
          if(output[2]==='object-map'&&!developmentReports)return json(res,404,{error:'요청한 경로를 찾을 수 없습니다.'});
          const job=jobs.get(ownerId,output[1]);if(!job)return missing(res);
          if(output[2]==='object-map'&&job.targetType==='html')return json(res,404,{error:'요청한 경로를 찾을 수 없습니다.'});
          if(job.exportResult?.status!=='generated')return json(res,409,{error:'완료된 변환 파일이 없습니다. 변환 결과를 확인해 주세요.'});
          if(output[2]==='object-map')return json(res,200,job.exportResult.objectMap);
          const html=job.targetType==='html',bytes=Buffer.from(html?job.exportResult.html:job.exportResult.pptxBytes);
          res.writeHead(200,{'content-type':html?mime['.html']:pptxMime,'content-length':bytes.length,'content-disposition':`attachment; filename="slides.${html?'html':'pptx'}"; filename*=UTF-8''${encodeURIComponent(downloadName(job))}`,'cache-control':'no-store','x-content-type-options':'nosniff'});
          return res.end(bytes);
        }
        const preview=/^\/api\/jobs\/([a-f0-9-]+)\/preview\/(\d+)\.png$/.exec(route);
        if(preview){const j=jobs.get(ownerId,preview[1]);if(!j)return missing(res);const s=j.measured?.model?.slides[Number(preview[2])-1],image=j.measured?.previews.find(p=>p.slideId===s?.id);if(!image)return missing(res);res.writeHead(200,{'content-type':'image/png','cache-control':'no-store','x-content-type-options':'nosniff'});return res.end(Buffer.from(image.bytes));}
        if(jobRoute){const j=jobs.get(ownerId,jobRoute[1]);return j?json(res,200,summarize(j)):missing(res);}
        const relative={'/':'web/html-to-pptx/index.html','/app.js':'web/html-to-pptx/app.js','/style.css':'web/html-to-pptx/style.css'}[route];
        if(typeof relative==='string'){
          const file=path.join(root,relative),bytes=await fs.readFile(file);
          res.writeHead(200,{'content-type':mime[path.extname(file)],'x-content-type-options':'nosniff','cache-control':'no-store','Content-Security-Policy':"default-src 'none'; img-src 'self' data:; style-src 'self'; script-src 'self'; connect-src 'self'; object-src 'none'; base-uri 'none'; frame-ancestors 'none'; form-action 'self'"});return res.end(bytes);
        }
      }
      return json(res,404,{error:'요청한 경로를 찾을 수 없습니다.'});
    }catch(e){
      const status=[400,408,409,413,429].includes(e.statusCode)?e.statusCode:500;
      const messages={JOB_TIMEOUT:'처리 시간을 초과해 작업을 중단했습니다. 파일을 확인한 뒤 다시 시도해 주세요.',JOB_CANCELLED:'작업을 취소했습니다.',JOB_MEMORY_LIMIT:'작업 메모리 기준을 초과해 중단했습니다.',JOB_BUSY:'다른 장표를 처리 중입니다. 잠시 뒤 다시 실행해 주세요.'};
      const message=messages[e.code]??(status===429?'보관 가능한 작업 용량을 초과했습니다. 이전 작업을 지운 뒤 다시 시도해 주세요.':status===413?'입력 크기 제한을 초과했습니다.':status===400?'올바른 입력이 필요합니다.':'분석 중 오류가 발생했습니다. 파일을 확인한 뒤 다시 시도해 주세요.');
      if(!res.headersSent)json(res,status,{error:message,diagnostics:[developmentReports?diagnostic('RENDER_INCOMPLETE',e.message):publicDiagnostic({code:'RENDER_INCOMPLETE'})]});else res.end();
    }
  });
  await new Promise((resolve,reject)=>{server.once('error',reject);server.listen(port,host,resolve);});
  url=`http://${host}:${server.address().port}`;
  let closing;
  return {url,jobs,close:()=>closing??=(async()=>{const stopped=new Promise(resolve=>server.close(resolve));await jobs.close();server.closeAllConnections();await stopped;})()};
}
if(process.argv[1]&&pathToFileURL(path.resolve(process.argv[1])).href===import.meta.url){
  const app=await createHtmlServer({port:Number(process.env.HTML_PPT_PORT??42337),developmentReports:process.env.HTML_PPT_DEV_REPORTS==='1'});console.log(`Slidehandoff: ${app.url}`);
  for(const signal of ['SIGINT','SIGTERM'])process.once(signal,()=>app.close().then(()=>process.exit(0)));
}
