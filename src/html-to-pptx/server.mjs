import http from 'node:http';
import fs from 'node:fs/promises';
import path from 'node:path';
import {fileURLToPath,pathToFileURL} from 'node:url';
import {randomUUID} from 'node:crypto';
import {createBrowserRuntime} from './browser-runtime.mjs';
import {readInput} from './input.mjs';
import {measureInput} from './measure.mjs';
import {diagnostic} from './model.mjs';
import {exportPptx} from './export.mjs';
import {publicDiagnostic} from './public-diagnostics.mjs';
const root=fileURLToPath(new URL('../../',import.meta.url));
const json=(res,status,data)=>{res.writeHead(status,{'content-type':'application/json; charset=utf-8','cache-control':'no-store','x-content-type-options':'nosniff'});res.end(JSON.stringify(data));};
const pptxMime='application/vnd.openxmlformats-officedocument.presentationml.presentation';
const mime={'.html':'text/html; charset=utf-8','.js':'text/javascript; charset=utf-8','.css':'text/css; charset=utf-8'};
const downloadName=job=>(job.fileName.replace(/[\\/]/g,'_').replace(/[\x00-\x1f\x7f<>:"|?*]/g,'_').replace(/\.html?$/i,'').slice(0,120)||'slides')+'.pptx';
async function readBody(req,limit) {
  let size=0;const chunks=[];
  for await(const c of req){size+=c.length;if(size>limit){const e=new Error('입력 크기 제한을 초과했습니다.');e.statusCode=413;throw e;}chunks.push(c);}
  try{return JSON.parse(Buffer.concat(chunks).toString('utf8'));}catch{const e=new Error('올바른 JSON 입력이 필요합니다.');e.statusCode=400;throw e;}
}
export function summarizeJob(job,{developmentReports=false}={}) {
  const measured=job.measured,failed=measured.diagnostics.some(d=>d.severity==='error');
  const diagnostics=job.exportResult?.diagnostics??measured.diagnostics;
  return {id:job.id,fileName:job.fileName,status:failed?'failed':'analyzed',profile:job.input.profile,sourceDigest:job.input.source.sha256,
    slideCount:measured.model?.slides.length??0,objectCount:measured.model?.slides.flatMap(s=>s.nodes).filter(n=>n.kind!=='group').length??0,
    diagnostics:developmentReports?diagnostics:diagnostics.map(d=>publicDiagnostic(d,measured.model?.slides)),exportAvailable:!failed&&!job.exporting,visualQuality:'unmeasured',
    exportStatus:job.exporting?'converting':job.exportResult?.status??'pending',
    downloadUrl:job.exportResult?.status==='generated'?`/api/jobs/${job.id}/download`:null,
    downloadFileName:downloadName(job),objectMapUrl:developmentReports&&job.exportResult?.status==='generated'?`/api/jobs/${job.id}/object-map`:null,
    slides:measured.model?.slides.map((s,i)=>({id:developmentReports?s.id:`slide-${i+1}`,number:i+1,widthCssPx:s.widthCssPx,heightCssPx:s.heightCssPx,objectCount:s.nodes.filter(n=>n.kind!=='group').length,
      previewUrl:measured.previews.some(p=>p.slideId===s.id)?`/api/jobs/${job.id}/preview/${i+1}.png`:null}))??[]};
}
export async function createHtmlServer({port=42337,host='127.0.0.1',maxBodyBytes=96*1024*1024,developmentReports=false}={}) {
  if(host!=='127.0.0.1')throw new Error('HTML 입력 서버는 로컬 루프백에서만 실행합니다.');
  const serveDev=developmentReports?(await import('./dev-reports.mjs')).serveDevelopmentRoute:null;
  const summarize=job=>summarizeJob(job,{developmentReports});
  const runtime=await createBrowserRuntime(),jobs=new Map();let active=0,url;
  const server=http.createServer(async(req,res)=>{
    try {
      if(req.headers.host!==new URL(url).host)return json(res,403,{error:'허용되지 않은 호스트입니다.'});
      if(req.headers.origin&&req.headers.origin!==url)return json(res,403,{error:'같은 로컬 화면에서 요청해 주세요.'});
      const route=new URL(req.url,url).pathname;
      for(const [id,j]of jobs)if(!j.exporting&&Date.now()-j.createdAt>30*60*1000)jobs.delete(id);
      if(req.method==='POST'&&route==='/api/analyze') {
        if(active>=2)return json(res,429,{error:'현재 입력을 분석 중입니다. 잠시 뒤 다시 실행해 주세요.'});
        active++;
        try {
          const body=await readBody(req,maxBodyBytes);
          if(!body||typeof body!=='object'||Array.isArray(body))return json(res,400,{error:'HTML과 자산 목록이 필요합니다.'});
          if(body.assets!==undefined&&!Array.isArray(body.assets))return json(res,400,{error:'자산 목록은 배열이어야 합니다.'});
          const assets=(body.assets??[]).map(a=>{
            if(!a||typeof a.base64!=='string'||!/^(?:[A-Za-z0-9+/]{4})*(?:[A-Za-z0-9+/]{2}==|[A-Za-z0-9+/]{3}=)?$/.test(a.base64)){const e=new Error('자산 base64가 올바르지 않습니다.');e.statusCode=400;throw e;}
            return {path:a.path,mime:a.mime||undefined,bytes:Buffer.from(a.base64,'base64')};
          });
          const input=await readInput({html:body.html,assets,slideSelector:body.slideSelector??null});
          const measured=await measureInput(input,runtime);
          const job={id:randomUUID(),createdAt:Date.now(),fileName:typeof body.fileName==='string'?body.fileName.slice(0,200):'입력 HTML',input,measured};
          while(jobs.size>=10){const old=[...jobs].find(([,j])=>!j.exporting);if(!old)break;jobs.delete(old[0]);}jobs.set(job.id,job);
          const summary=summarize(job);return json(res,summary.status==='failed'?422:200,summary);
        }finally{active--;}
      }
      if(req.method==='POST'&&route==='/api/export'){
        const body=await readBody(req,4096);
        if(!body||typeof body.jobId!=='string'||!body.jobId)return json(res,400,{error:'분석 작업 식별자가 필요합니다.'});
        const job=jobs.get(body.jobId);
        if(!job)return json(res,404,{error:'분석 결과가 없거나 만료됐습니다. 다시 분석해 주세요.'});
        if(job.exporting)return json(res,409,{error:'이미 PPT를 만드는 중입니다.'});
        if(job.measured.diagnostics.some(d=>d.severity==='error'))return json(res,422,summarize(job));
        if(job.exportResult?.status==='generated')return json(res,200,summarize(job));
        if(active>=2)return json(res,429,{error:'다른 장표를 처리 중입니다. 잠시 뒤 다시 실행해 주세요.'});
        active++;job.exporting=true;
        try{job.exportResult=await exportPptx(job.measured);job.createdAt=Date.now();}
        finally{job.exporting=false;active--;}
        return json(res,job.exportResult.status==='generated'?200:422,summarize(job));
      }
      if(req.method==='GET') {
        if(serveDev&&await serveDev({route,res,root}))return;
        const output=/^\/api\/jobs\/([a-f0-9-]+)\/(download|object-map)$/.exec(route);
        if(output){
          if(output[2]==='object-map'&&!developmentReports)return json(res,404,{error:'요청한 경로를 찾을 수 없습니다.'});
          const job=jobs.get(output[1]);if(!job)return json(res,404,{error:'분석 결과가 없거나 만료됐습니다.'});
          if(job.exportResult?.status!=='generated')return json(res,409,{error:'완료된 PPT가 없습니다. 변환 결과를 확인해 주세요.'});
          if(output[2]==='object-map')return json(res,200,job.exportResult.objectMap);
          res.writeHead(200,{'content-type':pptxMime,'content-length':job.exportResult.pptxBytes.length,'content-disposition':`attachment; filename="slides.pptx"; filename*=UTF-8''${encodeURIComponent(downloadName(job))}`,'cache-control':'no-store','x-content-type-options':'nosniff'});
          return res.end(Buffer.from(job.exportResult.pptxBytes));
        }
        const preview=/^\/api\/jobs\/([a-f0-9-]+)\/preview\/(\d+)\.png$/.exec(route);
        if(preview){const j=jobs.get(preview[1]),s=j?.measured.model?.slides[Number(preview[2])-1],image=j?.measured.previews.find(p=>p.slideId===s?.id);if(!image)return json(res,404,{error:'미리보기가 없거나 만료됐습니다.'});res.writeHead(200,{'content-type':'image/png','cache-control':'no-store','x-content-type-options':'nosniff'});return res.end(Buffer.from(image.bytes));}
        const jobRoute=/^\/api\/jobs\/([a-f0-9-]+)$/.exec(route);
        if(jobRoute){const j=jobs.get(jobRoute[1]);return j?json(res,200,summarize(j)):json(res,404,{error:'분석 결과가 만료됐습니다. 다시 분석해 주세요.'});}
        const relative={'/':'web/html-to-pptx/index.html','/app.js':'web/html-to-pptx/app.js','/style.css':'web/html-to-pptx/style.css'}[route];
        if(typeof relative==='string'){
          const file=path.join(root,relative),bytes=await fs.readFile(file);
          res.writeHead(200,{'content-type':mime[path.extname(file)],'x-content-type-options':'nosniff','cache-control':'no-store','Content-Security-Policy':"default-src 'none'; img-src 'self' data:; style-src 'self'; script-src 'self'; connect-src 'self'; object-src 'none'; base-uri 'none'; frame-ancestors 'none'; form-action 'self'"});return res.end(bytes);
        }
      }
      return json(res,404,{error:'요청한 경로를 찾을 수 없습니다.'});
    }catch(e){
      const status=[400,413].includes(e.statusCode)?e.statusCode:500;
      const message=status===413?'입력 크기 제한을 초과했습니다.':status===400?'올바른 JSON 입력이 필요합니다.':'분석 중 오류가 발생했습니다. 파일을 확인한 뒤 다시 시도해 주세요.';
      if(!res.headersSent)json(res,status,{error:message,diagnostics:[developmentReports?diagnostic('RENDER_INCOMPLETE',e.message):publicDiagnostic({code:'RENDER_INCOMPLETE'})]});else res.end();
    }
  });
  try{await new Promise((resolve,reject)=>{server.once('error',reject);server.listen(port,host,resolve);});}catch(e){await runtime.close();throw e;}
  url=`http://${host}:${server.address().port}`;
  return {url,jobs,runtime,close:async()=>{await new Promise(resolve=>server.close(resolve));await runtime.close();}};
}
if(process.argv[1]&&pathToFileURL(path.resolve(process.argv[1])).href===import.meta.url){
  const app=await createHtmlServer({port:Number(process.env.HTML_PPT_PORT??42337),developmentReports:process.env.HTML_PPT_DEV_REPORTS==='1'});console.log(`Slidehandoff: ${app.url}`);
  for(const signal of ['SIGINT','SIGTERM'])process.once(signal,()=>app.close().then(()=>process.exit(0)));
}
