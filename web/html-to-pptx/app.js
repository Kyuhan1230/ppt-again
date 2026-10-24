const $=id=>document.getElementById(id);
let revision=0,controller=null,currentJob=null;
const labels={idle:'파일 선택 대기',ready:'분석 준비',analyzing:'분석 중',analyzed:'입력 분석 완료',converting:'PPT 만드는 중',generated:'PPT 생성 완료',failed:'입력 확인 필요'};
function status(state,message){$('status-badge').dataset.state=state;$('status-badge').textContent=labels[state];$('status-message').textContent=message;}
function clearDownload(){$('download-button').disabled=true;$('download-button').textContent='편집 가능한 PPT 다운로드';}
function reset(){revision++;controller?.abort();currentJob=null;$('analyze-button').disabled=!$('html-file').files.length;$('analyze-button').textContent='입력 분석';clearDownload();$('preview-area').hidden=true;$('empty-preview').hidden=false;$('diagnostics').hidden=true;$('diagnostic-list').replaceChildren();$('file-label').textContent=$('html-file').files[0]?.name??'장표 HTML을 선택하세요';status($('html-file').files.length?'ready':'idle','파일과 자산을 선택한 뒤 입력 분석을 실행하세요.');}
for(const id of ['html-file','assets','asset-folder','selector'])$(id).addEventListener('change',reset);
const base64=buffer=>{const bytes=new Uint8Array(buffer);let binary='';for(let i=0;i<bytes.length;i+=32768)binary+=String.fromCharCode(...bytes.subarray(i,i+32768));return btoa(binary);};
function showSlide(index){const slide=currentJob.slides[index];if(!slide?.previewUrl)return;$('preview-image').src=slide.previewUrl;$('preview-image').alt=`${index+1}번 장표 미리보기`;$('slide-detail').textContent=`${index+1} / ${currentJob.slideCount}장 · 원본 ${slide.widthCssPx} × ${slide.heightCssPx}px · 객체 ${slide.objectCount}개`;for(const [i,button]of [...$('slide-list').children].entries())button.setAttribute('aria-current',String(i===index));}
function showResult(job){
  currentJob=job;const failed=job.status==='failed';status(job.status,failed?'아래 항목을 수정하거나 장표 영역을 지정한 뒤 다시 분석하세요.':`${job.fileName}에서 장표 ${job.slideCount}장과 객체 ${job.objectCount}개를 확인했습니다.`);
  $('download-button').disabled=!job.exportAvailable;
  $('diagnostic-list').replaceChildren();$('diagnostics').hidden=!job.diagnostics.length;
  for(const d of job.diagnostics){const li=document.createElement('li'),small=document.createElement('small');li.append(document.createTextNode(`${d.severity==='error'?'확인 필요':'지원 제한'} · ${d.message}`));small.textContent=d.slideNumber?`${d.slideNumber}번 장표`:'';li.append(small);$('diagnostic-list').append(li);}
  const visible=job.slides.some(s=>s.previewUrl);$('preview-area').hidden=!visible;$('empty-preview').hidden=visible;
  if(visible){$('slide-count').textContent=`${job.slideCount}장`;$('object-count').textContent=`분석 객체 ${job.objectCount}개`;$('slide-list').replaceChildren();job.slides.forEach((s,i)=>{const b=document.createElement('button');b.type='button';b.textContent=String(i+1);b.setAttribute('aria-label',`${i+1}번 장표`);b.addEventListener('click',()=>showSlide(i));$('slide-list').append(b);});showSlide(job.slides.findIndex(s=>s.previewUrl));}
}
$('analyze-form').addEventListener('submit',async event=>{
  event.preventDefault();const file=$('html-file').files[0];if(!file||$('analyze-button').disabled)return;
  const thisRevision=++revision;controller?.abort();controller=new AbortController();currentJob=null;clearDownload();$('analyze-button').disabled=true;$('analyze-button').textContent='장표를 분석하고 있습니다…';status('analyzing','글꼴과 이미지를 읽고 실제 장표 배치를 확인하고 있습니다.');$('diagnostics').hidden=true;
  try{
    if(file.size>16*1024*1024)throw new Error('HTML은 16MB 이하로 선택해 주세요.');
    const files=[...$('assets').files,...$('asset-folder').files];if(files.reduce((sum,f)=>sum+f.size,0)>64*1024*1024)throw new Error('자산의 전체 크기는 64MB 이하여야 합니다.');
    const html=new TextDecoder('utf-8',{fatal:true,ignoreBOM:true}).decode(await file.arrayBuffer());
    const assets=await Promise.all(files.map(async f=>({path:f.webkitRelativePath||f.name,mime:f.type,base64:base64(await f.arrayBuffer())})));
    const response=await fetch('/api/analyze',{method:'POST',headers:{'content-type':'application/json'},body:JSON.stringify({html,assets,fileName:file.name,slideSelector:$('selector').value.trim()||null}),signal:controller.signal});
    const job=await response.json();if(thisRevision!==revision)return;
    if(!job.id)throw new Error(job.error??'분석 결과를 받지 못했습니다.');showResult(job);
  }catch(error){if(error.name!=='AbortError'&&thisRevision===revision)status('failed',error.message);}
  finally{if(thisRevision===revision){$('analyze-button').disabled=false;$('analyze-button').textContent='입력 다시 분석';}}
});
$('download-button').addEventListener('click',async()=>{
  if(!currentJob||$('download-button').disabled)return;
  const thisRevision=revision,jobId=currentJob.id;controller?.abort();controller=new AbortController();
  $('download-button').disabled=true;$('download-button').textContent='PPT를 만들고 있습니다…';$('analyze-button').disabled=true;
  status('converting','글자·도형·그림을 PowerPoint 객체로 변환하고 있습니다.');
  try{
    const response=await fetch('/api/export',{method:'POST',headers:{'content-type':'application/json'},body:JSON.stringify({jobId}),signal:controller.signal});
    const result=await response.json();if(thisRevision!==revision)return;
    if(result.id){showResult(result);$('download-button').disabled=true;}
    if(!response.ok||result.exportStatus!=='generated')throw new Error(result.error??'PPT를 만들지 못했습니다. 아래 진단을 확인해 주세요.');
    const file=await fetch(result.downloadUrl,{signal:controller.signal});if(!file.ok)throw new Error('PPT 파일을 받지 못했습니다. 다시 다운로드해 주세요.');
    const blob=await file.blob();if(thisRevision!==revision)return;
    const url=URL.createObjectURL(blob),link=document.createElement('a');link.href=url;link.download=result.downloadFileName;document.body.append(link);link.click();link.remove();setTimeout(()=>URL.revokeObjectURL(url),1000);
    status('generated',`${result.slideCount}장 PPT를 만들었습니다. 브라우저의 다운로드 목록에서 파일을 확인하세요. 화면 품질 검증은 진행 중입니다.`);
  }catch(error){if(error.name!=='AbortError'&&thisRevision===revision)status('failed',error.message);}
  finally{if(thisRevision===revision){$('download-button').disabled=!currentJob?.exportAvailable;$('download-button').textContent='편집 가능한 PPT 다운로드';$('analyze-button').disabled=false;}}
});
