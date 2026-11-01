const $=id=>document.getElementById(id);
let revision=0,controller=null,currentJob=null,busy=false;
let sessionReady=null;
const longDeckLimit=20;
function durationNotice(count){const note=$('duration-note');note.hidden=!(count>longDeckLimit);note.textContent=note.hidden?'':`${count}장이라 분석과 변환에 오래 걸릴 수 있어요. 장표가 복잡할수록 시간이 더 걸립니다.`;}
function embeddedPageCount(html){
  const lower=html.toLowerCase();let count=null,pos=0;
  while(pos<lower.length){const start=lower.indexOf('<script',pos);if(start<0)break;const open=lower.indexOf('>',start+7);if(open<0)break;const end=lower.indexOf('</script',open+1);if(end<0)break;pos=end+8;
    if(!/\btype\s*=\s*(["'])application\/json\1/i.test(html.slice(start,open)))continue;try{const data=JSON.parse(html.slice(open+1,end));if(data&&Array.isArray(data.pages)&&data.sources&&typeof data.sources==='object'){if(count!==null)return null;count=data.pages.length;}}catch{}}
  return count;
}
async function apiPost(route,body,signal,method='POST'){
  sessionReady??=fetch('/api/session').then(async response=>{if(!response.ok)throw new Error('화면을 새로 연 뒤 다시 실행해 주세요.');return (await response.json()).csrfToken;}).catch(error=>{sessionReady=null;throw error;});
  const csrfToken=await sessionReady;
  return fetch(route,{method,headers:{'content-type':'application/json','x-csrf-token':csrfToken},body:body===undefined?undefined:JSON.stringify(body),signal});
}
const labels={idle:'파일 선택 대기',ready:'분석 준비',analyzing:'분석 중',analyzed:'입력 분석 완료',converting:'파일 만드는 중',generated:'파일 생성 완료',failed:'입력 확인 필요',cancelling:'작업 정리 중',cancelled:'취소 요청됨',deleted:'작업 삭제 완료'};
const sourceType=file=>/\.pptx$/i.test(file?.name??'')?'pptx':/\.html?$/i.test(file?.name??'')?'html':null;
const targetType=()=>currentJob?.targetType??(sourceType($('html-file').files[0])==='pptx'?'html':'pptx');
const targetLabel=()=>targetType()==='html'?'HTML':'PPT';
const downloadLabel=()=>sourceType($('html-file').files[0])?(targetType()==='html'?'HTML 다운로드':'편집 가능한 PPT 다운로드'):'변환 파일 다운로드';
function describeDirection(){
  const source=sourceType($('html-file').files[0]);$('html-options').hidden=source==='pptx';for(const id of ['assets','asset-folder','selector'])$(id).disabled=source==='pptx';
  $('conversion-direction').textContent=source==='pptx'?'PPTX → HTML · 현재 수정 내용을 읽습니다.':source==='html'?'HTML → PPTX · 여러 장을 한 PPT로 묶습니다.':'파일 형식에 맞춰 변환 방향을 정합니다.';
  $('output-help').textContent=source==='pptx'?'지원되는 현재 글자·도형·표·차트를 HTML로 저장합니다. 복잡한 서식과 미지원 객체는 분석에서 알려드립니다.':'지원되는 글자·도형·표·차트를 PowerPoint에서 수정할 수 있습니다. 확인이 필요한 항목은 분석 결과에 표시합니다.';
  $('quality-note').textContent=source==='pptx'?'이 미리보기는 변환된 HTML입니다. 원본 PowerPoint 화면과 차이가 있을 수 있습니다.':'이 미리보기는 입력 HTML입니다. 함께 제공한 지원 정적 TTF·WOFF2 글꼴은 PPT에 포함합니다. 포함되지 않은 글꼴은 사용할 컴퓨터에 설치해야 합니다.';
}
function status(state,message){document.body.dataset.phase=state;$('status-badge').dataset.state=state;$('status-badge').textContent=labels[state];$('status-message').textContent=message;}
function clearDownload(){$('download-button').disabled=true;$('download-button').hidden=false;$('download-button').textContent=downloadLabel();$('partial-download-button').hidden=true;$('partial-download-button').disabled=true;$('partial-download-note').hidden=true;}
function jobActions(){$('cancel-button').disabled=!busy;$('delete-button').disabled=!currentJob||busy;}
function reset(){revision++;controller?.abort();if(currentJob)void apiPost(`/api/jobs/${currentJob.id}`,undefined,undefined,'DELETE').catch(()=>{});currentJob=null;busy=false;durationNotice(null);jobActions();const file=$('html-file').files[0],type=sourceType(file);$('analyze-button').disabled=!type;$('analyze-button').textContent='다시 분석';describeDirection();clearDownload();$('preview-area').hidden=true;$('empty-preview').hidden=false;$('diagnostics').hidden=true;$('diagnostic-list').replaceChildren();$('preview-image').removeAttribute('src');$('file-label').textContent=file?.name??'파일을 여기에 놓거나 선택하세요';status(file&&!type?'failed':type?'ready':'idle',file&&!type?'HTML(.html·.htm) 또는 PPTX 파일을 선택해 주세요.':'파일을 선택하면 자동으로 분석합니다.');}
$('html-file').addEventListener('change',()=>{reset();if(sourceType($('html-file').files[0]))$('analyze-form').requestSubmit();});
for(const id of ['assets','asset-folder','selector'])$(id).addEventListener('change',reset);
const dropArea=document.querySelector('.file-area');
for(const event of ['dragenter','dragover'])dropArea.addEventListener(event,e=>{e.preventDefault();dropArea.classList.add('dragging');});
for(const event of ['dragleave','drop'])dropArea.addEventListener(event,e=>{e.preventDefault();dropArea.classList.remove('dragging');});
dropArea.addEventListener('drop',e=>{if(e.dataTransfer.files.length!==1){$('html-file').value='';reset();status('failed','HTML 또는 PPTX 파일 한 개를 선택해 주세요.');return;}$('html-file').files=e.dataTransfer.files;$('html-file').dispatchEvent(new Event('change',{bubbles:true}));});
const base64=buffer=>{const bytes=new Uint8Array(buffer);let binary='';for(let i=0;i<bytes.length;i+=32768)binary+=String.fromCharCode(...bytes.subarray(i,i+32768));return btoa(binary);};
function showSlide(index){const slide=currentJob.slides[index];if(!slide?.previewUrl)return;$('preview-image').src=slide.previewUrl;$('preview-image').alt=`${index+1}번 장표 미리보기`;$('slide-detail').textContent=`${index+1} / ${currentJob.slideCount}장`;for(const [i,button]of [...$('slide-list').children].entries())button.setAttribute('aria-current',String(i===index));}
function showResult(job){
  durationNotice(job.slideCount);
  currentJob=job;const failed=job.status==='failed';status(job.status,failed?(job.sourceType==='pptx'?'아래 지원 제한을 확인하고 PPT를 수정한 뒤 다시 선택해 주세요.':'아래 항목을 수정하거나 장표 영역을 지정한 뒤 다시 분석하세요.'):`${job.fileName}에서 장표 ${job.slideCount}장을 확인했습니다.`);
  describeDirection();$('download-button').textContent=downloadLabel();
  jobActions();
  $('download-button').disabled=!job.exportAvailable;$('download-button').hidden=Boolean(job.partialExportAvailable);
  $('partial-download-button').hidden=!job.partialExportAvailable;$('partial-download-button').disabled=!job.partialExportAvailable;$('partial-download-note').hidden=!job.partialExportAvailable;
  if(job.partialExportAvailable){$('partial-download-button').textContent=`변환 가능한 ${job.partialSlideCount}장만 받기`;$('partial-download-note').textContent=`전체 ${job.slideCount}장 중 오류가 있는 ${job.excludedSlideNumbers.join(', ')}번 장표는 빠집니다. 원래 순서를 유지합니다.`;}
  $('diagnostic-list').replaceChildren();$('diagnostics').hidden=!job.diagnostics.length;
  const groups=new Map();for(const d of job.diagnostics){const key=JSON.stringify([d.severity,d.code,d.message]);if(!groups.has(key))groups.set(key,{...d,numbers:new Set()});if(d.slideNumber)groups.get(key).numbers.add(d.slideNumber);}
  for(const d of groups.values()){const li=document.createElement('li'),small=document.createElement('small'),numbers=[...d.numbers].sort((a,b)=>a-b),ranges=[];for(let i=0;i<numbers.length;){let end=i;while(end+1<numbers.length&&numbers[end+1]===numbers[end]+1)end++;ranges.push(end===i?`${numbers[i]}`:`${numbers[i]}–${numbers[end]}`);i=end+1;}li.append(document.createTextNode(`${d.severity==='error'?'확인 필요':'지원 제한'} · ${d.message}`));small.textContent=ranges.length?`${ranges.join(', ')}번 장표`:'';li.append(small);$('diagnostic-list').append(li);}
  const visible=job.slides.some(s=>s.previewUrl);$('preview-area').hidden=!visible;$('empty-preview').hidden=visible;
  if(visible){$('slide-count').textContent=`${job.slideCount}장`;$('object-count').textContent=`${targetLabel()}로 저장`;$('slide-list').replaceChildren();job.slides.forEach((s,i)=>{const b=document.createElement('button');b.type='button';b.textContent=String(i+1);b.setAttribute('aria-label',`${i+1}번 장표`);b.addEventListener('click',()=>showSlide(i));$('slide-list').append(b);});showSlide(job.slides.findIndex(s=>s.previewUrl));}
}
$('analyze-form').addEventListener('submit',async event=>{
  event.preventDefault();const file=$('html-file').files[0];if(!file||$('analyze-button').disabled)return;
  const thisRevision=++revision,previousJob=currentJob;controller?.abort();controller=new AbortController();const signal=controller.signal;currentJob=null;busy=true;jobActions();clearDownload();$('preview-area').hidden=true;$('empty-preview').hidden=false;$('preview-image').removeAttribute('src');$('analyze-button').disabled=true;$('analyze-button').textContent='장표를 분석하고 있습니다…';status('analyzing','글꼴과 이미지를 읽고 실제 장표 배치를 확인하고 있습니다.');$('diagnostics').hidden=true;
  const noticeTimer=setTimeout(()=>{if(thisRevision===revision&&!signal.aborted&&$('duration-note').hidden){$('duration-note').textContent='장표를 분석하는 데 시간이 걸리고 있어요. 잠시만 기다려 주세요.';$('duration-note').hidden=false;}},20000);
  try{
    if(previousJob)await apiPost(`/api/jobs/${previousJob.id}`,undefined,undefined,'DELETE');
    const type=sourceType(file);if(!type)throw new Error('HTML 또는 PPTX 파일을 선택해 주세요.');
    if(file.size>(type==='pptx'?64:16)*1024*1024)throw new Error(type==='pptx'?'PPTX는 64MiB 이하로 선택해 주세요.':'HTML은 16MiB 이하로 선택해 주세요.');
    let input;
    if(type==='pptx')input={sourceType:'pptx',pptxBase64:base64(await file.arrayBuffer()),fileName:file.name};
    else{
      const files=[...$('assets').files,...$('asset-folder').files];if(files.reduce((sum,f)=>sum+f.size,0)>64*1024*1024)throw new Error('자산의 전체 크기는 64MiB 이하여야 합니다.');
      const html=new TextDecoder('utf-8',{fatal:true,ignoreBOM:true}).decode(await file.arrayBuffer());
      if(thisRevision===revision&&!signal.aborted)durationNotice(embeddedPageCount(html));
      const assets=await Promise.all(files.map(async f=>({path:f.webkitRelativePath||f.name,mime:f.type,base64:base64(await f.arrayBuffer())})));
      input={sourceType:'html',html,assets,fileName:file.name,slideSelector:$('selector').value.trim()||null};
    }
    if(thisRevision!==revision||signal.aborted)return;
    const response=await apiPost('/api/analyze',input,signal);
    const job=await response.json();if(thisRevision!==revision)return;
    if(!job.id)throw new Error(job.error??'분석 결과를 받지 못했습니다.');showResult(job);
  }catch(error){if(error.name!=='AbortError'&&thisRevision===revision)status('failed',error.message);}
  finally{clearTimeout(noticeTimer);if(thisRevision===revision){busy=false;jobActions();$('analyze-button').disabled=false;$('analyze-button').textContent='입력 다시 분석';}}
});
async function download(partial=false){
  const button=$(partial?'partial-download-button':'download-button');if(!currentJob||button.disabled)return;
  const thisRevision=revision,jobId=currentJob.id;controller?.abort();controller=new AbortController();const signal=controller.signal;
  busy=true;jobActions();
  const target=targetType(),label=targetLabel();
  $('download-button').disabled=true;$('partial-download-button').disabled=true;button.textContent=`${label} 파일을 만들고 있습니다…`;$('analyze-button').disabled=true;
  status('converting',target==='html'?'현재 내용과 배치를 HTML로 저장하고 있습니다.':'지원되는 글자·도형·표·차트를 PowerPoint 객체로 변환하고 있습니다.');
  try{
    const response=await apiPost('/api/export',{jobId,target,...(partial?{partial:true}:{})},signal);
    const result=await response.json();if(thisRevision!==revision)return;
    if(result.id){showResult(result);$('download-button').disabled=true;}
    if(!response.ok||result.exportStatus!=='generated')throw new Error(result.error??`${label} 파일을 만들지 못했습니다. 아래 진단을 확인해 주세요.`);
    const file=await fetch(result.downloadUrl,{signal});if(!file.ok)throw new Error(`${label} 파일을 받지 못했습니다. 다시 다운로드해 주세요.`);
    const blob=await file.blob();if(thisRevision!==revision)return;
    const url=URL.createObjectURL(blob),link=document.createElement('a');link.href=url;link.download=result.downloadFileName;document.body.append(link);link.click();link.remove();setTimeout(()=>URL.revokeObjectURL(url),1000);
    status('generated',`${result.downloadSlideCount??result.slideCount}장 ${label} 파일을 만들었습니다.${partial?' 오류가 있는 장표는 포함되지 않았습니다.':''} 브라우저의 다운로드 목록에서 파일을 확인하세요. 화면 품질 검증은 진행 중입니다.`);
  }catch(error){if(error.name!=='AbortError'&&thisRevision===revision)status('failed',error.message);}
  finally{if(thisRevision===revision){busy=false;jobActions();$('download-button').disabled=!currentJob?.exportAvailable;$('download-button').textContent=downloadLabel();$('partial-download-button').disabled=!currentJob?.partialExportAvailable;if(currentJob?.partialExportAvailable)$('partial-download-button').textContent=`변환 가능한 ${currentJob.partialSlideCount}장만 받기`;$('analyze-button').disabled=false;}}
}
$('download-button').addEventListener('click',()=>download(false));
$('partial-download-button').addEventListener('click',()=>download(true));
async function removeCurrentJob(){
  const thisRevision=++revision,job=currentJob,request=controller;busy=false;
  $('cancel-button').disabled=true;$('delete-button').disabled=true;$('analyze-button').disabled=true;clearDownload();status('cancelling','처리를 중단하고 서버에 보관된 작업을 정리하고 있습니다.');
  try{
    let deleted=false;
    // Keep the request open until DELETE confirms that the worker has stopped.
    if(job){const response=await apiPost(`/api/jobs/${job.id}`,undefined,undefined,'DELETE');if(!response.ok&&response.status!==404)throw new Error('작업 삭제를 확인하지 못했습니다. 다시 시도해 주세요.');deleted=response.status===204;}
    request?.abort();if(thisRevision!==revision)return;
    currentJob=null;$('preview-area').hidden=true;$('empty-preview').hidden=false;$('preview-image').removeAttribute('src');$('diagnostics').hidden=true;$('diagnostic-list').replaceChildren();
    status(deleted?'deleted':'cancelled',deleted?'서버의 작업과 결과를 삭제했습니다. 이미 다운로드한 파일은 남아 있습니다.':job?'작업을 찾을 수 없습니다. 이미 만료되었거나 정리 중일 수 있습니다.':'취소 요청을 보냈습니다. 서버에서 진행 중인 작업을 정리합니다.');
  }catch(error){if(thisRevision===revision)status('failed',error.message);}
  finally{if(thisRevision===revision){jobActions();$('analyze-button').disabled=!sourceType($('html-file').files[0]);$('analyze-button').textContent='입력 다시 분석';}}
}
$('cancel-button').addEventListener('click',removeCurrentJob);
$('delete-button').addEventListener('click',removeCurrentJob);
