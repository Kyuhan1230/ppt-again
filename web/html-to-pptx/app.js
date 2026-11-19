import {renderDiagnostics,validBounds} from './diagnostic-view.js';
import {selectedFileState,sourceType} from './input-policy.js';
const $=id=>document.getElementById(id);
$('quality-note').hidden=true;
const hints=[...document.querySelectorAll('.hint-wrap')];
for(const hint of hints){for(const event of ['pointerenter','focusin','click'])hint.addEventListener(event,()=>{delete hint.dataset.dismissed;});}
document.addEventListener('keydown',event=>{if(event.key==='Escape')for(const hint of hints)if(hint.matches(':hover')||hint.contains(document.activeElement))hint.dataset.dismissed='true';});
let revision=0,controller=null,currentJob=null,busy=false,selectedChoiceId=null,selectableJobId=null,activeRequestId=null;
const cleanupJobs=new Set(),cleanupRequests=new Set();let cleanupPromise=null;
let sessionReady=null;
const longDeckLimit=20;
function durationNotice(count){const note=$('duration-note');note.hidden=!(count>longDeckLimit);note.textContent=note.hidden?'':`${count}장이라 분석과 변환에 오래 걸릴 수 있어요. 장표가 복잡할수록 시간이 더 걸립니다.`;}
function embeddedPageCount(html){
  const lower=html.toLowerCase();let count=null,pos=0;
  while(pos<lower.length){const start=lower.indexOf('<script',pos);if(start<0)break;const open=lower.indexOf('>',start+7);if(open<0)break;const end=lower.indexOf('</script',open+1);if(end<0)break;pos=end+8;
    if(!/\btype\s*=\s*(["'])application\/json\1/i.test(html.slice(start,open)))continue;try{const data=JSON.parse(html.slice(open+1,end));if(data&&Array.isArray(data.pages)&&data.sources&&typeof data.sources==='object'){if(count!==null)return null;count=data.pages.length;}}catch{}}
  return count;
}
async function apiPost(route,body,signal,method='POST',requestId=null){
  sessionReady??=fetch('/api/session').then(async response=>{if(!response.ok)throw new Error('화면을 새로 연 뒤 다시 실행해 주세요.');return (await response.json()).csrfToken;}).catch(error=>{sessionReady=null;throw error;});
  const csrfToken=await sessionReady;
  return fetch(route,{method,headers:{'content-type':'application/json','x-csrf-token':csrfToken,...(requestId?{'x-job-request-id':requestId}:{})},body:body===undefined?undefined:JSON.stringify(body),signal});
}
async function postJob(route,body,signal){
  const token=crypto.randomUUID().replaceAll('-','');activeRequestId=token;
  try{const response=await apiPost(route,body,signal,'POST',token),data=await response.json();if(activeRequestId===token)activeRequestId=null;if(!data.id){cleanupRequests.add(token);void cleanupOutstanding().catch(()=>{});}return {response,data};}
  catch(error){if(activeRequestId===token){activeRequestId=null;cleanupRequests.add(token);void cleanupOutstanding().catch(()=>{});}throw error;}
}
const labels={idle:'파일 선택 대기',ready:'분석 준비',analyzing:'분석 중','selection-required':'영역 선택 필요',selecting:'선택 영역 분석 중',analyzed:'입력 분석 완료',converting:'파일 만드는 중',generated:'파일 생성 완료',failed:'입력 확인 필요',cancelling:'작업 정리 중',cancelled:'취소 요청됨',deleted:'작업 삭제 완료'};
const targetType=()=>currentJob?.targetType??(sourceType($('html-file').files[0])==='pptx'?'html':'pptx');
const targetLabel=()=>targetType()==='html'?'HTML':'PPT';
const downloadLabel=()=>sourceType($('html-file').files[0])?(targetType()==='html'?'HTML 다운로드':'편집 가능한 PPT 다운로드'):'변환 파일 다운로드';
const paginateDocument=()=>$('html-page-mode').value==='paginate';
function describeDirection(){
  const source=sourceType($('html-file').files[0]);
  $('html-options').hidden=source==='pptx';
  $('pptx-options').hidden=source!=='pptx';
  for(const id of ['assets','asset-folder','html-page-mode','slide-size-policy'])$(id).disabled=source==='pptx';
  for(const id of ['slide-width','slide-height'])$(id).disabled=source==='pptx'||$('slide-size-policy').value!=='custom';
  const paginating=paginateDocument();
  $('selector').disabled=source==='pptx'||paginating;
  $('html-selector-settings').hidden=paginating;
  $('html-pagination-help').hidden=!paginating;
  $('slide-size-policy').options[0].textContent=paginating?'와이드 16:9 (기본)':'원본 PPT 또는 첫 장 크기 (기본)';
  $('output-help').textContent=source==='pptx'?'지원되는 현재 글자·도형·표·차트를 HTML로 저장합니다. 복잡한 서식과 미지원 객체는 분석에서 알려드립니다.':'지원되는 글자·도형·표·차트를 PowerPoint에서 수정할 수 있습니다. 확인이 필요한 항목은 분석 결과에 표시합니다.';
  $('quality-note').textContent=source==='pptx'?'이 미리보기는 변환된 HTML입니다. 원본 PowerPoint 화면과 차이가 있을 수 있습니다.':'이 미리보기는 입력 HTML입니다. 함께 제공한 지원 정적 TTF·WOFF2 글꼴은 PPT에 포함합니다. 포함되지 않은 글꼴은 사용할 컴퓨터에 설치해야 합니다.';
}
function status(state,message){document.body.dataset.phase=state;$('status-badge').dataset.state=state;$('status-badge').textContent=labels[state];$('status-message').textContent=message;}
function clearDownload(){$('download-button').disabled=true;$('download-button').hidden=false;$('download-button').textContent=downloadLabel();$('partial-download-button').hidden=true;$('partial-download-button').disabled=true;$('partial-download-note').hidden=true;}
function jobActions(){$('cancel-button').disabled=!busy;$('delete-button').disabled=(!currentJob&&!cleanupJobs.size&&!cleanupRequests.size)||busy;}
function rememberCleanup(job){if(job?.id)cleanupJobs.add(job.id);}
function rememberRequest(){if(activeRequestId){cleanupRequests.add(activeRequestId);activeRequestId=null;}}
function cleanupOutstanding(){
  if(cleanupPromise)return cleanupPromise;
  cleanupPromise=(async()=>{let removed=false;while(cleanupRequests.size||cleanupJobs.size){const request=cleanupRequests.size>0,id=request?cleanupRequests.values().next().value:cleanupJobs.values().next().value,response=await apiPost(request?`/api/requests/${id}`:`/api/jobs/${id}`,undefined,undefined,'DELETE');if(!response.ok&&response.status!==404)throw new Error('이전 작업을 정리하지 못했습니다. 다시 시도해 주세요.');if(response.ok)removed=true;(request?cleanupRequests:cleanupJobs).delete(id);}return removed;})().then(removed=>{$('cleanup-note').hidden=true;return removed;}).catch(error=>{$('cleanup-note').textContent='이전 작업을 삭제하지 못했습니다. 작업 삭제를 누르거나 다시 분석해 주세요.';$('cleanup-note').hidden=false;throw error;}).finally(()=>{cleanupPromise=null;jobActions();});
  return cleanupPromise;
}
function reset(){revision++;rememberRequest();controller?.abort();rememberCleanup(currentJob);currentJob=null;if(cleanupJobs.size||cleanupRequests.size)void cleanupOutstanding().catch(()=>{});busy=false;selectedChoiceId=null;selectableJobId=null;$('region-reselect-button').hidden=true;$('slide-size-report').hidden=true;$('slide-size-details').hidden=true;$('slide-size-report').textContent='';showPagination();durationNotice(null);jobActions();const file=$('html-file').files[0],selection=selectedFileState(file),type=selection.type;$('analyze-button').disabled=!selection.valid;$('analyze-button').textContent='다시 분석';describeDirection();clearDownload();$('region-area').hidden=true;$('region-map-image').removeAttribute('src');$('preview-area').hidden=true;$('quality-note').hidden=true;$('empty-preview').hidden=false;$('empty-preview').querySelector('p').textContent='파일을 분석하면 미리보기가 표시됩니다.';$('diagnostics').hidden=true;$('diagnostics').open=false;$('diagnostic-list').replaceChildren();$('preview-image').removeAttribute('src');$('file-label').textContent=file?.name??'파일을 여기에 놓거나 선택하세요';status(file&&!selection.valid?'failed':selection.valid?'ready':'idle',selection.message);}
$('html-file').addEventListener('change',()=>{reset();if(selectedFileState($('html-file').files[0]).valid)$('analyze-form').requestSubmit();});
for(const id of ['assets','asset-folder','html-page-mode','selector','pptx-font-policy','pptx-allow-partial','slide-size-policy','slide-width','slide-height'])$(id).addEventListener('change',()=>{$('slide-size-custom').hidden=$('slide-size-policy').value!=='custom';reset();if(selectedFileState($('html-file').files[0]).valid)status('ready','설정이 바뀌었습니다. 다시 분석을 눌러 적용하세요.');});
function selectedSizePolicy(){
 const mode=$('slide-size-policy').value,policy={mode:'contain'};
 if(mode==='wide')Object.assign(policy,{targetWidthEmu:12192000,targetHeightEmu:6858000});
 if(mode==='standard')Object.assign(policy,{targetWidthEmu:9144000,targetHeightEmu:6858000});
 if(mode==='custom'){const w=Number($('slide-width').value),h=Number($('slide-height').value);if(!Number.isFinite(w)||!Number.isFinite(h)||w<1||h<1||w>56||h>56)throw new Error('PPT 너비와 높이는 각각 1~56인치여야 합니다.');Object.assign(policy,{targetWidthEmu:Math.round(w*914400),targetHeightEmu:Math.round(h*914400)});}
 return policy;
}
const dropArea=document.querySelector('.file-area');
for(const event of ['dragenter','dragover'])dropArea.addEventListener(event,e=>{e.preventDefault();dropArea.classList.add('dragging');});
for(const event of ['dragleave','drop'])dropArea.addEventListener(event,e=>{e.preventDefault();dropArea.classList.remove('dragging');});
dropArea.addEventListener('drop',e=>{if(e.dataTransfer.files.length!==1){$('html-file').value='';reset();status('failed','HTML 또는 PPTX 파일 한 개를 선택해 주세요.');return;}$('html-file').files=e.dataTransfer.files;$('html-file').dispatchEvent(new Event('change',{bubbles:true}));});
const base64=buffer=>{const bytes=new Uint8Array(buffer);let binary='';for(let i=0;i<bytes.length;i+=32768)binary+=String.fromCharCode(...bytes.subarray(i,i+32768));return btoa(binary);};
function showSlide(index,diagnostic=null){
 const slide=currentJob.slides[index];if(!slide?.previewUrl)return;
 $('quality-note').hidden=false;
 $('preview-image').src=slide.previewUrl;$('preview-image').alt=`${index+1}번 HTML 미리보기`;$('slide-detail').textContent=`${index+1} / ${currentJob.slideCount}장`;
 const overlay=$('problem-overlay'),note=$('problem-location-note');overlay.hidden=true;note.hidden=!diagnostic;note.textContent='';
 if(diagnostic){if(validBounds(diagnostic.bounds)){const b=diagnostic.bounds;Object.assign(overlay.style,{left:`${b.x*100}%`,top:`${b.y*100}%`,width:`${b.w*100}%`,height:`${b.h*100}%`});overlay.hidden=false;note.textContent=`HTML 미리보기 · ${diagnostic.objectName??'선택한 객체'}의 기록된 위치입니다. 원본 PowerPoint 화면은 아닙니다.`;}else note.textContent='이 객체의 위치 정보는 확인할 수 없습니다. 해당 장표의 HTML 미리보기만 표시합니다.';}
 for(const [i,button]of [...$('slide-list').children].entries())button.setAttribute('aria-current',String(i===index));
}
function resultMessages(job){
  const visible=job.slides.some(s=>s.previewUrl),pptx=job.sourceType==='pptx',failed=job.status==='failed';
  const inputStopped=pptx&&failed&&!job.slides.length&&!job.exportAvailable&&!job.partialExportAvailable;
  const htmlAvailable=pptx&&job.exportAvailable;
  const replaced=htmlAvailable&&(job.diagnostics??[]).some(d=>d.code==='PPTX_OBJECT_PLACEHOLDER');
  return {
    status:inputStopped?'입력 분석 중단: HTML과 미리보기 생성 안 됨. 아래 진단에서 원인을 확인하세요.':replaced?`HTML을 다운로드할 수 있지만 일부 객체가 경고 표시로 대체됩니다.${visible?' 미리보기와 진단을 확인하세요.':' 미리보기 렌더링도 실패했습니다. 아래 진단을 확인하세요.'}`:htmlAvailable&&!visible?'HTML 변환은 가능합니다. 미리보기 렌더링만 실패했습니다.':failed?(pptx?'변환을 막는 항목이 있습니다. 아래 진단에서 해당 장표와 원인을 확인하세요.':'아래 항목을 수정하거나 장표 영역을 지정한 뒤 다시 분석하세요.'):`${job.fileName}에서 장표 ${job.slideCount}장을 확인했습니다.`,
    empty:inputStopped?'입력 분석 중단: HTML과 미리보기 생성 안 됨. 원본 PowerPoint 이미지는 제공되지 않습니다.':replaced?'미리보기 렌더링에 실패했습니다. 다운로드할 HTML의 일부 객체도 경고 표시로 대체됩니다. 아래 진단을 확인하세요.':htmlAvailable?'미리보기 렌더링에 실패했지만 HTML은 다운로드할 수 있습니다. 다운로드 후 배치를 확인하세요.':job.partialExportAvailable?'일부 장표는 다운로드할 수 있지만 미리보기는 없습니다. 아래 진단과 제외 장표를 확인하세요.':'미리보기를 만들지 못했습니다. 아래 원인별 진단을 확인하세요. 원본 PowerPoint 이미지는 제공되지 않습니다.'
  };
}
const uncertaintyText={
  'hidden-content-not-observed':'숨겨진 내용은 확인하지 못했습니다.',
  'clipped-or-scroll-content-not-observed':'잘리거나 내부 스크롤 뒤에 있는 내용은 모두 확인하지 못했습니다.',
  'script-content-not-observed':'스크립트로 나중에 생성되는 내용은 확인하지 못했습니다.',
  'external-resources-blocked':'외부에서 불러오는 내용은 확인하지 못했습니다.',
  'generated-content-not-fully-observed':'CSS로 생성된 내용은 완전히 확인하지 못했습니다.',
  'css-compositing-not-fully-observed':'일부 겹침과 CSS 효과는 완전히 확인하지 못했습니다.',
  'render-pixels-not-fully-observed':'전체 화면을 한 번에 확인하지 못했습니다.',
  'layout-may-change':'화면 배치가 바뀔 수 있어 다시 확인해야 합니다.',
  'resource-not-ready':'이미지나 글꼴을 모두 확인하지 못했습니다.',
  'descendant-extends-outside-candidate':'후보 밖으로 나온 내용이 있습니다.',
  'overlapping-elements':'화면에서 겹치는 후보가 있습니다.',
  'repeated-layout-not-fully-observed':'반복된 영역의 구성을 모두 비교하지 못했습니다.',
  'small-elements-not-proposed':'작은 글자 요소는 장표 후보로 따로 표시하지 않았습니다.',
  'objects-inside-declared-slide-not-separate-candidates':'이미 확인된 장표 안의 객체는 장표 후보로 따로 표시하지 않았습니다.',
  'embedded-documents-have-separate-viewports':'각 페이지는 서로 다른 화면에서 확인했습니다.'
};
const reasonText={
  'root-size-mismatch':'장표 영역에 content-box와 패딩 또는 테두리가 함께 있습니다. box-sizing:border-box를 사용하거나 패딩과 테두리를 제거해 주세요.',
  'conversion-selection-not-connected':'이 범위는 아직 선택해 변환할 수 없습니다.',
  'different-size':'크기가 다른 영역을 함께 변환하는 기능은 아직 지원하지 않습니다.',
  'different-sizes':'크기가 다른 영역을 함께 변환하는 기능은 아직 지원하지 않습니다.',
  'hidden-or-unmeasured':'숨겨져 있거나 위치를 확인할 수 없는 영역입니다.',
  'size-not-fixed':'영역의 크기가 고정되지 않아 아직 변환할 수 없습니다.',
  'overlap-or-nesting':'겹치거나 포함 관계인 영역을 함께 변환할 수 없습니다.',
  'document-size-not-fixed':'문서의 크기가 고정되지 않아 아직 전체를 변환할 수 없습니다.',
  'document-bounds-do-not-match':'문서 크기와 실제 표시 범위가 달라 전체를 변환할 수 없습니다.',
  'profile-not-supported':'이 형식의 선택 범위는 아직 변환할 수 없습니다.',
  'too-many-slides':'한 번에 변환할 수 있는 장표 수를 초과했습니다.',
  'selection-too-large':'선택한 영역이 너무 커서 한 번에 확인할 수 없습니다.',
  'inspection-incomplete':'화면이나 자산을 완전히 확인하지 못해 이 영역을 변환할 수 없습니다. 다시 분석해 주세요.',
  'ancestor-appearance-not-supported':'이 영역의 상위 요소에 있는 배경이나 효과를 함께 보존할 수 없습니다.',
  'external-overlay-not-supported':'이 영역 위에 겹친 바깥 요소나 효과를 함께 보존할 수 없습니다.',
  'root-appearance-not-supported':'선택 영역의 배경이나 테두리 효과를 보존할 수 없습니다.',
  'document-scope-not-supported':'이 문서 전체를 한 장에 담을 수 없습니다.',
  'hidden':'숨겨진 영역은 변환할 수 없습니다.',
  'overlapping':'겹친 영역을 함께 변환할 수 없습니다.'
};
const regionBox=b=>b&&[b.x,b.y,b.width,b.height].every(Number.isFinite)&&b.width>0&&b.height>0;
function hideRegions(){$('region-area').hidden=true;$('region-map-image').removeAttribute('src');selectedChoiceId=null;}
function showRegions(job){
  showPagination();
  currentJob=job;selectedChoiceId=null;clearDownload();durationNotice(null);jobActions();
  $('slide-size-report').hidden=true;$('slide-size-details').hidden=true;
  $('region-reselect-button').hidden=true;
  $('preview-area').hidden=true;$('quality-note').hidden=true;$('empty-preview').hidden=true;$('diagnostics').hidden=true;
  $('region-area').hidden=false;status('selection-required','변환할 장표 범위를 선택해 주세요.');
  const inspection=job.inspection,regions=inspection.regions??[],choices=inspection.choices??[];
  const names=new Map(regions.map((r,i)=>[r.id,`영역 ${i+1}`]));
  const byId=new Map(regions.map(r=>[r.id,r]));
  const offscreen=regions.filter(r=>r.visibility==='offscreen').length,hidden=regions.filter(r=>r.visibility==='hidden').length;
  $('region-summary').textContent=`영역 ${regions.length}개 · 화면 밖 ${offscreen}개 · 숨김 ${hidden}개`;
  const suggested=inspection.suggestedViewport,fitButton=$('region-fit-button');
  fitButton.hidden=!suggested;
  if(suggested)fitButton.textContent=`장표 크기(${suggested.width} × ${suggested.height}px)로 다시 분석`;
  const notices=(inspection.uncertainties??[]).map(code=>uncertaintyText[code]??'일부 내용이나 배치는 확인하지 못했습니다.');
  $('region-uncertainties').textContent=[...new Set(notices)].join(' ');$('region-uncertainties').hidden=!notices.length;
  const list=$('region-list'),overlays=$('region-map-overlays');list.replaceChildren();overlays.replaceChildren();
  const selectChoice=id=>{if(busy)return;const choice=choices.find(c=>c.id===id);if(!choice?.convertible)return;selectedChoiceId=id;
    for(const button of list.querySelectorAll('button'))button.setAttribute('aria-pressed',String(button.dataset.choiceId===id));
    for(const button of overlays.querySelectorAll('button'))button.setAttribute('aria-pressed',String(choice.regionIds.includes(button.dataset.regionId)));
    $('region-select-button').disabled=false;
  };
  for(const choice of choices){
    const selected=(choice.regionIds??[]).map(id=>byId.get(id)).filter(Boolean),button=document.createElement('button');
    button.type='button';button.className='region-choice';button.dataset.choiceId=choice.id;button.setAttribute('aria-pressed','false');button.disabled=!choice.convertible;
    const title=document.createElement('strong');title.textContent=choice.scope==='document'?'문서 전체를 한 장으로 만들기':selected.length>1?`묶음 · ${selected.length}개 영역`:names.get(choice.regionIds?.[0])??'영역';button.append(title);
    const details=[];if(selected.length===1){const r=selected[0];if(r.parentId)details.push(`${names.get(r.parentId)??'다른 영역'} 안`);if(r.overlaps?.length)details.push(`다른 후보와 ${r.overlaps.length}개 겹침`);if(r.visibility==='offscreen')details.push('화면 밖');if(r.visibility==='hidden')details.push('숨김');if(r.visibility==='unmeasured')details.push('위치 확인 불가');}
    if(choice.evidence?.includes('repeated-sibling'))details.push('반복된 영역');
    if(Number.isFinite(choice.outsideContentCount))details.push(`범위 밖 내용 ${choice.outsideContentCountAccuracy==='estimate'?'약 ':''}${choice.outsideContentCount}개`);
    if(details.length){const detail=document.createElement('small');detail.textContent=details.join(' · ');button.append(detail);}
    if(!choice.convertible){const reason=document.createElement('small');reason.className='region-reason';reason.textContent=(choice.reasons??[]).map(r=>reasonText[r]??'이 범위는 아직 변환할 수 없습니다.').join(' ')||'이 범위는 아직 변환할 수 없습니다.';button.append(reason);}
    button.addEventListener('click',()=>selectChoice(choice.id));list.append(button);
  }
  const doc=choices.find(c=>c.scope==='document'&&regionBox(c.bounds))?.bounds;
  const canShow=Boolean(job.previewUrl&&doc&&doc.x===0&&doc.y===0);
  $('region-map-stage').hidden=!canShow;$('region-preview-note').hidden=false;
  $('region-preview-note').textContent=canShow?'1280 × 960 화면을 기준으로 조사한 전체 내용입니다. 화면 밖에 이미 배치된 영역도 표시합니다.':'전체 미리보기를 확인할 수 없습니다. 아래 목록에서 범위를 선택하세요.';
  const picture=$('region-map-image');if(canShow){picture.src=job.previewUrl;picture.onerror=()=>{$('region-map-stage').hidden=true;$('region-preview-note').hidden=false;$('region-preview-note').textContent='미리보기를 불러오지 못했습니다. 아래 목록에서 범위를 선택하세요.';};}else picture.removeAttribute('src');
  if(canShow)for(const [i,r]of regions.entries()){
    if(!regionBox(r.bounds)||r.bounds.x<0||r.bounds.y<0||r.bounds.x+r.bounds.width>doc.width+.01||r.bounds.y+r.bounds.height>doc.height+.01)continue;
    const choice=choices.find(c=>c.scope==='elements'&&c.regionIds?.length===1&&c.regionIds[0]===r.id)??choices.find(c=>c.scope==='elements'&&c.regionIds?.includes(r.id));
    const button=document.createElement('button');button.type='button';button.dataset.regionId=r.id;button.disabled=!choice?.convertible;button.setAttribute('aria-label',`${names.get(r.id)} 선택`);button.setAttribute('aria-pressed','false');
    Object.assign(button.style,{left:`${r.bounds.x/doc.width*100}%`,top:`${r.bounds.y/doc.height*100}%`,width:`${r.bounds.width/doc.width*100}%`,height:`${r.bounds.height/doc.height*100}%`});
    const number=document.createElement('span');number.textContent=String(i+1);button.append(number);button.addEventListener('click',()=>selectChoice(choice?.id));overlays.append(button);
  }
  $('region-select-button').disabled=true;
}
function showPagination(report){
  const notice=$('html-pagination-report'),details=$('html-pagination-details'),list=$('html-pagination-list');
  notice.hidden=details.hidden=!report;
  notice.textContent='';
  list.replaceChildren();
  if(!report){details.open=false;return;}
  const counts=[];
  if(report.editableTableCount>0)counts.push(`편집 가능한 표 조각 ${report.editableTableCount}개`);
  if(report.imageCount>0)counts.push(`그림 ${report.imageCount}개`);
  notice.textContent=`문서 전체를 ${report.pages.length}장으로 나눴습니다. 줄마다 편집 가능한 텍스트 상자를 만듭니다.${counts.length?' '+counts.join(' · '):''}`;
  const format=new Intl.NumberFormat('ko-KR',{maximumFractionDigits:2});
  for(const {fromY,toY} of report.pages){
    const item=document.createElement('li');
    item.textContent=`${format.format(fromY)}–${format.format(toY)}px`;
    list.append(item);
  }
}
function showResult(job){
  hideRegions();
  $('region-reselect-button').hidden=selectableJobId!==job.id;
  durationNotice(job.slideCount);
  currentJob=job;const messages=resultMessages(job);status(job.status,messages.status);
  describeDirection();showPagination(job.pagination);$('download-button').textContent=downloadLabel();
  const policy=job.sizePolicy,report=$('slide-size-report');report.hidden=!policy;
  if(policy){const source=job.pagination?.sizeSource==='default-16:9'?'문서 분할 기본 크기 16:9':({original:'원본 PPT 크기', 'first-slide':'첫 장 크기',custom:'사용자 선택'}[policy.source]??'선택한 크기');const inch=n=>Number((n/914400).toFixed(2)),color=(policy.marginColor??'#FFFFFF').toUpperCase();report.textContent=`PPT 크기 ${inch(policy.targetWidthEmu)} × ${inch(policy.targetHeightEmu)}인치 (${source}) · 비율 유지 · 가운데 배치 · 여백 ${color==='#FFFFFF'?'흰색':color}${policy.originalAspectMismatch?' · 원본 PPT 크기 정보와 장표 비율이 다름':''}`;const list=$('slide-size-list');list.replaceChildren();for(const [i,s] of (policy.slides??[]).entries()){const item=document.createElement('li'),scale=Number.isFinite(s.physicalScale)?`PPT 물리 크기 ${Math.round(s.physicalScale*100)}% (CSS 96dpi 기준)`:'PPT 물리 배율 확인 불가',margin=Number.isFinite(s.offsetXEmu)&&Number.isFinite(s.offsetYEmu)?`좌우 여백 각 ${inch(s.offsetXEmu)}인치, 위아래 여백 각 ${inch(s.offsetYEmu)}인치`:'여백 크기 확인 불가';item.textContent=`${i+1}번: ${scale}, ${margin}`;list.append(item);}$('slide-size-details').hidden=!list.children.length;}else $('slide-size-details').hidden=true;
  if(job.previewKind==='source')$('quality-note').textContent='입력 HTML 미리보기입니다. PPT의 크기와 여백은 위 안내를 확인해 주세요.';
  jobActions();
  $('download-button').disabled=!job.exportAvailable;$('download-button').hidden=Boolean(job.partialExportAvailable);
  $('partial-download-button').hidden=!job.partialExportAvailable;$('partial-download-button').disabled=!job.partialExportAvailable;$('partial-download-note').hidden=!job.partialExportAvailable;
  if(job.partialExportAvailable){$('partial-download-button').textContent=`변환 가능한 ${job.partialSlideCount}장만 받기`;$('partial-download-note').textContent=`전체 ${job.slideCount}장 중 오류가 있는 ${job.excludedSlideNumbers.join(', ')}번 장표는 빠집니다. 원래 순서를 유지합니다.`;}
  $('diagnostic-list').replaceChildren();$('diagnostics').hidden=!job.diagnostics.length;
  renderDiagnostics({document,container:$('diagnostic-list'),summary:$('diagnostic-summary'),diagnostics:job.diagnostics,slides:job.slides,onSelectSlide:(index,diagnostic)=>{showSlide(index,diagnostic);$('preview-area').scrollIntoView({block:'nearest',behavior:'smooth'});}});
  const visible=job.slides.some(s=>s.previewUrl);$('empty-preview').querySelector('p').textContent=visible?'장표 미리보기':job.pagination&&job.exportAvailable?'PPT를 다운로드해 장표별 배치를 확인할 수 있습니다.':messages.empty;$('preview-area').hidden=!visible;$('quality-note').hidden=!visible;$('empty-preview').hidden=visible;
  if(visible){$('slide-count').textContent=`${job.slideCount}장`;$('object-count').textContent=`${targetLabel()}로 저장`;$('slide-list').replaceChildren();job.slides.forEach((s,i)=>{const b=document.createElement('button');b.type='button';b.textContent=String(i+1);b.setAttribute('aria-label',`${i+1}번 장표`);b.addEventListener('click',()=>showSlide(i));$('slide-list').append(b);});showSlide(job.slides.findIndex(s=>s.previewUrl));}
}
async function analyzeFile(viewport){
  const file=$('html-file').files[0];if(!file||$('analyze-button').disabled)return;
  showPagination();
  const thisRevision=++revision,previousJob=currentJob;controller?.abort();controller=new AbortController();const signal=controller.signal;selectableJobId=null;$('region-reselect-button').hidden=true;$('slide-size-report').hidden=true;$('slide-size-details').hidden=true;busy=true;jobActions();clearDownload();hideRegions();$('preview-area').hidden=true;$('quality-note').hidden=true;$('empty-preview').hidden=false;$('empty-preview').querySelector('p').textContent='입력을 분석하고 있습니다.';$('preview-image').removeAttribute('src');$('analyze-button').disabled=true;$('analyze-button').textContent='장표를 분석하고 있습니다…';status('analyzing','글꼴과 이미지를 읽고 실제 장표 배치를 확인하고 있습니다.');$('diagnostics').hidden=true;$('diagnostics').open=false;
  const noticeTimer=setTimeout(()=>{if(thisRevision===revision&&!signal.aborted&&$('duration-note').hidden){$('duration-note').textContent='장표를 분석하는 데 시간이 걸리고 있어요. 잠시만 기다려 주세요.';$('duration-note').hidden=false;}},20000);
  try{
    rememberCleanup(previousJob);await cleanupOutstanding();
    if(thisRevision!==revision||signal.aborted)return;
    currentJob=null;
    const type=sourceType(file);if(!type)throw new Error('HTML 또는 PPTX 파일을 선택해 주세요.');
    const selection=selectedFileState(file);if(!selection.valid)throw new Error(selection.message);
    let input;
    if(type==='pptx')input={sourceType:'pptx',pptxBase64:base64(await file.arrayBuffer()),fileName:file.name,pptxOptions:{fontPolicy:$('pptx-font-policy').value,allowPartial:$('pptx-allow-partial').checked}};
    else{
      const files=[...$('assets').files,...$('asset-folder').files];if(files.reduce((sum,f)=>sum+f.size,0)>64*1024*1024)throw new Error('자산의 전체 크기는 64MiB 이하여야 합니다.');
      const html=new TextDecoder('utf-8',{fatal:true,ignoreBOM:true}).decode(await file.arrayBuffer());
      if(thisRevision===revision&&!signal.aborted)durationNotice(embeddedPageCount(html));
      const assets=await Promise.all(files.map(async f=>({path:f.webkitRelativePath||f.name,mime:f.type,base64:base64(await f.arrayBuffer())})));
      const pagination=paginateDocument();
      input={sourceType:'html',html,assets,fileName:file.name,sizePolicy:selectedSizePolicy(),...(viewport?{viewport}:{}),
        ...(pagination?{pageMode:'paginate'}:{slideSelector:$('selector').value.trim()||null})};
    }
    if(thisRevision!==revision||signal.aborted)return;
    const {data:job}=await postJob('/api/analyze',input,signal);
    if(thisRevision!==revision||signal.aborted){rememberCleanup(job);if(cleanupJobs.size||cleanupRequests.size)void cleanupOutstanding().catch(()=>{});return;}
    if(!job.id)throw new Error(job.error??'분석 결과를 받지 못했습니다.');
    if(type==='html'&&input.pageMode!=='paginate'&&(job.diagnostics??[]).some(d=>d.code==='SELECTION_REQUIRED')){
      currentJob=job;jobActions();status('analyzing','장표로 사용할 수 있는 영역을 찾고 있습니다.');
      const removed=await apiPost(`/api/jobs/${job.id}`,undefined,signal,'DELETE');if(thisRevision!==revision||signal.aborted)return;
      if(!removed.ok&&removed.status!==404)throw new Error('이전 분석을 정리하지 못했습니다. 다시 시도해 주세요.');currentJob=null;
      const inspectionInput={html:input.html,assets:input.assets,fileName:input.fileName,slideSelector:input.slideSelector,sizePolicy:input.sizePolicy,...(viewport?{viewport}:{})};
      const {response:inspected,data:candidateJob}=await postJob('/api/inspect',inspectionInput,signal);if(thisRevision!==revision||signal.aborted){rememberCleanup(candidateJob);if(cleanupJobs.size||cleanupRequests.size)void cleanupOutstanding().catch(()=>{});return;}
      if(!inspected.ok||!candidateJob.id||!candidateJob.inspection)throw new Error(candidateJob.error??'장표 후보를 확인하지 못했습니다. 다시 분석해 주세요.');
      showRegions(candidateJob);
    }else showResult(job);
  }catch(error){if(error.name!=='AbortError'&&thisRevision===revision){status('failed',error.message);$('empty-preview').querySelector('p').textContent='분석을 마치지 못했습니다. 위 안내를 확인한 뒤 다시 시도해 주세요.';}}
  finally{clearTimeout(noticeTimer);if(thisRevision===revision){busy=false;jobActions();$('analyze-button').disabled=false;$('analyze-button').textContent='입력 다시 분석';}}
}
$('analyze-form').addEventListener('submit',event=>{event.preventDefault();void analyzeFile();});
$('region-fit-button').addEventListener('click',()=>{
  const viewport=currentJob?.inspection?.suggestedViewport;
  if(viewport&&!busy)void analyzeFile(viewport);
});
$('region-select-button').addEventListener('click',async()=>{
  const job=currentJob,choiceId=selectedChoiceId;if(!job?.inspection||!choiceId||busy)return;
  const thisRevision=revision;controller?.abort();controller=new AbortController();const signal=controller.signal;
  busy=true;jobActions();clearDownload();$('region-select-button').disabled=true;$('analyze-button').disabled=true;
  $('diagnostics').hidden=true;$('diagnostics').open=false;$('diagnostic-list').replaceChildren();
  status('selecting','선택한 영역이 그대로인지 다시 확인하고 있습니다.');
  try{
    const response=await apiPost(`/api/jobs/${job.id}/select`,{snapshotId:job.inspection.snapshotId,choiceId},signal),result=await response.json();
    if(thisRevision!==revision||signal.aborted)return;
    if(!response.ok&&result.id&&result.status==='failed'&&Array.isArray(result.diagnostics)){
      $('diagnostics').hidden=!result.diagnostics.length;
      renderDiagnostics({document,container:$('diagnostic-list'),summary:$('diagnostic-summary'),diagnostics:result.diagnostics});
      status('selection-required','선택한 영역을 분석하지 못했습니다. 아래 진단을 확인하거나 다른 영역을 선택해 주세요.');return;
    }
    if(!response.ok||!result.id)throw new Error(result.error??'영역을 다시 확인하지 못했습니다. 다시 분석해 주세요.');
    selectableJobId=result.id;showResult(result);
  }catch(error){if(error.name!=='AbortError'&&thisRevision===revision)status('selection-required',error.message);}
  finally{if(thisRevision===revision){busy=false;jobActions();$('analyze-button').disabled=false;if(currentJob?.inspection)$('region-select-button').disabled=!selectedChoiceId;}}
});
$('region-reselect-button').addEventListener('click',()=>{if(!busy)$('analyze-form').requestSubmit();});
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
    rememberRequest();const hadJobs=Boolean(job||cleanupJobs.size||cleanupRequests.size);let deleted=false;
    // Keep the request open until DELETE confirms that the worker has stopped.
    rememberCleanup(job);if(hadJobs)deleted=await cleanupOutstanding();
    request?.abort();if(thisRevision!==revision)return;
    currentJob=null;selectableJobId=null;$('region-reselect-button').hidden=true;$('slide-size-report').hidden=true;$('slide-size-details').hidden=true;showPagination();hideRegions();$('preview-area').hidden=true;$('quality-note').hidden=true;$('empty-preview').hidden=false;$('empty-preview').querySelector('p').textContent='입력을 분석하고 있습니다.';$('preview-image').removeAttribute('src');$('diagnostics').hidden=true;$('diagnostic-list').replaceChildren();
    status(deleted?'deleted':'cancelled',deleted?'서버의 작업과 결과를 삭제했습니다. 이미 다운로드한 파일은 남아 있습니다.':hadJobs?'작업을 찾을 수 없습니다. 이미 만료되었거나 정리 중일 수 있습니다.':'취소 요청을 보냈습니다. 서버에서 진행 중인 작업을 정리합니다.');
  }catch(error){if(thisRevision===revision)status('failed',error.message);}
  finally{if(thisRevision===revision){jobActions();$('analyze-button').disabled=!selectedFileState($('html-file').files[0]).valid;$('analyze-button').textContent='입력 다시 분석';}}
}
$('cancel-button').addEventListener('click',removeCurrentJob);
$('delete-button').addEventListener('click',removeCurrentJob);
