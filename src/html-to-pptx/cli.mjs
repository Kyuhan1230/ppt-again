#!/usr/bin/env node
import fs from 'node:fs/promises';
import path from 'node:path';
import {parseArgs} from 'node:util';
import {randomUUID} from 'node:crypto';
import {runConversion} from './job-runner.mjs';
import {publicDiagnostic} from './public-diagnostics.mjs';
import {publicPaginationReport} from './pagination-report.mjs';
import {assetPath} from './assets.mjs';
import {writeErrorLog} from './error-log.mjs';

const MiB=1024*1024;
const logContext={operation:'cli',jobId:randomUUID()};
const help=`Usage: ppt-again convert <input.html|input.pptx> --output <output.pptx|output.html> [--selector <css>] [--assets <directory>] [--viewport <width>x<height>] [--slide-size auto|16:9|4:3|WIDTHxHEIGHT] [--margin-color '#RRGGBB'] [--inspection <file> --region <id> | --scope document --page-mode single|paginate] [--json] [--font-policy original|malgun-gothic|pretendard] [--allow-partial]\n       PPT 크기 직접 지정 단위: 인치 (각 변 1~56). 기본: 원본 PPT 크기, 없으면 첫 장 크기. 비율 유지, 가운데 정렬, 흰색 여백.\n       문서 전체를 여러 장으로 나누려면 --scope document --page-mode paginate를 지정하세요. 크기 기본값은 16:9입니다.\n       ppt-again inspect <input.html> [--assets <directory>] [--viewport <width>x<height>] [--output <file>] [--json]\n       ppt-again --help | --version`;
const fail=(code,message,exit=1)=>Object.assign(new Error(message),{code,exit});
async function readFileLimited(file,max){
  const handle=await fs.open(file,'r');
  try{const stat=await handle.stat();if(!stat.isFile()||stat.size>max)throw fail('INPUT_LIMIT','입력 파일의 크기 또는 형식을 확인해 주세요.');const chunks=[];let total=0;while(true){const chunk=Buffer.allocUnsafe(Math.min(65536,max+1-total));const {bytesRead}=await handle.read(chunk,0,chunk.length,null);if(!bytesRead)break;total+=bytesRead;if(total>max)throw fail('INPUT_LIMIT','입력 파일이 너무 큽니다.');chunks.push(chunk.subarray(0,bytesRead));}return Buffer.concat(chunks,total);}finally{await handle.close();}
}
async function readAssets(dir){
  const root=await fs.realpath(dir),rootStat=await fs.lstat(dir);if(rootStat.isSymbolicLink()||!rootStat.isDirectory())throw fail('ASSET_INVALID','자산 폴더를 확인해 주세요.');
  const assets=[],names=new Set();let total=0,entryCount=0;
  async function visit(folder,relative=''){
    const entries=[];for await(const entry of await fs.opendir(folder)){if(++entryCount>1024)throw fail('ASSET_LIMIT','자산 폴더는 최대 1024개 항목까지 사용할 수 있습니다.');entries.push(entry);}entries.sort((a,b)=>a.name.localeCompare(b.name));
    for(const entry of entries){
      const rel=relative?`${relative}/${entry.name}`:entry.name;const normalized=assetPath(rel);if(normalized!==rel||names.has(normalized))throw fail('ASSET_INVALID','자산 경로가 올바르지 않습니다.');names.add(normalized);
      const file=path.join(folder,entry.name),stat=await fs.lstat(file);if(stat.isSymbolicLink())throw fail('ASSET_INVALID','자산의 심볼릭 링크는 사용할 수 없습니다.');
      const real=await fs.realpath(file),inside=path.relative(root,real);if(inside.startsWith('..'+path.sep)||inside==='..'||path.isAbsolute(inside))throw fail('ASSET_INVALID','자산 경로가 폴더를 벗어났습니다.');
      if(stat.isDirectory()){await visit(file,rel);continue;}
      if(!stat.isFile()||stat.size>16*MiB||(total+=stat.size)>64*MiB)throw fail('ASSET_INVALID','자산 형식 또는 크기를 확인해 주세요.');
      const bytes=await readFileLimited(file,16*MiB);total+=bytes.length-stat.size;if(total>64*MiB)throw fail('ASSET_LIMIT','자산의 총 크기가 64MiB를 초과합니다.');assets.push({path:rel,bytes});
    }
  }
  await visit(root);return assets;
}
async function refuseOverwrite(output){
 try{await fs.lstat(output);throw fail('OUTPUT_EXISTS','출력 파일이 이미 존재합니다. 다른 경로를 지정해 주세요.');}catch(e){if(e.code!=='ENOENT')throw e;}
}
async function writeNew(output,data){
 let handle;try{handle=await fs.open(output,'wx');await handle.writeFile(data);await handle.close();handle=null;}catch(e){if(handle){await handle.close().catch(()=>{});await fs.unlink(output).catch(()=>{});}if(e.code==='EEXIST')throw fail('OUTPUT_EXISTS','출력 파일이 이미 존재합니다.');throw e;}
}
function parseViewport(value){
 if(value===undefined)return undefined;
 const match=/^(\d{1,4})x(\d{1,4})$/i.exec(value);if(!match)throw fail('USAGE',help,2);
 const viewport={width:Number(match[1]),height:Number(match[2])};
 if(viewport.width<1||viewport.height<1||viewport.width>8192||viewport.height>8192)throw fail('USAGE',help,2);
 return viewport;
}
function parseSizePolicy(value,color){
 const policy={mode:'contain'};
 if(value===undefined||value==='auto'){}
 else if(value==='16:9')Object.assign(policy,{targetWidthEmu:12192000,targetHeightEmu:6858000});
 else if(value==='4:3')Object.assign(policy,{targetWidthEmu:9144000,targetHeightEmu:6858000});
 else{const match=/^(\d+(?:\.\d+)?)x(\d+(?:\.\d+)?)$/i.exec(value);if(!match)throw fail('USAGE',help,2);const w=Number(match[1]),h=Number(match[2]);if(w<1||h<1||w>56||h>56)throw fail('USAGE',help,2);Object.assign(policy,{targetWidthEmu:Math.round(w*914400),targetHeightEmu:Math.round(h*914400)});}
 if(color!==undefined){if(!/^#[a-f0-9]{6}$/i.test(color))throw fail('USAGE',help,2);policy.marginColor=color;}
 return policy;
}
function paginationText(report){
 const format=new Intl.NumberFormat('ko-KR',{maximumFractionDigits:2});
 const pages=report.pages.map(({number,fromY,toY})=>`${number}번 장표 · ${format.format(fromY)}–${format.format(toY)}px`);
 const counts=[];
 if(report.editableTableCount>0)counts.push(`편집 가능한 표 조각 ${report.editableTableCount}개`);
 if(report.imageCount>0)counts.push(`그림 ${report.imageCount}개`);
 return '\n'+[
  `문서 전체를 ${report.pages.length}장으로 나눴습니다. 줄마다 편집 가능한 텍스트 상자를 만들었습니다.`,
  ...(counts.length?[counts.join(' · ')]:[]),
  '한 장보다 큰 객체나 안전하게 나눌 수 없는 경계에서는 변환을 중단합니다.',
  ...pages
 ].join('\n');
}
function candidatesText(inspection){
 const evidence={'explicit-selector':'직접 지정','declared-slide':'장표 표식','validated-svg':'검증된 SVG','validated-embedded-slide':'내장 장표','validated-embedded-slides':'내장 장표 묶음','repeated-sibling':'비슷한 반복 영역','visible-content':'보이는 내용','background':'배경','hidden':'숨김','display-contents':'박스 없는 컨테이너','document':'문서 전체'};
 const reason={'selection-not-supported':'현재 변환기가 지원하지 않는 범위','conversion-selection-not-connected':'아직 변환할 수 없는 범위','profile-not-supported':'이 입력 형식의 후보 선택은 지원하지 않음','too-many-slides':'선택한 장표가 너무 많음','hidden-or-unmeasured':'숨겨져 있거나 크기를 확인할 수 없음','size-not-fixed':'영역 크기가 고정되어 있지 않음','selection-too-large':'선택 영역이 지원 크기를 초과함','inspection-incomplete':'화면을 충분히 확인하지 못함','root-size-mismatch':'장표 영역에 content-box와 패딩 또는 테두리가 함께 있습니다. box-sizing:border-box를 사용하거나 패딩과 테두리를 제거해 주세요.','root-appearance-not-supported':'선택 영역의 배경이나 테두리 효과를 보존할 수 없음','ancestor-appearance-not-supported':'부모 배경이나 효과를 보존할 수 없음','external-overlay-not-supported':'영역에 겹친 외부 요소나 효과를 보존할 수 없음','different-sizes':'장표마다 크기가 다름','overlap-or-nesting':'선택 영역이 겹치거나 포함 관계임','document-size-not-fixed':'문서 크기가 고정되어 있지 않음','document-bounds-do-not-match':'문서 크기와 표시된 내용의 범위가 다름'};
 const uncertainty={'hidden-content-not-observed':'숨겨진 내용','clipped-or-scroll-content-not-observed':'잘리거나 내부 스크롤에 가려진 내용','css-compositing-not-fully-observed':'일부 CSS 효과','layout-may-change':'변할 수 있는 화면 배치','generated-content-not-fully-observed':'CSS로 생성된 내용','repeated-layout-not-fully-observed':'일부 반복 배치','overlapping-elements':'서로 겹치는 영역','descendant-extends-outside-candidate':'영역 밖으로 돌출한 자손','render-pixels-not-fully-observed':'일부 화면 이미지','resource-not-ready':'아직 준비되지 않은 이미지 또는 글꼴','script-content-not-observed':'스크립트로 생성되는 내용','external-resources-blocked':'차단된 외부 자원','small-elements-not-proposed':'후보로 제시하지 않은 작은 요소','objects-inside-declared-slide-not-separate-candidates':'기존 장표 내부 객체는 별도 후보에서 제외','embedded-documents-have-separate-viewports':'내장 장표가 각각 별도 화면을 사용함'};
 const visibility={visible:'화면 안',offscreen:'화면 밖',hidden:'숨김',unmeasured:'위치 확인 불가'};
 const regions=new Map((inspection.regions??[]).map(region=>[region.id,region]));
 const rows=inspection.choices.map((choice,index)=>{
  const members=choice.regionIds.map(id=>{const r=regions.get(id);if(!r)return id;const b=r.bounds;return `${id} ${visibility[r.visibility]??r.visibility}${b?` (${Math.round(b.x)}, ${Math.round(b.y)}, ${Math.round(b.width)}×${Math.round(b.height)})`:''}${r.parentId?` 부모 ${r.parentId}`:''}${r.overlaps?.length?` 겹침 ${r.overlaps.join(', ')}`:''}`;});
  const why=(choice.evidence??[]).map(item=>evidence[item]??item).join(', ')||'영역';
  const outside=Number.isFinite(choice.outsideContentCount)?` · 범위 밖 내용 약 ${choice.outsideContentCount}개`:'';
  const unavailable=choice.convertible?'':` · 변환 불가: ${(choice.reasons??[]).map(item=>reason[item]??item).join(', ')||'지원하지 않는 범위'}`;
  return `${index+1}. ${choice.id} — ${why}${outside}${unavailable}${members.length?'\n   '+members.join('\n   '):''}`;
 });
 return `장표 후보 ${rows.length}개\n${rows.join('\n')}${inspection.uncertainties?.length?'\n확인 한계: '+inspection.uncertainties.map(item=>uncertainty[item]??item).join(', '):''}`;
}
async function main(){
 let values,positionals;try{({values,positionals}=parseArgs({allowPositionals:true,options:{output:{type:'string'},selector:{type:'string'},assets:{type:'string'},json:{type:'boolean'},help:{type:'boolean'},version:{type:'boolean'},'font-policy':{type:'string'},'allow-partial':{type:'boolean'},viewport:{type:'string'},inspection:{type:'string'},region:{type:'string'},scope:{type:'string'},'page-mode':{type:'string'},'slide-size':{type:'string'},'margin-color':{type:'string'}}}));}catch{throw fail('USAGE',help,2);}
 if(values.help&&positionals.length===0){console.log(help);return;}
 if(values.version&&positionals.length===0){const pkg=JSON.parse(await fs.readFile(new URL('../../package.json',import.meta.url),'utf8'));console.log(pkg.version);return;}
 if(positionals.length!==2||!['convert','inspect'].includes(positionals[0])||values.help||values.version)throw fail('USAGE',help,2);
 const command=positionals[0],inputPath=positionals[1],output=values.output,ext=path.extname(inputPath).toLowerCase(),outExt=output?path.extname(output).toLowerCase():null;
 const sourceType=ext==='.pptx'?'pptx':ext==='.html'||ext==='.htm'?'html':null,targetType=sourceType==='html'?'pptx':'html';
 const viewport=parseViewport(values.viewport);
 if(!sourceType||command==='inspect'&&sourceType!=='html'||command==='convert'&&!output)throw fail('USAGE',help,2);
 if(command==='inspect'&&(values.selector!==undefined||values.inspection!==undefined||values.region!==undefined||values.scope!==undefined||values['page-mode']!==undefined||values['font-policy']!==undefined||values['allow-partial']!==undefined||values['slide-size']!==undefined||values['margin-color']!==undefined))throw fail('USAGE',help,2);
 if((values.inspection===undefined)!==(values.region===undefined)||values.region!==undefined&&!values.region.trim())throw fail('USAGE',help,2);
 const pageMode=values['page-mode'],paginating=pageMode==='paginate';
 if((values.scope===undefined)!==(pageMode===undefined)||values.scope!==undefined&&(values.scope!=='document'||!['single','paginate'].includes(pageMode)))throw fail('USAGE',help,2);
 if([values.selector!==undefined,values.inspection!==undefined,values.scope!==undefined].filter(Boolean).length>1)throw fail('USAGE',help,2);
 if(sourceType==='html'&&(values['font-policy']!==undefined||values['allow-partial']!==undefined)||values['font-policy']!==undefined&&!['original','malgun-gothic','pretendard'].includes(values['font-policy']))throw fail('USAGE',help,2);
 if(command==='convert'&&outExt!==`.${targetType}`||(sourceType==='pptx'&&(values.selector!==undefined||values.assets!==undefined||viewport||values.inspection!==undefined||values.scope!==undefined||values['slide-size']!==undefined||values['margin-color']!==undefined))||values.selector!==undefined&&!values.selector.trim())throw fail('USAGE',help,2);
 const sizePolicy=sourceType==='html'&&command==='convert'?parseSizePolicy(values['slide-size'],values['margin-color']):undefined;
 if(output)await refuseOverwrite(output);
 let bytes;try{bytes=await readFileLimited(inputPath,(sourceType==='html'?16:64)*MiB);}catch(e){if(e.code==='ENOENT')throw fail('USAGE','입력 파일을 찾을 수 없습니다.',2);throw e;}
 const input=sourceType==='html'?{html:bytes.toString('utf8'),assets:values.assets?await readAssets(values.assets):[],slideSelector:values.selector??null,allowTableOverlay:true,...(viewport?{viewport}:{}),...(sizePolicy?{sizePolicy}:{}),...(paginating?{pageMode}:{})}:{bytes,options:{fontPolicy:values['font-policy']??'original',allowPartial:values['allow-partial']??false}};
 const controller=new AbortController(),abort=()=>controller.abort();process.on('SIGINT',abort);process.on('SIGTERM',abort);
 try{
  Object.assign(logContext,{operation:sourceType==='html'?'analyze-html':'analyze-pptx',sourceType,targetType,inputBytes:bytes.byteLength});
  const inspect=async()=>{logContext.operation='inspect-html';const result=await runConversion({operation:'inspect-html',input,signal:controller.signal});logContext.resourceUsage=result.resourceUsage;return result.inspection;};
  if(command==='inspect'){
   const inspection=await inspect();if(output)await writeNew(output,JSON.stringify(inspection,null,2)+'\n');
   console.log(values.json?JSON.stringify(inspection):candidatesText(inspection));return;
  }
  let selection;
  if(values.inspection!==undefined){
   let saved;try{saved=JSON.parse((await readFileLimited(values.inspection,16*MiB)).toString('utf8'));}catch(e){if(e.code==='INPUT_LIMIT')throw e;throw fail('INPUT_INVALID','후보 조사 파일을 읽을 수 없습니다.');}
   if(typeof saved?.snapshotId!=='string'||saved.snapshotId.length>128||!values.region||values.region.length>128)throw fail('INPUT_INVALID','후보 조사 파일을 확인해 주세요.');
   selection={snapshotId:saved.snapshotId,choiceId:values.region};
  }else if(values.scope==='document'&&pageMode==='single'){
   const inspection=await inspect(),choice=inspection.choices.find(item=>item.scope==='document');
   if(!choice?.convertible)throw Object.assign(fail('UNSUPPORTED_SCOPE','이 문서 범위는 아직 변환할 수 없습니다.'),{inspection});
   selection={snapshotId:inspection.snapshotId,choiceId:choice.id};
  }
  logContext.operation=selection?'analyze-html-selection':sourceType==='html'?'analyze-html':'analyze-pptx';
  const {measured,resourceUsage}=await runConversion({operation:logContext.operation,input:selection?{...input,...selection}:input,signal:controller.signal});
  logContext.resourceUsage=resourceUsage;
  let diagnostics=measured.diagnostics??[];
  if(diagnostics.some(d=>d.severity==='error')){
   const code=diagnostics.find(d=>d.severity==='error').code;
   const inspection=code==='SELECTION_REQUIRED'&&sourceType==='html'&&!paginating?await inspect():undefined;
   throw Object.assign(fail(code,code==='SELECTION_REQUIRED'?'장표 후보를 확인하고 범위를 선택해 주세요.':'변환할 수 없는 내용이 있습니다.'),{diagnostics,slides:measured.model?.slides,inspection});
  }
  const previewFailures=diagnostics.filter(item=>item.code==='PPTX_PREVIEW_UNAVAILABLE');
  if(previewFailures.length)await writeErrorLog({...logContext,operation:'preview-pptx',diagnostics:previewFailures,slides:measured.model?.slides});
  const {previews,...data}=measured;
  const pagination=paginating?publicPaginationReport(data.pagination,data.model?.slides.length??0):null;
  if(paginating&&!pagination)throw fail('PAGINATION_INVALID','분할한 문서 범위를 확인할 수 없습니다.');
  logContext.operation=targetType==='pptx'?'export-pptx':'export-html';
  logContext.resourceUsage=undefined;
  const result=await runConversion({operation:logContext.operation,input:targetType==='pptx'?data:data.model,signal:controller.signal});
  logContext.resourceUsage=result.resourceUsage;
  diagnostics=[...new Map([...diagnostics,...(result.diagnostics??[])].map(d=>[JSON.stringify(d),d])).values()];
  if(result.status!=='generated'||diagnostics.some(d=>d.severity==='error'))throw Object.assign(fail(diagnostics.find(d=>d.severity==='error')?.code??'EXPORT_FAILED','변환 파일을 만들지 못했습니다.'),{diagnostics,slides:data.model?.slides});
  if(controller.signal.aborted)throw fail('JOB_CANCELLED','변환이 취소됐습니다.');
  logContext.operation='cli';
  await writeNew(output,targetType==='pptx'?result.pptxBytes:result.html);
  const policy=data.sizePolicy?{...data.sizePolicy,slides:(data.sizePolicy.slides??[]).map((s,i)=>({...s,slideId:`slide-${i+1}`}))}:undefined;
  const report={status:'generated',output,sourceType,targetType,...(policy?{sizePolicy:policy}:{}),...(pagination?{pagination}:{}),diagnostics:diagnostics.map(d=>publicDiagnostic(d,data.model?.slides))};
  const sizeSource=pagination?.sizeSource==='default-16:9'?'문서 분할 기본 크기 16:9':({original:'원본 PPT', 'first-slide':'첫 장',custom:'직접 지정'})[policy?.source];
  const sizeText=policy?`\nPPT 크기 ${(policy.targetWidthEmu/914400).toFixed(2)} × ${(policy.targetHeightEmu/914400).toFixed(2)}인치 (${sizeSource}) · 비율 유지 · 가운데 배치 · 여백 ${policy.marginColor}${policy.originalAspectMismatch?' · 원본 PPT 크기 정보와 장표 비율이 다름':''}`+(policy.slides??[]).map((s,i)=>`\n${i+1}번: PPT 물리 크기 ${Math.round(s.physicalScale*100)}% (CSS 96dpi 기준), 좌우 여백 각 ${(s.offsetXEmu/914400).toFixed(2)}인치, 위아래 여백 각 ${(s.offsetYEmu/914400).toFixed(2)}인치`).join(''):'';
  console.log(values.json?JSON.stringify(report):`변환 완료: ${output}${sizeText}${pagination?paginationText(pagination):''}${report.diagnostics.length?'\n'+report.diagnostics.map(d=>d.message).join('\n'):''}`);
 }finally{process.off('SIGINT',abort);process.off('SIGTERM',abort);}
}
main().catch(async e=>{
  await writeErrorLog({...logContext,error:e,diagnostics:e.diagnostics,slides:e.slides,resourceUsage:e.resourceUsage??logContext.resourceUsage});
  const message={STALE_INSPECTION:'입력 또는 화면이 달라졌거나 이전 선택을 확인할 수 없습니다. 후보를 다시 조사해 주세요.',REGION_UNSUPPORTED:'선택한 범위는 아직 변환할 수 없습니다.',UNSUPPORTED_SCOPE:'이 문서 범위는 아직 변환할 수 없습니다.',SELECTION_REQUIRED:'장표 후보를 확인하고 범위를 선택해 주세요.'}[e.code];
  const report={status:'failed',output:null,code:e.code??'IO_ERROR',...(message?{message}:{}),diagnostics:(e.diagnostics??[]).map(d=>publicDiagnostic(d,e.slides)),...(e.inspection?{inspection:e.inspection}:{})};
  if(process.argv.includes('--json'))console.log(JSON.stringify(report));
  else console.error(`${e.code&&e.exit?e.message:e.code==='STALE_INSPECTION'?'화면 또는 입력이 달라졌습니다. 후보를 다시 조사해 주세요.':e.code==='REGION_UNSUPPORTED'?'선택한 범위는 아직 변환할 수 없습니다. 다른 후보를 선택해 주세요.':'변환에 실패했습니다. 입력과 출력 경로를 확인해 주세요.'}${report.diagnostics.length?'\n'+report.diagnostics.map(d=>d.message).join('\n'):''}${e.inspection?'\n'+candidatesText(e.inspection):''}`);
  process.exitCode=e.exit??1;
});
