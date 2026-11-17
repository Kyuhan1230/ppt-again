const rangeTolerance=0.000001;

// Report only geometry and fixed policy values. Text, node IDs and source
// selectors stay with the worker result.
export function publicPaginationReport(report,slideCount){
  if(report?.mode!=='paginate'||report.oversizePolicy!=='stop'||report.textEditable!==true||
     report.textObjectGranularity!=='line'||!Array.isArray(report.pages)||
     report.pages.length!==slideCount||slideCount<1||slideCount>200)return null;

  const {sourceWidth,sourceHeight,pageHeight}=report;
  if(![sourceWidth,sourceHeight,pageHeight].every(n=>Number.isFinite(n)&&n>0))return null;

  let nextY=0;
  const pages=[];
  for(const page of report.pages){
    const {fromY,toY}=page??{};
    if(!Number.isFinite(fromY)||!Number.isFinite(toY)||fromY!==nextY||
       toY<=fromY||toY>sourceHeight||toY-fromY>pageHeight+rangeTolerance)return null;
    pages.push({number:pages.length+1,fromY,toY});
    nextY=toY;
  }
  if(nextY!==sourceHeight)return null;

  const counts={};
  for(const key of ['editableTableCount','imageCount']){
    if(Number.isSafeInteger(report[key])&&report[key]>=0)counts[key]=report[key];
  }
  const sizeSource=['default-16:9','configured'].includes(report.sizeSource)?report.sizeSource:null;
  return {mode:'paginate',sourceWidth,sourceHeight,pageHeight,oversizePolicy:'stop',
    textEditable:true,textObjectGranularity:'line',pages,...(sizeSource?{sizeSource}:{}),...counts};
}
