export const FILE_LIMIT_BYTES=Object.freeze({html:16*1024*1024,pptx:64*1024*1024});
export const sourceType=file=>/\.pptx$/i.test(file?.name??'')?'pptx':/\.html?$/i.test(file?.name??'')?'html':null;
export function selectedFileState(file){
 const type=sourceType(file),limit=FILE_LIMIT_BYTES[type];
 if(!file)return {type:null,valid:false,message:'파일을 선택하면 자동으로 분석합니다.'};
 if(!type)return {type:null,valid:false,message:'HTML(.html·.htm) 또는 PPTX 파일을 선택해 주세요.'};
 if(!Number.isSafeInteger(file.size)||file.size<0)return {type,valid:false,message:'파일 크기를 확인할 수 없습니다. 다시 선택해 주세요.'};
 if(file.size>limit)return {type,valid:false,message:`파일 크기 ${file.size.toLocaleString('ko-KR')}바이트 (${(file.size/1024/1024).toFixed(2)}MiB) · ${type.toUpperCase()} 최대 ${limit/1024/1024}MiB (${limit.toLocaleString('ko-KR')}바이트)를 초과했습니다. 더 작은 파일을 선택하세요.`};
 return {type,valid:true,message:'파일을 선택하면 자동으로 분석합니다.'};
}
