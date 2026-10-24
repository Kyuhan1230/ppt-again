const messages={
  INPUT_INVALID:'HTML 파일의 형식과 크기를 확인해 주세요.',
  SLIDE_ROOT_REQUIRED:'장표 영역을 찾지 못했습니다. 크기가 정해진 장표 HTML을 사용하거나 장표 선택자를 지정해 주세요.',
  ASSET_MISSING:'필요한 이미지나 글꼴 파일을 찾지 못했습니다. HTML에서 사용하는 자산을 함께 선택해 주세요.',
  FONT_UNAVAILABLE:'사용할 수 없는 글꼴 또는 굵기가 있습니다. 해당 글꼴을 설치하거나 지원되는 글꼴로 바꿔 주세요.',
  SCRIPT_REQUIRED:'스크립트 실행이 필요한 HTML입니다. 장표 내용이 포함된 정적 HTML로 저장해 주세요.',
  UNSUPPORTED_OBJECT:'지원하지 않는 표현이 있습니다. 해당 장표의 복잡한 효과를 제거하거나 기본 글자와 도형으로 바꿔 주세요.',
  RENDER_INCOMPLETE:'장표를 끝까지 읽지 못했습니다. 파일과 자산을 확인한 뒤 다시 시도해 주세요.',
  EXPORT_FAILED:'PPT를 만들지 못했습니다. 입력 파일을 확인한 뒤 다시 시도해 주세요.'
};
const warnings={
  FONT_UNAVAILABLE:'일부 글꼴 굵기를 설치된 굵기에 맞춰 출력합니다. PowerPoint에서 글자와 줄바꿈을 확인해 주세요.',
  UNSUPPORTED_OBJECT:'일부 효과는 그림으로 출력되어 내부 요소를 개별 수정할 수 없습니다. 다운로드한 PPT를 확인해 주세요.',
  SCRIPT_REQUIRED:'HTML에 포함된 스크립트는 실행하지 않습니다. 필요한 내용이 미리보기에 있는지 확인해 주세요.'
};

// Only fixed messages cross the UI boundary. Input text, selectors and exception
// details remain in the developer record, never in a customer diagnostic.
export function publicDiagnostic(item,slides=[]){
  const code=Object.hasOwn(messages,item.code)?item.code:'RENDER_INCOMPLETE';
  const severity=item.severity==='warning'?'warning':'error';
  const index=slides.findIndex(s=>s.id===item.slideId);
  return {code,severity,message:(severity==='warning'&&warnings[code])||messages[code],slideNumber:index<0?null:index+1};
}
