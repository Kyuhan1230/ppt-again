const messages={
  INPUT_INVALID:'입력 파일의 형식과 크기를 확인해 주세요.',
  CHART_DATA_REQUIRED:'차트의 원자료가 필요합니다. 차트 영역과 같은 ID의 JSON 데이터를 함께 넣어 주세요.',
  CHART_DATA_MISMATCH:'차트 표시와 원자료가 일치하지 않습니다. 원자료와 차트 표시를 일치시킨 뒤 다시 저장해 주세요.',
  LIMIT_EXCEEDED:'문서의 글자나 차트 데이터가 지원 한도를 초과합니다. 내용을 나누어 변환해 주세요.',
  SLIDE_ROOT_REQUIRED:'장표 영역을 찾지 못했습니다. 크기가 정해진 장표 HTML을 사용하거나 장표 선택자를 지정해 주세요.',
  ASSET_MISSING:'필요한 이미지나 글꼴 파일을 찾지 못했습니다. HTML에서 사용하는 자산을 함께 선택해 주세요.',
  FONT_UNAVAILABLE:'사용할 수 없는 글꼴 또는 굵기가 있습니다. 해당 글꼴을 설치하거나 지원되는 글꼴로 바꿔 주세요.',
  FONT_SUBSTITUTED:'원래 글꼴을 사용할 수 없어 Pretendard로 바꿨습니다. 줄바꿈과 배치를 확인해 주세요.',
  FONT_PREVIEW_SUBSTITUTED:'이 컴퓨터에 없는 글꼴이 있어 미리보기에서 다른 글꼴로 표시됩니다. HTML에는 원래 글꼴 이름을 유지하며 줄바꿈과 배치를 확인해 주세요.',
  FONT_NOT_PORTABLE:'글꼴 파일은 출력에 포함하지 않습니다. 다른 PC에서 수정하려면 같은 글꼴이 필요합니다.',
  FONT_INVALID:'글꼴 파일을 읽을 수 없습니다. 손상되지 않은 지원 글꼴 파일을 다시 선택해 주세요.',
  FONT_FORMAT_UNSUPPORTED:'이 글꼴 형식은 아직 PPT에 포함할 수 없습니다. 지원되는 정적 TTF 파일을 제공해 주세요.',
  FONT_VARIABLE_UNSUPPORTED:'가변 글꼴은 아직 포함할 수 없습니다. 같은 글꼴의 정적 TTF 파일을 제공해 주세요.',
  FONT_STYLE_UNSUPPORTED:'지원하지 않는 글꼴 굵기·기울임·폭입니다. 일반 또는 굵게의 정적 글꼴을 사용해 주세요.',
  FONT_STYLE_MISMATCH:'문서의 굵기·기울임과 제공 글꼴 파일이 다릅니다. 맞는 스타일의 파일을 선택해 주세요.',
  FONT_GLYPH_MISSING:'제공 글꼴에 문서의 일부 글자가 없습니다. 해당 문자를 포함한 글꼴 파일을 선택해 주세요.',
  FONT_SOURCE_UNSUPPORTED:'사용한 글꼴 파일을 확정할 수 없습니다. 스타일마다 한 개의 글꼴 파일을 제공해 주세요.',
  FONT_SUBSET_UNSUPPORTED:'여러 조각으로 나뉜 글꼴 또는 Unicode 범위는 아직 지원하지 않습니다. 필요한 문자를 포함한 단일 글꼴 파일을 제공해 주세요.',
  FONT_ASSET_LIMIT:'글꼴은 최대 32개, 전체 32MiB까지 사용할 수 있습니다. 사용하지 않는 글꼴을 제거해 주세요.',
  FONT_EMBEDDING_RESTRICTED:'PPT에 포함할 권한이 없는 글꼴입니다. 문서에 포함할 수 있는 글꼴을 사용해 주세요.',
  FONT_EMBEDDING_INVALID:'글꼴의 포함 권한 정보가 올바르지 않습니다. 권한을 확인할 수 있는 글꼴을 제공해 주세요.',
  FONT_EMBEDDING_UNSUPPORTED:'이 글꼴의 포함 권한을 확인할 수 없습니다. 지원되는 다른 글꼴을 제공해 주세요.',
  FONT_PREVIEW_ONLY:'이 글꼴은 보기·인쇄만 허용되어 편집용 PPT에 포함할 수 없습니다. 편집용 포함이 허용된 글꼴을 사용해 주세요.',
  FONT_BITMAP_ONLY:'그림 형태만 허용된 글꼴은 편집용 PPT에 포함할 수 없습니다. 윤곽 글꼴 포함이 허용된 파일을 사용해 주세요.',
  FONT_SUBSETTING_RESTRICTED:'이 글꼴은 일부 문자만 포함할 수 없습니다. 전체 글꼴 파일을 사용해 주세요.',
  TABLE_LAYOUT_APPROXIMATED:'복잡한 표의 배경은 그림으로, 문구는 별도 글자 상자로 변환합니다. PPT 표 셀로 편집할 수 없으며, 배경과 글자는 함께 이동하지 않습니다.',
  PARAGRAPH_SPACING_APPROXIMATED:'문단 앞뒤 간격은 그대로 옮기지 못했습니다. HTML과 다시 만든 PPT의 문단 간격을 확인해 주세요.',
  SVG_TEXT_APPROXIMATED:'글자 너비를 비슷하게 맞췄습니다. PPT에서 글자 모양과 겹침을 확인해 주세요.',
  SCRIPT_REQUIRED:'스크립트 실행이 필요한 HTML입니다. 장표 내용이 포함된 정적 HTML로 저장해 주세요.',
  UNSUPPORTED_OBJECT:'지원하지 않는 표현이 있습니다. 해당 장표의 복잡한 효과를 제거하거나 기본 글자와 도형으로 바꿔 주세요.',
  SVG_TEXT_NOT_EDITABLE:'장표 안의 SVG 글자를 편집 가능한 글자로 읽는 기능은 아직 지원하지 않습니다. 해당 글자를 HTML 글자로 바꾼 정적 장표를 사용해 주세요.',
  RENDER_INCOMPLETE:'장표를 끝까지 읽지 못했습니다. 파일과 자산을 확인한 뒤 다시 시도해 주세요.',
  EXPORT_FAILED:'변환 파일을 만들지 못했습니다. 입력 파일을 확인한 뒤 다시 시도해 주세요.'
};
const warnings={
  FONT_UNAVAILABLE:'일부 글꼴 굵기를 설치된 굵기에 맞춰 출력합니다. PowerPoint에서 글자와 줄바꿈을 확인해 주세요.',
  FONT_FORMAT_UNSUPPORTED:'이 글꼴은 PPT 파일에 포함되지 않습니다. 이 경고만으로 변환이 막히지는 않지만 다른 컴퓨터에서는 글꼴이나 줄바꿈이 달라질 수 있습니다.',
  UNSUPPORTED_OBJECT:'일부 효과는 그림으로 출력되어 내부 요소를 개별 수정할 수 없습니다. 다운로드한 PPT를 확인해 주세요.',
  SCRIPT_REQUIRED:'HTML에 포함된 스크립트는 실행하지 않습니다. 필요한 내용이 미리보기에 있는지 확인해 주세요.'
};
const knownObjectErrors=new Map([
  ['SVG image 객체는 아직 편집 가능한 객체로 변환할 수 없습니다.','SVG 안의 이미지를 아직 PPT 객체로 변환하지 못합니다.'],
  ['회전·개별 글자 위치·여러 줄 위치가 지정된 SVG 글자는 아직 지원하지 않습니다.','SVG의 회전되거나 글자별 위치가 지정된 텍스트를 아직 변환하지 못합니다.'],
  ['중첩 SVG의 글자·참조·효과 또는 뷰포트는 아직 보존할 수 없습니다.','겹쳐진 SVG 안의 글자 또는 효과를 아직 보존하지 못합니다.'],
  ['중첩 표·그림·입력 요소·캡션은 일반 표에 넣을 수 없습니다.','표 안의 그림이나 중첩된 요소를 아직 편집 가능한 표로 변환하지 못합니다.'],
  ['CSS ::after의 생성 콘텐츠는 명시적인 HTML 요소로 넣어 주세요.','CSS로 만들어진 화면 내용은 아직 변환하지 못합니다.']
]);

// Only fixed messages cross the UI boundary. Input text, selectors and exception
// details remain in the developer record, never in a customer diagnostic.
export function publicDiagnostic(item,slides=[]){
  const code=Object.hasOwn(messages,item.code)?item.code:'RENDER_INCOMPLETE';
  const severity=item.severity==='warning'?'warning':'error';
  const index=slides.findIndex(s=>s.id===item.slideId);
  const known=code==='UNSUPPORTED_OBJECT'&&severity==='error'?knownObjectErrors.get(item.message):null;
  return {code,severity,message:known||(severity==='warning'&&warnings[code])||messages[code],slideNumber:index<0?null:index+1};
}
