const messages={
  PPTX_FONT_REPLACED:'선택한 글꼴로 HTML 문구를 바꿨습니다. 글꼴 파일은 포함하지 않으며 줄바꿈과 배치가 달라질 수 있습니다.',
  PPTX_OBJECT_PLACEHOLDER:'지원하지 않는 내용을 표시 상자로 남겼습니다. 원본과 비교해 보완해 주세요.',
  SOURCE_PLACEHOLDER_PRESENT:'이 HTML에는 원본에서 변환하지 못한 객체의 표시가 있습니다. 다시 만든 PPT에는 상세 설명이 포함되지 않으므로 원본과 비교해 보완해 주세요.',
  PPTX_PREVIEW_UNAVAILABLE:'미리보기를 만들지 못했지만 읽어 낸 내용을 HTML로 내려받을 수 있습니다. 배치를 확인해 주세요.',
  PPTX_PREVIEW_LAYOUT_APPROXIMATED:'미리보기 일부 배치가 원본과 다를 수 있습니다. 원본과 비교해 주세요.',
  PPTX_GROUP_SCALE_APPROXIMATED:'배율이 적용된 그룹의 위치와 크기는 옮겼지만 글자 크기와 선 굵기, 셀 여백은 원본과 다를 수 있습니다.',
  PPTX_SHAPE_IMAGE_APPROXIMATED:'HTML에서는 SVG 도형을 편집할 수 있지만, 다시 만든 PPT에서는 그림으로 저장되어 개별 편집할 수 없습니다.',
  PPTX_IMAGE_APPEARANCE_APPROXIMATED:'사진은 보존하지만 색상·자유형 마스크·그림자 등 일부 효과가 생략되어 원본과 다를 수 있습니다.',
  PPTX_IMAGE_DUOTONE_BAKED:'듀오톤 색상 효과를 그림에 적용했습니다. 효과를 별도 설정으로 편집할 수 없습니다.',
  PPTX_ANIMATION_OMITTED:'애니메이션과 화면 전환은 정적 HTML에서 재현되지 않습니다.',
  PPTX_MATH_IMAGE_FALLBACK:'수식을 그림으로 보존했으며 수식 자체는 편집할 수 없습니다. 배치가 원본과 다를 수 있습니다.',
  PPTX_MATH_TABLE_BORDER_OMITTED:'수식 그림에 포함되지 않은 표 테두리는 빠질 수 있습니다. 원본과 비교해 주세요.',
  PPTX_TABLE_STYLE_APPROXIMATED:'표 내용은 보존했지만 일부 셀의 배경·테두리 또는 서식은 원본과 다를 수 있습니다.',
  PPTX_SHADOW_OMITTED:'바깥 그림자는 표시하지 않았습니다. 원본과 비교해 주세요.',
  PPTX_SHADOW_APPROXIMATED:'바깥 그림자는 CSS로 근사했습니다. PPT로 다시 저장할 때는 빠질 수 있습니다.',
  PPTX_ARROWHEAD_OMITTED:'삼각형 화살표 끝은 표시하지 않았습니다. 원본과 비교해 주세요.',

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
  TABLE_HEIGHT_EXPANDED:'표가 셀 내용에 맞춰 늘어나 표시된 높이로 저장됩니다. 원본과 비교해 주세요.',
  PARAGRAPH_SPACING_APPROXIMATED:'문단 들여쓰기나 앞뒤 간격이 일부 달라질 수 있습니다. 변환 결과를 확인해 주세요.',
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

// Fixed reasons plus narrowly bounded labels cross the UI boundary. Never
// forward exception text, selectors, source paths, or an arbitrary message.
const reasons={
  HTML_PPTX_PARAGRAPH_LAYOUT_APPROXIMATED:'PPT로 다시 저장할 때 문단 들여쓰기·간격·정렬이 일부 달라질 수 있습니다.',
  HTML_PPTX_SHADOW_OMITTED:'HTML에 표시된 바깥 그림자는 PPT로 다시 저장할 때 빠집니다.',
  PPTX_CHART_DATA_MISMATCH:'차트 표시 데이터와 연결된 내장 원자료가 일치하지 않습니다.',
  PPTX_CHART_STRUCTURE_INVALID:'차트 또는 내장 원자료의 구조가 올바르지 않아 검증을 완료하지 못했습니다.',
  PPTX_CHART_LIMIT_EXCEEDED:'차트 데이터가 처리 한도를 초과했습니다.',
  PPTX_CHART_UNSUPPORTED:'이 차트는 현재 지원 범위에서 안전하게 검증하지 못해 부분 변환할 수 없습니다.',
  PPTX_CHART_VALIDATION_FAILED:'차트 또는 내장 원자료 검증에 실패했습니다.',
  PPTX_CHART_PRESENTATION_UNSUPPORTED:'차트 데이터와 내장 연결은 검증했지만 이 차트의 스타일 또는 막대 간격은 아직 재현하지 못합니다.',
  PPTX_INVALID_COORDINATES:'객체의 위치·크기·회전 수치가 없거나 올바르지 않습니다.',
  PPTX_INVALID_CROP:'그림의 자르기 비율이 유효 범위를 벗어났습니다.',
  PPTX_UNUSED_EMBEDDED_PACKAGE:'현재 차트에 연결되지 않은 내장 문서가 있어 안전하게 가져올 수 없습니다.',
  PPTX_UNEXPECTED_PACKAGE_RELATIONSHIP:'검증된 차트 데이터 외의 내장 패키지 연결이 있어 가져올 수 없습니다.',
  PPTX_XML_INVALID:'PPTX 내부 XML 문서의 구조가 손상되었거나 올바르지 않습니다.',
  PPTX_REQUIRED_PART_MISSING:'PPTX에서 참조하는 필수 구성 파일이 없습니다.',
  PPTX_XML_UNSAFE:'DTD 또는 엔터티 선언이 포함되어 안전하게 가져올 수 없습니다.',
  PPTX_EXTERNAL_REFERENCE:'외부 파일 또는 네트워크를 참조하는 연결이 있어 가져올 수 없습니다.',
  PPTX_MODEL_INVALID:'읽어 낸 장표 구조가 검증 조건을 충족하지 못했습니다.',
  PPTX_PARSE_FATAL:'PPTX 내부 구성요소를 안전하게 읽지 못했습니다. 알려진 지원 제한만으로 원인을 확정할 수 없습니다.',
  PPTX_INHERITED_TEXT_UNRESOLVED:'마스터·레이아웃에서 상속된 글꼴 또는 글자 크기를 아직 해석하지 못합니다. 첨자나 취소선 때문이라는 뜻은 아닙니다.',
  PPTX_PARAGRAPH_INDENT_APPROXIMATED:'문단 들여쓰기와 글머리 배치를 원본과 똑같이 옮기지 못했습니다. 글머리 기호의 글꼴과 색상도 달라질 수 있습니다. HTML의 줄바꿈과 배치를 확인해 주세요.',
  PPTX_LINE_SPACING_APPROXIMATED:'줄 간격 0은 기본 간격으로 표시했습니다. 원본과 비교해 주세요.',
  PPTX_TEXT_BASELINE_APPROXIMATED:'위첨자·아래첨자는 일반 글자로 표시했습니다.',
  PPTX_PARAGRAPH_ALIGNMENT_APPROXIMATED:'서로 다른 문단 정렬을 첫 문단 정렬로 표시했습니다.',
  PPTX_INHERITED_SIZE_APPROXIMATED:'상속 글자 크기를 확인하지 못해 18pt로 표시했습니다. 원본과 비교해 주세요.',
  PPTX_AUTONUM_APPROXIMATED:'자동 번호를 글자로 옮겼습니다. 번호와 들여쓰기 배치를 확인해 주세요.',
  PPTX_TEXT_TAB_APPROXIMATED:'탭은 유지했지만 위치가 원본과 다를 수 있습니다.',
  PPTX_TEXT_HIGHLIGHT_APPROXIMATED:'강조 색을 표시하지 않았습니다. 원본을 확인해 주세요.',
  PPTX_TABLE_TEXT_COLOR_APPROXIMATED:'표의 교차 셀 글자색은 원본과 다를 수 있습니다.',
  PPTX_TABLE_STYLE_APPROXIMATED:'표 내용은 보존했지만 일부 셀의 배경·테두리 또는 서식은 원본과 다를 수 있습니다.',
  PPTX_SHADOW_OMITTED:'바깥 그림자는 표시하지 않았습니다. 원본과 비교해 주세요.',
  PPTX_SHADOW_APPROXIMATED:'바깥 그림자는 CSS로 근사했습니다. PPT로 다시 저장할 때는 빠질 수 있습니다.',
  PPTX_ARROWHEAD_OMITTED:'삼각형 화살표 끝은 표시하지 않았습니다. 원본과 비교해 주세요.',
  PPTX_THEME_COLOR_UNSUPPORTED:'테마 또는 시스템 색상 참조를 명시적인 색으로 아직 해석하지 못합니다.',
  PPTX_INHERITED_SHAPE_STYLE_UNSUPPORTED:'테마에서 상속한 도형 채우기·선·효과를 아직 해석하지 못합니다.',
  PPTX_ZERO_SIZE_OBJECT_UNSUPPORTED:'너비나 높이가 0인 객체는 보이는 내용이 없다고 단정할 수 없습니다. 부분 변환에서는 표시 상자로 남깁니다.',

  PPTX_THEME_BACKGROUND_UNSUPPORTED:'단색이 아닌 테마 배경 또는 색상 효과를 아직 옮기지 못합니다.',
  PPTX_MASTER_ARTWORK_UNSUPPORTED:'마스터에 들어 있는 그림이나 도형을 아직 옮기지 못합니다.',
  PPTX_ANIMATION_UNSUPPORTED:'애니메이션이나 화면 전환은 HTML로 재현하지 않습니다.',
  PPTX_OBJECT_ROTATION_UNSUPPORTED:'회전 또는 뒤집힌 객체를 아직 옮기지 못합니다.',
  PPTX_OBJECT_POSITION_UNSUPPORTED:'객체의 좌표 또는 크기를 지원 범위에서 읽지 못했습니다.',
  PPTX_SHAPE_GRADIENT_UNSUPPORTED:'도형의 그라데이션 채우기를 아직 옮기지 못합니다.',
  PPTX_CUSTOM_GEOMETRY_UNSUPPORTED:'자유형 경로로 만든 도형을 아직 SVG나 HTML 도형으로 옮기지 못합니다.',
  PPTX_TABLE_STRUCTURE_INVALID:'표의 행·열·셀 구조 또는 네임스페이스가 잘못되어 안전하게 읽을 수 없습니다.',
  PPTX_TABLE_LIMIT_EXCEEDED:'표의 행·열·셀 수가 처리 한도를 초과했습니다.',
  PPTX_TABLE_STYLE_UNSUPPORTED:'표 안의 글머리표·글자 효과·특수 서식을 아직 옮기지 못합니다.',
  PPTX_TABLE_UNSUPPORTED:'이 표의 병합·서식·내용 구성을 아직 옮기지 못합니다.',
  PPTX_SHAPE_EFFECT_UNSUPPORTED:'그라데이션·자유형 경로·복잡한 도형 효과를 아직 옮기지 못합니다.',
  PPTX_TEXT_STYLE_UNSUPPORTED:'첨자·취소선·글자 효과 등 일부 텍스트 서식을 아직 옮기지 못합니다.',
  PPTX_GROUP_TRANSFORM_UNSUPPORTED:'배율이나 변형이 적용된 그룹을 아직 보존하지 못합니다.',
  PPTX_GROUP_SCALE_APPROXIMATED:'배율이 적용된 그룹의 위치와 크기는 옮겼지만 글자 크기와 선 굵기, 셀 여백은 원본과 다를 수 있습니다.',
  PPTX_SHAPE_IMAGE_APPROXIMATED:'HTML에서는 SVG 도형을 편집할 수 있지만, 다시 만든 PPT에서는 그림으로 저장되어 개별 편집할 수 없습니다.',
  PPTX_IMAGE_APPEARANCE_APPROXIMATED:'사진은 보존하지만 색상·자유형 마스크·그림자 등 일부 효과가 생략되어 원본과 다를 수 있습니다.',
  PPTX_IMAGE_DUOTONE_BAKED:'듀오톤 색상 효과를 그림에 적용했습니다. 효과를 별도 설정으로 편집할 수 없습니다.',
  PPTX_MATH_IMAGE_FALLBACK:'수식을 그림으로 보존했으며 수식 자체는 편집할 수 없습니다. 배치가 원본과 다를 수 있습니다.',
  PPTX_MATH_TABLE_BORDER_OMITTED:'수식 대체 표의 테두리는 별도 선으로 옮기지 않아 이미지에 포함되지 않은 부분이 빠질 수 있습니다.',
  PPTX_IMAGE_UNSUPPORTED:'그림 형식 또는 자르기·채우기를 아직 지원하지 않습니다.',
  PPTX_OBJECT_UNSUPPORTED:'이 객체의 표현을 지원하지 않아 원본 확인이 필요합니다.',
  PPTX_FONT_SELECTED:'HTML의 글꼴을 사용자가 선택한 글꼴로 변경했습니다.',
  PPTX_EMBEDDED_FONT_IGNORED:'명시한 대체 글꼴을 사용하며 내장 글꼴 파일은 가져오거나 배포하지 않습니다.',
  PPTX_EMBEDDED_FONT_LOCAL_ONLY:'원본 글꼴 이름을 유지하지만 내장 글꼴 파일은 포함하지 않습니다. 글꼴이 설치되지 않은 컴퓨터에서는 모양과 줄바꿈이 달라질 수 있습니다.',
  PREVIEW_UNAVAILABLE:'미리보기 렌더러를 사용할 수 없습니다. HTML 파일에서 배치를 확인해 주세요.'
};
const objectTypes=new Set(['text','shape','image','group','table','chart','background','font','line']);
function safeLabel(value){
  if(typeof value!=='string')return null;
  const label=value.replace(/[\u0000-\u001f\u007f]/g,' ').trim();
  if(!label||label.length>120||/[<>]|(?:[a-z][a-z0-9+.-]*:\/\/)|(?:^[a-z]:[\\/])|(?:^[/\\])|(?:\.\.[/\\])/i.test(label))return null;
  return label;
}
export function publicDiagnostic(item,slides=[]){
  const code=Object.hasOwn(messages,item.code)?item.code:'RENDER_INCOMPLETE';
  const severity=item.severity==='warning'?'warning':'error';
  const index=slides.findIndex(s=>s.id===item.slideId);
  const explicit=Number.isInteger(item.slideNumber)&&item.slideNumber>0&&item.slideNumber<=200?item.slideNumber:null;
  const known=code==='UNSUPPORTED_OBJECT'&&severity==='error'?knownObjectErrors.get(item.message):null;
  const result={code,severity,message:known||(severity==='warning'&&warnings[code])||messages[code],slideNumber:index<0?explicit:index+1};
  if(item.blocking===true&&severity==='error'){result.blocking=true;result.message='변환 중단: '+result.message;}
  if(Object.hasOwn(reasons,item.reasonKey)){
    result.reasonKey=item.reasonKey;result.reason=reasons[item.reasonKey];
    if(['HTML_PPTX_PARAGRAPH_LAYOUT_APPROXIMATED','HTML_PPTX_SHADOW_OMITTED','PPTX_PARAGRAPH_INDENT_APPROXIMATED','PPTX_LINE_SPACING_APPROXIMATED','PPTX_TEXT_BASELINE_APPROXIMATED','PPTX_PARAGRAPH_ALIGNMENT_APPROXIMATED','PPTX_INHERITED_SIZE_APPROXIMATED','PPTX_AUTONUM_APPROXIMATED','PPTX_TEXT_TAB_APPROXIMATED','PPTX_TEXT_HIGHLIGHT_APPROXIMATED','PPTX_TABLE_TEXT_COLOR_APPROXIMATED','PPTX_EMBEDDED_FONT_LOCAL_ONLY'].includes(item.reasonKey)&&severity==='warning')result.message=reasons[item.reasonKey];
    if(item.blocking===true&&severity==='error'){result.blocking=true;result.message='변환 중단: '+reasons[item.reasonKey];}
    const name=safeLabel(item.objectName),font=safeLabel(item.fontFamily);
    if(name)result.objectName=name;
    if(font)result.fontFamily=font;
    if(objectTypes.has(item.objectType))result.objectType=item.objectType;
    if(result.slideNumber&&Number.isInteger(item.objectIndex)&&item.objectIndex>0&&item.objectIndex<=10000)result.objectIndex=item.objectIndex;
    const b=item.bounds;
    if(b&&['x','y','w','h'].every(k=>typeof b[k]==='number'&&Number.isFinite(b[k]))&&b.x>=0&&b.y>=0&&b.w>0&&b.h>0&&b.x+b.w<=1&&b.y+b.h<=1)result.bounds={x:b.x,y:b.y,w:b.w,h:b.h};
  }
  return result;
}
