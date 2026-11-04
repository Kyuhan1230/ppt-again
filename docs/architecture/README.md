# PPT, Again? 시스템 구성

[시스템 구성도](architecture.html) · [양방향 변환 흐름](conversion-flows.html) · 원본: [시스템](architecture-v4.json), [변환 흐름](conversion-flows.json)

구성도는 구성 요소 사이의 주요 관계를 묶어 보여준다. HTML 파일을 내려받아 브라우저에서 열면 확대, 검색, 테마 전환과 이미지 내보내기를 사용할 수 있다. GitHub의 파일 화면에서는 HTML이 실행되지 않는다. 구성도 본문은 한국어이며 Archify 뷰어의 고정 메뉴는 영어로 표시된다.

## 진입점과 작업 관리

- **웹 화면**: [app.js](../../web/html-to-pptx/app.js)는 파일 선택, 요청, 미리보기, 다운로드와 진단을 관리한다. [diagnostic-view.js](../../web/html-to-pptx/diagnostic-view.js)는 공개 진단을 원인별로 묶는다.
- **로컬 HTTP 서버**: [server.mjs](../../src/html-to-pptx/server.mjs)는 기본적으로 `127.0.0.1:42337`에서 실행된다. `/api/session`, `/api/analyze`, `/api/export`, 작업별 미리보기·다운로드·삭제 요청을 처리한다. [session.mjs](../../src/html-to-pptx/session.mjs)는 서명된 세션 쿠키와 쓰기 요청용 CSRF 토큰을 관리한다.
- **CLI**: [cli.mjs](../../src/html-to-pptx/cli.mjs)는 로컬 파일과 자산을 읽어 같은 작업 실행기를 사용하며, 서버를 띄울 필요가 없다.
- **작업 보관**: [job-store.mjs](../../src/html-to-pptx/job-store.mjs)는 결과와 작업 상태를 프로세스 메모리의 `Map`에 세션별로 보관한다. 기본 만료 시간은 30분이며 개수·크기 제한을 적용한다. 데이터베이스나 영구 저장소는 아니다. 삭제·만료·서버 종료 시 작업을 정리한다.
- **실행기와 워커**: [job-runner.mjs](../../src/html-to-pptx/job-runner.mjs)는 허용된 작업과 입력 크기를 검사하고 [conversion-worker.mjs](../../src/html-to-pptx/conversion-worker.mjs)를 자식 프로세스로 실행한다. 시간 제한, 관측 메모리 제한, 취소와 프로세스 정리를 다룬다. 메모리 관측은 OS가 강제하는 하드 제한이 아니다.

## PPTX → HTML

1. 서버 또는 CLI가 PPTX 바이트와 글꼴·부분 변환 옵션을 워커에 전달한다.
2. [read-pptx.mjs](../../src/html-to-pptx/read-pptx.mjs)가 PPTX 패키지를 읽어 장표 모델과 진단을 만든다. [package-budget.mjs](../../src/html-to-pptx/package-budget.mjs)는 패키지 압축 해제 한도를 관리하고 [read-table.mjs](../../src/html-to-pptx/read-table.mjs), [read-chart.mjs](../../src/html-to-pptx/read-chart.mjs)가 표와 차트 읽기를 돕는다.
3. [export-html.mjs](../../src/html-to-pptx/export-html.mjs)가 모델에서 HTML을 생성한다. 웹 분석에서는 생성된 HTML을 다시 측정해 장표 미리보기를 시도한다. 미리보기를 만들지 못해도 안전하게 생성한 HTML은 내려받을 수 있다.
4. 지원되지 않는 객체는 옵션과 오류 유형에 따라 진단이나 표시 상자로 남을 수 있다. 모든 서식과 편집 기능의 완전한 왕복을 보장하지 않는다.

## HTML → PPTX

1. [input.mjs](../../src/html-to-pptx/input.mjs)가 HTML과 제공된 자산을 읽어 입력 구조를 준비한다.
2. [browser-runtime.mjs](../../src/html-to-pptx/browser-runtime.mjs)가 Playwright와 Edge 런타임을 사용하고 [measure.mjs](../../src/html-to-pptx/measure.mjs)가 장표와 객체의 배치를 측정해 모델과 진단을 만든다.
3. [export.mjs](../../src/html-to-pptx/export.mjs)가 PptxGenJS로 모델을 PPTX로 내보낸다. [write-table.mjs](../../src/html-to-pptx/write-table.mjs)와 [write-chart.mjs](../../src/html-to-pptx/write-chart.mjs)는 지원되는 표와 차트를 작성한다.
4. 웹 서버는 결과를 작업에 보관하고 다운로드 요청에 응답한다. CLI는 결과를 사용자가 지정한 경로에 쓴다.

## 보안과 진단의 경계

- 브라우저 렌더링에서 입력 스크립트 실행과 서비스 워커를 끄고, 검증된 문서·글꼴 요청 외의 네트워크 요청을 차단한다. 이 조치는 [browser-runtime.mjs](../../src/html-to-pptx/browser-runtime.mjs)에 구현되어 있다.
- 서버는 세션과 CSRF 토큰을 검사하고, 워커 실행 전 입력·자산·모델 크기를 제한한다. 자세한 범위와 남은 과제는 [SECURITY.md](../../SECURITY.md)를 참고한다. 개인 로컬 실행용이며 인터넷 공개나 다중 사용자 운영이 검증된 구성은 아니다.
- [public-diagnostics.mjs](../../src/html-to-pptx/public-diagnostics.mjs)는 내부 진단을 공개 가능한 문구로 변환한다. 화면에서의 오류·주의는 변환 중단과 표현 제한을 구분해서 읽어야 한다.

구성도는 현재 코드의 책임과 주요 경로를 추상화한 것이며, 화살표 하나가 모든 함수 호출이나 데이터 반환 방향을 빠짐없이 표현하는 것은 아니다. 구체적인 입력과 기능 제한은 [지원 범위](../supported-inputs.md)를 확인한다.
