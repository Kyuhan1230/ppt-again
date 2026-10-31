# 개발 안내

Node.js 24와 Microsoft Edge를 설치한 Windows에서 시작합니다.

```sh
npm ci
npm test
npm start
```

## 자동 검사와 검증본

GitHub Actions는 변경 사항을 올리거나 풀 리퀘스트를 열면 Windows에서 Node.js 24로 의존성을 설치하고 테스트를 실행합니다. `main`의 검사가 통과하면 해당 소스의 ZIP 파일을 실행 결과의 아티팩트로 14일간 보관합니다. Actions 화면에서 직접 검사를 다시 실행할 수도 있습니다. 이 ZIP은 소스 검증본이며, 자동 공개 릴리스나 설치 프로그램은 아닙니다.

## 코드 구조

```text
src/html-to-pptx/  입력 해석, 브라우저 측정, PPT 생성, 로컬 서버
web/html-to-pptx/  파일 선택과 다운로드 화면
tests/            메모리에서 만든 가상 입력·다운로드·접근 차단·왕복 검사
docs/             사용 범위와 설계 이유
```

HTML 처리 순서는 `readInput → measureInput → exportPptx`입니다. 입력 JavaScript를 실행하지 않고, 격리된 브라우저에서 측정한 위치와 서식을 PPT 객체로 옮깁니다. 반대 방향은 `readPptx → exportHtml`이며, 미리보기 측정값으로 원래 PPT의 현재 값을 바꾸지 않습니다. 서버는 세션별 접근 확인과 작업별 프로세스 안에서 두 흐름을 실행합니다.

## PPT에서 HTML로

앱에서 PPTX를 선택하면 HTML로 내려받을 수 있습니다. 다음은 같은 변환을 Node.js 모듈로 실행하는 방법입니다. 입력과 출력 경로는 개인 폴더로 정하고 원본을 덮어쓰지 마세요.

```js
import fs from 'node:fs/promises';
import {readPptx} from './src/html-to-pptx/read-pptx.mjs';
import {exportHtml} from './src/html-to-pptx/export-html.mjs';

const inputPath = process.argv[2];
const outputPath = process.argv[3];
if (!inputPath || !outputPath) throw new Error('입력 PPTX와 출력 HTML 경로가 필요합니다.');
const parsed = await readPptx(await fs.readFile(inputPath));
if (parsed.status !== 'parsed') throw new Error(JSON.stringify(parsed.diagnostics));
const output = exportHtml(parsed.model);
if (output.status !== 'generated') throw new Error(JSON.stringify(output.diagnostics));
await fs.writeFile(outputPath, output.html, {flag: 'wx'});
```

생성 HTML은 `readInput → measureInput → exportPptx`로 다시 변환할 수 있습니다. 공통 모델의 `pptxSizeEmu`는 출처 PPT 크기를 기록합니다. 현재 문구와 배치는 HTML DOM에만 저장하므로 이전 문구로 되돌리는 숨은 사본을 사용하지 않습니다.

## 변경할 때 확인할 것

변환기를 수정하면 가상 입력의 모든 문구가 출력에 남는지, 기본 객체가 이미지 한 장으로 바뀌지 않았는지 확인합니다. 다른 비율의 HTML과 SVG도 함께 검사합니다. 실제 화면과 PowerPoint에서의 수정·저장·재열기는 패키지 구조 검사와 별개입니다. 입력 파일은 커밋하지 않고 테스트 코드에서 가상 자료를 만듭니다.

화면 비교에서는 내용·줄바꿈·정렬이 같고 새 잘림이나 겹침이 없어야 합니다. 이 조건과 육안 확인을 통과한 경우에만 위치·외곽 1px 이하 차이를 허용합니다. 구조 좌표 검사 통과를 화면 품질 통과로 부르지 않습니다.

예제나 이슈에 실제 회사 파일, 개인 경로, 자격 증명을 넣지 마세요. 기본 서버에는 문서와 결과 파일을 임의로 읽는 정적 경로를 추가하지 않습니다.

브랜치는 `feat/...`, `fix/...`, `docs/...`, `test/...`, `chore/...` 등 소문자 기능 이름을 사용합니다. `main`은 기준 브랜치입니다. 테스트 실패나 미측정 항목은 변경 설명에 남깁니다.
