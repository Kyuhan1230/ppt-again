# 오류 로그와 Windows 자원 감시

## 로그 위치와 보관

변환과 서버 시작 중 발생한 오류는 프로젝트의 `.local/logs/errors-YYYY-MM-DD.log`에 UTF-8 JSON Lines 형식으로 기록합니다. 파일 이름의 날짜는 UTC 기준입니다. 이 디렉터리는 Git에서 제외합니다.

각 로그 파일의 크기는 최대 1MiB입니다. 크기를 넘기면 같은 날짜의 이전 파일 하나를 `.1.log`로 보관합니다. 새 기록을 쓸 때 최근 7일 범위 밖의 로그를 정리합니다. 로그 쓰기 순서와 파일 회전은 단일 앱 프로세스를 기준으로 보장하며, 여러 프로세스 간 조정은 지원하지 않습니다.

로그에는 작업·요청 ID, 처리 단계, 오류 코드, 허용된 프로젝트 소스 위치, 장표 번호와 처리 시간·메모리 측정값을 기록할 수 있습니다. 문서 내용, 이미지, 원본 파일명, 절대 경로, 토큰과 원본 오류 메시지는 기록하지 않습니다. 진단 설명에는 정해진 안내 문구를 사용합니다.

## 로그 확인

오류를 재현한 뒤 최신 파일을 확인합니다.

```powershell
$latestErrorLog = Get-ChildItem -LiteralPath '.local/logs' -Filter 'errors-*.log' -ErrorAction SilentlyContinue |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 1

if ($null -ne $latestErrorLog) {
    Get-Content -LiteralPath $latestErrorLog.FullName -Encoding UTF8 -Tail 20
}
```

`jobId`, `requestId`, `time`으로 같은 작업을 연결하고 `operation`으로 실패 단계를 구분합니다. `error`와 `error.cause`에는 허용된 오류 코드와 종료 정보가, `resourceUsage`에는 측정 횟수, 최대 메모리와 처리 시간이 포함될 수 있습니다.

## Windows 프로세스 감시

Windows에서는 프로세스 ID, 부모 ID, 메모리 사용량과 생성 시각을 조회합니다. 각 조회의 제한 시간은 5초이며, 작업 시작 전 조회가 시간 초과되면 한 번 재시도합니다. 실행 중 감시가 실패하면 작업을 중단하고 확인된 하위 프로세스를 정리합니다.

작업자가 결과를 보낸 뒤 정상 종료한 경우, 종료 중 발생한 감시 시간 초과에 한해 프로세스를 다시 조회합니다. 관련 프로세스가 모두 종료됐다고 확인한 경우에만 결과를 반환합니다. 확인에 실패하거나 프로세스가 남아 있으면 오류를 반환하고 정리를 시도합니다.

감시 실패 기록에서는 `error.observerPhase`와 `error.cause.code`를 확인합니다.

| 값 | 의미 |
| --- | --- |
| `initial` | 작업자에게 입력을 보내기 전 조회 |
| `active` | 작업자 종료 전 조회 실패 |
| `closing` | 작업자 종료 중 진행하던 조회 실패 |
| `OBSERVER_ACCESS_DENIED` | 프로세스 조회 권한 거부 |
| `OBSERVER_TIMEOUT` | 프로세스 조회 시간 초과 |
| `OBSERVER_IDENTITY_UNAVAILABLE` | 초기 조회에서 작업자의 생성 시각을 확인하지 못함 |
| `OBSERVER_QUERY_FAILED` | 조회 결과 또는 메모리 값이 유효하지 않음 |

## 확인과 테스트

필요하면 실패한 Windows 환경에서 다음 명령으로 CIM 조회 권한과 실행 시간을 확인합니다. 프로세스의 명령행과 경로는 출력하지 않습니다.

```powershell
Get-CimInstance Win32_Process -Property ProcessId,ParentProcessId,WorkingSetSize,CreationDate -ErrorAction Stop |
    Measure-Object |
    Select-Object -ExpandProperty Count

Measure-Command {
    Get-CimInstance Win32_Process -Property ProcessId,ParentProcessId,WorkingSetSize,CreationDate -ErrorAction Stop |
        Measure-Object | Out-Null
} | Select-Object TotalMilliseconds
```

로그·감시 검사와 PPTX 가져오기 회귀 검사는 가상 자료로 실행합니다.

```powershell
node --test --test-concurrency=1 tests/error-log.test.mjs tests/observer-retry.test.mjs
node --test --test-concurrency=1 --test-name-pattern='PPT input downloads current content' tests/server.test.mjs
```

전체 검사는 `node --test --test-concurrency=1 tests/*.test.mjs`로 실행합니다.
