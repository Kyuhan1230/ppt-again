# PPT ↔ HTML 리팩토링 진행도

아래 그림은 요청받은 원본 구조다. 분홍색은 해당 역할의 구현과 검증이 끝난 부분만 표시한다. 일부만 끝난 역할은 전체 노드를 칠하지 않고 세부 진행 표에 기록한다.

```mermaid
flowchart TB
  subgraph CURRENT["현재 구조"]
    direction TB
    CP["PPT 입력"] --> CPR["PPT 읽기 · read-pptx.mjs<br/>파일 검사 + 장표 목록 + 객체 순회<br/>그림 · 도형/글자 · 그룹 · 표/차트"]
    CH["HTML 입력"] --> CHI["HTML 입력 분석 · input.mjs<br/>장표 표식 · 내장 데이터 · 일부 SVG"]
    CHI --> CHM["HTML 읽기·측정 · measure.mjs<br/>후보 선택 + 객체 인식 + 위치·줄 측정"]
    CPR --> CM["기존 공통 장표 모델"]
    CHM --> CM
    CM -->|"PPT → HTML"| CHW["HTML 생성 · export-html.mjs"]
    CHW --> CHO["HTML 출력"]
    CM -->|"HTML → PPT"| CPW["PPT 생성 · export.mjs"]
    CHM -.->|"글자 줄·기준선 측정값"| CPW
    CPW --> CPO["PPT 출력"]
  end

  subgraph PROPOSED["개선 후 구조 · 제안"]
    direction TB
    NP["PPT 입력"] --> NPC["① PPT 파일 검사<br/>압축 · 구조 · 외부 참조"]
    NPC --> NPS["② PPT 장표 확인<br/>기록된 목록 · 크기 · 순서"]
    NPS --> NPO["③ PPT 객체 인식<br/>그림 · 도형/글자 · 그룹 · 표/차트<br/>원본 객체 관계 보존"]
    NH["HTML 입력"] --> NHR["④ HTML 화면·구조 조사<br/>화면 크기 · 자산 · 화면 안과 밖"]
    NHR --> NHC["⑤ 장표 후보 찾기<br/>표식 · 크기/위치 · 반복 구조<br/>부모/자식 · 겹침 · 숨김"]
    NHC -.-> ST1["1단계 완료<br/>자동 후보 밖 내용 누락 방지"]
    NHC --> ND{"장표 범위가<br/>확실한가?"}
    ND -->|"검증된 장표 정보<br/>또는 명시된 선택"| NS["⑥ 변환 범위 확정<br/>한 영역 · 여러 영역 · 문서 범위"]
    ND -->|"불확실하거나<br/>후보가 겹침"| NUI["UI<br/>테두리 · 축소 지도<br/>클릭해서 선택"]
    ND -->|"불확실하거나<br/>후보가 겹침"| NCLI["CLI<br/>후보 목록 · JSON<br/>선택 또는 미리 정한 방식"]
    NUI --> NS
    NCLI --> NS
    NS --> NHO["⑦ HTML 객체 인식·측정<br/>글자 · 그림 · 배경/도형 · 표/차트<br/>위치 · 앞뒤 순서 · 줄바꿈"]
    NHO --> NI["⑧ 원본 장표별 중간 결과<br/>각 장표의 크기와 측정값"]
    NI --> NFIT["⑨ 출력 크기·분할 결정<br/>원본 크기 / 비율 유지와 여백<br/>필요하면 명시적으로 분할"]
    NFIT --> NM["공통 장표 모델 확장<br/>객체 · 위치 · 내용 · 서식<br/>원본 객체 관계"]
    NPO --> NM
    NM -->|"PPT → HTML"| NHW["⑩ HTML 객체별 생성<br/>글자 · 그림 · 도형 · 표/차트"]
    NHW --> NHOOUT["HTML 출력"]
    NM -->|"HTML → PPT"| NPW["⑪ PPT 객체별 생성<br/>글자 · 그림 · 도형 · 표/차트"]
    NI -.->|"정규화한 줄·기준선 측정값"| NPW
    NPW --> NPOUT["PPT 출력"]
    NHC -.-> REPORT["결과 보고<br/>선택 이유 · 범위 밖 내용<br/>편집 제한 · 확인 못 한 내용"]
    NHO -.-> REPORT
    NPW -.-> REPORT
  end

  CPR -.->|"역할 분리"| NPC
  CHM -.->|"후보와 객체 읽기 분리"| NHR
  classDef input fill:#f3f4f6,stroke:#7d8790,color:#283142
  classDef current fill:#fff3db,stroke:#c18124,color:#3f2b12
  classDef proposed fill:#e8f1ff,stroke:#557fb3,color:#102a48
  classDef common fill:#e6f5e9,stroke:#4e8b66,color:#183c27
  classDef choice fill:#f4edff,stroke:#8767ad,color:#36254b
  classDef report fill:#fff0f0,stroke:#b35b5b,color:#4a2020
  classDef done fill:#ffd0e6,stroke:#c42e78,color:#46102b
  class CP,CH,CHO,CPO,NP,NH,NHOOUT,NPOUT input
  class CPR,CHI,CHM current
  class NPC,NPS,NPO,NHR,NHC,ND,NS,NHO,NI,NFIT proposed
  class NUI,NCLI choice
  class CM,CHW,CPW,NM,NHW,NPW common
  class REPORT report
  class ST1 done
```

| 단계 | 상태 | 완료 범위 |
|---|---|---|
| 1. 영역 밖 내용 누락 방지 | 완료 | 자동 선택은 범위 밖 내용이 발견되면 오류, 명시 선택은 경고; 확인하지 못한 내용은 별도 알림 |
| 2. 후보 조사 | 대기 | 후보·관계·화면 안팎 조사 |
| 3. UI·CLI 선택 | 대기 | 동일 후보 표시 및 선택 |
| 4. HTML 객체 인식 분리 | 대기 | 글자·그림·도형·표/차트 분리 |
| 5. 원본 객체 관계 보존 | 대기 | PPT 한 객체의 왕복 보존 |
| 6. PPT 읽기 분리 | 대기 | 파일 검사·장표·객체 읽기 분리 |
| 7. 양쪽 생성기 분리 | 대기 | HTML·PPT 객체별 출력 |
| 8. 가변 크기 | 대기 | 장표별 중간 결과와 크기 정규화 |
| 9. 문서 분할 | 대기 | 명시적 페이지 분할 |
| 10. 통합 검증 | 대기 | 양방향·UI·CLI·지원 제한 확인 |
## 1단계 검토

자동으로 선택된 장표 밖에 표시된 내용이 있으면 변환을 중단한다. 사용자가 영역을 명시한 경우에는 선택한 영역을 유지하고 경고한다. 숨겨진 내용이나 CSS 합성 등 조사하지 못한 내용은 별도로 알린다. 화면 크기 변경 전후의 관측 결과를 함께 사용한다.

검증: 변경 전 실패 사례를 확인한 뒤 범위 밖 내용, 화면 크기 변경, 배경 및 기존 장표 입력에 대한 집중 테스트 80개를 모두 통과했다. 독립 검토에서 문서 배경과 반응형 화면의 누락을 재현하고 수정 후 다시 확인했다. Windows 전체 테스트는 이 단계 PR의 CI에서 확인한다.

남은 한계: 복잡한 CSS 합성과 스크립트로 생성되는 내용은 완전하게 판정할 수 없다. 다음 단계에서 장표 후보와 범위를 별도로 조사한다.
