# Hangul — remaining

본문 읽기·문단/선택 수정·누름틀·표 칸·표 생성·본문/셀 서식·웹 검색·채팅 저장은 됨.
4000자 본문 한도는 rhwp `applyTextCommand` 계약이라 풀지 않음.
호스트는 스튜디오 Document `_request`로만 호출. Hangul WASM을 호스트에서 우회 호출하지 말 것.

## 편집

- [x] **한글 AI 글쓰기 검증** — `insert_content`는 스튜디오 `insertFilledParagraphs` 한 번으로 문단을 만들고 `applyTextCommand`로 채움. 잠긴 첫 문단은 건너뜀. 주간보고·계획서가 한 번에 들어감
- [x] **표 생성 / 서식** — `insert_table`, `apply_format`(굵게·색·글꼴·크기·정렬·줄간격·들여쓰기·글머리표/번호). 표 셀은 `table`+`row`(·`col`). HTML 한 방에 넣는 Docs와 다름. 본문 쓴 뒤 도구로 적용
- [ ] **표 구조** — WASM: `insertTableRow` / `insertTableColumn` / `deleteTableRow` / `deleteTableColumn` / `mergeTableCells` / `splitTableCell`. “행 추가”, “열 삭제”, “헤더 합치기”
- [ ] **칸·표 모양** — WASM: `setCellProperties` / `setTableProperties`. 배경색·테두리·세로 정렬·표 너비. 글자 서식(`applyCharFormatInCell`)과 별개
- [ ] **용지·단** — WASM: `setPageDef` / `setColumnDef`. 여백, 가로/세로, 2단
- [ ] 머리글/바닥글/각주 읽기·수정 — WASM에 `getHeaderFooter` / `insertTextInHeaderFooter` / `insertFootnote` 있음. 공개 SDK 없음. 호스트에서 WASM 직접 호출하면 undo/레이아웃이 깨짐

## 다른 문서 AI에 있고 한글에 없는 것

- [ ] 첨부 파일 읽기 (`files-skill` / `read_attachment`)
- [ ] 이미지 검색·생성·삽입 — 일부러 보류. WASM `insertPicture`는 있음
- [ ] `create_document` — 필요 없음. 다시 넣지 말 것

## 호스트 / 셸

- [ ] 페이지 넘김 — canvas2d에서 다음 페이지가 끊김. 예전 패치는 `eb3c4f8`에서 되돌림. 다시 넣지 말 것
- [ ] 인쇄 / PDF — 후순위. 스튜디오에 `file:print` / `file:print-to-pdf`가 있음. `print.html`을 스냅샷에 넣고 셸 File 메뉴에서 호출하면 됨

## 하지 않음

- 본문 4000자 한도 해제
- 댓글/교정 — 한글 스튜디오에 호스트 경로 없음
- 차트 — WASM `listCharts` / `setChartDataByIndex` 있음. 호스트 경로 없음. 붙이지 말 것
- 표 수식 — WASM `evaluateTableFormula` 있음. 쓰임이 좁음
