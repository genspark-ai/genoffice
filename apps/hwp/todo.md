# Hangul AI — remaining

본문 문단/선택/누름틀/표 칸 편집과 웹 검색·채팅 저장은 됨.
아래는 Docs/Markdown AI와 맞추려면 남은 것. 4000자 본문 한도는 rhwp `applyTextCommand` 계약이라 풀지 않음.

## 편집 (스튜디오 표면이 더 필요)

- [ ] 머리글/각주 읽기·수정 — Docs의 `set_header_footer`에 해당. 공개 API 없음. WASM 직접 호출은 undo/레이아웃이 깨질 수 있음
- [ ] 새 문단 삽입 / 빈 문서에 글 쓰기 — Docs/Markdown `insert_content`. `applyTextCommand`가 `\n`을 거절해서 기존 문단 교체만 가능

## 다른 문서 AI에 있고 한글에 없는 것

- [ ] 첨부 파일 읽기 (`files-skill` / `read_attachment`) — Docs·Slides·Sheets에 있음
- [ ] 이미지 검색·생성·삽입 — Docs·Markdown에 있음. 한글 쪽은 일부러 보류
- [ ] `create_document` — Docs처럼 AI가 새 문서를 만들어 열기

## 하지 않음

- 본문 4000자 한도 해제 — rhwp SDK 제한
- 댓글/교정/차트 — Docs 전용. 한글 스튜디오에 호스트 경로 없음
- 인쇄/PDF — AI가 아니라 셸 File 메뉴. 아직 한글 탭에 안 연결됨
