<p align="center">
  <a href="https://genoffice.ai/">
    <picture>
      <source srcset="../assets/readme/hero-dark.webp" media="(prefers-color-scheme: dark)">
      <img src="../assets/readme/hero.webp" alt="GenOffice — 내장 AI 패널을 갖춘 오픈소스 AI 오피스 제품군: Docs, Sheets, Slides, PDF, Markdown, HTML" width="100%">
    </picture>
  </a>
</p>

<h1 align="center">GenOffice</h1>

<p align="center"><b>세계 최초의 완전한 기능을 갖춘 오픈소스 AI 오피스 스위트.</b><br>
Word, Excel, PowerPoint, PDF 파일을 당신과 당신의 AI가 함께 편집하고, 실제 포맷 그대로 저장합니다.</p>

<p align="center">
  <a href="../../LICENSE"><img src="https://img.shields.io/github/license/genspark-ai/genoffice" alt="라이선스: Apache-2.0"></a>
  <a href="https://github.com/genspark-ai/genoffice/releases/latest"><img src="https://img.shields.io/github/v/release/genspark-ai/genoffice" alt="최신 릴리스"></a>
  <a href="https://github.com/genspark-ai/genoffice/releases"><img src="https://img.shields.io/github/downloads/genspark-ai/genoffice/total" alt="다운로드 수"></a>
  <a href="https://github.com/genspark-ai/genoffice/stargazers"><img src="https://img.shields.io/github/stars/genspark-ai/genoffice?style=flat" alt="GitHub 스타 수"></a>
</p>

<p align="center"><a href="../../README.md">English</a> · <a href="README.es.md">Español</a> · <a href="README.pt-BR.md">Português (Brasil)</a> · <a href="README.de.md">Deutsch</a> · <a href="README.fr.md">Français</a> · <a href="README.zh-CN.md">简体中文</a> · <a href="README.zh-TW.md">繁體中文</a> · <b>한국어</b> · <a href="README.ja.md">日本語</a> · <a href="README.ar.md">العربية</a> · <a href="README.ru.md">Русский</a> · <a href="README.it.md">Italiano</a> · <a href="README.nl.md">Nederlands</a> · <a href="README.pl.md">Polski</a> · <a href="README.cs.md">Čeština</a> · <a href="README.id.md">Bahasa Indonesia</a> · <a href="README.ms.md">Bahasa Melayu</a> · <a href="README.th.md">ไทย</a> · <a href="README.hi.md">हिन्दी</a> · <a href="README.he.md">עברית</a></p>

<p align="center">
  <a href="#download"><b>다운로드</b></a> ·
  <a href="https://genoffice.ai/"><b>웹사이트</b></a> ·
  <a href="https://genoffice.ai/join"><b>커뮤니티</b></a> ·
  <a href="../../PRIVACY.md"><b>개인정보</b></a>
</p>

GenOffice는 macOS, Windows, Linux에서 사용할 수 있는 무료 오픈소스
마이크로소프트 오피스 대체 프로그램입니다. 네이티브 `.docx`, `.xlsx`,
`.pptx` 파일을 열고 저장하며, PDF, Markdown, HTML을 편집하고, 모든 문서
옆에 AI 에이전트를 배치합니다 — 옆에 덧붙인 챗봇이 아니라, 파일을 읽고
변경을 수행한 뒤 정확히 무엇을 수정했는지 보여주는 편집기입니다.

- **실제 포맷을 그대로, 바이트 단위로 보존합니다.** 편집한 부분만 다시
  작성되고, 파일의 나머지는 바이트 단위로 그대로 유지되어 문서가 Word,
  Excel, PowerPoint에서도 계속 정상적으로 열립니다.
- **검토할 수 있는 AI입니다.** 편집 내용은 변경 내용 추적과 diff로
  반영되며 원클릭 롤백이 가능합니다. 스프레드시트에는 붙여넣은 값이
  아니라 실시간으로 동작하는 수식이 들어갑니다. 슬라이드와 페이지는
  캔버스에 직접 생성되어 계속 자유롭게 편집할 수 있습니다.
- **로컬 우선으로 설계되었습니다.** 파일은 사용자의 기기에서 열리고,
  편집되고, 저장되고, 변환됩니다. PDF → Word / Excel / PowerPoint,
  Markdown → Word, HTML → Word 변환 모두 기기 내에서 실행됩니다. 오직
  AI 호출만 사용자가 선택한 공급자로 전송됩니다.
- **API 키가 있어도, 없어도 됩니다.** Genspark로 로그인하면 키 없이 바로
  사용할 수 있고, 원한다면 Claude, OpenAI, Gemini, DeepSeek, Kimi, GLM,
  Qwen, Doubao, MiniMax, Grok, Mistral, OpenRouter는 물론 OpenAI
  호환 엔드포인트(로컬 서버 포함)까지 직접 가져올 수 있습니다.

**다운로드:** [macOS](https://github.com/genspark-ai/genoffice/releases/latest) (Apple Silicon 및 Intel) ·
[Windows](https://github.com/genspark-ai/genoffice/releases/latest) (x64 및 Arm) ·
[Linux](https://github.com/genspark-ai/genoffice/releases/latest) (deb, rpm, AppImage) —
자세한 내용과 요구 사항은 [다운로드](#download) 섹션을 참고하세요.

## 데모

여섯 개의 앱, 하나의 AI 패널. 모든 스크린샷은 macOS에서 실제로 구동되는
앱이며, 패널에서 확인할 수 있는 프롬프트로 AI를 직접 실행한 결과입니다.

### 1 · Docs — 검토 가능한 AI로 `.docx`를 열고 편집합니다

<table>
<tr>
<td width="50%"><img src="../assets/readme/docs-report.webp" alt="GenOffice Docs가 전체 폭 표지 이미지, 음영 처리된 KPI 표, 머리글과 바닥글이 있는 2단 연차 보고서 페이지를 80% 확대 비율로 렌더링하고 있으며 AI 패널은 접혀 있는 화면"></td>
<td width="50%"><img src="../assets/readme/docs-ai.webp" alt="GenOffice Docs: 배너 이미지가 있는 회사 개요 문서에서 AI가 개요 섹션을 다듬고 새로운 글머리 기호 섹션을 삽입했으며, 패널에서 원클릭 롤백을 제공하는 화면"></td>
</tr>
<tr>
<td><b>Word가 배치한 그대로 파일을 엽니다</b> — 2단 구성, 여백 없는 전체 이미지, 음영 표, 머리글과 바닥글, Word의 줄 단위 계량 기준을 따르는 페이지 나누기까지 그대로입니다. 스타일, 댓글, 변경 내용 추적, 수식, 잉크 주석도 손상 없이 그대로 유지됩니다.</td>
<td><b>원하는 편집을 요청하세요</b> — AI가 필요한 블록을 읽어 개요를 다시 작성하고 새로운 글머리 기호 섹션을 삽입합니다. AI가 수행한 모든 작업은 롤백 가능한 스냅샷으로 남으며, <b>변경 내용 추적</b>을 켜면 편집 내용이 Word 스타일의 수정 내용으로 표시됩니다.</td>
</tr>
</table>

### 2 · Sheets — 붙여넣은 값이 아니라 실시간 수식과 차트가 있는 `.xlsx`

<table>
<tr>
<td width="50%"><img src="../assets/readme/sheets-ai.webp" alt="GenOffice Sheets: AI가 SUMIF 수식을 사용해 지역 및 카테고리별 매출이 담긴 Summary 시트와 세로 막대형 차트를 추가하고, 실행 취소 버튼과 함께 43건의 변경 사항이 적용되었다고 알려주는 화면"></td>
<td width="50%"><img src="../assets/readme/sheets-qa.webp" alt="GenOffice Sheets: 2분기 매출을 어느 지역이 주도했는지 묻자 AI가 카테고리별 세부 내역과 함께 유럽이라고 답하고, 사용한 셀을 링크 형태의 출처로 인용하는 화면(Orders 시트 옆에 표시됨)"></td>
</tr>
<tr>
<td><b>직접 만들어 줍니다</b> — 한 문장의 요청만으로 에이전트가 지역 및 카테고리별 실제 <code>SUMIF</code> 수식이 담긴 Summary 시트를 추가하고, 세로 막대형 차트를 삽입하며, 43건의 변경을 실행 취소 가능한 하나의 배치로 적용합니다.</td>
<td><b>물어보세요</b> — 워크북에 대한 질문에는 근거와 함께, 사용한 정확한 셀이 클릭 가능한 출처로 돌아옵니다. 내부적으로는 자체 개발한 Rust <code>.xlsx</code> 엔진, 피벗 테이블, 슬라이서, 조건부 서식, 수식 추적 기능이 동작합니다.</td>
</tr>
</table>

### 3 · Slides — 프롬프트 한 줄로 완성하는 `.pptx` 덱

<img src="../assets/readme/slides-generate.webp" alt="GenOffice Slides가 Aurora Home 투자자용 덱을 생성하는 타임랩스: AI가 패널에서 스토리라인을 구상하고, 슬라이드가 캔버스에 차례로 나타나며, 완성된 덱은 마무리 제안으로 끝난다" width="100%">

<table>
<tr>
<td width="50%"><img src="../assets/readme/slides-cover.webp" alt="GenOffice Slides: 캔버스에 표시된 AI 생성 Aurora Home 투자자 덱의 표지 슬라이드, 패널에는 원래 입력한 한 줄 프롬프트와 AI가 만든 내용에 대한 요약이 보이는 화면"></td>
<td width="50%"><img src="../assets/readme/slides-ai.webp" alt="GenOffice Slides: 같은 11장짜리 덱의 디자인이 완료된 마지막 슬라이드, 왼쪽에는 축소판 슬라이드 목록이, AI 패널에는 전체 스토리라인 요약이 표시된 화면"></td>
</tr>
<tr>
<td><b>한 줄만 입력하세요</b> — "Aurora Home을 위한 10장짜리 투자자 피치 덱을 만들어줘…"처럼요. GenOffice가 스토리라인을 구성하고, 필요한 숫자를 조사하고, 실제 <code>.pptx</code>로 모든 슬라이드를 캔버스에 초안 작성합니다.</td>
<td><b>완성된 덱이 나옵니다</b> — 일관된 타이포그래피, 이미지, 마무리 행동 유도 문구까지 디자인이 완료된 11장의 슬라이드. 마스터, 레이아웃, 스마트 가이드, 비파괴 자르기로 계속 편집하거나, 패널에 스타일 변경, 재작성, 순서 재배치를 요청할 수 있습니다.</td>
</tr>
</table>

### 4 · PDF — PDF 텍스트를 그 자리에서 편집하고, 기기 내에서 Word로 변환합니다

<table>
<tr>
<td width="50%"><img src="../assets/readme/pdf-edit.webp" alt="GenOffice PDF: 텍스트 편집 모드가 페이지의 모든 텍스트 블록을 표시해 바로 편집할 수 있게 하고, 동시에 AI 패널이 보고서에 대한 질문에 페이지 출처와 함께 답변하는 화면"></td>
<td width="50%"><img src="../assets/readme/pdf-convert.webp" alt="GenOffice Docs가 Helios 분기 보고서 PDF에서 로컬로 변환된 Word 문서를, 원본 PDF 옆의 두 번째 탭에 열어 보여주는 화면"></td>
</tr>
<tr>
<td><b>페이지 안에서 바로 편집합니다</b> — 텍스트 편집 모드가 모든 텍스트 블록을 표시해 그 자리에서 다시 입력할 수 있게 하며, 위에 덮어씌우는 주석이 아니라 PDFium을 통해 원본 폰트를 유지한 채 콘텐츠 스트림 자체를 다시 작성합니다. 긴 보고서에 대해 AI에게 물으면 페이지 출처와 함께 답변을 받을 수 있습니다.</td>
<td><b>기기 내에서 변환합니다</b> — <b>PDF 변환기 → PDF를 Word로</b>를 실행하면 편집 가능한 <code>.docx</code>가 만들어지고, 원본 옆에 Docs에서 열리며 제목, 통계 행, 문단이 그대로 유지됩니다. Excel과 PowerPoint 변환도 동일하게 동작하며, 스캔한 페이지는 시스템 OCR을 거칩니다.</td>
</tr>
</table>

### 5 · HTML — 디자인 브리프를 먼저 제시하는 AI 페이지 및 UI 빌더

페이지의 목적과 대상을 알려주세요. AI는 먼저 **디자인 브리프**를
제안합니다 — 후킹 문구, 색상 팔레트, 타이포그래피, 스타일 방향까지 —
그런 다음 이 토큰들을 기준으로 완전히 독립된 `.html` 파일 하나를 만듭니다.

<img src="../assets/readme/html-restyle-motion.webp" alt="GenOffice HTML이 Lumen 랜딩 페이지를 리스타일링하는 타임랩스: 패널에서의 한 번의 Restyle 요청으로 어두운 Midnight Studio 페이지가 따뜻한 Solar Daybreak 버전으로 바뀌지만, 모든 섹션과 문구는 그대로 유지된다" width="100%">

<table>
<tr>
<td width="50%"><img src="../assets/readme/html-ai.webp" alt="GenOffice HTML: 어두운 Midnight Studio 스타일로 생성된 태양광 스탠드 랜딩 페이지가 실시간 미리보기에 표시되고, AI 패널이 방금 만든 페이지를 요약해 보여주는 화면"></td>
<td width="50%"><img src="../assets/readme/html-restyle.webp" alt="같은 Lumen 랜딩 페이지를 AI가 따뜻한 느낌의 Solar Daybreak 스타일로 다시 디자인한 화면: 종이 질감 배경, 세리프 헤드라인, 오렌지 강조색을 사용했지만 모든 섹션과 문구는 그대로 유지된 모습"></td>
</tr>
<tr>
<td><b>프롬프트 한 번으로 생성됩니다</b> — Midnight Studio 스타일로 만들어진 강렬한 히어로 섹션, 기능 카드, 요금 안내, Lumen의 대기자 명단 폼까지. 요소를 클릭하면 스타일을 바꿀 수 있고, 더블클릭하면 텍스트를 편집할 수 있으며, CodeMirror 소스 보기로도 전환할 수 있습니다.</td>
<td><b>디자인은 같게, 방향만 다르게</b> — <b>스타일 변경</b> 요청 한 번으로 브리프의 토큰이 바뀌고 페이지도 그에 따라 변합니다: 따뜻한 종이 질감, 에디토리얼 세리프체, 선명한 오렌지 강조색까지, 내용은 그대로입니다. 전체 화면으로 발표하거나 PDF, 편집 가능한 네이티브 Word 문서로 내보낼 수 있습니다.</td>
</tr>
</table>
<table>
<tr>
<td width="50%"><img src="../assets/readme/html-dashboard.webp" alt="GenOffice HTML: 프리랜서 디자이너를 위해 생성된 개인 대시보드 UI가 따뜻한 리넨 스타일로 표시되고, 왼쪽 레일, 세리프체 인사말, 네 개의 지표 카드가 있는 화면"></td>
<td width="50%"><img src="../assets/readme/html-report.webp" alt="GenOffice HTML: 신문 형식으로 생성된 전기차 시장 데이터 보고서, 세리프체 제호, 1,730만이라는 헤드라인 숫자, 통계 행이 표시된 화면"></td>
</tr>
<tr>
<td><b>UI 목업</b> — "개인 대시보드" 스타터는 페르소나를 실제로 동작하는 레이아웃으로 바꿔줍니다: 왼쪽 레일, 인사말, 청구 가능 시간 스파크라인, 인보이스와 사용률 카드까지 모두 개발자에게 그대로 넘길 수 있는 실제 HTML입니다.</td>
<td><b>데이터 스토리</b> — "데이터 보고서" 스타터는 신문 형식의 에디토리얼을 만듭니다: 세리프체 제호, 하나의 헤드라인 숫자, 구분선으로 나뉜 통계 행, 인라인 SVG 차트, 방법론 설명까지 포함됩니다.</td>
</tr>
</table>

### 6 · Markdown — Ask AI를 갖춘, 순수 `.md` 위의 블록 에디터

<table>
<tr>
<td width="50%"><img src="../assets/readme/markdown-ai.webp" alt="GenOffice Markdown: 선택한 문단에 Ask AI 팝오버가 나타나 입력한 지시문과 함께 다듬기, 더 간결하게, 확장하기, 문법 수정 같은 제안 칩, 그리고 지금 보내기와 대기열에 추가 버튼이 표시된 화면"></td>
<td width="50%"><img src="../assets/readme/markdown-render.webp" alt="GenOffice Markdown이 표, Mermaid 순서도, 할 일 목록이 포함된 출시 공지 문서를 렌더링하고, 왼쪽에는 AI 패널의 시작 프롬프트가 표시된 화면"></td>
</tr>
<tr>
<td><b>선택한 부분에 대해 AI에게 물어보세요</b> — 원하는 부분을 선택하면 <b>Ask AI</b> 칩이 나타납니다: 직접 지시문을 입력하거나 제안 중 하나를 고르고, 바로 실행하거나 여러 개의 위치 지정 편집을 대기열에 쌓아 한 번에 실행할 수 있습니다. 동일한 진입점이 모든 앱에 있습니다.</td>
<td><b>렌더링되지만 순수 Markdown으로 저장됩니다</b> — Tiptap 블록 에디터에서 제목, 목록, 표, 이미지, 코드 블록, Mermaid 다이어그램을 다루면서도 저장 시에는 순수 <code>.md</code>로 그대로 기록되며, 완전히 로컬에서 동작하는 <b>Markdown → Word</b> 내보내기도 지원합니다.</td>
</tr>
</table>

## GenOffice를 선택하는 이유

- **오픈소스**, Apache-2.0 라이선스로 GitHub에서 공개 개발됩니다.
- **내 기기에서 직접 실행.** macOS, Windows, Linux용 네이티브 앱; 파일은 디스크
  에 그대로 남고 모든 편집, 저장, 변환이 내 기기에서 이루어집니다.
- **진짜 Office 파일.** 네이티브 `.docx`, `.xlsx`, `.pptx`, 바이트 단위 보존:
  손대지 않은 부분은 그대로 복사됩니다.
- **문서를 직접 편집하는 AI.** Docs에서는 변경 내용 추적, Sheets에서는 실시간으로
  동작하는 수식과 차트, 슬라이드는 캔버스 위에 그려지며, AI가 한 번 작업할 때마다
  되돌릴 수 있는 스냅샷이 남습니다.
- **내 모델, 내 키.** Genspark로 로그인하거나 Claude, OpenAI, Gemini, DeepSeek
  등의 키를 사용할 수 있으며, 로컬 서버와 모든 OpenAI 호환 엔드포인트도 지원합니
  다.
- **PDF도 제대로.** 페이지 안에서 텍스트를 직접 편집하고, 기기 내에서 PDF를
  Word, Excel, PowerPoint로 변환하며, 스캔 문서는 시스템 OCR을 사용합니다.
- **Markdown과 HTML도 지원**하며, 동일한 AI 패널과 로컬 Word 내보내기를 사용합
  니다.
- **무료**로 개인과 팀 모두 사용할 수 있습니다.

## AI 백엔드

**Genspark로 로그인**하면 별도 설정이 전혀 필요 없습니다. 모델 호출은
Genspark 프록시(Claude, GPT, Gemini 계열)를 거치고, 에이전트는 웹 및
이미지 검색, 이미지 생성, 이미지/오디오/비디오 분석 기능까지 사용할 수
있습니다.

**직접 키를 가져올 수도 있습니다.** 설정 → AI에는 Claude, OpenAI, Gemini,
DeepSeek, Kimi, GLM, Qwen, Doubao, MiniMax, Grok, Mistral, OpenRouter,
OpenCode Zen/Go가 있으며, 로컬 모델 서버를 포함한 모든 OpenAI 호환
엔드포인트(기본 URL + 키)를 위한 사용자 지정 슬롯도 제공됩니다. 검색과
미디어는 **AI 미디어 및 검색**에서 기능별로 별도 공급자를 설정할 수
있습니다: 웹 검색은 Serper 또는 Tavily, 이미지 생성과 이미지/비디오
분석은 OpenAI, Gemini, Doubao/Seedream, GLM, Grok, Qwen, MiniMax 또는
OpenAI 호환 이미지 엔드포인트 중에서 선택합니다.

전체 제품군은 라이트, 다크, 시스템 테마를 지원합니다. 테마는 화면에
보이는 모습만 바꿀 뿐입니다: 내보내기, 인쇄, 저장된 파일은 항상 문서
고유의 색상을 그대로 유지합니다.

<a id="download"></a>

## 다운로드

| 플랫폼                               | 요구 사항                                             | 다운로드                                                                                |
| ------------------------------------ | ----------------------------------------------------- | --------------------------------------------------------------------------------------- |
| **macOS** — Apple Silicon (arm64)    | macOS 11+                                             | [최신 `.dmg` (arm64)](https://github.com/genspark-ai/genoffice/releases/latest)         |
| **macOS** — Intel (x64)              | macOS 11+                                             | [최신 `.dmg` (x64)](https://github.com/genspark-ai/genoffice/releases/latest)           |
| **Windows** (x64, 대부분의 PC)       | Windows 10+, Intel/AMD                                | [최신 `-x64.exe` 설치 파일](https://github.com/genspark-ai/genoffice/releases/latest)   |
| **Windows** on Arm (ARM64)           | Windows 11 on Arm (Snapdragon X 및 유사 기종)         | [최신 `-arm64.exe` 설치 파일](https://github.com/genspark-ai/genoffice/releases/latest) |
| **Linux** — Debian / Ubuntu          | x86_64, glibc 2.34+ (Ubuntu 22.04 이상)               | [최신 `.deb`](https://github.com/genspark-ai/genoffice/releases/latest)                 |
| **Linux** — Fedora / RHEL / openSUSE | x86_64, glibc 2.34+ (Fedora 35+, RHEL 9+, Leap 15.6+) | [최신 `.rpm`](https://github.com/genspark-ai/genoffice/releases/latest)                 |
| **Linux** — 기타 배포판              | x86_64, glibc 2.34+, FUSE 2                           | [최신 `.AppImage`](https://github.com/genspark-ai/genoffice/releases/latest)            |

모든 빌드는 `main` 브랜치에서 만들어지며, macOS와 Windows 설치 파일에는
서명이 되어 있습니다. 이전 버전은 [릴리스](https://github.com/genspark-ai/genoffice/releases) 페이지에서 확인할 수 있습니다.

<details>
<summary><b>Linux에 설치하기</b></summary>

deb 패키지는 apt로 설치합니다 — 필요한 의존성을 함께 설치하고 GenOffice를
애플리케이션 메뉴에 추가합니다:

```bash
sudo apt install ./genoffice_<version>_amd64.deb
```

Fedora / RHEL 계열 / openSUSE에서는 대신 rpm을 설치합니다:

```bash
sudo dnf install ./genoffice-<version>.x86_64.rpm     # Fedora / RHEL family
sudo zypper install ./genoffice-<version>.x86_64.rpm  # openSUSE
```

AppImage는 별도 설치 없이 그 자리에서 실행됩니다: FUSE 2 런타임을
설치하고(`sudo apt install libfuse2`; Ubuntu 24.04에서는 패키지 이름이
`libfuse2t64`), 파일에 실행 권한을 부여한 뒤 실행하세요:

```bash
chmod +x GenOffice-<version>.AppImage
./GenOffice-<version>.AppImage
```

</details>

## 동작 방식

Docs, Sheets, Slides, PDF, Markdown, HTML, 그리고 탭 기반 셸까지 일곱
개의 Electron 앱은 순수 TypeScript 패키지로 이루어진 하나의 엔진
레이어와 `.xlsx`용 Rust 사이드카를 함께 사용합니다. 원본 파일은 언제나
신뢰할 수 있는 원본(source of truth)이며, 편집은 좁은 범위의 패치로만
적용되어 편집기가 손대지 않은 부분은 왕복 과정에서도 그대로 유지됩니다.

```
open docx ─► archive original by hash (never touched)
          ─► parse word/document.xml into a block tree, each block anchored to its original XML
          ─► Tiptap editor (manual + AI editing, dirty tracking)
save      ─► dirty blocks → OOXML fragments (referencing existing styles only)
          ─► splice into the original document.xml; untouched blocks keep their bytes
          ─► repack the zip; every other entry is copied byte-for-byte
```

패키지별 상세 설명(docx/pptx 엔진, `pdf2docx`, `html2docx`, 에이전트
코어와 프로바이더)은 [CONTRIBUTING.md](../../CONTRIBUTING.md#engine-packages)에 있습니다.

## 개발

```bash
npm install
npm run fixtures     # generate test .docx fixtures
npm test             # engine + app unit tests (docs/sheets/slides need no display)
npm run typecheck    # tsc --noEmit across every workspace
npm run dev          # all six editors + shell against Vite dev servers
npm run dev:docs     # a single app (same pattern works per workspace)
npm run dist:mac     # package macOS dmg (regenerates third-party notices)
npm run dist:win     # package Windows nsis installer
npm run dist:linux   # package Linux AppImage + deb + rpm
```

sheets 앱은 xlsx 사이드카를 위해 Rust 툴체인이 추가로 필요합니다
(`cargo`가 PATH에 있어야 함). `npm run build -w @genoffice/sheets`를
실행하면 자동으로 컴파일됩니다. 모든 변경 사항이 통과해야 하는 검사와
풀 리퀘스트가 병합되는 과정은 [CONTRIBUTING.md](../../CONTRIBUTING.md)를 참고하세요.

## 커뮤니티

GenOffice는 활발히 개발되고 있으며, 여러분의 피드백이 그 방향을
만들어갑니다.

- **버그를 신고하거나 기능을 요청**하려면
  [GitHub Issues](https://github.com/genspark-ai/genoffice/issues)를 이용하세요.
- **GenOffice 그룹 채팅에 참여**하려면
  [GenTeam](https://genoffice.ai/join)에서 팀 및 다른 사용자들과 대화할 수 있습니다.
- GenOffice가 유용하다고 느끼셨다면 **저장소에 스타를 눌러주세요** —
  프로젝트를 지원하는 가장 좋은 방법입니다.

## 자주 묻는 질문

<details>
<summary><b>GenOffice는 무료인가요?</b></summary>

네. GenOffice는 Apache-2.0 라이선스 아래 무료로 제공되는 오픈소스
소프트웨어입니다 — 체험판도 없고, 앱 자체에 유료 등급도 없습니다.

</details>

<details>
<summary><b>GenOffice에서 마이크로소프트 Word, Excel, PowerPoint 파일을 열 수 있나요?</b></summary>

네. GenOffice는 네이티브 `.docx`, `.xlsx`, `.pptx` 파일을 열고 저장할 수
있습니다. 저장은 바이트 단위로 보존되므로, 손대지 않은 부분은 그대로
다시 기록되어 문서가 마이크로소프트 오피스에서도 계속 정상적으로
열립니다.

</details>

<details>
<summary><b>GenOffice는 오프라인에서도 동작하나요?</b></summary>

문서 편집은 완전히 로컬에서 이루어집니다 — 파일을 열고, 편집하고,
저장하고, 변환하는 과정에서 파일이 기기를 벗어나지 않습니다. AI
기능(에이전트, 검색, 이미지 도구)은 네트워크 연결이 필요하며, Genspark
로그인 또는 직접 발급받은 모델 API 키가 있어야 합니다.

</details>

<details>
<summary><b>GenOffice에서 PDF 파일을 편집할 수 있나요?</b></summary>

네 — 위에 덮어씌우는 주석이 아니라, 원본 폰트를 유지한 채 페이지
콘텐츠 스트림 자체를 다시 작성하는 실제 PDF 텍스트/이미지 편집 기능을
제공합니다.

</details>

<details>
<summary><b>GenOffice에서 PDF를 Word, Excel, PowerPoint로 변환할 수 있나요?</b></summary>

네 — 전적으로 기기 내에서 처리됩니다: PDFium을 이용한 문자 단위 추출과
기하학 기반 레이아웃 분석만 사용하며, 클라우드 서비스나 업로드가 필요
없습니다. 스캔된 페이지도 지원합니다: macOS와 Windows에서는 시스템 OCR이
텍스트를 인식해 페이지 이미지가 아니라 편집 가능한 텍스트로 변환됩니다.

</details>

<details>
<summary><b>제 AI 모델이나 API 키를 직접 사용할 수 있나요?</b></summary>

네. 키가 필요 없는 Genspark 로그인 외에도, GenOffice는 Claude, OpenAI,
Gemini, DeepSeek, Kimi, GLM, Qwen, Doubao, MiniMax, Grok, Mistral,
OpenRouter, OpenCode Zen/Go를 위한 직접 키 사용을 지원하며, 로컬 모델
서버를 포함한 모든 OpenAI 호환 엔드포인트도 사용할 수 있습니다. 검색,
이미지 생성, 이미지/비디오 분석은 설정 → AI 미디어 및 검색에서 별도의
키를 설정합니다.

</details>

<details>
<summary><b>GenOffice에서 HTML을 Word로 변환할 수 있나요?</b></summary>

네 — HTML 앱의 Word로 내보내기 기능은 전적으로 기기 내에서 네이티브
편집 가능한 `.docx`를 생성합니다. 페이지는 내장 Chromium에서 렌더링된
뒤 실제 Word 구조로 변환됩니다: 제목, 문단, 목록, 표, 카드, KPI 행,
양식 필드, 페이지 배경까지 포함됩니다. Word에 대응 요소가 없는 시각
요소(차트, 아이콘, 장식된 박스)만 이미지로 삽입됩니다.

</details>

<details>
<summary><b>GenOffice가 데이터를 수집하나요?</b></summary>

공식 패키지 빌드는 기본적으로 제한된 사용 통계를 전송하며, 설정 →
일반에서 언제든지 이 보고를 끌 수 있습니다. 통계에는 문서 내용, 파일
이름, 파일 경로, 계정 신원, 이메일 주소가 절대 포함되지 않습니다.
전체 이벤트와 데이터 공개 내용은 [GenOffice Privacy](../../PRIVACY.md)에서 확인할 수 있습니다.

</details>

## 보안

프로세스 보안 방침(렌더러 샌드박싱, IPC 검증, 외부 링크 차단)과 AI
생성 콘텐츠에 대한 위협 모델은 [SECURITY.md](../../SECURITY.md)를 참고하세요.

## 감사의 말

GenOffice는 다음의 오픈소스 프로젝트가 없었다면 만들어질 수 없었습니다:

- [Electron](https://www.electronjs.org/) — 모든 앱의 데스크톱 런타임입니다.
- [Univer](https://github.com/dream-num/univer) (Apache-2.0) — Sheets가
  확장해서 사용하는 스프레드시트 UI 코어입니다.
- [PDFium](https://pdfium.googlesource.com/pdfium/) (BSD-3-Clause,
  [@embedpdf/pdfium](https://github.com/embedpdf/embed-pdf-viewer)을 통해
  포함됨) — 실제 PDF 텍스트 및 이미지 편집을 지원하는 콘텐츠 스트림
  엔진입니다.
- [pdf.js](https://github.com/mozilla/pdf.js) (Apache-2.0)와
  [pdf-lib](https://github.com/Hopding/pdf-lib) (MIT) — PDF 렌더링과
  문서 조합을 담당합니다.
- [Tiptap](https://tiptap.dev/) / [ProseMirror](https://prosemirror.net/) —
  Docs와 Markdown의 블록 에디터입니다.
- [CodeMirror](https://codemirror.net/) (MIT) — HTML 앱의 소스
  에디터입니다.
- [Konva](https://konvajs.org/) — Slides와 Sheets 차트의 캔버스
  렌더링을 담당합니다.
- [HarfBuzz](https://github.com/harfbuzz/harfbuzz) (wasm) — 복잡한
  문자 체계를 위한 텍스트 셰이핑 계량을 제공합니다.
- [calamine](https://github.com/tafia/calamine)과
  [IronCalc](https://github.com/ironcalc/IronCalc) — Rust xlsx
  사이드카의 읽기 및 계산 레이어입니다.
- [libeot](https://github.com/umanwizard/libeot) (MPL-2.0) — 임베디드
  PowerPoint 폰트를 위한 MicroType Express 디코더를 TypeScript로 포팅한
  것입니다.
- [React](https://react.dev/) (MIT) — 모든 앱의 UI 레이어입니다.
- [Mermaid](https://mermaid.js.org/) (MIT)와 [KaTeX](https://katex.org/)
  (MIT) — Markdown과 Docs의 다이어그램과 수식을 담당합니다.
- [opentype.js](https://opentype.js.org/) (MIT) — 메트릭 확인과 글리프
  조회를 위한 폰트 파싱을 담당합니다.
- [JSZip](https://stuk.github.io/jszip/) (MIT)와
  [fast-xml-parser](https://github.com/NaturalIntelligence/fast-xml-parser)
  (MIT) — OOXML 컨테이너와 XML 레이어입니다.
- [Fluent UI System Icons](https://github.com/microsoft/fluentui-system-icons)
  (MIT) — 각 리본에서 사용하는 아이콘 세트입니다.
- [electron-updater](https://www.electron.build/) (MIT) — 앱 내 업데이트를
  담당합니다.
- Liberation, Carlito, Caladea, Noto CJK 폰트(OFL/Apache-2.0) — 함께
  제공되는 문서용 폰트입니다.

`npm run notices`는 함께 제공되는 서드파티 라이선스 요약
(`tools/gen-third-party-notices.mjs`)을 다시 생성합니다. 모든 런타임
의존성은 MIT/Apache-2.0/BSD-3-Clause/OFL 라이선스를 따릅니다.

## 라이선스

GenOffice는 [Apache License 2.0](../../LICENSE)에 따라 라이선스가
부여됩니다. 단, `ee/` 디렉터리는 향후 엔터프라이즈 모듈을 위해
예약되어 있으며 [GenOffice Enterprise License](../../ee/LICENSE)가
적용됩니다.

GenOffice와 Genspark의 이름과 로고는 Mainfunc, Inc.의 상표입니다.
Apache-2.0 라이선스는 이를 사용할 권리를 부여하지 않으며(6항 참고),
포크는 자체 브랜딩을 사용해야 합니다.
