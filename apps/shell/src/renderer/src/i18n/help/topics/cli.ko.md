# 명령줄과 에이전트

모든 설치본에는 `genoffice` 명령이 들어 있으며, 창이 쓰는 바로 그 엔진을 구동합니다. 같은 파서, 같은 writer, 같은 렌더러입니다. 앱이 저장한 파일과 명령이 쓴 파일은 같은 파일이고, 앱의 AI 패널이 통과한 검사는 명령도 통과합니다.

```sh
genoffice --help          # every command
genoffice guide slides    # the op reference, for writing your own
```

그게 전부입니다, 화면 한 번에 다 들어 있습니다 — 각 명령 옆에는 하는 일을 한 줄로 적은 설명이 붙고, 그다음 전역 옵션과 종료 코드가 나옵니다:

![genoffice --help의 실제 출력: 버전 배너, 각 명령에 대한 한 줄 설명, 그리고 전역 옵션과 종료 코드](img/cli.png)

## 명령 준비하기

macOS와 Windows에서는 앱 번들 안에 있습니다. 이름으로 부르려면 `genoffice install-cli`을 한 번 실행하세요. 번들에 들어 있는 바이너리를 `/usr/local/bin`에, Windows에서는 사용자 `PATH`에 심볼릭 링크로 연결합니다.

## 알아 둘 만한 명령

| 명령              | 하는 일                                                                                                                                                                            |
| ----------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `open`            | 앱에서 문서를 엽니다. 앱이 실행 중이 아니면 실행합니다.                                                                                                                            |
| `convert`         | 앱 자체 엔진으로 형식을 변환합니다.                                                                                                                                                |
| `create`          | 구조화된 내용으로 문서를 만듭니다.                                                                                                                                                 |
| `render`          | 렌더러가 배치한 그대로 한 페이지에 한 장씩 PNG로 내보냅니다.                                                                                                                       |
| `pdf`             | 앱 프로세스를 띄우지 않고 PDF의 텍스트 층을 한 페이지씩 읽습니다.                                                                                                                  |
| `info`            | 문서의 메타데이터와 구조 요약.                                                                                                                                                     |
| `search`          | 앱에 설정된 제공자를 통해 웹이나 이미지를 검색합니다.                                                                                                                              |
| `image` / `media` | 이미지를 만들거나, 이미지·동영상·오디오 파일을 설명하고 질문하게 합니다.                                                                                                           |
| `merge`           | `.docx`, `.pptx`, `.xlsx` 템플릿의 `{{key}}` 자리표시자를 채웁니다.                                                                                                                |
| `capabilities`    | 이 컴퓨터에서 클라우드 기능이 어떻게 설정되어 있는지 보고합니다.                                                                                                                   |
| `guide`           | op 레퍼런스와 디자인 가이드입니다. executor가 검증하는 바로 그 정의에서 만들어지므로 `apply`가 받아들이는 것과 어긋날 수 없습니다. `--json`을 주면 각 op의 스키마와 함께 나옵니다. |
| `install-cli`     | `genoffice`를 `PATH`에 넣습니다.                                                                                                                                                   |
| `skill`           | 이 컴퓨터에서 찾은 코딩 에이전트를 나열하고, 그 안에 GenOffice 스킬을 설치하거나 갱신합니다.                                                                                       |
| `mcp`             | 모든 명령을 Model Context Protocol 도구로 제공합니다. **코딩 에이전트 연결**를 보세요.                                                                                             |

## 편집: 문서, 스프레드시트, 슬라이드

`genoffice docs`, `genoffice sheet`, `genoffice slides`는 앱이 쓰는 것과 같은 쓰기 경로로 읽고 편집합니다. 어휘도 하나를 나눕니다. **op**은 한 번의 편집이고, **spec**은 op을 순서대로 적용한 목록입니다.

```sh
genoffice sheet read report.xlsx
genoffice slides apply deck.pptx --ops ops.json --dry-run
genoffice docs apply brief.docx --ops ops.json
```

`--dry-run`은 그 묶음이 무엇을 할지 보고만 하고 아무것도 쓰지 않습니다. spec을 적용하기 전에 싸게 확인하는 방법입니다. 앱의 AI 패널이 도는 것도 바로 이 op이므로, AI에게 시킬 수 있는 건 전부 스크립트로 할 수 있습니다.

## Model Context Protocol

`genoffice mcp`는 모든 명령을 MCP 도구로 제공하고, `genoffice mcp install <agent|all>`는 코딩 에이전트 자체 설정에 등록합니다. 그쪽은 **코딩 에이전트 연결**를 보세요.

## 명령이 하지 않는 것

파일은 읽고 씁니다. 앱 그 자체는 아닙니다. 창이 없고, 앱 내 업데이트 대화상자도 적용되지 않습니다. 창이 필요한 일 — AI 패널, 렌더링한 슬라이드에 대한 QC 검사 — 은 파일을 열 때까지 기다려야 합니다. `genoffice render`는 창 없이 픽셀을 얻게 해 주고, `genoffice slides`는 덱의 레이아웃을 스스로 검사합니다.
