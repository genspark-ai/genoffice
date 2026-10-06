# 코딩 에이전트 연결

GenOffice는 Model Context Protocol을 말하므로, 코딩 에이전트가 앱이 쓰는 바로 그 엔진으로 문서를 읽고, 쓰고, 렌더링할 수 있습니다. 에이전트가 파일 형식을 추측하는 일은 없습니다. executor가 검증하는 바로 그 정의에서 타입이 붙은 op 스키마를 받습니다.

## 등록하기

보통은 명령 하나로 끝납니다.

```sh
genoffice mcp install all
```

이 컴퓨터에 있는 코딩 에이전트(Claude Code, Codex, Cursor, Gemini CLI, Copilot CLI, OpenCode, Windsurf)를 찾아, 각자의 설정에 stdio 서버 항목을 씁니다. 그 파일의 나머지는 있는 그대로 둡니다.

```sh
genoffice mcp list             # where each agent stands
genoffice mcp install cursor    # just one
genoffice mcp uninstall cursor  # take it back out
```

평소와 다른 곳에 설치한 에이전트에는 `--dir <path>`를 줍니다. `--force`는 이미 있는 항목을 다시 씁니다.

## 직접 실행하기

다른 컴퓨터에 있는 클라이언트라면 HTTP로 내보내세요.

```sh
genoffice mcp --http 8765 --token <secret>
```

`--host <addr>`는 어디에서 받을지를 바꿉니다. 클라이언트에는 같은 토큰을 넘기세요.

HTTP로도 파일이 오갑니다. `PUT /files/<name>`으로 하나를 올리고, 모든 도구는 경로 대신 `http(s)` URL을 받습니다. 출력은 다운로드 URL로 돌아오고, 충분히 작으면 포함 리소스로도 돌아옵니다. op, spec, Markdown은 어느 쪽이든 인라인으로 넘깁니다.

## 에이전트가 얻는 것

모든 명령이 도구입니다. 그중 흥미로운 것들:

- **`docs`, `sheet`, `slides`** — 앱 자체 쓰기 경로로 파일을 읽고 편집합니다. 한 번에 **op** 하나씩. 새 덱은 `deck_start`, `deck_page`, `deck_build`로 갑니다.
- **`render`** — 한 페이지에 한 장씩 PNG를, 앱의 렌더러가 배치한 그대로 내보내므로 에이전트가 슬라이드를 추측하지 않고 볼 수 있습니다.
- **`pdf`** — 앱 프로세스 없이 PDF의 텍스트 층을 한 페이지씩 읽습니다.
- **`info`** — 메타데이터와 구조 요약으로, 낯선 파일에 대한 첫 호출로는 거의 항상 가장 쌉니다.
- **`search`, `image`, `media`** — 앱에 설정된 제공자들이라 에이전트가 따로 키를 준비할 필요가 없습니다.
- **`merge`** — `{{key}}` 템플릿을 채웁니다.

## 스키마와 더 작은 예산

`apply`와 `create`는 `ops`, `cells`, `data` 매개변수를 op마다 타입이 붙은 스키마로 알립니다. 그 스키마는 `genoffice guide <domain> --json`에서 만들어집니다. 정확하지만 크지는 않습니다. 컨텍스트 창이 좁은 클라이언트는 대신 단순 배열을 요청할 수 있습니다:

```sh
genoffice mcp --compact-schemas
GENOFFICE_MCP_COMPACT_SCHEMAS=1 genoffice mcp
```

## 스킬

op 어휘를 모르는 에이전트는 추측합니다. `genoffice skill`은 찾은 에이전트들에 GenOffice 스킬을 설치하며, 레퍼런스와 디자인 가이드를 함께 넣습니다. `genoffice guide`가 출력하는 것과 같은 내용입니다.

## 이것이 아닌 것

MCP 서버는 파일을 읽고 씁니다. 창이 아닙니다. AI 패널이 없고, 앱 내 업데이트 대화상자도 적용되지 않습니다. 창이 필요한 단계라면 파일을 여세요.
