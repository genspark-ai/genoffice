

Rất nhiều anh em công nghệ đang chuyển từ các sản phẩm Office bản quyền qua nguồn mở: Đây là một bộ Microsoft Office mã nguồn mở tích hợp AI ngay từ đầu

Sau làn sóng AI Coding và AI Agent, giờ đây đến lượt AI Office trở thành xu hướng mới. GenOffice là một dự án mã nguồn mở với tham vọng xây dựng bộ ứng dụng văn phòng thế hệ mới, nơi AI không chỉ hỗ trợ mà còn trực tiếp hoàn thành công việc cho người dùng.

Điểm nổi bật của dự án:

* Tích hợp đầy đủ Docs, Sheets, Slides và PDF trong một ứng dụng.
* AI được tích hợp sẵn để viết tài liệu, phân tích dữ liệu và tạo bài thuyết trình.
* Có thể nghiên cứu thông tin, tổng hợp nội dung và sinh tài liệu chỉ từ một yêu cầu.
* Mã nguồn mở, cho phép cộng đồng cùng phát triển và mở rộng.
* Hỗ trợ trên cả Windows, macOS và Linux.
* Thiết kế hướng tới trải nghiệm AI-first thay vì chỉ bổ sung AI vào các ứng dụng văn phòng truyền thống. (Reddit⁠￼)

Ứng dụng thực tế:

* Soạn thảo báo cáo bằng AI.
* Phân tích dữ liệu và tạo bảng tính thông minh.
* Tự động xây dựng slide thuyết trình.
* Tóm tắt và chỉnh sửa tài liệu PDF.
* Xây dựng AI Workspace cho cá nhân hoặc doanh nghiệp.

Điều thú vị nhất là GenOffice không chỉ thêm một nút “AI” vào Office, mà được thiết kế ngay từ đầu theo triết lý AI-native. Thay vì phải chuyển qua lại giữa chatbot và Word, Excel hay PowerPoint, AI có thể trực tiếp nghiên cứu, phân tích dữ liệu, viết tài liệu và tạo slide ngay trong cùng một không gian làm việc. Đây có thể sẽ là hướng phát triển của các bộ ứng dụng văn phòng trong vài năm tới.


# GenOffice

An AI-native office suite for macOS and Windows: word processor, spreadsheet,
presentations, and PDF — five Electron apps sharing one engine layer, built
around AI editing as a first-class workflow rather than a bolted-on chat box.

[![Meet GenOffice — the world's first full-featured open-source AI Office (video)](https://img.youtube.com/vi/B2pLdMX95v4/maxresdefault.jpg)](https://www.youtube.com/watch?v=B2pLdMX95v4)

[Watch the demo video on YouTube](https://www.youtube.com/watch?v=B2pLdMX95v4)

## Download

Signed installers built from `main`:

- **macOS** (Apple Silicon): [GenOffice-0.5.83-arm64.dmg](https://github.com/genspark-ai/genoffice/releases/download/v0.5.83/GenOffice-0.5.83-arm64.dmg)
- **Windows** (x64): [GenOfficeSetup-v0.5.79.exe](https://github.com/genspark-ai/genoffice/releases/download/v0.5.83/GenOfficeSetup-v0.5.79.exe)

Previous version:

- **macOS** (Apple Silicon): [GenOffice-0.5.1-arm64.dmg](https://github.com/genspark-ai/genoffice/releases/download/v0.5.1/GenOffice-0.5.1-arm64.dmg)
- **Windows** (x64): [GenOfficeSetup-v0.5.1.exe](https://github.com/genspark-ai/genoffice/releases/download/v0.5.1/GenOfficeSetup-v0.5.1.exe)

Other versions are on the [Releases](https://github.com/genspark-ai/genoffice/releases) page.

## Apps

| App           | Product              | What it is                                                                                                                                                                                                                                                                                                                                                    |
| ------------- | -------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `apps/docs`   | **GenOffice Docs**   | `.docx` word processor. Byte-preserving round trip: only dirty paragraphs are regenerated (paragraph patch), everything else in the original file is kept byte-for-byte, so opening and saving never breaks layout in Word. Paginated view whose line metrics reproduce the original document's layout, tracked changes, comments, styles, equations, ink.    |
| `apps/sheets` | **GenOffice Sheets** | `.xlsx` spreadsheet. UI built on the open-source [Univer](https://github.com/dream-num/univer) core (Apache-2.0) with a large layer of in-house extensions; `.xlsx` import/export runs through an in-house Rust sidecar (calamine + IronCalc), charts are rendered in-house (Konva), plus pivot tables, slicers, conditional formatting, and formula tracing. |
| `apps/slides` | **GenOffice Slides** | `.pptx` presentations. In-house `.pptx` parse/render/edit engine with masters, charts, cropping, ink, and text shaping (HarfBuzz metrics).                                                                                                                                                                                                                    |
| `apps/pdf`    | **GenOffice PDF**    | `.pdf` viewer/editor on pdf.js + pdf-lib: annotations, forms, outlines, stamps, signatures, page operations, and printing support.                                                                                                                                                                                                                            |
| `apps/shell`  | **GenOffice**        | The suite shell: home screen, tabbed hosting of the four editors, auto-update.                                                                                                                                                                                                                                                                                |

Every app embeds the same AI panel: block-granular AI editing with version
snapshots and diffs in docs, a tool-calling agent over workbook/slide/PDF
state in the others.

**AI providers.** The apps sign in to a Genspark account and route model
calls through the Genspark service side; no model API key is stored locally.

## Engine packages

All pure TypeScript, no Electron dependency, unit-tested (except the UI kit):

- `packages/docx-engine` — docx parsing → block tree (with `docxIndex`
  anchors and passthrough), OOXML fragment generation, byte-level paragraph
  patching.
- `packages/pptx-engine` / `packages/pptx-render` — pptx model and rendering.
- `packages/file-parse` — text extraction for AI attachments (office formats,
  text formats).
- `packages/agent-core` — the AI agent loop and skill composition shared by
  every app.
- `packages/ai-provider` — provider abstraction and streaming for the model
  backends.
- `packages/ai-search` — Genspark auth + web/image search tools.
- `packages/i18n`, `packages/ui`, `packages/project-store`,
  `packages/electron-utils` — shared i18n core, React UI kit, recent-files
  store, and Electron main-process helpers.

## Development

```bash
npm install
npm run fixtures     # generate test .docx fixtures
npm test             # engine + app unit tests (docs/sheets/slides need no display)
npm run typecheck    # tsc --noEmit across every workspace
npm run dev          # all four editors + shell against Vite dev servers
npm run dev:docs     # a single app (same pattern works per workspace)
npm run dist:mac     # package macOS dmg (regenerates third-party notices)
npm run dist:win     # package Windows nsis installer
```

The sheets app additionally needs a Rust toolchain for its xlsx sidecar
(`cargo` on PATH); `npm run build -w @genoffice/sheets` compiles it
automatically.

Local UI/e2e driver scripts (Playwright + Electron, for local acceptance, not
committed by default) live in [`scripts/drivers/`](scripts/drivers/README.md).

## Architecture notes (docx round trip)

```
open docx ─► archive original by hash (never touched)
          ─► docx-engine parses word/document.xml top-level elements (w:p / w:tbl / …)
          ─► Block tree, each block anchored by docxIndex + original XML slice
          ─► Tiptap streaming editor (manual + AI editing, dirty tracking)
save      ─► dirty blocks → OOXML fragments (referencing existing styles only)
          ─► splice into original document.xml (untouched blocks keep original bytes)
          ─► repack zip; all other entries copied byte-for-byte
```

The same philosophy holds in sheets and slides: the original file is the
source of truth, edits are applied as narrow patches, and everything the
editor didn't touch survives the round trip untouched.

## Security

See [SECURITY.md](SECURITY.md) for the process security posture (renderer
sandboxing, IPC validation, external-link gating) and the threat models for
AI-generated content.

## Third-party notices

`npm run notices` regenerates the bundled third-party license summary
(`tools/gen-third-party-notices.mjs`); all runtime dependencies are
MIT/Apache-2.0/OFL, and the bundled fonts (Liberation, Carlito, Caladea, Noto
CJK subsets) are OFL/Apache.

## License

GenOffice is licensed under the [Apache License 2.0](LICENSE), with one
exception: the `ee/` directory is reserved for future enterprise modules and
is covered by the [GenOffice Enterprise License](ee/LICENSE).

The GenOffice and Genspark names and logos are trademarks of Mainfunc, Inc.
The Apache-2.0 license does not grant permission to use them (see section 6);
forks should use their own branding.
