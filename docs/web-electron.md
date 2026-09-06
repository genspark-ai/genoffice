# Web + Electron dual protocol

GenOffice runs the same renderer build as an Electron desktop app and a local
web app. Electron renderers call main-process handlers through `ipcRenderer`;
the browser calls the same handlers through the loopback HTTP bridge. Main-to-
renderer pushes use an isolated server-sent events stream.

## Architecture

- `@genoffice/ipc-bridge` wraps `ipcMain.handle`, `removeHandler`, and `on`
  registrations once, so existing and future channels are available over both
  protocols without per-channel registration.
- The browser transport maps `invoke`/`send` to `POST /api/ipc/:channel` and
  listener registration to `GET /api/ipc/events?session=<id>` (SSE).
- Binary arguments/results are encoded as tagged base64, preserving
  `ArrayBuffer` and typed-array payloads across JSON.
- Each HTTP caller owns a session. `event.sender.send` frames route only to
  that session's SSE stream. Pending frames are buffered for reconnection and
  expire after 60 seconds.
- Development pages are served by Vite and proxy `/api` to the bridge.
- Production web mode is served by the bridge from each app's `out/renderer`.

## Start the web version

Start a standalone app (this starts both the Electron main process and Vite):

```sh
npm run dev -w @genoffice/docs     # Vite http://localhost:5173, bridge 5273
npm run dev -w @genoffice/sheets   # Vite http://localhost:5174, bridge 5274
npm run dev -w @genoffice/slides   # Vite http://localhost:5175, bridge 5275
npm run dev -w @genoffice/pdf      # Vite http://localhost:5176, bridge 5276
npm run dev -w @genoffice/markdown # Vite http://localhost:5177, bridge 5277
npm run dev -w @genoffice/shell    # Vite http://localhost:5199, bridge 5299
```

Open the Vite URL in Chromium or Chrome. The desktop main process must keep
running: it is the backend for the web version. Use `DOCS_IPC_PORT`,
`SHEETS_IPC_PORT`, `SLIDES_IPC_PORT`, `PDF_IPC_PORT`, `MARKDOWN_IPC_PORT`, or
`SHELL_IPC_PORT` to change a bridge port.

The unified shell serves all module handlers on port `5299`:

```sh
npm run build:all
npm run start -w @genoffice/shell
curl http://127.0.0.1:5299/api/ipc/health
curl -X POST http://127.0.0.1:5299/api/ipc/home:get-language -H 'content-type: application/json' -d '{}'
```

## Real-launch verification

The web version is verified with a real browser against the documented startup
commands (no mocks, no spawned test servers). Start all six apps, then run the
Playwright specs that drive the flows over HTTP:

```sh
npm run dev -w @genoffice/docs     # terminal 1: Vite 5173 + bridge 5273
npm run dev -w @genoffice/sheets   # terminal 2: Vite 5174 + bridge 5274
npm run dev -w @genoffice/slides   # terminal 3: Vite 5175 + bridge 5275
npm run dev -w @genoffice/pdf      # terminal 4: Vite 5176 + bridge 5276
npm run dev -w @genoffice/markdown # terminal 5: Vite 5177 + bridge 5277
npm run dev -w @genoffice/shell    # terminal 6: Vite 5199 + bridge 5299
npx playwright test --config e2e/playwright.config.ts e2e/web-launch-all.spec.ts
```

The spec requires the dev servers to already be running and verifies, in a real
Chromium page over HTTP:

- docs: create → edit → save → reopen.
- markdown: edit → save → reopen.
- sheets: open a workbook (`workbook:open-path`) and read a cell back.
- slides: new blank → render → save, plus the web-native image-export overrides.
- pdf: open via `#open=<path>` → read → save.
- shell: home data and tab management over HTTP.
- native-only channels: every desktop-only channel has a browser equivalent
  wired in the web bridge (no `WEB_UNSUPPORTED` from the renderer's own flows).

`e2e/web-launch-verify.spec.ts` covers the docs/markdown subset, and
`e2e/web-bridge-shell.spec.ts` verifies the built shell's dual protocol
(IPC/HTTP parity, SSE pushes, and the `WEB_UNSUPPORTED` safety net for direct
calls to native-dialog channels).

Smoke checks for the running bridge:

```sh
curl http://127.0.0.1:5273/api/ipc/health   # {"ok":true,"channels":54}
curl -X POST http://127.0.0.1:5273/api/ipc/app:get-language \
  -H 'content-type: application/json' -d '{"args":[]}'  # {"ok":true,"result":"zh"}
```

## Protocol

Invoke with a JSON argument array and read the JSON result:

```sh
curl -X POST http://127.0.0.1:5273/api/ipc/app:get-language \
  -H 'content-type: application/json' \
  -d '{"args":[]}'
```

Success is `{"ok":true,"result":...}`. Handler failures preserve the message and
return an HTTP 500 response. Unknown channels return `IPC_NO_HANDLER`; blocked
desktop-only channels return `WEB_UNSUPPORTED`.

For production web builds, the bridge can serve the built renderer directly,
keeping API calls same-origin. Build first, then launch the desktop process.

## Capability matrix

Web calls run in the local desktop main process, so file storage and local
compute remain available. Interactive native OS UI and desktop-only resources
get browser equivalents in each app's web bridge (the same transport-agnostic
API factories the preload exposes, bound to HTTP/SSE instead of `ipcRenderer`):

| Capability                                          | Web behavior                                                       |
| --------------------------------------------------- | ------------------------------------------------------------------ |
| Data conversion and document compute                | Available — runs in the local main process.                        |
| Local project/chat storage and file persistence     | Available — main process has the same filesystem access.           |
| AI requests, search, and stream chunks              | Available — main process performs requests; chunks arrive over SSE.|
| Open / file selection                               | Browser `<input type=file>` → `web:write-temp-file` → main process opens the temp path. |
| Save / export                                       | Main process returns bytes → browser Blob download.                |
| Print / PDF export                                  | `window.print()` (save-as-PDF) or PDF bytes downloaded.            |
| Screen capture                                      | `getDisplayMedia` → canvas frame.                                  |
| Window / tab management                             | `window.open` browser tabs; shell tabs resolve to no-ops.          |
| Clipboard                                           | `navigator.clipboard` (async Clipboard API).                      |
| Font metrics / font install                         | Canvas `measureText`; `FontFace` where applicable.                |
| Fullscreen                                          | `requestFullscreen`.                                               |
| Drag-and-drop path resolution and `md-asset://`     | Unsupported — requires Electron protocol/native file binding.      |

Direct HTTP calls to channels that would open a native dialog (for example
`docs:open`, `home:browse`) still receive a structured `WEB_UNSUPPORTED` error
instead of opening an invisible dialog; the renderer's own flows never hit
that path because the web bridge overrides them with the browser equivalents
above. Electron behavior is unchanged: preload exposes transport-agnostic API
factories bound to `ipcRenderer`, while browser bootstraps expose the same
factories bound to HTTP/SSE.
