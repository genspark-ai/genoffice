/**
 * Web dual-protocol bridge — verification against the REAL built shell.
 *
 * The shell aggregates every editor module's register*Ipc() calls, so its
 * single process serves all channels over both transports: IPC (Electron
 * renderers) and HTTP/SSE (browsers, via @genoffice/ipc-bridge on 127.0.0.1:5399).
 * These checks drive real HTTP requests against the bridge of the launched app
 * and compare results with what the same app answers over IPC.
 *
 * Requires `npm run build:all` first (same as the rest of the e2e suite).
 * Covers acceptance items A1 (server dual protocol), A2 (dual-path parity),
 * A4 (real-time push over SSE), A5 (structured desktop-only degradation).
 *
 * NOTE on the launcher: helpers.launchShell's ready-check evaluates on the
 * first window, whose CDP session can be wedged by the screencast attaching
 * mid-navigation (evaluate hangs forever, page.url() stays empty — the same
 * symptom the Linux note in helpers describes). Waiting for a window whose URL
 * already points at the built renderer keeps the session healthy everywhere.
 */
import {
  _electron as electron,
  test,
  expect,
  type ElectronApplication,
  type Page,
} from '@playwright/test'
import { mkdtemp, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import { createRequire } from 'node:module'

// The built shell's bridge binds SHELL_IPC_PORT (5399 here) so this spec never
// collides with the dev shell's bridge on 5299 when both run at the same time.
const BRIDGE = 'http://127.0.0.1:5399'
const SHELL_DIR = resolve(__dirname, '../apps/shell')

interface Launched {
  app: ElectronApplication
  page: Page
}

async function launchShellLite(): Promise<Launched> {
  const require2 = createRequire(join(SHELL_DIR, 'package.json'))
  const userDataDir = await mkdtemp(join(tmpdir(), 'web-bridge-e2e-'))
  await writeFile(join(userDataDir, 'app-settings.json'), JSON.stringify({ onboardingSeen: true }))
  // ELECTRON_RUN_AS_NODE (set by some hosts) would boot Electron as plain Node
  const { ELECTRON_RUN_AS_NODE: _runAsNode, ...env } = process.env
  const app = await electron.launch({
    executablePath: require2('electron'),
    args: [SHELL_DIR],
    env: { ...env, GENOFFICE_USER_DATA: userDataDir, GENOFFICE_LANG: 'en', SHELL_IPC_PORT: '5399' },
    timeout: 30_000,
  })
  // poll window URLs instead of evaluating on a possibly-wedged about:blank
  const deadline = Date.now() + 60_000
  for (;;) {
    const win = app.windows().find((w) => w.url().includes('out/renderer/index.html'))
    if (win) return { app, page: win }
    if (Date.now() > deadline) {
      throw new Error(
        `shell renderer window never appeared; urls: ${app.windows().map((w) => w.url())}`,
      )
    }
    await new Promise((r) => setTimeout(r, 200))
  }
}

interface InvokeResult {
  status: number
  body: { ok?: boolean; result?: unknown; error?: { code?: string; message?: string } }
}

async function invoke(
  channel: string,
  args: unknown[] = [],
  session?: string,
): Promise<InvokeResult> {
  const response = await fetch(`${BRIDGE}/api/ipc/${encodeURIComponent(channel)}`, {
    method: 'POST',
    headers: {
      'content-type': 'application/json',
      ...(session ? { 'x-ipc-session': session } : {}),
    },
    body: JSON.stringify({ args }),
    signal: AbortSignal.timeout(10_000),
  })
  return { status: response.status, body: await response.json() }
}

/** Reads `count` data frames from the bridge's SSE push stream for a session. */
async function collectSseFrames(
  session: string,
  count: number,
): Promise<Array<{ channel: string; args: unknown[] }>> {
  const response = await fetch(`${BRIDGE}/api/ipc/events?session=${encodeURIComponent(session)}`, {
    signal: AbortSignal.timeout(30_000),
  })
  expect(response.status).toBe(200)
  const reader = response.body!.getReader()
  const decoder = new TextDecoder()
  let buffered = ''
  const frames: Array<{ channel: string; args: unknown[] }> = []
  while (frames.length < count) {
    const { done, value } = await reader.read()
    if (done) break
    buffered += decoder.decode(value, { stream: true })
    let boundary = buffered.indexOf('\n\n')
    while (boundary !== -1 && frames.length < count) {
      const raw = buffered.slice(0, boundary)
      buffered = buffered.slice(boundary + 2)
      if (raw.startsWith('data: ')) frames.push(JSON.parse(raw.slice('data: '.length)))
      boundary = buffered.indexOf('\n\n')
    }
  }
  return frames
}

test.describe('web bridge on the built shell', () => {
  let launched: Launched | null = null
  let launchPromise: Promise<Launched> | null = null

  // launched lazily inside the first test (this playwright version's beforeAll
  // has no timeout parameter — the launch must run under the test's own timeout)
  function getLaunched(): Promise<Launched> {
    launchPromise ??= launchShellLite()
    return launchPromise
  }

  test.afterAll(async () => {
    if (launched) await launched.app.close().catch(() => {})
  })

  test('A1: the bridge is live and serves real channels over HTTP', async () => {
    launched = await getLaunched()
    const health = (await (await fetch(`${BRIDGE}/api/ipc/health`)).json()) as {
      ok: boolean
      channels: number
    }
    expect(health.ok).toBe(true)
    // the shell registers home/docs/ai/project/tabs channels — a hundred plus
    expect(health.channels).toBeGreaterThan(50)

    const language = await invoke('app:get-language')
    expect(language.status).toBe(200)
    expect(language.body.ok).toBe(true)
    // the launcher pins GENOFFICE_LANG=en — the HTTP answer matches the app UI
    expect(language.body.result).toBe('en')
  })

  test('A2: same channel + payload returns identical results over IPC and HTTP', async () => {
    launched = await getLaunched()
    const { page } = launched

    const [ipcVersion, ipcTheme, ipcLang] = await Promise.all([
      page.evaluate(() =>
        (
          window as unknown as { aiOffice: { getAppVersion(): Promise<string> } }
        ).aiOffice.getAppVersion(),
      ),
      page.evaluate(() =>
        (window as unknown as { aiOffice: { getTheme(): Promise<string> } }).aiOffice.getTheme(),
      ),
      page.evaluate(() =>
        (
          window as unknown as { aiOffice: { getLanguage(): Promise<string> } }
        ).aiOffice.getLanguage(),
      ),
    ])

    const httpVersion = await invoke('home:get-app-version')
    const httpTheme = await invoke('home:get-theme')
    const httpLang = await invoke('home:get-language')

    expect(httpVersion.status).toBe(200)
    expect(httpVersion.body.result).toBe(ipcVersion)
    expect(httpTheme.body.result).toBe(ipcTheme)
    expect(httpLang.body.result).toBe(ipcLang)
    // sanity: the values are real, not three coincidental nulls
    expect(String(ipcVersion)).toMatch(/^\d+\.\d+/)
    expect(['system', 'light', 'dark']).toContain(ipcTheme)
  })

  test('A5: native-only channels get structured WEB_UNSUPPORTED over HTTP', async () => {
    await getLaunched()
    // native open dialog — the bridge refuses before the handler could open an
    // invisible dialog and hang a web caller
    const docsOpen = await invoke('docs:open')
    expect(docsOpen.status).toBe(400)
    expect(docsOpen.body.error?.code).toBe('WEB_UNSUPPORTED')
    expect(docsOpen.body.error?.message).toContain('desktop-only')

    // home file browser — same for the shell's own dialog channel
    const browse = await invoke('home:browse')
    expect(browse.status).toBe(400)
    expect(browse.body.error?.code).toBe('WEB_UNSUPPORTED')

    // unknown channels stay distinguishable from unsupported ones
    const missing = await invoke('no:such-channel')
    expect(missing.status).toBe(404)
    expect(missing.body.error?.code).toBe('IPC_NO_HANDLER')
  })

  test('A4: main→renderer push arrives over SSE in real time (ai:stream chunk)', async () => {
    await getLaunched()
    const session = 'e2e-shell-sse'
    const framesPromise = collectSseFrames(session, 1)
    // ai:stream without a configured key answers with a real push frame on
    // 'ai:stream-chunk' — exactly what the renderer's AI panel consumes
    const started = await invoke(
      'ai:stream',
      [
        {
          requestId: 'e2e-sse-1',
          settings: { provider: 'openai' },
          system: '',
          messages: [{ role: 'user', content: 'ping' }],
        },
      ],
      session,
    )
    expect(started.status).toBe(200)
    const frames = await framesPromise
    expect(frames).toHaveLength(1)
    expect(frames[0].channel).toBe('ai:stream-chunk')
    const chunk = frames[0].args[0] as { requestId: string; type: string }
    expect(chunk.requestId).toBe('e2e-sse-1')
    expect(chunk.type).toBe('error')
  })

  test('the Electron side keeps working while the bridge serves HTTP', async () => {
    launched = await getLaunched()
    const { page } = launched
    // the home screen rendered through the untouched IPC stack
    await expect(page.locator('body')).toContainText(/GenOffice|文档|Home/i)
  })
})
