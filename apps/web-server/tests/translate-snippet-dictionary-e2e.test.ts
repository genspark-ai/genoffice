/**
 * Snippet translation E2E — terminology provenance.
 *
 * `home:translate-snippet` is what the Settings → 翻译知识库 pane's paste box
 * calls. It has to report *which* rules shaped the answer, because the pane
 * renders "KB · N" / "Dictionary · N" badges off that data, and it has to let
 * the user reuse the `--dictionary` they just generated for a file instead of
 * re-entering the same terms by hand.
 *
 * The provider is a local OpenAI-compatible stub, so this exercises the real
 * prompt assembly, the real KB/dictionary resolution and the real enforcement
 * pass — not a mocked translation-core.
 */
import { afterAll, beforeAll, describe, expect, it } from 'vitest'
import { spawn, type ChildProcess } from 'node:child_process'
import { createServer, type Server } from 'node:http'
import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { tmpdir } from 'node:os'

interface IpcResult<T = unknown> {
  ok: boolean
  result: T
}

async function ipc<T = unknown>(base: string, channel: string, args: unknown[] = []): Promise<IpcResult<T>> {
  const res = await fetch(`${base}/api/ipc/${channel}`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ args }),
  })
  return (await res.json()) as IpcResult<T>
}

async function waitForHealth(base: string, timeoutMs = 30_000): Promise<void> {
  const deadline = Date.now() + timeoutMs
  while (Date.now() < deadline) {
    try {
      const res = await fetch(`${base}/health`)
      if (res.ok) return
    } catch {
      /* keep polling */
    }
    await new Promise((r) => setTimeout(r, 250))
  }
  throw new Error(`web-server did not become healthy within ${timeoutMs}ms`)
}

/** The stub echoes the source text back, so a term the "model" left untranslated
 *  is observable: the enforcement pass must substitute it. */
function startFakeProvider(): Promise<{ server: Server; port: number; prompts: string[] }> {
  const prompts: string[] = []
  return new Promise((resolve) => {
    const server = createServer((req, res) => {
      let body = ''
      req.on('data', (chunk) => (body += chunk))
      req.on('end', () => {
        try {
          const parsed = JSON.parse(body) as { messages?: Array<{ content?: string }> }
          const system = parsed.messages?.[0]?.content ?? ''
          const user = parsed.messages?.[1]?.content ?? ''
          prompts.push(`${system}\n---\n${user}`)
          const source = /<source_text>\n([\s\S]*?)\n<\/source_text>/.exec(user)?.[1] ?? ''
          res.writeHead(200, { 'Content-Type': 'application/json' })
          res.end(
            JSON.stringify({
              choices: [{ message: { role: 'assistant', content: `译文：${source}` } }],
            }),
          )
        } catch {
          res.writeHead(500).end('{}')
        }
      })
    })
    server.listen(0, '127.0.0.1', () => {
      const address = server.address()
      resolve({ server, port: typeof address === 'object' && address ? address.port : 0, prompts })
    })
  })
}

describe('snippet translation + dictionary reuse E2E', () => {
  let server: ChildProcess | undefined
  let fake: { server: Server; port: number; prompts: string[] }
  let base: string
  let dataDir: string
  let dictionaryPath: string

  beforeAll(async () => {
    dataDir = mkdtempSync(join(tmpdir(), 'genoffice-tkb-snippet-e2e-'))
    fake = await startFakeProvider()
    const port = 21000 + Math.floor(Math.random() * 8000)
    base = `http://127.0.0.1:${port}`
    const bundle = join(__dirname, '..', 'dist', 'bundle', 'index.js')
    server = spawn(process.execPath, [bundle], {
      env: {
        ...process.env,
        PORT: String(port),
        DATA_DIR: dataDir,
        GENOFFICE_WEB_DATA_DIR: dataDir,
        GENOFFICE_TRANSLATION_KB: join(dataDir, 'translation-kb.json'),
        NO_OPEN: '1',
      },
      stdio: ['ignore', 'pipe', 'pipe'],
    })
    server.stdout?.on('data', () => {})
    server.stderr?.on('data', () => {})
    await waitForHealth(base)

    // Point the active provider at the stub. `config.baseUrl` overrides the
    // vendor endpoint, which is how the app supports regional mirrors.
    await ipc(base, 'ai:set-settings', [
      {
        provider: 'openai',
        providers: {
          openai: {
            apiKey: 'test-key',
            model: 'gpt-4o-mini',
            baseUrl: `http://127.0.0.1:${fake.port}`,
          },
        },
      },
    ])

    // A KB term so the two provenances can be told apart.
    await ipc(base, 'ai:translation-kb-upsert', [
      {
        id: 'e2e-fabric-weight',
        scope: 'company',
        priority: 50,
        sourceTerm: 'fabric weight',
        targetTerm: '克重',
        sourceLang: 'en-US',
        targetLang: 'zh-CN',
      },
    ])

    // The dictionary the pane writes after a file pass.
    dictionaryPath = join(dataDir, 'translation-dictionaries', 'e2e.json')
    mkdirSync(join(dataDir, 'translation-dictionaries'), { recursive: true })
    writeFileSync(dictionaryPath, JSON.stringify({ 'Oxford cloth': '牛津布' }, null, 2), 'utf8')
  }, 90_000)

  afterAll(() => {
    server?.kill('SIGTERM')
    fake?.server.close()
    try {
      rmSync(dataDir, { recursive: true, force: true })
    } catch {
      /* ignore */
    }
  })

  it('translates a snippet and reports the KB terms that applied', async () => {
    const { ok, result } = await ipc<{
      ok: boolean
      translation?: string
      matchedTerms?: string[]
      dictionaryHits?: string[]
      status?: string
    }>(base, 'home:translate-snippet', [
      {
        text: 'The fabric weight is heavy.',
        sourceLang: 'en-US',
        targetLang: 'zh-CN',
        useDictionary: false,
      },
    ])
    expect(ok).toBe(true)
    expect(result.ok).toBe(true)
    // The stub left the term in English; the KB enforcement pass replaced it.
    expect(result.translation).toBe('译文：The 克重 is heavy.')
    expect(result.matchedTerms).toEqual(['fabric weight'])
    expect(result.dictionaryHits).toEqual([])
  })

  it('applies a generated dictionary and reports the hits separately', async () => {
    const { result } = await ipc<{
      ok: boolean
      translation?: string
      matchedTerms?: string[]
      dictionaryHits?: string[]
      dictionary?: { path: string; terms: number; hits: number } | null
    }>(base, 'home:translate-snippet', [
      {
        text: 'Oxford cloth is durable.',
        sourceLang: 'en-US',
        targetLang: 'zh-CN',
        dictionaryPath,
        useDictionary: true,
      },
    ])
    expect(result.ok).toBe(true)
    expect(result.translation).toBe('译文：牛津布 is durable.')
    expect(result.dictionaryHits).toEqual(['Oxford cloth'])
    expect(result.matchedTerms).toEqual([])
    expect(result.dictionary?.path).toBe(dictionaryPath)
    expect(result.dictionary?.terms).toBe(1)
    expect(result.dictionary?.hits).toBe(1)
    // The term had to reach the model as an instruction, not only post-hoc.
    expect(fake.prompts.at(-1)).toContain('Oxford cloth => 牛津布')
  })

  it('skips the dictionary when the pane turns reuse off', async () => {
    // A distinct source text so the memory written by the previous test does
    // not short-circuit this one.
    const text = 'Our Oxford cloth ships weekly.'
    const { result } = await ipc<{
      ok: boolean
      translation?: string
      dictionaryHits?: string[]
      dictionary?: unknown
    }>(base, 'home:translate-snippet', [
      { text, sourceLang: 'en-US', targetLang: 'zh-CN', dictionaryPath, useDictionary: false },
    ])
    expect(result.ok).toBe(true)
    expect(result.translation).toBe(`译文：${text}`)
    expect(result.dictionaryHits).toEqual([])
    expect(result.dictionary).toBeNull()
    expect(fake.prompts.at(-1)).not.toContain('牛津布')
  })

  it('serves the second call from translation memory', async () => {
    const text = 'Merino Oxford cloth is warm.'
    const first = await ipc<{ status?: string }>(base, 'home:translate-snippet', [
      { text, sourceLang: 'en-US', targetLang: 'zh-CN', dictionaryPath, useDictionary: true },
    ])
    expect(first.result.status).toBe('translated')
    const before = fake.prompts.length
    const { result } = await ipc<{
      ok: boolean
      status?: string
      matchedTerms?: string[]
      dictionaryHits?: string[]
    }>(base, 'home:translate-snippet', [
      { text, sourceLang: 'en-US', targetLang: 'zh-CN', dictionaryPath, useDictionary: true },
    ])
    expect(result.status).toBe('memory-hit')
    // A memory hit still reports where the terminology came from.
    expect(result.dictionaryHits).toEqual(['Oxford cloth'])
    expect(result.matchedTerms).toEqual([])
    expect(fake.prompts.length).toBe(before)
  })
})
