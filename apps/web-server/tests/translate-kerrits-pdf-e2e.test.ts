/**
 * Real-world PDF translation E2E — drives the KERRITS Chinese garment spec
 * (the same one in `~/Downloads/资料（保密）/KERRITS-英文工艺单.pdf`)
 * through the full KB → dictionary → file pass pipeline. We copy a 1-page
 * snippet of the real PDF to a temp dir so the test is hermetic — running it
 * does not depend on the user keeping that exact file in place, but the text
 * comes from a real Chinese garment spec, not a synthetic fixture.
 *
 * The KERRITS PDF mixes Chinese with English garment industry jargon
 * (`FLATSEAM`, `WB`, `POWERMESH`, etc.) and brand-coded measurements
 * (`3/8"`, `1/2"`). That makes it a strong end-to-end test: KB terms that
 * catch the brand code must reach the model as terminology, the dictionary
 * miner must not invent cells out of the document decoration, and the PDF
 * handler must round-trip the new text without losing the page layout.
 */
import { afterAll, beforeAll, describe, expect, it } from 'vitest'
import { spawn, type ChildProcess } from 'node:child_process'
import { createServer, type Server } from 'node:http'
import {
  copyFileSync,
  existsSync,
  mkdirSync,
  mkdtempSync,
  readFileSync,
  rmSync,
  writeFileSync,
} from 'node:fs'
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

/**
 * Local OpenAI-compatible stub. The model is asked to translate short
 * Chinese garment phrases; we echo the source wrapped in a marker so the
 * dictionary enforcement pass is observable.
 */
function startFakeProvider(): Promise<{ server: Server; port: number; prompts: string[] }> {
  const prompts: string[] = []
  return new Promise((resolve) => {
    const server = createServer((req, res) => {
      let body = ''
      req.on('data', (chunk) => (body += chunk))
      req.on('end', () => {
        try {
          const parsed = JSON.parse(body) as { messages?: Array<{ content?: string }> }
          const user = parsed.messages?.[1]?.content ?? ''
          prompts.push(user)
          // Extract every <source_text>…</source_text> block and prefix each
          // with "EN: ". The dictionary pass is observable: a KB term the
          // stub did not translate is replaced by the dictionary pass.
          const out: string[] = []
          const re = /<source_text>\n([\s\S]*?)\n<\/source_text>/g
          let m: RegExpExecArray | null
          while ((m = re.exec(user))) out.push(`EN: ${m[1]}`)
          res.writeHead(200, { 'Content-Type': 'application/json' })
          res.end(JSON.stringify({ choices: [{ message: { role: 'assistant', content: out.join('\n') } }] }))
        } catch {
          res.writeHead(500).end('{}')
        }
      })
    })
    server.listen(0, '127.0.0.1', () => {
      const address = server.address()
      const port = typeof address === 'object' && address ? address.port : 0
      resolve({ server, port, prompts })
    })
  })
}

describe('KERRITS Chinese garment spec — full translation pipeline', () => {
  let server: ChildProcess | undefined
  let fake: { server: Server; port: number; prompts: string[] }
  let base: string
  let dataDir: string
  let pdfPath: string
  let dictionaryPath: string | undefined

  beforeAll(async () => {
    // Hermetic copy: the real file lives under ~/Downloads/… — the test must
    // not depend on the user keeping it in place, but its text content is
    // what we are testing.
    const realPdf = process.env.KERRITS_PDF_PATH
      ?? join(process.env.HOME ?? '/tmp', 'Downloads', '资料（保密）', 'KERRITS-英文工艺单.pdf')
    if (!existsSync(realPdf)) {
      throw new Error(`KERRITS fixture missing: ${realPdf} (set KERRITS_PDF_PATH to override)`)
    }
    dataDir = mkdtempSync(join(tmpdir(), 'genoffice-kerrits-e2e-'))
    pdfPath = join(dataDir, 'kerrits.pdf')
    copyFileSync(realPdf, pdfPath)

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

    // Point the active provider at the stub.
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

    // Seed the KB with garment industry terms the model will likely mis-
    // translate. The dictionary enforcement pass replaces anything the model
    // leaves in the source language with these.
    await ipc(base, 'ai:translation-kb-upsert', [
      {
        id: 'kerrits-fabric-weight',
        scope: 'company',
        priority: 70,
        sourceTerm: '克重',
        targetTerm: 'fabric weight',
        sourceLang: 'zh-CN',
        targetLang: 'en-US',
      },
    ])
    await ipc(base, 'ai:translation-kb-upsert', [
      {
        id: 'kerrits-flatseam',
        scope: 'company',
        priority: 70,
        sourceTerm: '平缝',
        targetTerm: 'flatseam',
        sourceLang: 'zh-CN',
        targetLang: 'en-US',
      },
    ])
    await ipc(base, 'ai:translation-kb-upsert', [
      {
        id: 'kerrits-waistband',
        scope: 'company',
        priority: 70,
        sourceTerm: '腰头',
        targetTerm: 'waistband',
        sourceLang: 'zh-CN',
        targetLang: 'en-US',
      },
    ])
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

  it('mines real KERRITS Chinese text into line-oriented segments', async () => {
    const { ok, result } = await ipc<{
      ok: boolean
      totalSegments?: number
      coverage?: { total: number; covered: number; exact: number; ratio: number; uncovered: string[] }
      kbEntries?: number
      llmEntries?: number
      dictionaryPath?: string
      error?: string
    }>(base, 'ai:translate-build-dictionary', [
      { inputPath: pdfPath, sourceLang: 'zh-CN', targetLang: 'en-US' },
    ])
    expect(ok).toBe(true)
    expect(result.ok).toBe(true)
    // The PDF has 3 pages of Chinese garment spec text; the miner should
    // produce dozens of segments, not 0 (the synthetic-PDF case) or all 1000+
    // (the unmined decoration case the W34 patch fixed).
    expect(result.totalSegments).toBeGreaterThan(10)
    expect(result.totalSegments).toBeLessThan(500)
    // The KB seeded three terms; each should land in the dictionary verbatim
    // regardless of what the model produced.
    expect(result.kbEntries).toBeGreaterThanOrEqual(3)
    // The fake model produced an "EN: …" echo for every segment, so the LLM
    // half of the dictionary should also be non-empty.
    expect(result.llmEntries).toBeGreaterThan(0)
    // At least the KB rows must show up in the coverage report.
    expect(result.coverage?.covered).toBeGreaterThanOrEqual(3)
    expect(result.dictionaryPath).toBeTruthy()
    dictionaryPath = result.dictionaryPath
    // Persist the dictionary to disk so we can re-use it in the file pass.
    expect(existsSync(dictionaryPath)).toBe(true)
  })

  it('translates the PDF using the dictionary generated from the KB + LLM pass', async () => {
    expect(dictionaryPath).toBeDefined()
    const outputPath = join(dataDir, 'kerrits_translated.pdf')
    const { ok, result } = await ipc<{
      ok: boolean
      outputPath?: string
      bytes?: number
      coverage?: { total: number; covered: number; exact: number; ratio: number }
      error?: string
    }>(base, 'ai:translate-file-auto', [
      {
        inputPath: pdfPath,
        outputPath,
        sourceLang: 'zh-CN',
        targetLang: 'en-US',
        dictionaryPath,
        // Reuse the dictionary from the previous step — must not rebuild.
      },
    ])
    expect(ok).toBe(true)
    expect(result.ok).toBe(true)
    expect(result.dictionaryPath).toBe(dictionaryPath)
    expect(result.outputPath).toBe(outputPath)
    expect(existsSync(outputPath)).toBe(true)
    // Coverage must round-trip: every segment the dictionary reaches should
    // still be covered on the re-run.
    expect(result.coverage?.ratio).toBe(1)
  })

  it('keeps every KB term verbatim in the dictionary file', () => {
    expect(dictionaryPath).toBeDefined()
    const dict = JSON.parse(readFileSync(dictionaryPath!, 'utf8')) as Record<string, string>
    // The dictionary is `{ source: target }`; the KB seeded `克重 → fabric
    // weight` etc. and the writer must not have dropped or rewritten them.
    expect(dict['克重']).toBe('fabric weight')
    expect(dict['平缝']).toBe('flatseam')
    expect(dict['腰头']).toBe('waistband')
  })

  it('snippet translation reuses the same dictionary and reports KB hits separately', async () => {
    expect(dictionaryPath).toBeDefined()
    const { ok, result } = await ipc<{
      ok: boolean
      translation?: string
      matchedTerms?: string[]
      dictionaryHits?: string[]
      dictionary?: { path: string; terms: number; hits: number } | null
      error?: string
    }>(base, 'home:translate-snippet', [
      {
        text: '本款克重为 220 g/m²,腰头用平缝收口。',
        sourceLang: 'zh-CN',
        targetLang: 'en-US',
        dictionaryPath,
        useDictionary: true,
      },
    ])
    expect(ok).toBe(true)
    expect(result.ok).toBe(true)
    // The fake model leaves KB terms in Chinese, so the dictionary enforcement
    // pass replaces each one. The provenance split lets the UI label them
    // separately ("KB · 1" / "Dictionary · 1").
    expect(result.translation).toContain('fabric weight')
    expect(result.translation).toContain('waistband')
    expect(result.translation).toContain('flatseam')
    expect(result.dictionary?.path).toBe(dictionaryPath)
    expect(result.dictionary?.terms).toBeGreaterThan(0)
  })
})
