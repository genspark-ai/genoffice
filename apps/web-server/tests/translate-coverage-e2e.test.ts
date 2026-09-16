/**
 * Coverage + gap-filling E2E — the "is this dictionary actually enough?" pass.
 *
 * The upstream format handlers are pure dictionary rewriters: they swap the
 * terms they are given and leave everything else in the source language, and
 * the untranslated report they *do* print is tuned for zh-CN targets, so a
 * zh -> en pass reports "0 untranslated" while translating nothing. The pane
 * therefore cannot trust the handler's stdout and has to compute coverage
 * itself.
 *
 * This exercises the real web-server bundle for the three pieces that make up
 * the loop the Settings -> 翻译知识库 pane drives:
 *
 *   1. `ai:translate-build-dictionary` reports how much of the file its
 *      dictionary reaches, including partly-matched (mixed-language) segments.
 *   2. `ai:translate-fill-gaps` sends only the gaps to the model and writes an
 *      extended dictionary next to the original, which stays untouched.
 *   3. `ai:translate-file-auto` can re-run a pass from a dictionary alone —
 *      no provider configured, no rebuild, no lost gap-filling work.
 *
 * The model is a local OpenAI-compatible stub and the file handler is a shell
 * stub, so the test stays hermetic: what is under test is GenOffice's wiring,
 * not `python-docx`.
 */
import { afterAll, beforeAll, describe, expect, it } from 'vitest'
import { spawn, type ChildProcess } from 'node:child_process'
import { createServer, type Server } from 'node:http'
import { existsSync, mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { tmpdir } from 'node:os'
import JSZip from 'jszip'

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
 * A minimal but valid DOCX: one paragraph per line, plus a one-row table.
 *
 * The table matters — `@genoffice/file-parse` joins its cells with ` | `, and
 * the handler rewrites cells individually, so mining the joined row would
 * produce keys that can never match.
 */
async function writeDocx(path: string, paragraphs: string[], tableRow: string[]): Promise<void> {
  const cell = (text: string) => `<w:tc><w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p></w:tc>`
  const body =
    paragraphs.map((text) => `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`).join('') +
    `<w:tbl><w:tr>${tableRow.map(cell).join('')}</w:tr></w:tbl>`
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
      '</Types>',
  )
  zip.file(
    '_rels/.rels',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
      '</Relationships>',
  )
  zip.file(
    'word/document.xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}</w:body></w:document>`,
  )
  writeFileSync(path, await zip.generateAsync({ type: 'nodebuffer' }))
}

/**
 * Stand in for `~/.lumos/bundled-skills/<hash>/translate/scripts/translate.py`.
 *
 * `resolvePython` accepts any existing executable, and the handler is spawned as
 * `<python> <script> <input> <output> [--dictionary <path>]`, so a POSIX shell
 * script driven by `/bin/sh` is a drop-in handler: it copies the input to the
 * output, which is all the file-pass assertions need.
 */
function writeStubTranslateSkill(dir: string): void {
  mkdirSync(join(dir, 'scripts'), { recursive: true })
  writeFileSync(
    join(dir, 'scripts', 'translate.py'),
    ['#!/bin/sh', '# stub handler: copy input -> output', 'cp "$1" "$2"', ''].join('\n'),
    { mode: 0o755 },
  )
}

/** OpenAI-compatible stub: echoes the source so a translated term is observable. */
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
            JSON.stringify({ choices: [{ message: { role: 'assistant', content: `EN: ${source}` } }] }),
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

const PARAGRAPHS = [
  '产品验收报告已提交。',
  '请在一周内完成复核。',
  '本批次货物共 120 件。',
  '供应商需提供质检报告。',
  '付款条件为 30 天账期。',
]
const TABLE_ROW = ['物料', '单价']
/** What the dictionary has to reach: paragraphs and table cells, one by one. */
const SEGMENTS = [...PARAGRAPHS, ...TABLE_ROW]

describe('dictionary coverage + gap filling E2E', () => {
  let server: ChildProcess | undefined
  let fake: { server: Server; port: number; prompts: string[] }
  let base: string
  let dataDir: string
  let skillDir: string
  let docxPath: string
  let builtDictionaryPath: string | undefined
  let completeDictionaryPath: string | undefined

  beforeAll(async () => {
    dataDir = mkdtempSync(join(tmpdir(), 'genoffice-translate-coverage-e2e-'))
    skillDir = join(dataDir, 'skills')
    writeStubTranslateSkill(skillDir)
    docxPath = join(dataDir, 'supplier-notes.docx')
    await writeDocx(docxPath, PARAGRAPHS, TABLE_ROW)

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
        GENOFFICE_TRANSLATE_SKILLS_DIR: skillDir,
        GENOFFICE_PYTHON: '/bin/sh',
        NO_OPEN: '1',
      },
      stdio: ['ignore', 'pipe', 'pipe'],
    })
    server.stdout?.on('data', () => {})
    server.stderr?.on('data', () => {})
    await waitForHealth(base)

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

    // One KB term that only *partly* covers a paragraph — the case the handlers
    // turn into mixed-language output ("Product Acceptance Report已提交。").
    await ipc(base, 'ai:translation-kb-upsert', [
      {
        id: 'e2e-acceptance-report',
        scope: 'company',
        priority: 50,
        sourceTerm: '产品验收报告',
        targetTerm: 'Product Acceptance Report',
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

  it('reports coverage, including the segments a KB term only partly covers', async () => {
    const { ok, result } = await ipc<{
      ok: boolean
      totalSegments?: number
      kbEntries?: number
      dictionaryPath?: string
      coverage?: {
        total: number
        covered: number
        exact: number
        partial: string[]
        uncovered: string[]
        ratio: number
      }
    }>(base, 'ai:translate-build-dictionary', [
      {
        inputPath: docxPath,
        sourceLang: 'zh-CN',
        targetLang: 'en-US',
        useLlm: false,
      },
    ])
    expect(ok).toBe(true)
    expect(result.ok).toBe(true)
    // The table row is mined as two cells, not as the `物料 | 单价` line the
    // extractor prints.
    expect(result.totalSegments).toBe(SEGMENTS.length)
    expect(result.kbEntries).toBe(1)
    const coverage = result.coverage!
    expect(coverage.total).toBe(SEGMENTS.length)
    // A KB-only build translates nothing whole: the one paragraph it touches is
    // a substring hit, which the UI has to show as "partly matched" rather than
    // count as done.
    expect(coverage.exact).toBe(0)
    expect(coverage.partial).toEqual([PARAGRAPHS[0]])
    expect(coverage.uncovered).toEqual([...PARAGRAPHS.slice(1), ...TABLE_ROW])
    expect(coverage.covered).toBe(1)
    expect(coverage.ratio).toBeCloseTo(1 / SEGMENTS.length, 5)

    // The provider is configured but a KB-only build must not touch it.
    expect(fake.prompts.length).toBe(0)

    builtDictionaryPath = result.dictionaryPath
    expect(builtDictionaryPath).toBeTruthy()
  })

  it('fills the gaps from the dictionary alone and writes an extended copy', async () => {
    const { ok, result } = await ipc<{
      ok: boolean
      dictionaryPath?: string
      added?: number
      stillUncovered?: string[]
      coverageBefore?: { total: number; covered: number; partial: string[]; uncovered: string[] }
      coverageAfter?: { total: number; covered: number; exact: number; partial: string[]; uncovered: string[]; ratio: number }
    }>(base, 'ai:translate-fill-gaps', [
      {
        inputPath: docxPath,
        sourceLang: 'zh-CN',
        targetLang: 'en-US',
        dictionaryPath: builtDictionaryPath,
      },
    ])
    expect(ok).toBe(true)
    expect(result.ok).toBe(true)
    completeDictionaryPath = result.dictionaryPath
    expect(completeDictionaryPath).toBeTruthy()
    // Only the gaps are sent: one partly-matched paragraph plus four untouched.
    expect(result.added).toBe(SEGMENTS.length)
    expect(result.coverageBefore?.covered).toBe(1)
    expect(result.coverageAfter?.exact).toBe(SEGMENTS.length)
    expect(result.coverageAfter?.ratio).toBe(1)
    expect(result.stillUncovered).toEqual([])

    const extended = JSON.parse(readFileSync(completeDictionaryPath!, 'utf8')) as Record<string, string>
    // The partly-matched paragraph becomes an *exact* key so the handler's
    // exact-match branch wins over the substring one and the mixed-language
    // output disappears.
    expect(extended[PARAGRAPHS[0]]).toContain('EN: ')
    expect(extended['产品验收报告']).toBe('Product Acceptance Report')
    for (const segment of SEGMENTS) expect(extended[segment]).toBeTruthy()

    // The dictionary the pass started from is left alone so the two can be diffed.
    const original = JSON.parse(readFileSync(builtDictionaryPath!, 'utf8')) as Record<string, string>
    expect(Object.keys(original)).toEqual(['产品验收报告'])
  })

  it('re-runs a file pass from a dictionary with no provider configured', async () => {
    const outputPath = join(dataDir, 'supplier-notes_translated.docx')
    const { ok, result } = await ipc<{
      ok: boolean
      outputPath?: string
      dictionaryReused?: boolean
      dictionaryPath?: string
      coverage?: { total: number; ratio: number; uncovered: string[] }
      error?: string
    }>(base, 'ai:translate-file-auto', [
      {
        inputPath: docxPath,
        outputPath,
        sourceLang: 'zh-CN',
        targetLang: 'en-US',
        dictionaryPath: completeDictionaryPath,
        // Deliberately unusable: re-running with an existing dictionary is a
        // pure rewrite and must not depend on a model being reachable.
        settings: { provider: 'not-configured', providers: {} },
      },
    ])
    expect(result.error).toBeUndefined()
    expect(ok).toBe(true)
    expect(result.ok).toBe(true)
    expect(result.dictionaryReused).toBe(true)
    expect(result.dictionaryPath).toBe(completeDictionaryPath)
    expect(result.coverage?.total).toBe(SEGMENTS.length)
    expect(result.coverage?.ratio).toBe(1)
    expect(result.coverage?.uncovered).toEqual([])
    expect(result.outputPath).toBe(outputPath)
    expect(existsSync(outputPath)).toBe(true)
  })

  it('fails the re-run when the dictionary cannot be read', async () => {
    const { result } = await ipc<{ ok: boolean; stage?: string; error?: string }>(
      base,
      'ai:translate-file-auto',
      [
        {
          inputPath: docxPath,
          sourceLang: 'zh-CN',
          targetLang: 'en-US',
          dictionaryPath: join(dataDir, 'does-not-exist.json'),
        },
      ],
    )
    expect(result.ok).toBe(false)
    expect(result.stage).toBe('dictionary')
    expect(result.error).toContain('does-not-exist.json')
  })
})
