/**
 * Translate-suite E2E — the LumosAI translate skills ported into GenOffice.
 *
 * Verifies the two integration surfaces end-to-end against the real
 * web-server bundle:
 *
 *   1. The marketplace exposes the six translate-* entries (translate,
 *      translate-config, translate-docx, translate-pdf, translate-ppt,
 *      translate-xls) under the `translation` category, and search finds
 *      them by tag / tool name.
 *
 *   2. The translation knowledge base IPC round-trips an entry through
 *      upsert → list → resolve → remove. This is the same data the
 *      system prompt gets injected with on every `ai:translate` call.
 */
import { afterAll, beforeAll, describe, expect, it } from 'vitest'
import { spawn, type ChildProcess } from 'node:child_process'
import { mkdtempSync, rmSync } from 'node:fs'
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

const TRANSLATE_IDS = [
  'translate',
  'translate-config',
  'translate-docx',
  'translate-pdf',
  'translate-ppt',
  'translate-xls',
] as const

describe('translate suite E2E', () => {
  let server: ChildProcess | undefined
  let base: string
  let dataDir: string
  let port: number

  beforeAll(async () => {
    dataDir = mkdtempSync(join(tmpdir(), 'genoffice-translate-e2e-'))
    port = 21000 + Math.floor(Math.random() * 8000)
    base = `http://127.0.0.1:${port}`
    const bundle = join(__dirname, '..', 'dist', 'bundle', 'index.js')
    server = spawn(process.execPath, [bundle], {
      env: {
        ...process.env,
        PORT: String(port),
        DATA_DIR: dataDir,
        GENOFFICE_WEB_DATA_DIR: dataDir,
        // Point the translation KB at the scratch dir so the test never
        // touches the user's real ~/.genoffice/translation-kb.json.
        GENOFFICE_TRANSLATION_KB: join(dataDir, 'translation-kb.json'),
        NO_OPEN: '1',
      },
      stdio: ['ignore', 'pipe', 'pipe'],
    })
    server.stdout?.on('data', () => {})
    server.stderr?.on('data', () => {})
    await waitForHealth(base)
  }, 60_000)

  afterAll(() => {
    server?.kill('SIGTERM')
    try {
      rmSync(dataDir, { recursive: true, force: true })
    } catch {
      /* ignore */
    }
  })

  it('exposes the six translate-* marketplace entries', async () => {
    const { ok, result } = await ipc<{
      skills?: Array<{ id: string; category: string; tools: string[]; tags: string[] }>
    }>(base, 'home:marketplace-search', [{ category: 'translation' }])
    expect(ok).toBe(true)
    const skills = result.skills ?? []
    const ids = new Set(skills.map((e) => e.id))
    for (const id of TRANSLATE_IDS) {
      expect(ids.has(id), `missing marketplace entry for ${id}`).toBe(true)
    }
    const translate = skills.find((e) => e.id === 'translate')
    expect(translate?.tools).toContain('translate_file')
    expect(translate?.tags).toContain('unified-entry')
  })

  it('searches the translate suite by tag and by tool name', async () => {
    const byTag = await ipc<{ skills?: Array<{ id: string }> }>(base, 'home:marketplace-search', [
      { q: 'soffice' },
    ])
    expect(byTag.ok).toBe(true)
    expect((byTag.result.skills ?? []).map((e) => e.id)).toContain('translate-xls')

    const byTool = await ipc<{ skills?: Array<{ id: string }> }>(base, 'home:marketplace-search', [
      { q: 'translate_pptx' },
    ])
    expect(byTool.ok).toBe(true)
    expect((byTool.result.skills ?? []).map((e) => e.id)).toContain('translate-ppt')
  })

  it('round-trips a knowledge base entry through the IPC handlers', async () => {
    const id = `e2e-term-${Date.now().toString(36)}`
    const upsert = await ipc<{ ok: boolean }>(base, 'ai:translation-kb-upsert', [
      {
        id,
        scope: 'company',
        priority: 3,
        sourceTerm: '克重',
        targetTerm: 'GSM',
        sourceLang: 'zh-CN',
        targetLang: 'en-US',
      },
    ])
    expect(upsert.ok).toBe(true)
    expect(upsert.result.ok).toBe(true)

    const list = await ipc<{ ok: boolean; entries: Array<{ id: string; sourceTerm?: string }> }>(
      base,
      'ai:translation-kb-list',
      [{ schema: 'term' }],
    )
    expect(list.ok).toBe(true)
    const found = list.result.entries.find((e) => e.id === id)
    expect(found?.sourceTerm).toBe('克重')

    const resolved = await ipc<{ ok: boolean; terms: Array<{ sourceTerm: string }>; promptBlock: string }>(
      base,
      'ai:translation-kb-resolve',
      [{ sourceLang: 'zh-CN', targetLang: 'en-US' }],
    )
    expect(resolved.ok).toBe(true)
    expect(resolved.result.terms.map((t) => t.sourceTerm)).toContain('克重')
    expect(resolved.result.promptBlock).toContain('GSM')

    const removed = await ipc<{ ok: boolean; removed: boolean }>(base, 'ai:translation-kb-remove', [id])
    expect(removed.ok).toBe(true)
    expect(removed.result.removed).toBe(true)

    const after = await ipc<{ ok: boolean; entries: Array<{ id: string }> }>(base, 'ai:translation-kb-list', [
      { schema: 'term' },
    ])
    expect(after.result.entries.find((e) => e.id === id)).toBeUndefined()
  })

  it('reports KB stats grouped by schema', async () => {
    const stats = await ipc<{ ok: boolean; total: number; bySchema: Record<string, number> }>(
      base,
      'ai:translation-kb-stats',
      [],
    )
    expect(stats.ok).toBe(true)
    expect(typeof stats.result.total).toBe('number')
    expect(stats.result.bySchema).toBeTypeOf('object')
  })
})
