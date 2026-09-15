/**
 * End-to-end marketplace flow test.
 *
 * Boots the full web-server bundle on a random port and drives the IPC
 * channels directly to verify the complete upload → install → pi loader
 * → uninstall chain works without any stubs.
 *
 * This is the regression test for the W34 marketplace fixes:
 *   1. skillMarketCatalog uses allMarketplaceSkills (curated + uploaded)
 *   2. market.install is idempotent (re-installing doesn't throw)
 *   3. home:uninstall-skill always calls market.uninstall
 *   4. home:install-skill always writes SKILL.md (not just for uploads)
 *   5. SKILL.md satisfies pi's strict frontmatter rules
 */
import { afterAll, beforeAll, describe, expect, it } from 'vitest'
import { spawn, type ChildProcess } from 'node:child_process'
import { mkdtempSync, readFileSync, rmSync, existsSync } from 'node:fs'
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

describe('marketplace E2E flow', () => {
  let server: ChildProcess | undefined
  let base: string
  let dataDir: string
  let port: number

  beforeAll(async () => {
    dataDir = mkdtempSync(join(tmpdir(), 'genoffice-e2e-'))
    port = 18000 + Math.floor(Math.random() * 1000)
    base = `http://127.0.0.1:${port}`
    const bundle = join(__dirname, '..', 'dist', 'bundle', 'index.js')
    server = spawn(process.execPath, [bundle], {
      env: {
        ...process.env,
        PORT: String(port),
        HOST: '127.0.0.1',
        GENOFFICE_DATA_DIR: dataDir,
      },
      stdio: 'pipe',
    })
    server.stderr?.on('data', () => {})
    server.stdout?.on('data', () => {})
    await waitForHealth(base)
  }, 60_000)

  afterAll(() => {
    if (server) {
      server.kill('SIGKILL')
    }
    try {
      rmSync(dataDir, { recursive: true, force: true })
    } catch {
      /* ignore */
    }
  })

  it('runs the full upload → install → pi-loader → uninstall flow', async () => {
    const id = `e2e-test-${Date.now()}`

    // 1. Upload
    const upload = await ipc<{ ok: boolean; reviewStatus: string }>(
      base,
      'home:marketplace-upload',
      [
        {
          kind: 'skill',
          payload: {
            id,
            name: 'E2E Test Skill',
            description: 'Regression test for marketplace upload → install → pi-loader flow',
            author: 'E2E',
            version: '1.0.0',
            tools: ['e2e_tool'],
            scopes: ['files:read'],
            category: 'dev',
            tags: ['e2e'],
          },
        },
      ],
    )
    expect(upload.ok).toBe(true)
    expect(upload.result.ok).toBe(true)
    expect(upload.result.reviewStatus).toBe('pending')

    // 2. Search
    const search = await ipc<{ skills: { id: string }[]; plugins: unknown[]; total: number }>(
      base,
      'home:marketplace-search',
      [{ q: id }],
    )
    expect(search.ok).toBe(true)
    expect(search.result.skills.some((s) => s.id === id)).toBe(true)

    // 3. Install
    const install = await ipc<{ ok: boolean; piInstalled: boolean; installed: { id: string } }>(
      base,
      'home:install-skill',
      [{ id }],
    )
    expect(install.ok).toBe(true)
    expect(install.result.ok).toBe(true)
    expect(install.result.piInstalled).toBe(true)
    expect(install.result.installed.id).toBe(id)

    // 4. Verify SKILL.md on disk
    const skillPath = join(dataDir, 'pi-skills', id, 'SKILL.md')
    expect(existsSync(skillPath)).toBe(true)
    const content = readFileSync(skillPath, 'utf-8')
    expect(content).toMatch(new RegExp(`^name: ${id}$`, 'm')) // pi slug in frontmatter
    expect(content).toMatch(/^description: /m) // required by pi
    expect(content).toMatch(/^display_name: E2E Test Skill$/m) // human name

    // 5. Verify pi loader actually sees it
    const piSkills = await ipc<{
      piSkills: { name: string; filePath: string }[]
      diagnostics: { type: string; path: string }[]
      records: { name: string }[]
    }>(base, 'home:list-pi-skills', [])
    expect(piSkills.ok).toBe(true)
    expect(piSkills.result.piSkills.some((s) => s.name === id)).toBe(true)
    expect(
      piSkills.result.diagnostics.some((d) => d.path.includes(id)),
    ).toBe(false)
    expect(piSkills.result.records.some((r) => r.name === id)).toBe(true)

    // 6. Uninstall
    const uninstall = await ipc<{ ok: boolean }>(base, 'home:uninstall-skill', [{ id }])
    expect(uninstall.ok).toBe(true)
    expect(uninstall.result.ok).toBe(true)

    // 7. Verify cleanup
    expect(existsSync(skillPath)).toBe(false)
    const indexPath = join(dataDir, 'pi-skills', '.index.json')
    if (existsSync(indexPath)) {
      const indexContent = readFileSync(indexPath, 'utf-8')
      expect(indexContent).not.toContain(`"${id}"`)
    }
    const afterUninstall = await ipc<{ piSkills: { name: string }[] }>(
      base,
      'home:list-pi-skills',
      [],
    )
    expect(afterUninstall.result.piSkills.some((s) => s.name === id)).toBe(false)
  }, 60_000)
})

import { describe, it, expect } from 'vitest'

describe('marketplace search sort regression (W35+)', () => {
  // These tests pin the documented sort semantics. The previous implementation
  // sorted 'newest' by entry.name — a real bug that meant newly uploaded
  // extensions could never appear at the top of the newest tab. The fix
  // exposes uploadedAt on every entry and sorts by it descending, with
  // download count as tiebreaker.
  it("'newest' sort orders by uploadedAt descending, not by name", async () => {
    const { searchMarketplace } = await import('../src/shell/skills')
    const result = searchMarketplace({ sort: 'newest' })
    // skills and plugins are sorted as independent lists — verify each one.
    for (const list of [result.skills, result.plugins]) {
      expect(list.length).toBeGreaterThan(0)
      for (let i = 1; i < list.length; i++) {
        const prev = Date.parse(list[i - 1].uploadedAt ?? '') || 0
        const cur = Date.parse(list[i].uploadedAt ?? '') || 0
        expect(prev).toBeGreaterThanOrEqual(cur)
      }
    }
  })

  it("'popular' sort orders by downloads descending", async () => {
    const { searchMarketplace } = await import('../src/shell/skills')
    const result = searchMarketplace({ sort: 'popular' })
    for (const list of [result.skills, result.plugins]) {
      for (let i = 1; i < list.length; i++) {
        expect(list[i - 1].downloads).toBeGreaterThanOrEqual(list[i].downloads)
      }
    }
  })

  it("'rating' sort orders by rating descending with downloads tiebreaker", async () => {
    const { searchMarketplace } = await import('../src/shell/skills')
    const result = searchMarketplace({ sort: 'rating' })
    for (const list of [result.skills, result.plugins]) {
      for (let i = 1; i < list.length; i++) {
        const prev = list[i - 1]
        const cur = list[i]
        if (prev.rating === cur.rating) {
          expect(prev.downloads).toBeGreaterThanOrEqual(cur.downloads)
        } else {
          expect(prev.rating).toBeGreaterThanOrEqual(cur.rating)
        }
      }
    }
  })

  it("'name' sort orders by display name ascending", async () => {
    const { searchMarketplace } = await import('../src/shell/skills')
    const result = searchMarketplace({ sort: 'name' })
    for (const list of [result.skills, result.plugins]) {
      for (let i = 1; i < list.length; i++) {
        expect(list[i - 1].name.localeCompare(list[i].name)).toBeLessThanOrEqual(0)
      }
    }
  })

  it('every uploaded entry exposes uploadedAt as a valid ISO timestamp', async () => {
    const { searchMarketplace } = await import('../src/shell/skills')
    const result = searchMarketplace({})
    const all = [...result.skills, ...result.plugins]
    // When uploadedAt is present (community uploads) it must parse as a
    // valid ISO timestamp. Curated entries are allowed to omit it; the
    // 'newest' sort treats those as epoch 0 and sinks them to the bottom.
    for (const e of all) {
      if (typeof e.uploadedAt === 'string') {
        const t = Date.parse(e.uploadedAt)
        expect(Number.isFinite(t)).toBe(true)
      }
    }
  })
})
