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
    // Pick a port well away from the dev default (18081) so a locally running
    // web-server never collides with the one this suite spawns. Range
    // 20000-28999 is unassigned for the usual dev tooling on macOS/Linux.
    port = 20000 + Math.floor(Math.random() * 9000)
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
    expect(content).toMatch(/^# E2E Test Skill$/m) // human name lives in the H1 per pi.dev

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

  // -------------------------------------------------------------------------
  // Extensions are pi resources — a plugin install must be a pi package
  // install, and a skill install must be visible to pi's loader. These tests
  // drive the real IPC surface of the spawned server and then read the files
  // and settings pi would read, so a regression in the wiring fails here
  // instead of silently shipping a UI-only "install".
  // -------------------------------------------------------------------------

  /** A minimal but real pi extension module. */
  const EXTENSION_SOURCE = [
    'export default function (pi) {',
    '  pi.registerTool({',
    "    name: 'e2e_plugin_tool',",
    "    label: 'E2E Plugin Tool',",
    "    description: 'Registered by the e2e plugin package',",
    '    parameters: {},',
    '    async execute() {',
    "      return { content: [{ type: 'text', text: 'e2e' }], details: {} }",
    '    },',
    '  })',
    '}',
    '',
  ].join('\n')

  function publishPlugin(id: string, withArtifact: boolean) {
    return ipc<{ ok: boolean; artifact?: { kind: string; filename: string } | null; error?: string }>(
      base,
      'home:marketplace-upload',
      [
        {
          kind: 'plugin',
          payload: {
            id,
            name: 'E2E Pi Plugin',
            description: 'End-to-end test of the pi package install path',
            author: 'E2E',
            version: '1.0.0',
            tools: ['e2e_plugin_tool'],
            scopes: ['files:read'],
            requirements: [],
            category: 'dev',
            tags: ['e2e', 'pi'],
            ...(withArtifact
              ? { artifact: { filename: 'index.ts', content: EXTENSION_SOURCE } }
              : {}),
          },
        },
      ],
    )
  }

  it('installs a plugin as a pi package that pi itself resolves', async () => {
    const id = `e2e-plugin-${Date.now()}`
    const upload = await publishPlugin(id, true)
    expect(upload.result.ok).toBe(true)
    expect(upload.result.artifact?.kind).toBe('extension')

    const install = await ipc<{
      ok: boolean
      pi: { mode: string; packageDir: string; hasCode: boolean; extensions: string[] }
    }>(base, 'home:install-plugin', [{ id }])
    expect(install.result.ok).toBe(true)
    expect(install.result.pi.mode).toBe('local-package')
    expect(install.result.pi.hasCode).toBe(true)

    const dir = join(dataDir, 'pi-plugins', id)
    const extPath = join(dir, 'extensions', 'index.ts')
    const skillPath = join(dir, 'skills', id, 'SKILL.md')
    expect(existsSync(join(dir, 'package.json'))).toBe(true)
    expect(existsSync(extPath)).toBe(true)
    expect(existsSync(skillPath)).toBe(true)
    expect(readFileSync(extPath, 'utf-8')).toContain('registerTool')

    // pi's settings carry the package source — this is what makes the install
    // survive a restart and be visible to any pi session using this agent dir.
    const settings = JSON.parse(readFileSync(join(dataDir, 'pi-agent', 'settings.json'), 'utf-8')) as {
      packages?: unknown[]
    }
    expect(settings.packages).toContain(dir)

    const report = await ipc<{
      extensions: { path: string; enabled: boolean; managed: boolean }[]
      skills: { path: string; enabled: boolean }[]
      packages: { source: string }[]
    }>(base, 'home:list-pi-resources', [])
    expect(report.result.extensions.some((e) => e.path === extPath && e.enabled && e.managed)).toBe(true)
    expect(report.result.skills.some((s) => s.path === skillPath && s.enabled)).toBe(true)
    expect(report.result.packages.some((p) => p.source === dir)).toBe(true)

    // Disabling must change what pi resolves, not just a UI flag.
    const off = await ipc<{ ok: boolean }>(base, 'home:toggle-plugin', [{ id, enabled: false }])
    expect(off.result.ok).toBe(true)
    const disabled = await ipc<{ extensions: { path: string }[]; skills: { path: string }[] }>(
      base,
      'home:list-pi-resources',
      [],
    )
    expect(disabled.result.extensions.some((e) => e.path === extPath)).toBe(false)
    expect(disabled.result.skills.some((s) => s.path === skillPath)).toBe(false)

    const on = await ipc<{ ok: boolean }>(base, 'home:toggle-plugin', [{ id, enabled: true }])
    expect(on.result.ok).toBe(true)
    const reenabled = await ipc<{ extensions: { path: string }[] }>(base, 'home:list-pi-resources', [])
    expect(reenabled.result.extensions.some((e) => e.path === extPath)).toBe(true)

    const removed = await ipc<{ ok: boolean; piRemoved: boolean }>(
      base,
      'home:uninstall-plugin',
      [{ id }],
    )
    expect(removed.result.ok).toBe(true)
    expect(existsSync(dir)).toBe(false)
    const after = JSON.parse(readFileSync(join(dataDir, 'pi-agent', 'settings.json'), 'utf-8')) as {
      packages?: string[]
    }
    expect(after.packages ?? []).not.toContain(dir)
  }, 60_000)

  it('installs a plugin without an artifact as guidance only (no fake tools)', async () => {
    const id = `e2e-plugin-solo-${Date.now()}`
    const upload = await publishPlugin(id, false)
    expect(upload.result.ok).toBe(true)
    expect(upload.result.artifact ?? null).toBeNull()

    const install = await ipc<{ ok: boolean; pi: { hasCode: boolean; extensions: string[] } }>(
      base,
      'home:install-plugin',
      [{ id }],
    )
    expect(install.result.ok).toBe(true)
    expect(install.result.pi.hasCode).toBe(false)
    expect(install.result.pi.extensions).toEqual([])

    const dir = join(dataDir, 'pi-plugins', id)
    expect(existsSync(join(dir, 'extensions'))).toBe(false)
    const body = readFileSync(join(dir, 'skills', id, 'SKILL.md'), 'utf-8')
    // The generated guidance must say the tools are not registered.
    expect(body).toContain('NOT registered')
    await ipc(base, 'home:uninstall-plugin', [{ id }])
  }, 60_000)

  it('rejects an uploaded module that would register nothing', async () => {
    const id = `e2e-plugin-bad-${Date.now()}`
    const upload = await ipc<{ ok: boolean; error?: string }>(base, 'home:marketplace-upload', [
      {
        kind: 'plugin',
        payload: {
          id,
          name: 'Broken Plugin',
          description: 'A plugin module with no default export at all',
          author: 'E2E',
          version: '1.0.0',
          tools: ['nope'],
          scopes: [],
          requirements: [],
          category: 'dev',
          tags: ['e2e'],
          artifact: { filename: 'index.ts', content: 'export const nothing = 1\n' },
        },
      },
    ])
    expect(upload.result.ok).toBe(false)
    expect(upload.result.error).toMatch(/export default/)
  })

  it('rejects a SKILL.md whose frontmatter name does not match the entry id', async () => {
    const id = `e2e-skill-bad-${Date.now()}`
    const upload = await ipc<{ ok: boolean; error?: string }>(base, 'home:marketplace-upload', [
      {
        kind: 'skill',
        payload: {
          id,
          name: 'Mismatched Skill',
          description: 'Frontmatter name does not match the marketplace id',
          author: 'E2E',
          version: '1.0.0',
          tools: ['some_tool'],
          scopes: [],
          category: 'dev',
          tags: ['e2e'],
          artifact: {
            filename: 'SKILL.md',
            content: '---\nname: something-else\ndescription: Mismatched frontmatter name\n---\n\nBody.\n',
          },
        },
      },
    ])
    expect(upload.result.ok).toBe(false)
    expect(upload.result.error).toMatch(/must be/)
  })

  it('publishes a real SKILL.md and installs that exact file', async () => {
    const id = `e2e-skill-md-${Date.now()}`
    const body = `---\nname: ${id}\ndescription: A publisher-authored skill used by the e2e suite\n---\n\nAlways answer in one sentence.\n`
    const upload = await ipc<{ ok: boolean; artifact?: { kind: string } | null }>(
      base,
      'home:marketplace-upload',
      [
        {
          kind: 'skill',
          payload: {
            id,
            name: 'Publisher Skill',
            description: 'A publisher-authored skill used by the e2e suite',
            author: 'E2E',
            version: '1.0.0',
            tools: ['none'],
            scopes: [],
            category: 'dev',
            tags: ['e2e'],
            artifact: { filename: 'SKILL.md', content: body },
          },
        },
      ],
    )
    expect(upload.result.ok).toBe(true)
    expect(upload.result.artifact?.kind).toBe('skill-md')

    const install = await ipc<{ ok: boolean; piInstalled: boolean }>(base, 'home:install-skill', [
      { id },
    ])
    expect(install.result.ok).toBe(true)
    expect(install.result.piInstalled).toBe(true)

    const skillPath = join(dataDir, 'pi-skills', id, 'SKILL.md')
    // The uploaded body wins over any body the server could synthesize.
    expect(readFileSync(skillPath, 'utf-8')).toBe(body)

    // pi's settings must point at the skills dir so a pi session loads it.
    const settings = JSON.parse(readFileSync(join(dataDir, 'pi-agent', 'settings.json'), 'utf-8')) as {
      skills?: string[]
    }
    expect(settings.skills).toContain(join(dataDir, 'pi-skills'))

    // Toggling the skill off must move it out of pi's discovery path.
    const off = await ipc<{ ok: boolean; piMoved: boolean; piPath: string }>(
      base,
      'home:toggle-skill',
      [{ id, enabled: false }],
    )
    expect(off.result.ok).toBe(true)
    expect(off.result.piMoved).toBe(true)
    expect(existsSync(skillPath)).toBe(false)
    expect(existsSync(join(dataDir, 'pi-skills-disabled', id, 'SKILL.md'))).toBe(true)

    const listed = await ipc<{ piSkills: { name: string; enabled: boolean }[] }>(
      base,
      'home:list-pi-skills',
      [],
    )
    expect(listed.result.piSkills.find((s) => s.name === id)?.enabled).toBe(false)

    const on = await ipc<{ ok: boolean; piMoved: boolean }>(base, 'home:toggle-skill', [
      { id, enabled: true },
    ])
    expect(on.result.piMoved).toBe(true)
    expect(existsSync(skillPath)).toBe(true)

    await ipc(base, 'home:uninstall-skill', [{ id }])
    expect(existsSync(skillPath)).toBe(false)
  }, 60_000)

  it('unpublishing removes the catalog entry, its artifact and the install', async () => {
    const id = `e2e-unpublish-${Date.now()}`
    const upload = await publishPlugin(id, true)
    expect(upload.result.ok).toBe(true)
    await ipc(base, 'home:install-plugin', [{ id }])
    expect(existsSync(join(dataDir, 'pi-plugins', id))).toBe(true)

    const search = await ipc<{ skills: { id: string }[]; plugins: { id: string }[] }>(
      base,
      'home:marketplace-search',
      [{ q: id }],
    )
    expect(search.result.plugins.some((p) => p.id === id)).toBe(true)

    const del = await ipc<{ ok: boolean; uninstalled: boolean }>(
      base,
      'home:marketplace-delete-upload',
      [{ kind: 'plugin', id }],
    )
    expect(del.result.ok).toBe(true)
    expect(del.result.uninstalled).toBe(true)
    expect(existsSync(join(dataDir, 'pi-plugins', id))).toBe(false)

    const gone = await ipc<{ plugins: { id: string }[] }>(base, 'home:marketplace-search', [{ q: id }])
    expect(gone.result.plugins.some((p) => p.id === id)).toBe(false)
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

describe('marketplace search ranking (W35+)', () => {
  // The old search was a single substring test over a joined haystack.
  // The W35+ rework tokenizes on whitespace with AND semantics and scores
  // hits by field weight (id > name > tags > description > author), so a
  // multi-word query narrows instead of widening and an id/name hit beats
  // a body-only mention.
  it('multi-token query uses AND semantics (every token must match)', async () => {
    const { searchMarketplace } = await import('../src/shell/skills')
    // "pdf ocr" appears in pdf-ocr-pro's id; both tokens are present there.
    const both = searchMarketplace({ q: 'pdf ocr' })
    const hitsBoth = [...both.skills, ...both.plugins]
    expect(hitsBoth.length).toBeGreaterThan(0)
    expect(hitsBoth.some((e) => e.id === 'pdf-ocr-pro')).toBe(true)

    // A query whose second token appears in no entry must return nothing,
    // proving the tokens are AND-ed rather than OR-ed.
    const impossible = searchMarketplace({ q: 'ocr zzzz-no-such-token' })
    expect(impossible.total).toBe(0)
  })

  it('an id/name hit outranks a body-only mention', async () => {
    const { searchMarketplace } = await import('../src/shell/skills')
    const res = searchMarketplace({ q: 'ocr' })
    const hits = [...res.skills, ...res.plugins]
    expect(hits.length).toBeGreaterThan(0)
    // The OCR extension itself must be first; any extension that merely
    // mentions "ocr" in its description sorts below it.
    expect(hits[0].id).toBe('pdf-ocr-pro')
  })

  it('relevance leads under the default sort, so name hits surface first', async () => {
    const { searchMarketplace } = await import('../src/shell/skills')
    // 'popular' is the UI default → relevance leads when a query is present.
    const res = searchMarketplace({ q: 'sync' })
    const hits = [...res.skills, ...res.plugins]
    expect(hits.length).toBeGreaterThan(0)
    // Both notion-sync and linear-sync match by name/id; a description-only
    // mention of "sync" must sort below them.
    const topIds = hits.slice(0, 2).map((e) => e.id)
    expect(topIds.some((id) => /sync/i.test(id))).toBe(true)
  })

  it('an explicit sort is respected verbatim even with a query', async () => {
    const { searchMarketplace } = await import('../src/shell/skills')
    // The user explicitly asked for rating order — the control must win.
    // skills and plugins are sorted as independent lists, so check each.
    const byRating = searchMarketplace({ q: 'sync', sort: 'rating' })
    for (const list of [byRating.skills, byRating.plugins]) {
      for (let i = 1; i < list.length; i++) {
        expect(list[i - 1].rating).toBeGreaterThanOrEqual(list[i].rating)
      }
    }
  })

  it('does not leak the internal _score field over IPC', async () => {
    const { searchMarketplace } = await import('../src/shell/skills')
    const res = searchMarketplace({ q: 'ocr' })
    for (const e of [...res.skills, ...res.plugins]) {
      expect(Object.prototype.hasOwnProperty.call(e, '_score')).toBe(false)
    }
  })

  it('empty query keeps the chosen sort intact', async () => {
    const { searchMarketplace } = await import('../src/shell/skills')
    const res = searchMarketplace({ sort: 'name' })
    for (const list of [res.skills, res.plugins]) {
      for (let i = 1; i < list.length; i++) {
        expect(list[i - 1].name.localeCompare(list[i].name)).toBeLessThanOrEqual(0)
      }
    }
  })
})
