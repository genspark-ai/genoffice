/**
 * Agent Skills spec compliance regression test.
 *
 * Reference: https://agentskills.io/specification (the spec Pi implements).
 * - `name`: 1-64 chars, lowercase a-z0-9 and hyphens only, no leading/trailing
 *   hyphen, no consecutive hyphens, must match the parent directory name.
 * - `description`: 1-1024 chars, non-empty.
 * - `allowed-tools` (optional): space-separated STRING (not a YAML list).
 * - Unknown frontmatter fields are ignored by Pi but allowed.
 *
 * This test guards against the YAML-list form slipping back in.
 */
import { mkdtempSync, existsSync, readdirSync, readFileSync, rmSync } from 'node:fs'
import { tmpdir, homedir } from 'node:os'
import { join } from 'node:path'
import { afterAll, describe, expect, it } from 'vitest'

const TMP_DATA = mkdtempSync(join(tmpdir(), 'genoffice-agent-skills-spec-'))
process.env.DATA_DIR = TMP_DATA
process.env.GENOFFICE_DATA_DIR = TMP_DATA
process.env.LUMOS_HOME = join(homedir(), '.lumos')

const piResources = await import('../src/shell/pi-resources')
const { PI_SKILLS_DIR, LUMOS_SKILLS_WRAPPER_DIR, ensureBuiltInSkillsMaterialized, ensureLumosSkillsRegistered } =
  piResources

function parseFrontmatter(body: string): Record<string, string> {
  const m = body.match(/^---\n([\s\S]*?)\n---\n?/)
  if (!m) return {}
  const out: Record<string, string> = {}
  // naive frontmatter parser good enough for our single-line scalars
  for (const line of m[1].split('\n')) {
    const idx = line.indexOf(':')
    if (idx === -1) continue
    out[line.slice(0, idx).trim()] = line.slice(idx + 1).trim()
  }
  return out
}

function isValidName(name: string): boolean {
  if (name.length < 1 || name.length > 64) return false
  if (!/^[a-z0-9-]+$/.test(name)) return false
  if (name.startsWith('-') || name.endsWith('-')) return false
  if (name.includes('--')) return false
  return true
}

describe('Agent Skills spec compliance', () => {
  afterAll(() => {
    try { rmSync(TMP_DATA, { recursive: true, force: true }) } catch { /* ignore */ }
  })

  it('every built-in skill name matches its parent directory and satisfies the spec', () => {
    const result = ensureBuiltInSkillsMaterialized()
    expect(result.written.length).toBeGreaterThan(0)
    const ids = readdirSync(PI_SKILLS_DIR)
    for (const id of ids) {
      const skillPath = join(PI_SKILLS_DIR, id, 'SKILL.md')
      if (!existsSync(skillPath)) continue
      const body = readFileSync(skillPath, 'utf-8')
      const fm = parseFrontmatter(body)
      expect(isValidName(id), `dir "${id}" must be a valid skill name`).toBe(true)
      expect(fm.name, `dir "${id}": SKILL.md missing name`).toBeTruthy()
      expect(isValidName(fm.name), `SKILL.md name "${fm.name}" violates spec`).toBe(true)
      expect(fm.name, `name must match parent directory "${id}"`).toBe(id)
      expect(fm.description, `description required (dir "${id}")`).toBeTruthy()
      expect(fm.description.length).toBeGreaterThanOrEqual(1)
      expect(fm.description.length).toBeLessThanOrEqual(1024)
    }
  })

  it('allowed-tools is a space-separated string (not a YAML list)', () => {
    ensureBuiltInSkillsMaterialized()
    const ids = readdirSync(PI_SKILLS_DIR)
    for (const id of ids) {
      const skillPath = join(PI_SKILLS_DIR, id, 'SKILL.md')
      if (!existsSync(skillPath)) continue
      const body = readFileSync(skillPath, 'utf-8')
      // Find the allowed-tools line; the spec says it must be a single
      // scalar string. A YAML list looks like:
      //   allowed-tools:
      //     - foo
      //     - bar
      // which would make the line end with `:` not a value.
      const lines = body.split('\n')
      const idx = lines.findIndex((l) => l.trimStart().startsWith('allowed-tools:'))
      if (idx === -1) continue
      const line = lines[idx]
      expect(
        line,
        `${id}: allowed-tools must be a single-line space-separated string per spec, got "${line}"`,
      ).toMatch(/^allowed-tools:\s+\S+/)
      // The line after must NOT be a YAML list item (`  - `)
      const next = lines[idx + 1] ?? ''
      expect(
        next,
        `${id}: allowed-tools must not be a YAML list — found continuation "${next}"`,
      ).not.toMatch(/^\s+-\s+/)
    }
  })

  it('every translate-related built-in skill lists the 6 translate_* tools', () => {
    ensureBuiltInSkillsMaterialized()
    const skillPath = join(PI_SKILLS_DIR, 'translate-skill', 'SKILL.md')
    const body = readFileSync(skillPath, 'utf-8')
    const fm = parseFrontmatter(body)
    expect(fm['allowed-tools']).toBeTruthy()
    const tools = fm['allowed-tools'].split(/\s+/)
    expect(tools).toEqual(expect.arrayContaining([
      'translate_text',
      'translate_file',
      'build_dictionary',
      'kb_search',
      'kb_upsert',
      'kb_remove',
    ]))
  })

  it('lumos wrappers also satisfy name rules when present', async () => {
    const result = await ensureLumosSkillsRegistered()
    if (result.registered.length === 0 && result.alreadyHad.length === 0) return
    const ids = readdirSync(LUMOS_SKILLS_WRAPPER_DIR)
    for (const id of ids) {
      const skillPath = join(LUMOS_SKILLS_WRAPPER_DIR, id, 'SKILL.md')
      if (!existsSync(skillPath)) continue
      const body = readFileSync(skillPath, 'utf-8')
      const fm = parseFrontmatter(body)
      expect(isValidName(id), `lumos wrapper dir "${id}" violates spec`).toBe(true)
      expect(isValidName(fm.name ?? ''), `lumos wrapper SKILL.md name "${fm.name}" violates spec`).toBe(true)
    }
  })
})
