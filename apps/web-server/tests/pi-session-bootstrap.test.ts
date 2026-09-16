/**
 * pi-session-bootstrap regression test.
 *
 * The GenOffice shell's pi integration has four moving parts that must all stay
 * wired up together:
 *
 *   1. `ensureBuiltInSkillsMaterialized()` writes a SKILL.md per built-in skill
 *      into DATA_DIR/pi-skills/<id>/. Without it the agent has no idea the
 *      tools (`read_blocks`, `web_search`, etc.) exist.
 *   2. `ensureLumosSkillsRegistered()` mirrors every LumosAI bundled skill
 *      into DATA_DIR/lumos-skill-wrappers/<id>/ and adds the wrapper root to
 *      pi's settings.json. Without it the upstream translate scripts are
 *      invisible to pi.
 *   3. `ensureSkillDirRegistered()` keeps pi's settings pointed at the
 *      marketplace skills dir.
 *   4. `createOfficeSession()` builds a real `AgentSession` and
 *      `resourceLoader.getSkills()` returns the merged view.
 *
 * Set DATA_DIR + LUMOS_HOME *before* the dynamic import so DATA_DIR is
 * computed against the temp directory from the very first line.
 */
import { mkdtempSync, existsSync, readdirSync, readFileSync, rmSync } from 'node:fs'
import { tmpdir, homedir } from 'node:os'
import { join } from 'node:path'
import { afterAll, beforeAll, describe, expect, it } from 'vitest'

const TMP_DATA = mkdtempSync(join(tmpdir(), 'genoffice-pi-bootstrap-'))
process.env.DATA_DIR = TMP_DATA
process.env.GENOFFICE_DATA_DIR = TMP_DATA
process.env.LUMOS_HOME = join(homedir(), '.lumos')

// Dynamic imports so DATA_DIR is computed AFTER the env vars are set.
const piResources = await import('../src/shell/pi-resources')
const piSession = await import('../src/shell/pi-session')

const {
  PI_AGENT_DIR,
  PI_SKILLS_DIR,
  LUMOS_SKILLS_WRAPPER_DIR,
  ensureBuiltInSkillsMaterialized,
  ensureLumosSkillsRegistered,
  ensureSkillDirRegistered,
} = piResources
const { getPiSession, invalidatePiSession } = piSession

describe('pi-session bootstrap', () => {
  beforeAll(() => {
    // Sanity: we are pointed at the temp DATA_DIR.
    expect(process.env.DATA_DIR).toBe(TMP_DATA)
    expect(PI_SKILLS_DIR).toBe(join(TMP_DATA, 'pi-skills'))
  })

  afterAll(async () => {
    invalidatePiSession()
    try { rmSync(TMP_DATA, { recursive: true, force: true }) } catch { /* ignore */ }
  })

  it('materialises SKILL.md files for every built-in skill', () => {
    const result = ensureBuiltInSkillsMaterialized()
    expect(result.written.length).toBeGreaterThanOrEqual(13)
    const ids = readdirSync(PI_SKILLS_DIR)
    expect(ids.length).toBeGreaterThanOrEqual(13)
    for (const id of ids) {
      const skillPath = join(PI_SKILLS_DIR, id, 'SKILL.md')
      if (existsSync(skillPath)) {
        const body = readFileSync(skillPath, 'utf-8')
        expect(body.startsWith('---')).toBe(true)
        expect(body).toMatch(/^name:\s+\S/m)
        expect(body).toMatch(/^description:\s+/m)
      }
    }
    // Idempotent: second call writes nothing.
    const second = ensureBuiltInSkillsMaterialized()
    expect(second.written.length).toBe(0)
    expect(second.skipped.length).toBe(result.written.length)
  })

  it('writes pi-compatible wrapper SKILL.md for each LumosAI bundled skill', async () => {
    const result = await ensureLumosSkillsRegistered()
    if (result.registered.length === 0 && result.alreadyHad.length === 0) {
      // LUMOS_HOME may not exist on the CI host — treat as opt-in.
      return
    }
    const wrapperIds = readdirSync(LUMOS_SKILLS_WRAPPER_DIR)
    expect(wrapperIds.length).toBeGreaterThan(0)
    const first = wrapperIds.find((id) => existsSync(join(LUMOS_SKILLS_WRAPPER_DIR, id, 'SKILL.md')))
    expect(first).toBeTruthy()
    const body = readFileSync(join(LUMOS_SKILLS_WRAPPER_DIR, first!, 'SKILL.md'), 'utf-8')
    expect(body.startsWith('---')).toBe(true)
    expect(body).toMatch(/^name:\s+[a-z0-9-]+/m)
    expect(body).toMatch(/^description:\s+/m)
  })

  it('registers both skill dirs in pi settings.json', async () => {
    await ensureSkillDirRegistered()
    const settings = JSON.parse(readFileSync(join(PI_AGENT_DIR, 'settings.json'), 'utf-8')) as { skills: string[] }
    expect(settings.skills).toContain(PI_SKILLS_DIR)
    if (existsSync(LUMOS_SKILLS_WRAPPER_DIR)) {
      expect(settings.skills).toContain(LUMOS_SKILLS_WRAPPER_DIR)
    }
  })

  it('creates a real pi AgentSession that sees the bootstrapped skills', async () => {
    const session = await getPiSession()
    expect(session.session).toBeTruthy()
    const skills = session.resourceLoader.getSkills()
    const names = skills.skills.map((s) => s.name)
    expect(names).toContain('docs-skill')
    expect(names).toContain('web-search')
    const tools = session.session.getAllTools()
    expect(tools.length).toBeGreaterThan(0)
    const toolNames = tools.map((t) => t.name)
    expect(toolNames).toContain('bash')
  }, 30_000)
})
