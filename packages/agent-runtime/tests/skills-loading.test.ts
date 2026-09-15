/**
 * pi resource wiring tests.
 *
 * A GenOffice marketplace install is only real if the *agent* sees it. These
 * tests pin the two halves of that contract:
 *   1. `additionalSkillPaths` are handed to pi's DefaultResourceLoader, so a
 *      SKILL.md written outside pi's own directories is loaded.
 *   2. `reloadResources()` re-runs discovery, so a skill installed while a
 *      session is alive shows up without rebuilding the session.
 */

import { describe, expect, it, beforeAll, afterAll } from 'vitest'
import { mkdtempSync, mkdirSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { createOfficeSession, type OfficeSession } from '../src/index'

function writeSkill(dir: string, name: string, description: string): string {
  const root = join(dir, name)
  mkdirSync(root, { recursive: true })
  const file = join(root, 'SKILL.md')
  writeFileSync(file, `---\nname: ${name}\ndescription: ${description}\n---\n\nBody for ${name}.\n`, 'utf-8')
  return file
}

describe('pi skill discovery through createOfficeSession', () => {
  let root: string
  let agentDir: string
  let skillsDir: string
  let session: OfficeSession | null = null

  beforeAll(async () => {
    root = mkdtempSync(join(tmpdir(), 'genoffice-pi-skills-'))
    agentDir = join(root, 'agent')
    skillsDir = join(root, 'marketplace-skills')
    mkdirSync(agentDir, { recursive: true })
    mkdirSync(skillsDir, { recursive: true })
    writeSkill(skillsDir, 'wired-skill', 'Loaded through additionalSkillPaths')
    session = await createOfficeSession({
      cwd: root,
      agentDir,
      additionalSkillPaths: [skillsDir],
    })
  }, 60_000)

  afterAll(() => {
    session?.dispose()
    rmSync(root, { recursive: true, force: true })
  })

  it('loads a skill that lives outside pi default directories', () => {
    const { skills, diagnostics } = session!.session.resourceLoader.getSkills()
    const found = skills.find((s) => s.name === 'wired-skill')
    expect(found).toBeDefined()
    expect(found!.description).toBe('Loaded through additionalSkillPaths')
    expect(diagnostics.filter((d) => d.path?.includes('wired-skill'))).toEqual([])
  })

  it('sees a skill installed while the session is running after a reload', async () => {
    const before = await session!.reloadResources()
    const file = writeSkill(skillsDir, 'hot-installed', 'Appeared after a marketplace install')
    const after = await session!.reloadResources()

    expect(after.skills).toBeGreaterThan(before.skills)
    const skills = session!.session.resourceLoader.getSkills().skills
    expect(skills.some((s) => s.name === 'hot-installed')).toBe(true)
    expect(skills.find((s) => s.name === 'hot-installed')!.filePath).toBe(file)
  })
})
