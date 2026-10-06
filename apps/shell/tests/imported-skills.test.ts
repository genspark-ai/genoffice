import { describe, expect, it } from 'vitest'
import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import {
  agentSkillDirs,
  findSkills,
  readSkillBody,
  skillsRoot,
  usableSkills,
} from '../src/main/imported-skills'

/** a scratch tree: <root>/save/skills/<name>/SKILL.md and <root>/home/.claude/skills/… */
/** a scratch tree: <root>/save/skills/<name>/SKILL.md and <root>/home/.agents/skills/<name>/SKILL.md */
function scaffold(ours: Record<string, string> = {}, theirs: Record<string, string> = {}) {
  const root = mkdtempSync(join(tmpdir(), 'skills-scan-'))
  const save = join(root, 'save')
  const h = join(root, 'home')
  mkdirSync(save, { recursive: true })
  mkdirSync(h, { recursive: true })
  const put = (base: string, skills: Record<string, string>) => {
    for (const [name, body] of Object.entries(skills)) {
      mkdirSync(join(base, name), { recursive: true })
      writeFileSync(join(base, name, 'SKILL.md'), body)
    }
  }
  put(join(save, 'skills'), ours)
  put(join(h, '.agents', 'skills'), theirs)
  return { root, save, home: h, cleanup: () => rmSync(root, { recursive: true, force: true }) }
}

const SKILL = (name: string, description: string, body = 'Do the thing.') =>
  `---\nname: ${name}\ndescription: ${description}\n---\n\n${body}\n`

const SHARED_BODY = SKILL('shared', 'Write a docx')

describe('findSkills', () => {
  it('returns nothing when there is no skills folder and no agent home', () => {
    const { save, home, cleanup } = scaffold()
    // home is passed explicitly on purpose: this machine really does have a
    // populated ~/.agents/skills, and a bare findSkills() would read it
    expect(findSkills(save, {}, home)).toEqual([])
    cleanup()
  })

  it('reads a skill we put in our own folder, and does not mark it as an agent one', () => {
    const { save, cleanup } = scaffold({ 'my-skill': SKILL('my-skill', 'Make a docx report') })
    const [skill] = usableSkills(save)
    expect(skill.name).toBe('my-skill')
    expect(skill.source).toBe('genoffice')
    expect(skill.agent).toBeUndefined()
    expect(skill.relevance.relevant).toBe(true)
    cleanup()
  })

  it('offers an agent skill as a candidate rather than as usable', () => {
    const { save, home, cleanup } = scaffold(
      {},
      { 'pptx-builder': SKILL('pptx-builder', 'Build a pptx') },
    )
    const all = findSkills(save, {}, home)
    expect(all).toHaveLength(1)
    expect(all[0].source).toBe('agent')
    expect(all[0].agent).toBe('agents')
    // ours-only: nothing was imported by merely looking
    expect(usableSkills(save)).toEqual([])
    cleanup()
  })

  it('keeps our copy when an agent directory has the same name', () => {
    const { save, home, cleanup } = scaffold({ shared: SHARED_BODY }, { shared: SHARED_BODY })
    const all = findSkills(save, {}, home)
    expect(all).toHaveLength(1)
    expect(all[0].source).toBe('genoffice')
    cleanup()
  })

  it('classifies a skill that is about deployment as not relevant, without hiding it', () => {
    const { save, home, cleanup } = scaffold(
      {},
      { deploy: SKILL('deploy', 'Roll back a Kubernetes release') },
    )
    const [skill] = findSkills(save, {}, home)
    expect(skill.relevance.relevant).toBe(false)
    cleanup()
  })

  it('skips a directory with no SKILL.md and one whose frontmatter is unusable', () => {
    const { save, cleanup } = scaffold({ good: SKILL('good', 'Convert pdf to html') })
    mkdirSync(join(save, 'skills', 'empty'), { recursive: true })
    mkdirSync(join(save, 'skills', 'broken'), { recursive: true })
    writeFileSync(join(save, 'skills', 'broken', 'SKILL.md'), 'no frontmatter here\n')
    mkdirSync(join(save, 'skills', 'noname'), { recursive: true })
    writeFileSync(join(save, 'skills', 'noname', 'SKILL.md'), '---\ndescription: x\n---\n\nbody\n')
    expect(usableSkills(save).map((s) => s.name)).toEqual(['good'])
    cleanup()
  })

  it('reads the home from the argument, not the real one', () => {
    const { save, home, cleanup } = scaffold({}, { x: SKILL('x', 'A pdf thing') })
    // a home with nothing in it must not reach the developer's real ~/.claude
    expect(findSkills(save, {}, join(home, 'nowhere'))).toEqual([])
    cleanup()
  })

  it('survives an unreadable directory by skipping it', () => {
    const { save, home, cleanup } = scaffold({}, { a: SKILL('a', 'docx') })
    mkdirSync(join(home, '.agents', 'skills', 'not-a-dir'), { recursive: true })
    writeFileSync(join(home, '.agents', 'skills', 'not-a-dir', 'SKILL.md'), SKILL('b', 'xlsx'))
    expect(() => findSkills(save, {}, home)).not.toThrow()
    cleanup()
  })
})

describe('agentSkillDirs', () => {
  it('lists the cross-agent directory before the vendor ones', () => {
    const dirs = agentSkillDirs({}, '/h')
    expect(dirs[0].agent).toBe('agents')
    expect(dirs.map((d) => d.dir)).toContain(join('/h', '.claude', 'skills'))
  })

  it('honours the config-dir environment variables those agents use', () => {
    const dirs = agentSkillDirs(
      { CLAUDE_CONFIG_DIR: '/cfg/claude', CODEX_HOME: '/cfg/codex' },
      '/h',
    )
    expect(dirs.find((d) => d.agent === 'claude-code')!.dir).toBe('/cfg/claude/skills')
    expect(dirs.find((d) => d.agent === 'codex')!.dir).toBe('/cfg/codex/skills')
  })

  it('ignores an empty config dir rather than joining onto it', () => {
    const dirs = agentSkillDirs({ CLAUDE_CONFIG_DIR: '' }, '/h')
    expect(dirs.find((d) => d.agent === 'claude-code')!.dir).toBe('/h/.claude/skills')
  })
})

describe('skillsRoot / readSkillBody', () => {
  it('puts skills under the save directory', () => {
    expect(skillsRoot('/save/dir')).toBe(join('/save/dir', 'skills'))
  })

  it('returns the body with the frontmatter stripped', () => {
    const { save, cleanup } = scaffold({ s: SKILL('s', 'A docx thing', 'Step one.\nStep two.') })
    const [skill] = usableSkills(save)
    expect(readSkillBody(skill.path)).toBe('Step one.\nStep two.\n')
    cleanup()
  })
})
