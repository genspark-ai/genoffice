import { describe, expect, it } from 'vitest'
import { existsSync, mkdirSync, mkdtempSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import {
  agentSkillDirs,
  findSkills,
  importSkill,
  isKnownSkillPath,
  knownSkillRoots,
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
  it('puts skills under userData, not where documents live', () => {
    // Home lists every .md under the save directory as a document, so a skills
    // folder there would show every imported SKILL.md to the user
    expect(skillsRoot('/app/userData')).toBe(join('/app/userData', 'skills'))
  })

  it('returns the body with the frontmatter stripped', () => {
    const { save, cleanup } = scaffold({ s: SKILL('s', 'A docx thing', 'Step one.\nStep two.') })
    const [skill] = usableSkills(save)
    expect(readSkillBody(skill.path)).toBe('Step one.\nStep two.\n')
    cleanup()
  })
})

describe('importSkill', () => {
  it('copies the whole skill folder, not just SKILL.md', () => {
    const { save, home, cleanup } = scaffold({}, { 'pptx-builder': SKILL('pptx', 'Build a pptx') })
    const src = join(home, '.agents', 'skills', 'pptx-builder')
    // the reference files a skill ships alongside its instructions
    mkdirSync(join(src, 'reference'), { recursive: true })
    writeFileSync(join(src, 'reference', 'layout.md'), 'grid rules')
    writeFileSync(join(src, 'notes.md'), 'side notes')

    const imported = importSkill(join(src, 'SKILL.md'), save)
    expect(imported).toMatchObject({ name: 'pptx-builder', source: 'genoffice' })
    expect(imported.agent).toBeUndefined()
    const dest = join(save, 'skills', 'pptx-builder')
    expect(existsSync(join(dest, 'reference', 'layout.md'))).toBe(true)
    expect(existsSync(join(dest, 'notes.md'))).toBe(true)
    // and the copy is now one of ours, not a candidate
    expect(usableSkills(save).map((s) => s.name)).toEqual(['pptx-builder'])
    cleanup()
  })

  it('classifies what it imported, and reads the file that landed', () => {
    const { save, home, cleanup } = scaffold({}, { x: SKILL('x', 'Turn a pdf into html') })
    const src = join(home, '.agents', 'skills', 'x')
    const imported = importSkill(join(src, 'SKILL.md'), save)
    expect(imported.relevance.relevant).toBe(true)
    expect(readSkillBody(imported.path)).toBe('Do the thing.\n')
    cleanup()
  })

  it('refuses to overwrite a skill already in our folder, and leaves it alone', () => {
    const { save, home, cleanup } = scaffold(
      { shared: SKILL('shared', 'Mine', 'mine wins') },
      { shared: SKILL('shared', 'Theirs', 'theirs loses') },
    )
    const src = join(home, '.agents', 'skills', 'shared')
    expect(() => importSkill(join(src, 'SKILL.md'), save)).toThrow('already there')
    expect(readSkillBody(join(save, 'skills', 'shared', 'SKILL.md'))).toBe('mine wins\n')
    cleanup()
  })

  it('names the folder after the directory it came from', () => {
    const { save, home, cleanup } = scaffold(
      {},
      { 'pdf-tools': SKILL('different-name', 'A pdf thing') },
    )
    const src = join(home, '.agents', 'skills', 'pdf-tools')
    // frontmatter says one thing, the folder says another; the folder wins,
    // because the folder is what skillsRoot is keyed on
    expect(importSkill(join(src, 'SKILL.md'), save).name).toBe('pdf-tools')
    cleanup()
  })
})

describe('isKnownSkillPath', () => {
  it('accepts every path the scan hands out, ours and the agents', () => {
    const { save, home, cleanup } = scaffold(
      { mine: SKILL('mine', 'A docx thing') },
      { theirs: SKILL('theirs', 'A pdf thing') },
    )
    const roots = knownSkillRoots(save, {}, home)
    const found = findSkills(save, {}, home)
    expect(found).toHaveLength(2)
    for (const skill of found) {
      expect(isKnownSkillPath(skill.path, roots)).toBe(true)
    }
    cleanup()
  })

  it('rejects a file outside every skills root', () => {
    expect(isKnownSkillPath('/etc/passwd', ['/save/skills'])).toBe(false)
    expect(isKnownSkillPath('/Users/me/.ssh/id_rsa', ['/save/skills'])).toBe(false)
  })

  it('rejects another file sitting beside a SKILL.md', () => {
    // the guard is the file name as much as the directory: a skills folder can
    // legitimately hold reference files, and those are not skills
    expect(isKnownSkillPath('/save/skills/mine/notes.md', ['/save/skills'])).toBe(false)
    expect(isKnownSkillPath('/save/skills/mine/scripts/evil.sh', ['/save/skills'])).toBe(false)
  })

  it('rejects a SKILL.md reached by climbing out of a skills root', () => {
    expect(isKnownSkillPath('/save/skills/../../.ssh/SKILL.md', ['/save/skills'])).toBe(false)
    expect(isKnownSkillPath('/save/skills/../../../../../../etc/SKILL.md', ['/save/skills'])).toBe(
      false,
    )
  })

  it('accepts a climb that lands back inside the root: it is the same skill', () => {
    // the traversal test above is only about leaving; one that comes back is
    // just an awkward spelling of a real path
    expect(isKnownSkillPath('/save/skills/../skills/x/SKILL.md', ['/save/skills'])).toBe(true)
    expect(isKnownSkillPath('/save/skills/a/../b/SKILL.md', ['/save/skills'])).toBe(true)
  })

  it('rejects a root that merely starts with the same characters', () => {
    expect(isKnownSkillPath('/save/skills-backup/x/SKILL.md', ['/save/skills'])).toBe(false)
  })

  it('rejects a skills folder with the same last component but a different parent', () => {
    // matching on the folder name alone would read any 'skills' directory on
    // the machine, which is exactly the leak this guard exists to close
    expect(isKnownSkillPath('/elsewhere/skills/mine/SKILL.md', ['/save/skills'])).toBe(false)
    expect(isKnownSkillPath('/Users/me/.claude/skills/mine/SKILL.md', ['/save/skills'])).toBe(false)
    expect(isKnownSkillPath('/private/tmp/agent/skills/x/SKILL.md', ['/save/skills'])).toBe(false)
  })

  it('rejects anything that is not an absolute path', () => {
    for (const bad of ['', 'skills/mine/SKILL.md', undefined, null, 42, {}]) {
      expect(isKnownSkillPath(bad, ['/save/skills'])).toBe(false)
    }
  })

  it('normalizes the incoming path, so an equivalent spelling is the same skill', () => {
    expect(isKnownSkillPath('/save/./skills/mine/SKILL.md', ['/save/skills'])).toBe(true)
    expect(isKnownSkillPath('/save/skills-backup/../skills/mine/SKILL.md', ['/save/skills'])).toBe(
      true,
    )
  })

  it('resolves the roots it compares against', () => {
    // a root with a trailing slash or a dot segment is still the same root
    const path = '/save/skills/mine/SKILL.md'
    expect(isKnownSkillPath(path, ['/save/./skills/'])).toBe(true)
  })
})
