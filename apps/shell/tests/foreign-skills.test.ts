import { describe, expect, it } from 'vitest'
import { existsSync, mkdirSync, mkdtempSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import {
  deriveSkillMeta,
  importForeignSkill,
  readSkillBody,
  scanForeignSkills,
  skillsRoot,
  usableSkills,
} from '../src/main/imported-skills'

const SKILL = (name: string, description: string, body = 'Do the thing.') =>
  `---\nname: ${name}\ndescription: ${description}\n---\n\n${body}\n`

/** a fake home dir holding the agent folders a scan reads, plus our save dir */
function foreignTree() {
  const root = mkdtempSync(join(tmpdir(), 'skills-foreign-'))
  const home = join(root, 'home')
  const save = join(root, 'save')
  mkdirSync(home, { recursive: true })
  mkdirSync(save, { recursive: true })
  const put = (rel: string, name: string, text: string) => {
    const dir = join(home, ...rel.split('/'))
    mkdirSync(dir, { recursive: true })
    writeFileSync(join(dir, name), text)
  }
  return { home, save, put, cleanup: () => rmSync(root, { recursive: true, force: true }) }
}

describe('deriveSkillMeta', () => {
  it('names a frontmatter-less file from its sanitized stem and describes it from its first line', () => {
    expect(
      deriveSkillMeta('Trip Planner', 'Turn notes into a docx itinerary.\n\nStep one.\n'),
    ).toEqual({
      name: 'trip-planner',
      description: 'Turn notes into a docx itinerary.',
      body: 'Turn notes into a docx itinerary.\n\nStep one.\n',
      extra: {},
    })
  })

  it('uses a full skill frontmatter as-is, name included', () => {
    const meta = deriveSkillMeta('anything', SKILL('pdf-tools', 'Convert pdf to html'))
    expect(meta?.name).toBe('pdf-tools')
    expect(meta?.description).toBe('Convert pdf to html')
    expect(meta?.body).toBe('Do the thing.\n')
  })

  it('strips markdown markers from a heading or bulleted first line', () => {
    expect(deriveSkillMeta('a', '# Build a **docx** report\n')?.description).toBe(
      'Build a docx report',
    )
    expect(deriveSkillMeta('b', '- Convert pdf to html\n')?.description).toBe('Convert pdf to html')
    expect(deriveSkillMeta('c', '1. Write the memo\n')?.description).toBe('Write the memo')
  })

  it('reads a description-only frontmatter block, as a Cursor rule carries', () => {
    expect(
      deriveSkillMeta(
        'ts-rules',
        '---\ndescription: Use type imports\nglobs: **/*.ts\nalwaysApply: false\n---\n\nUse `import type`.\n',
      ),
    ).toEqual({
      name: 'ts-rules',
      description: 'Use type imports',
      body: 'Use `import type`.\n',
      extra: { globs: '**/*.ts', alwaysApply: 'false' },
    })
  })

  it('replaces the characters a folder name cannot carry', () => {
    expect(deriveSkillMeta('Draft: v2?', 'words\n')?.name).toBe('draft-v2')
  })

  it('caps a description squeezed out of a long first line', () => {
    const meta = deriveSkillMeta('long', `${'word '.repeat(50)}\n`)
    expect(meta?.description.length).toBeLessThanOrEqual(161)
    expect(meta?.description.endsWith('…')).toBe(true)
  })

  it('falls back to the name when the file has nothing to say', () => {
    expect(deriveSkillMeta('blank', '\n\n')).toEqual({
      name: 'blank',
      description: 'blank',
      body: '\n\n',
      extra: {},
    })
  })

  it('returns null for a stem nothing usable can be built from', () => {
    for (const stem of ['???', '...', '  ', ':', '*']) {
      expect(deriveSkillMeta(stem, 'body\n')).toBeNull()
    }
  })
})

describe('scanForeignSkills', () => {
  it('returns nothing when the home dir holds no agent folders', async () => {
    const { home, save, cleanup } = foreignTree()
    expect(await scanForeignSkills(home, save)).toEqual([])
    cleanup()
  })

  it('finds claude commands and skill folders, codex prompts, cursor rules', async () => {
    const { home, save, put, cleanup } = foreignTree()
    put('.claude/commands', 'review.md', 'Review a docx diff carefully.\n')
    put('.claude/skills/pdf-tools', 'SKILL.md', SKILL('pdf-tools', 'Convert pdf to html'))
    put('.codex/prompts', 'y.md', 'Plan a docx itinerary.\n')
    put(
      '.cursor/rules',
      'z.mdc',
      '---\ndescription: Use type imports\nglobs: **/*.ts\n---\n\nUse `import type`.\n',
    )
    const list = await scanForeignSkills(home, save)
    expect(list.map((c) => [c.source, c.name, c.description])).toEqual([
      ['claude', 'pdf-tools', 'Convert pdf to html'],
      ['claude', 'review', 'Review a docx diff carefully.'],
      ['codex', 'y', 'Plan a docx itinerary.'],
      ['cursor', 'z', 'Use type imports'],
    ])
    for (const c of list) {
      expect(c.sourcePath.startsWith(home)).toBe(true)
      expect(c.alreadyImported).toBe(false)
    }
    cleanup()
  })

  it('marks a candidate whose name already sits in our folder', async () => {
    const { home, save, put, cleanup } = foreignTree()
    put('.codex/prompts', 'y.md', 'Plan a docx itinerary.\n')
    mkdirSync(join(skillsRoot(save), 'y'), { recursive: true })
    writeFileSync(join(skillsRoot(save), 'y', 'SKILL.md'), SKILL('y', 'mine'))
    expect((await scanForeignSkills(home, save))[0]?.alreadyImported).toBe(true)
    cleanup()
  })

  it('reads both windsurf locations, both opencode spellings, both pi shapes', async () => {
    const { home, save, put, cleanup } = foreignTree()
    put('.windsurf/workflows', 'a.md', 'words\n')
    put('.codeium/windsurf/workflows', 'b.md', 'words\n')
    put('.config/opencode/command', 'c.md', 'words\n')
    put('.config/opencode/commands', 'd.md', 'words\n')
    put('.pi/agent/commands', 'e.md', 'words\n')
    put('.pi/agent/skills/pi-skill', 'SKILL.md', SKILL('pi-skill', 'A docx thing'))
    expect((await scanForeignSkills(home, save)).map((c) => [c.source, c.name])).toEqual([
      ['opencode', 'c'],
      ['opencode', 'd'],
      ['pi', 'e'],
      ['pi', 'pi-skill'],
      ['windsurf', 'a'],
      ['windsurf', 'b'],
    ])
    cleanup()
  })

  it('skips dotfiles, other extensions, and a stem nothing usable can be built from', async () => {
    const { home, save, put, cleanup } = foreignTree()
    put('.codex/prompts', '.hidden.md', 'words\n')
    put('.codex/prompts', 'notes.txt', 'words\n')
    put('.codex/prompts', '...md', 'words\n')
    expect(await scanForeignSkills(home, save)).toEqual([])
    cleanup()
  })
})

describe('importForeignSkill', () => {
  it('refuses a path the scan did not report', async () => {
    const { home, save, put, cleanup } = foreignTree()
    put('.claude/commands', 'review.md', 'Review a docx diff carefully.\n')
    const outside = join(home, 'secret.md')
    writeFileSync(outside, 'not offered\n')
    for (const bad of [
      outside,
      `${home}/.claude/commands/../commands/evil.md`,
      '/etc/SKILL.md',
      'review.md',
    ]) {
      await expect(importForeignSkill(bad, home, save)).rejects.toThrow('not offered by the scan')
    }
    expect(existsSync(join(skillsRoot(save), 'secret'))).toBe(false)
    expect(existsSync(join(skillsRoot(save), 'evil'))).toBe(false)
    expect(existsSync(join(skillsRoot(save), 'review'))).toBe(false)
    cleanup()
  })

  it('imports a scanned file as a synthesized skill, listed like any other', async () => {
    const { home, save, put, cleanup } = foreignTree()
    put('.codex/prompts', 'trip.md', 'Plan a docx itinerary.\n\nStep one.\n')
    const [candidate] = await scanForeignSkills(home, save)
    const imported = await importForeignSkill(candidate.sourcePath, home, save)
    expect(imported.name).toBe('trip')
    expect(imported.description).toBe('Plan a docx itinerary.')
    expect(readSkillBody(imported.path)).toBe('Plan a docx itinerary.\n\nStep one.\n')
    expect((await usableSkills(save)).map((s) => s.name)).toEqual(['trip'])
    cleanup()
  })

  it('imports a claude skill folder whole, extra files and all', async () => {
    const { home, save, put, cleanup } = foreignTree()
    put('.claude/skills/pdf-tools', 'SKILL.md', SKILL('pdf-tools', 'Convert pdf to html'))
    put('.claude/skills/pdf-tools/reference', 'layout.md', 'grid rules')
    const [candidate] = await scanForeignSkills(home, save)
    const imported = await importForeignSkill(candidate.sourcePath, home, save)
    expect(imported.name).toBe('pdf-tools')
    expect(readSkillBody(imported.path)).toBe('Do the thing.\n')
    expect(existsSync(join(skillsRoot(save), 'pdf-tools', 'reference', 'layout.md'))).toBe(true)
    // the source folder is read, never consumed
    expect(existsSync(join(home, '.claude', 'skills', 'pdf-tools', 'reference', 'layout.md'))).toBe(
      true,
    )
    cleanup()
  })

  it('refuses when the name is already taken, and says so', async () => {
    const { home, save, put, cleanup } = foreignTree()
    put('.codex/prompts', 'trip.md', 'Plan a docx itinerary.\n')
    const [candidate] = await scanForeignSkills(home, save)
    await importForeignSkill(candidate.sourcePath, home, save)
    await expect(importForeignSkill(candidate.sourcePath, home, save)).rejects.toThrow(
      'already there',
    )
    expect((await scanForeignSkills(home, save))[0]?.alreadyImported).toBe(true)
    cleanup()
  })

  it('imports a candidate reached through a climb that lands on the same file', async () => {
    const { home, save, put, cleanup } = foreignTree()
    put('.codex/prompts', 'trip.md', 'Plan a docx itinerary.\n')
    const [candidate] = await scanForeignSkills(home, save)
    const climbed = `${join(home, '.codex', 'prompts')}/../prompts/trip.md`
    expect(resolve(climbed)).toBe(resolve(candidate.sourcePath))
    const imported = await importForeignSkill(climbed, home, save)
    expect(imported.name).toBe('trip')
    cleanup()
  })
})
