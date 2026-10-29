import { describe, expect, it, vi } from 'vitest'

/** lets one test make the final rename of an install fail, without touching the disk */
const h = vi.hoisted(() => ({ failRenameTo: '', failRmOf: '' }))
vi.mock('node:fs', async (importOriginal) => {
  const actual = await importOriginal<typeof import('node:fs')>()
  return {
    ...actual,
    rmSync: (path: import('node:fs').PathLike, opts?: import('node:fs').RmOptions) => {
      if (h.failRmOf && new RegExp(h.failRmOf).test(String(path))) {
        const e = new Error('resource busy') as NodeJS.ErrnoException
        e.code = 'EBUSY'
        throw e
      }
      return actual.rmSync(path, opts)
    },
    renameSync: (from: import('node:fs').PathLike, to: import('node:fs').PathLike) => {
      if (h.failRenameTo && String(to) === h.failRenameTo) {
        // fail once: the restore of the parked copy is a plain rename to the same place
        h.failRenameTo = ''
        throw new Error('disk full')
      }
      return actual.renameSync(from, to)
    },
  }
})
import {
  existsSync,
  mkdirSync,
  mkdtempSync,
  readFileSync,
  readdirSync,
  rmSync,
  symlinkSync,
  writeFileSync,
} from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import JSZip from 'jszip'
import {
  importSkillDir,
  importSkillMd,
  importSkillZip,
  installRecommendedSkill,
  isKnownSkillPath,
  readSkillBody,
  recommendedSkills,
  removeSkill,
  skillRawText,
  skillVersion,
  skillsRoot,
  usableSkills,
  writeSkill,
} from '../src/main/imported-skills'

/** a scratch tree: <root>/save/skills/<name>/SKILL.md (ours) and <root>/src/<name>/… (to import) */
function scaffold(ours: Record<string, string> = {}, picked: Record<string, string> = {}) {
  const root = mkdtempSync(join(tmpdir(), 'skills-import-'))
  const save = join(root, 'save')
  const src = join(root, 'src')
  mkdirSync(save, { recursive: true })
  mkdirSync(src, { recursive: true })
  const put = (base: string, skills: Record<string, string>) => {
    for (const [name, body] of Object.entries(skills)) {
      mkdirSync(join(base, name), { recursive: true })
      writeFileSync(join(base, name, 'SKILL.md'), body)
    }
  }
  put(join(save, 'skills'), ours)
  put(src, picked)
  return { root, save, src, cleanup: () => rmSync(root, { recursive: true, force: true }) }
}

const SKILL = (name: string, description: string, body = 'Do the thing.') =>
  `---\nname: ${name}\ndescription: ${description}\n---\n\n${body}\n`

async function zipOf(files: Record<string, string | Buffer>, dest: string): Promise<string> {
  const zip = new JSZip()
  for (const [name, data] of Object.entries(files)) zip.file(name, data)
  writeFileSync(dest, await zip.generateAsync({ type: 'nodebuffer' }))
  return dest
}

describe('usableSkills', () => {
  it('returns nothing when there is no skills folder', async () => {
    const { save, cleanup } = scaffold()
    expect(await usableSkills(save)).toEqual([])
    cleanup()
  })

  it('reads a skill in our own folder and classifies it', async () => {
    const { save, cleanup } = scaffold({ 'my-skill': SKILL('my-skill', 'Make a docx report') })
    const [skill] = await usableSkills(save)
    expect(skill.name).toBe('my-skill')
    expect(skill.relevance.relevant).toBe(true)
    cleanup()
  })

  it('never looks in a coding agent directory', async () => {
    const { root, save, cleanup } = scaffold()
    const agents = join(root, 'home', '.agents', 'skills', 'deploy')
    mkdirSync(agents, { recursive: true })
    writeFileSync(join(agents, 'SKILL.md'), SKILL('deploy', 'Roll back a release'))
    expect(await usableSkills(save)).toEqual([])
    cleanup()
  })

  it('skips a directory with no SKILL.md and one whose frontmatter is unusable', async () => {
    const { save, cleanup } = scaffold({ good: SKILL('good', 'Convert pdf to html') })
    mkdirSync(join(save, 'skills', 'empty'), { recursive: true })
    mkdirSync(join(save, 'skills', 'broken'), { recursive: true })
    writeFileSync(join(save, 'skills', 'broken', 'SKILL.md'), 'no frontmatter here\n')
    mkdirSync(join(save, 'skills', 'noname'), { recursive: true })
    writeFileSync(join(save, 'skills', 'noname', 'SKILL.md'), '---\ndescription: x\n---\n\nbody\n')
    expect((await usableSkills(save)).map((s) => s.name)).toEqual(['good'])
    cleanup()
  })
})

describe('skillsRoot / readSkillBody', () => {
  it('puts skills under userData, not where documents live', () => {
    expect(skillsRoot('/app/userData')).toBe(join('/app/userData', 'skills'))
  })

  it('returns the body with the frontmatter stripped', async () => {
    const { save, cleanup } = scaffold({ s: SKILL('s', 'A docx thing', 'Step one.\nStep two.') })
    const [skill] = await usableSkills(save)
    expect(readSkillBody(skill.path)).toBe('Step one.\nStep two.\n')
    cleanup()
  })
})

describe('importSkillDir', () => {
  it('copies the whole skill folder, not just SKILL.md', async () => {
    const { save, src, cleanup } = scaffold({}, { 'pptx-builder': SKILL('pptx', 'Build a pptx') })
    const dir = join(src, 'pptx-builder')
    mkdirSync(join(dir, 'reference'), { recursive: true })
    writeFileSync(join(dir, 'reference', 'layout.md'), 'grid rules')
    writeFileSync(join(dir, 'notes.md'), 'side notes')

    const imported = importSkillDir(dir, save)
    expect(imported.name).toBe('pptx-builder')
    expect(imported.relevance.relevant).toBe(true)
    const dest = join(save, 'skills', 'pptx-builder')
    expect(existsSync(join(dest, 'reference', 'layout.md'))).toBe(true)
    expect(existsSync(join(dest, 'notes.md'))).toBe(true)
    expect((await usableSkills(save)).map((s) => s.name)).toEqual(['pptx-builder'])
    cleanup()
  })

  it('refuses a folder without a usable SKILL.md', () => {
    const { save, src, cleanup } = scaffold()
    mkdirSync(join(src, 'plain'))
    writeFileSync(join(src, 'plain', 'README.md'), 'not a skill')
    expect(() => importSkillDir(join(src, 'plain'), save)).toThrow('no SKILL.md')
    mkdirSync(join(src, 'bare'))
    writeFileSync(join(src, 'bare', 'SKILL.md'), 'no frontmatter')
    expect(() => importSkillDir(join(src, 'bare'), save)).toThrow('frontmatter')
    expect(existsSync(join(save, 'skills'))).toBe(false)
    cleanup()
  })

  it('refuses to overwrite a skill already in our folder, and leaves it alone', () => {
    const { save, src, cleanup } = scaffold(
      { shared: SKILL('shared', 'Mine', 'mine wins') },
      { shared: SKILL('shared', 'Theirs', 'theirs loses') },
    )
    expect(() => importSkillDir(join(src, 'shared'), save)).toThrow('already there')
    expect(readSkillBody(join(save, 'skills', 'shared', 'SKILL.md'))).toBe('mine wins\n')
    cleanup()
  })

  it('refuses a skill directory that is itself a symlink', () => {
    const { save, src, cleanup } = scaffold({}, { real: SKILL('real', 'A docx thing') })
    symlinkSync(join(src, 'real'), join(src, 'linked'))
    expect(() => importSkillDir(join(src, 'linked'), save)).toThrow('symlink')
    expect(existsSync(join(save, 'skills', 'linked'))).toBe(false)
    cleanup()
  })

  it('refuses a skill that ships a symlink instead of copying the reach-through', () => {
    const { save, src, cleanup } = scaffold({}, { linked: SKILL('linked', 'A docx thing') })
    const dir = join(src, 'linked')
    writeFileSync(join(dir, 'outside.md'), 'somewhere else')
    symlinkSync(join(dir, 'outside.md'), join(dir, 'reference.md'))
    expect(() => importSkillDir(dir, save)).toThrow('symlink')
    expect(existsSync(join(save, 'skills', 'linked'))).toBe(false)
    cleanup()
  })

  it('refuses a skill over the import size cap', () => {
    const { save, src, cleanup } = scaffold({}, { heavy: SKILL('heavy', 'A docx thing') })
    mkdirSync(join(src, 'heavy', 'corpus'), { recursive: true })
    writeFileSync(join(src, 'heavy', 'corpus', 'big.bin'), Buffer.alloc(11 * 1024 * 1024, 0))
    expect(() => importSkillDir(join(src, 'heavy'), save)).toThrow('import cap')
    expect(existsSync(join(save, 'skills', 'heavy'))).toBe(false)
    cleanup()
  })

  it('names the folder after the directory it came from, not the frontmatter', () => {
    const { save, src, cleanup } = scaffold({}, { 'pdf-tools': SKILL('other', 'A pdf thing') })
    expect(importSkillDir(join(src, 'pdf-tools'), save).name).toBe('pdf-tools')
    cleanup()
  })
})

describe('importSkillZip', () => {
  it('unpacks a zip with one top-level folder, named after the folder', async () => {
    const { root, save, cleanup } = scaffold()
    const zip = await zipOf(
      {
        'pptx-builder/SKILL.md': SKILL('pptx', 'Build a pptx'),
        'pptx-builder/reference/layout.md': 'grid rules',
      },
      join(root, 'anything.zip'),
    )
    const imported = await importSkillZip(zip, save)
    expect(imported.name).toBe('pptx-builder')
    expect(existsSync(join(save, 'skills', 'pptx-builder', 'reference', 'layout.md'))).toBe(true)
    expect(readSkillBody(imported.path)).toBe('Do the thing.\n')
    cleanup()
  })

  it('unpacks a zip with SKILL.md at the root, named after the file', async () => {
    const { root, save, cleanup } = scaffold()
    const zip = await zipOf(
      { 'SKILL.md': SKILL('x', 'A docx thing'), 'notes.md': 'side' },
      join(root, 'docx-helper.zip'),
    )
    expect((await importSkillZip(zip, save)).name).toBe('docx-helper')
    expect(existsSync(join(save, 'skills', 'docx-helper', 'notes.md'))).toBe(true)
    cleanup()
  })

  it('imports a single markdown file, named from its frontmatter', () => {
    const { root, save, cleanup } = scaffold()
    const md = join(root, 'anything.md')
    writeFileSync(md, SKILL('trip-planner', 'Plan a docx itinerary'))
    const imported = importSkillMd(md, save)
    expect(imported.name).toBe('trip-planner')
    expect(imported.path).toBe(join(save, 'skills', 'trip-planner', 'SKILL.md'))
    expect(readSkillBody(imported.path)).toBe('Do the thing.\n')
    expect(() => importSkillMd(md, save)).toThrow(/already there/)
    cleanup()
  })

  it('falls back to the file stem when the frontmatter name is not a folder name', () => {
    const { root, save, cleanup } = scaffold()
    const md = join(root, 'travel.md')
    writeFileSync(md, SKILL('a/b', 'Plan a docx itinerary'))
    expect(importSkillMd(md, save).name).toBe('travel')
    cleanup()
  })

  it('imports a markdown file without frontmatter from its sanitized stem', () => {
    const { root, save, cleanup } = scaffold()
    const md = join(root, 'Trip Planner.md')
    writeFileSync(md, 'Turn notes into a docx itinerary.\n\nStep one.\n')
    const imported = importSkillMd(md, save)
    expect(imported.name).toBe('trip-planner')
    expect(imported.path).toBe(join(save, 'skills', 'trip-planner', 'SKILL.md'))
    expect(imported.description).toBe('Turn notes into a docx itinerary.')
    expect(readFileSync(imported.path, 'utf8')).toBe(
      '---\nname: trip-planner\ndescription: Turn notes into a docx itinerary.\n---\n\nTurn notes into a docx itinerary.\n\nStep one.\n',
    )
    expect(readSkillBody(imported.path)).toBe('Turn notes into a docx itinerary.\n\nStep one.\n')
    expect(() => importSkillMd(md, save)).toThrow(/already there/)
    cleanup()
  })

  it('derives a description from a first line that carries markdown markers', () => {
    const { root, save, cleanup } = scaffold()
    writeFileSync(join(root, 'heading.md'), '# Build a **docx** report\n\nStep one.\n')
    expect(importSkillMd(join(root, 'heading.md'), save).description).toBe('Build a docx report')
    writeFileSync(join(root, 'listing.md'), '- Convert pdf to html\n')
    expect(importSkillMd(join(root, 'listing.md'), save).description).toBe('Convert pdf to html')
    cleanup()
  })

  it('imports a Cursor .mdc rule: description from its frontmatter, other keys kept', () => {
    const { root, save, cleanup } = scaffold()
    const mdc = join(root, 'ts-rules.mdc')
    writeFileSync(
      mdc,
      '---\ndescription: Use type imports\nglobs: **/*.ts\nalwaysApply: false\n---\n\nUse `import type`.\n',
    )
    const imported = importSkillMd(mdc, save)
    expect(imported.name).toBe('ts-rules')
    expect(imported.description).toBe('Use type imports')
    expect(readFileSync(imported.path, 'utf8')).toBe(
      '---\nname: ts-rules\ndescription: Use type imports\nglobs: **/*.ts\nalwaysApply: false\n---\n\nUse `import type`.\n',
    )
    expect(readSkillBody(imported.path)).toBe('Use `import type`.\n')
    cleanup()
  })

  it('imports a .mdc rule with no frontmatter like any plain note', () => {
    const { root, save, cleanup } = scaffold()
    const mdc = join(root, 'plain-rule.mdc')
    writeFileSync(mdc, 'Prefer tables over images in a xlsx export.\n')
    const imported = importSkillMd(mdc, save)
    expect(imported.name).toBe('plain-rule')
    expect(imported.description).toBe('Prefer tables over images in a xlsx export.')
    cleanup()
  })

  it('falls back to the name as a description when the file says nothing', () => {
    const { root, save, cleanup } = scaffold()
    writeFileSync(join(root, 'silent.md'), '\n\n')
    expect(importSkillMd(join(root, 'silent.md'), save).description).toBe('silent')
    cleanup()
  })

  it('refuses a stem nothing usable can be derived from', () => {
    const { root, save, cleanup } = scaffold()
    writeFileSync(join(root, '...md'), 'words\n')
    expect(() => importSkillMd(join(root, '...md'), save)).toThrow('cannot be used as a skill name')
    expect(existsSync(join(save, 'skills'))).toBe(false)
    cleanup()
  })

  it('removes a listed skill and refuses anything else', () => {
    const { save, cleanup } = scaffold({ mine: SKILL('mine', 'docx') })
    writeFileSync(join(save, 'stray.txt'), 'x')
    expect(() => removeSkill('../stray.txt', save)).toThrow(/unknown skill/)
    expect(() => removeSkill('absent', save)).toThrow(/unknown skill/)
    removeSkill('mine', save)
    expect(existsSync(join(save, 'skills', 'mine'))).toBe(false)
    expect(existsSync(join(save, 'stray.txt'))).toBe(true)
    cleanup()
  })

  it.skipIf(process.platform === 'win32')(
    'removes a folder-imported skill whose name a zip would refuse',
    async () => {
      const { save, src, cleanup } = scaffold({}, { 'draft:v2': SKILL('draft', 'docx') })
      importSkillDir(join(src, 'draft:v2'), save)
      expect((await usableSkills(save)).map((s) => s.name)).toEqual(['draft:v2'])
      removeSkill('draft:v2', save)
      expect(existsSync(join(save, 'skills', 'draft:v2'))).toBe(false)
      cleanup()
    },
  )

  it('ignores the resource fork folder macOS adds', async () => {
    const { root, save, cleanup } = scaffold()
    const zip = await zipOf(
      {
        'tool/SKILL.md': SKILL('tool', 'A pdf thing'),
        '__MACOSX/tool/._SKILL.md': 'junk',
      },
      join(root, 'a.zip'),
    )
    expect((await importSkillZip(zip, save)).name).toBe('tool')
    expect(existsSync(join(save, 'skills', 'tool', '__MACOSX'))).toBe(false)
    cleanup()
  })

  it('refuses a zip with no SKILL.md where one is expected', async () => {
    const { root, save, cleanup } = scaffold()
    const two = await zipOf(
      { 'a/SKILL.md': SKILL('a', 'docx'), 'b/SKILL.md': SKILL('b', 'pdf') },
      join(root, 'two.zip'),
    )
    await expect(importSkillZip(two, save)).rejects.toThrow('no SKILL.md')
    const none = await zipOf({ 'a/README.md': 'hi' }, join(root, 'none.zip'))
    await expect(importSkillZip(none, save)).rejects.toThrow('no SKILL.md')
    cleanup()
  })

  it('refuses an entry that climbs out, and writes nothing', async () => {
    const { root, save, cleanup } = scaffold()
    const zip = await zipOf(
      { 'tool/SKILL.md': SKILL('tool', 'docx'), '../escape.txt': 'out' },
      join(root, 'bad.zip'),
    )
    // jszip itself may fold the `..` away; either way it is refused and nothing lands
    await expect(importSkillZip(zip, save)).rejects.toThrow()
    expect(existsSync(join(root, 'escape.txt'))).toBe(false)
    expect(existsSync(join(save, 'skills', 'tool'))).toBe(false)
    cleanup()
  })

  it('refuses a symlink entry', async () => {
    const { root, save, cleanup } = scaffold()
    const z = new JSZip()
    z.file('tool/SKILL.md', SKILL('tool', 'docx'))
    z.file('tool/link', '/etc/passwd', { unixPermissions: 0o120777 })
    const zip = join(root, 'link.zip')
    writeFileSync(zip, await z.generateAsync({ type: 'nodebuffer', platform: 'UNIX' }))
    await expect(importSkillZip(zip, save)).rejects.toThrow('symlink')
    expect(existsSync(join(save, 'skills', 'tool'))).toBe(false)
    cleanup()
  })

  it('refuses a zip over the import cap and leaves no scratch behind', async () => {
    const { root, save, cleanup } = scaffold()
    const zip = await zipOf(
      { 'heavy/SKILL.md': SKILL('heavy', 'docx'), 'heavy/big.bin': Buffer.alloc(11 * 1024 * 1024) },
      join(root, 'heavy.zip'),
    )
    await expect(importSkillZip(zip, save)).rejects.toThrow('import cap')
    expect(readdirSync(join(save, 'skills'))).toEqual([])
    cleanup()
  })

  it('refuses to overwrite an installed skill', async () => {
    const { root, save, cleanup } = scaffold({ tool: SKILL('tool', 'Mine', 'mine wins') })
    const zip = await zipOf({ 'tool/SKILL.md': SKILL('tool', 'Theirs') }, join(root, 't.zip'))
    await expect(importSkillZip(zip, save)).rejects.toThrow('already there')
    expect(readSkillBody(join(save, 'skills', 'tool', 'SKILL.md'))).toBe('mine wins\n')
    cleanup()
  })

  it('refuses a file that is not a zip', async () => {
    const { root, save, cleanup } = scaffold()
    const file = join(root, 'not.zip')
    writeFileSync(file, 'plain text')
    await expect(importSkillZip(file, save)).rejects.toThrow('not a zip')
    cleanup()
  })
})

describe('writeSkill', () => {
  it('creates a skill folder with the built SKILL.md and lists it', async () => {
    const { save, cleanup } = scaffold()
    const skill = writeSkill(save, 'meeting-notes', {
      description: 'Turn notes into a docx report',
      body: '1. Read the notes.\n2. Write the report.\n',
    })
    expect(skill.name).toBe('meeting-notes')
    expect(skill.path).toBe(join(save, 'skills', 'meeting-notes', 'SKILL.md'))
    expect(skill.description).toBe('Turn notes into a docx report')
    expect(readFileSync(skill.path, 'utf8')).toBe(
      '---\nname: meeting-notes\ndescription: Turn notes into a docx report\n---\n\n1. Read the notes.\n2. Write the report.\n',
    )
    expect(readSkillBody(skill.path)).toBe('1. Read the notes.\n2. Write the report.\n')
    expect((await usableSkills(save)).map((s) => s.name)).toEqual(['meeting-notes'])
    cleanup()
  })

  it('keeps a wrapped description on one frontmatter line', () => {
    const { save, cleanup } = scaffold()
    const skill = writeSkill(save, 'wrapped', {
      description: 'Turn notes into\na docx report',
      body: 'body',
    })
    expect(skill.description).toBe('Turn notes into a docx report')
    expect(readFileSync(skill.path, 'utf8')).toContain(
      'description: Turn notes into a docx report\n',
    )
    cleanup()
  })

  it('refuses to overwrite an existing skill without the flag, and leaves it alone', () => {
    const { save, cleanup } = scaffold({ mine: SKILL('mine', 'Mine', 'mine wins') })
    expect(() => writeSkill(save, 'mine', { description: 'Theirs', body: 'theirs' })).toThrow(
      'already there',
    )
    expect(readSkillBody(join(save, 'skills', 'mine', 'SKILL.md'))).toBe('mine wins\n')
    cleanup()
  })

  it('overwrites with the flag, keeping the folder’s other files', () => {
    const { save, src, cleanup } = scaffold({}, { 'folder-skill': SKILL('x', 'A docx thing') })
    importSkillDir(join(src, 'folder-skill'), save)
    writeFileSync(join(save, 'skills', 'folder-skill', 'reference.md'), 'grid rules')
    const skill = writeSkill(
      save,
      'folder-skill',
      { description: 'New', body: 'new body' },
      { allowOverwrite: true },
    )
    expect(readSkillBody(skill.path)).toBe('new body')
    expect(existsSync(join(save, 'skills', 'folder-skill', 'reference.md'))).toBe(true)
    expect(readdirSync(join(save, 'skills')).filter((n) => n.startsWith('.'))).toEqual([])
    cleanup()
  })

  it('refuses names a created folder cannot carry', () => {
    const { save, cleanup } = scaffold()
    for (const bad of [
      '',
      'a/b',
      'a\\b',
      'with:colon',
      'ends?',
      '.hidden',
      '..',
      'ends.',
      'ends ',
      ' pad',
    ]) {
      expect(() => writeSkill(save, bad, { description: 'd', body: 'b' })).toThrow(
        'cannot be used as a skill name',
      )
    }
    expect(existsSync(join(save, 'skills'))).toBe(false)
    cleanup()
  })

  it('requires a description and caps the size', () => {
    const { save, cleanup } = scaffold()
    expect(() => writeSkill(save, 'no-desc', { description: '  ', body: 'b' })).toThrow(
      'a description is required',
    )
    expect(() =>
      writeSkill(save, 'heavy', {
        description: 'd',
        body: Buffer.alloc(11 * 1024 * 1024, 97).toString(),
      }),
    ).toThrow('import cap')
    expect(existsSync(join(save, 'skills'))).toBe(false)
    cleanup()
  })

  it('leaves no scratch behind when the final rename fails', () => {
    const { save, cleanup } = scaffold()
    const dest = join(save, 'skills', 'fresh')
    h.failRenameTo = dest
    try {
      expect(() => writeSkill(save, 'fresh', { description: 'd', body: 'b' })).toThrow('disk full')
    } finally {
      h.failRenameTo = ''
    }
    expect(readdirSync(join(save, 'skills'))).toEqual([])
    cleanup()
  })

  it('restores the previous copy when an overwrite swap fails half-way', () => {
    const { save, cleanup } = scaffold({ mine: SKILL('mine', 'Mine', 'mine wins') })
    const dest = join(save, 'skills', 'mine')
    h.failRenameTo = dest
    try {
      expect(() =>
        writeSkill(save, 'mine', { description: 'New', body: 'theirs' }, { allowOverwrite: true }),
      ).toThrow('disk full')
    } finally {
      h.failRenameTo = ''
    }
    expect(readSkillBody(join(dest, 'SKILL.md'))).toBe('mine wins\n')
    expect(readdirSync(join(save, 'skills')).filter((n) => n.startsWith('.'))).toEqual([])
    cleanup()
  })
})

describe('skillRawText', () => {
  it('returns the whole SKILL.md of a listed skill, frontmatter included', async () => {
    const { save, cleanup } = scaffold({ mine: SKILL('mine', 'A docx thing', 'Step one.') })
    expect(await skillRawText(save, 'mine')).toBe(SKILL('mine', 'A docx thing', 'Step one.'))
    cleanup()
  })

  it('refuses names the listing never showed', async () => {
    const { save, root, cleanup } = scaffold({ mine: SKILL('mine', 'A docx thing') })
    writeFileSync(join(root, 'secret.md'), 'not a skill')
    for (const bad of ['nope', '../secret', 'a/b', '.hidden', '', 'mine.md']) {
      await expect(skillRawText(save, bad)).rejects.toThrow('unknown skill')
    }
    expect(readFileSync(join(root, 'secret.md'), 'utf8')).toBe('not a skill')
    cleanup()
  })
})

describe('isKnownSkillPath', () => {
  it('accepts every path the listing hands out', async () => {
    const { save, cleanup } = scaffold({ mine: SKILL('mine', 'A docx thing') })
    for (const skill of await usableSkills(save)) {
      expect(isKnownSkillPath(skill.path, skillsRoot(save))).toBe(true)
    }
    cleanup()
  })

  it('rejects a file outside the skills root', () => {
    expect(isKnownSkillPath('/etc/passwd', '/save/skills')).toBe(false)
    expect(isKnownSkillPath('/Users/me/.ssh/id_rsa', '/save/skills')).toBe(false)
  })

  it('rejects another file sitting beside a SKILL.md', () => {
    expect(isKnownSkillPath('/save/skills/mine/notes.md', '/save/skills')).toBe(false)
    expect(isKnownSkillPath('/save/skills/mine/scripts/evil.sh', '/save/skills')).toBe(false)
  })

  it('rejects a SKILL.md reached by climbing out of the skills root', () => {
    expect(isKnownSkillPath('/save/skills/../../.ssh/SKILL.md', '/save/skills')).toBe(false)
    expect(isKnownSkillPath('/save/skills/../../../../../../etc/SKILL.md', '/save/skills')).toBe(
      false,
    )
  })

  it('accepts a climb that lands back inside the root: it is the same skill', () => {
    expect(isKnownSkillPath('/save/skills/../skills/x/SKILL.md', '/save/skills')).toBe(true)
    expect(isKnownSkillPath('/save/skills/a/../b/SKILL.md', '/save/skills')).toBe(true)
  })

  it('rejects a root that merely starts with the same characters', () => {
    expect(isKnownSkillPath('/save/skills-backup/x/SKILL.md', '/save/skills')).toBe(false)
  })

  it('rejects a skills folder with the same last component but a different parent', () => {
    expect(isKnownSkillPath('/elsewhere/skills/mine/SKILL.md', '/save/skills')).toBe(false)
    expect(isKnownSkillPath('/Users/me/.claude/skills/mine/SKILL.md', '/save/skills')).toBe(false)
  })

  it('rejects anything that is not an absolute path', () => {
    for (const bad of ['', 'skills/mine/SKILL.md', undefined, null, 42, {}]) {
      expect(isKnownSkillPath(bad, '/save/skills')).toBe(false)
    }
  })

  it('normalizes the incoming path and the root', () => {
    expect(isKnownSkillPath('/save/./skills/mine/SKILL.md', '/save/skills')).toBe(true)
    expect(isKnownSkillPath('/save/skills/mine/SKILL.md', '/save/./skills/')).toBe(true)
  })
})

const VERSIONED = (name: string, version: string, body = 'Do the thing.') =>
  `---\nname: ${name}\ndescription: Write a docx report\nmetadata:\n  version: ${version}\n---\n\n${body}\n`

describe('recommendedSkills / installRecommendedSkill', () => {
  it('reads the version from the metadata block, empty when absent', () => {
    const { root, cleanup } = scaffold()
    writeFileSync(join(root, 'a.md'), VERSIONED('a', '1.2.0'))
    writeFileSync(join(root, 'b.md'), SKILL('b', 'docx'))
    expect(skillVersion(join(root, 'a.md'))).toBe('1.2.0')
    expect(skillVersion(join(root, 'b.md'))).toBe('')
    expect(skillVersion(join(root, 'missing.md'))).toBe('')
    cleanup()
  })

  it('lists the shipped skills with missing / installed / outdated against our folder', async () => {
    const { root, save, cleanup } = scaffold({
      same: VERSIONED('same', '1.0.0'),
      old: VERSIONED('old', '0.9.0', 'user edited'),
    })
    const rec = join(root, 'recommended')
    for (const [n, v] of [
      ['same', '1.0.0'],
      ['old', '1.0.0'],
      ['fresh', '1.0.0'],
    ]) {
      mkdirSync(join(rec, n), { recursive: true })
      writeFileSync(join(rec, n, 'SKILL.md'), VERSIONED(n, v))
    }
    const list = await recommendedSkills(rec, save)
    expect(list.map((s) => [s.name, s.state])).toEqual([
      ['fresh', 'missing'],
      ['old', 'outdated'],
      ['same', 'installed'],
    ])
    // listing copies nothing and leaves the user's edited copy alone
    expect((await usableSkills(save)).map((s) => s.name)).toEqual(['old', 'same'])
    expect(readSkillBody(join(save, 'skills', 'old', 'SKILL.md'))).toBe('user edited\n')
    cleanup()
  })

  it('returns nothing when the recommended directory is absent', async () => {
    const { root, save, cleanup } = scaffold()
    expect(await recommendedSkills(join(root, 'nowhere'), save)).toEqual([])
    cleanup()
  })

  it('installs by name only from the recommended directory, replacing an outdated copy', async () => {
    const { root, save, cleanup } = scaffold({ old: VERSIONED('old', '0.9.0', 'user edited') })
    const rec = join(root, 'recommended')
    mkdirSync(join(rec, 'old', 'reference'), { recursive: true })
    writeFileSync(join(rec, 'old', 'SKILL.md'), VERSIONED('old', '1.0.0', 'shipped'))
    writeFileSync(join(rec, 'old', 'reference', 'notes.md'), 'extra')
    const installed = installRecommendedSkill('old', rec, save)
    expect(installed.name).toBe('old')
    expect(readSkillBody(installed.path)).toBe('shipped\n')
    expect(existsSync(join(save, 'skills', 'old', 'reference', 'notes.md'))).toBe(true)
    expect((await recommendedSkills(rec, save))[0]?.state).toBe('installed')
    expect(readdirSync(join(save, 'skills')).filter((n) => n.startsWith('.'))).toEqual([])
    cleanup()
  })

  it('keeps the installed copy when the swap fails half-way', () => {
    const { root, save, cleanup } = scaffold({ old: VERSIONED('old', '0.9.0', 'user edited') })
    const rec = join(root, 'recommended')
    mkdirSync(join(rec, 'old'), { recursive: true })
    writeFileSync(join(rec, 'old', 'SKILL.md'), VERSIONED('old', '1.0.0', 'shipped'))
    const dest = join(save, 'skills', 'old')
    h.failRenameTo = dest
    try {
      expect(() => installRecommendedSkill('old', rec, save)).toThrow('disk full')
    } finally {
      h.failRenameTo = ''
    }
    expect(readSkillBody(join(dest, 'SKILL.md'))).toBe('user edited\n')
    expect(readdirSync(join(save, 'skills')).filter((n) => n.startsWith('.'))).toEqual([])
    cleanup()
  })

  it('a locked leftover does not fail the install, and the leftover is not listed as a skill', async () => {
    const { root, save, cleanup } = scaffold({ old: VERSIONED('old', '0.9.0', 'user edited') })
    const rec = join(root, 'recommended')
    mkdirSync(join(rec, 'old'), { recursive: true })
    writeFileSync(join(rec, 'old', 'SKILL.md'), VERSIONED('old', '1.0.0', 'shipped'))
    h.failRmOf = '\\.previous$'
    try {
      expect(installRecommendedSkill('old', rec, save).name).toBe('old')
    } finally {
      h.failRmOf = ''
    }
    expect(readSkillBody(join(save, 'skills', 'old', 'SKILL.md'))).toBe('shipped\n')
    const leftovers = readdirSync(join(save, 'skills')).filter((n) => n.startsWith('.'))
    expect(leftovers).toHaveLength(1)
    expect((await usableSkills(save)).map((s) => s.name)).toEqual(['old'])
    cleanup()
  })

  it('a locked scratch copy still restores the parked skill when the swap fails', () => {
    const { root, save, cleanup } = scaffold({ old: VERSIONED('old', '0.9.0', 'user edited') })
    const rec = join(root, 'recommended')
    mkdirSync(join(rec, 'old'), { recursive: true })
    writeFileSync(join(rec, 'old', 'SKILL.md'), VERSIONED('old', '1.0.0', 'shipped'))
    const dest = join(save, 'skills', 'old')
    h.failRenameTo = dest
    h.failRmOf = '\\.install-[^/\\\\]+$'
    try {
      expect(() => installRecommendedSkill('old', rec, save)).toThrow('disk full')
    } finally {
      h.failRenameTo = ''
      h.failRmOf = ''
    }
    expect(readSkillBody(join(dest, 'SKILL.md'))).toBe('user edited\n')
    cleanup()
  })

  it('refuses a name that is not a shipped skill, including path-shaped ones', () => {
    const { root, save, cleanup } = scaffold()
    const rec = join(root, 'recommended')
    mkdirSync(rec, { recursive: true })
    for (const bad of ['nope', '../save', '/etc', '.hidden', 'a/b']) {
      expect(() => installRecommendedSkill(bad, rec, save)).toThrow('unknown skill')
    }
    expect(existsSync(join(save, 'skills'))).toBe(false)
    cleanup()
  })
})
