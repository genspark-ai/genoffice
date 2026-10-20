import { describe, expect, it } from 'vitest'
import { classifySkill, parseSkillFrontmatter, SKILL_FORMAT_TOKENS } from '../src'

const SKILL = `---
name: genoffice
description: Create, convert, read and edit Office documents locally with GenOffice's command line.
metadata:
  version: 2.67.1
  cli: '>=0.9.0'
---

# genoffice

Run \`genoffice --version\` first.
`

describe('parseSkillFrontmatter', () => {
  it('reads name, description and body, and keeps unknown keys', () => {
    const out = parseSkillFrontmatter(SKILL)!
    expect(out.name).toBe('genoffice')
    expect(out.description).toContain('Office documents')
    expect(out.body.startsWith('# genoffice')).toBe(true)
    expect(out.body).not.toContain('---')
    // nested YAML is not modelled: `metadata:`'s children are kept as its text
    expect(out.extra.metadata).toContain('version: 2.67.1')
    expect(out.extra.metadata).toContain("cli: '>=0.9.0'")
  })

  it('rejects a file with no frontmatter', () => {
    expect(parseSkillFrontmatter('# just prose\n')).toBeNull()
    expect(parseSkillFrontmatter('')).toBeNull()
  })

  it('rejects frontmatter whose opening fence is missing', () => {
    // a hand-edited or concatenated file can end up with the two keys and the
    // closing rule but no opening `---`. Reading it would hand the model a
    // "skill" the author never wrote, so the leading fence is required.
    const halfOpen = [
      'name: not-really',
      'description: also not really',
      '---',
      '',
      'More prose.',
    ].join('\n')
    expect(parseSkillFrontmatter(halfOpen)).toBeNull()
  })

  it('rejects frontmatter that never closes, rather than half-reading it', () => {
    const open = '---\nname: x\ndescription: y\n\n# body with no fence\n'
    expect(parseSkillFrontmatter(open)).toBeNull()
  })

  it('rejects frontmatter missing the two keys it needs', () => {
    expect(parseSkillFrontmatter('---\nname: x\n---\n\nbody\n')).toBeNull()
    expect(parseSkillFrontmatter('---\ndescription: y\n---\n\nbody\n')).toBeNull()
  })

  it('joins a wrapped description onto one value', () => {
    const wrapped = '---\nname: x\ndescription: first line\n  second line\n---\n\nbody\n'
    expect(parseSkillFrontmatter(wrapped)!.description).toBe('first line second line')
  })

  it('strips quotes and unescapes a doubled apostrophe', () => {
    const quoted = "---\nname: 'x'\ndescription: \"it's fine\"\n---\n\nbody\n"
    expect(parseSkillFrontmatter(quoted)!.description).toBe("it's fine")
  })

  it('survives a byte-order mark, which a hand-edited file can carry', () => {
    expect(parseSkillFrontmatter('﻿' + SKILL)!.name).toBe('genoffice')
  })
})

describe('classifySkill', () => {
  it('keeps a skill that names a format GenOffice opens', () => {
    const out = classifySkill('genoffice', 'Create and convert docx, xlsx, pptx and pdf files')
    expect(out.relevant).toBe(true)
    expect(out.matched).toEqual(['docx', 'pdf', 'pptx', 'xlsx'])
  })

  it('keeps a conversion skill that names a format pair, not a product', () => {
    const out = classifySkill('pdf2html', 'Convert PDF files into clean HTML')
    expect(out.relevant).toBe(true)
    expect(out.matched).toContain('pdf')
    expect(out.matched).toContain('html')
  })

  it('keeps a skill that names only the product', () => {
    expect(classifySkill('word-report', 'Draft a Word report from notes').relevant).toBe(true)
  })

  it('drops a skill about work GenOffice does not do', () => {
    const out = classifySkill('k8s-deploy', 'Deploy to Kubernetes, roll back releases, read logs')
    expect(out.relevant).toBe(false)
    expect(out.matched).toEqual([])
  })

  it('drops a git skill even though it mentions "documents"', () => {
    // "document" is in the token list for *user-facing* phrasing, so this one is
    // the case that decides the rule: a skill that only says "documents" is
    // ambiguous, and the honest answer is to keep it and let the user decide.
    const out = classifySkill('git-helper', 'Commit and branch your repository')
    expect(out.relevant).toBe(false)
  })

  it('does not match a short extension inside a longer word', () => {
    // `md` must not fire on `cmd`, and `doc` must not fire on `doctor`
    expect(classifySkill('cmd-builder', 'Build a command palette').relevant).toBe(false)
    expect(classifySkill('clinic-bot', 'A doctor assistant for clinics').relevant).toBe(false)
  })

  it('matches an extension written with a dot or as part of a filename', () => {
    expect(classifySkill('convert', 'Turn report.docx into report.pdf').relevant).toBe(true)
    expect(classifySkill('x', 'writes .xlsx').relevant).toBe(true)
  })

  it('matches the name as well as the description', () => {
    expect(classifySkill('pptx-builder', 'nothing relevant here').relevant).toBe(true)
  })

  it('is case-insensitive', () => {
    expect(classifySkill('DOCX', 'Make a PPTX').matched).toEqual(['docx', 'pptx'])
  })

  it('reports every match once, sorted', () => {
    const out = classifySkill('s', 'pdf then docx then pdf again')
    expect(out.matched).toEqual(['docx', 'pdf'])
  })

  it('recognises every format the app itself opens, including html', () => {
    // the list is the feature's contract with the user: a skill about any of
    // these survives the filter. Deleting one token from the table must fail
    // here, so the list cannot rot into a subset.
    for (const ext of [
      'docx',
      'xlsx',
      'pptx',
      'pdf',
      'markdown',
      'html',
      'csv',
      'doc',
      'xls',
      'ppt',
    ]) {
      const out = classifySkill('a-skill', `Convert files to ${ext}`)
      expect(out.relevant, `${ext} should keep the skill`).toBe(true)
      expect(out.matched, `${ext} should be reported as the match`).toContain(ext)
    }
  })

  it('keeps a skill that only says markdown, and reports the word it matched', () => {
    // `md` left the table on purpose: as a token it fires on almost any prose
    // skill, which flattened the relevance flag into noise. The full word
    // still matches, which is what a Markdown skill actually says.
    expect(classifySkill('md-tidy', 'Normalise a Markdown file').matched).toEqual(['markdown'])
  })

  it('no longer fires on the generic words the tokens used to share', () => {
    // office / document / report / letter / cv described nearly every skill,
    // so a Git skill "documented things" and every letter template was
    // "relevant"; the flag is for file formats, not for prose about prose
    expect(classifySkill('git-flow', 'Document your release process in a report').relevant).toBe(
      false,
    )
    expect(classifySkill('cover-letter', 'Write a job application letter').relevant).toBe(false)
  })

  it('has no duplicate or empty entries in the token list', () => {
    expect(new Set(SKILL_FORMAT_TOKENS).size).toBe(SKILL_FORMAT_TOKENS.length)
    expect(SKILL_FORMAT_TOKENS.every((t) => t && t === t.toLowerCase())).toBe(true)
  })
})
