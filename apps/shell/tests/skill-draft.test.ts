import { describe, expect, it } from 'vitest'
import {
  SKILL_CREATOR_SYSTEM,
  SKILL_PARAMS,
  draftFieldErrors,
  draftFromInterview,
  insertSkillParam,
  interviewToUserMessage,
  legalSkillName,
  parseSkillDoc,
  skillDocFromReply,
  suggestSkillName,
} from '../src/renderer/src/skill-draft'

/**
 * The pure half of the skill creator: pulling a SKILL.md out of a model reply
 * (fenced or bare), and the deterministic draft the no-model path falls back
 * to. Everything here is a function over strings — no window, no IPC. The
 * meta-prompt asks for one fenced block; these tests hold the reply parser to
 * the loose version of that contract a real model produces.
 */

/** the document the meta-prompt asks the model for */
const SKILL_MD = (
  name = 'meeting-notes-to-docx',
  description = 'Use when the user asks to turn meeting notes into a formatted Word document.',
) =>
  `---\nname: ${name}\ndescription: ${description}\n---\n\n1. Read the notes.\n2. Write the report.\n`

describe('skillDocFromReply', () => {
  it('reads the skill out of a fenced block with prose around it', () => {
    const reply = [
      'Here is the skill you asked for:',
      '',
      '```markdown',
      SKILL_MD(),
      '```',
      '',
      'Let me know if you want any changes.',
    ].join('\n')
    const draft = skillDocFromReply(reply)
    expect(draft).not.toBeNull()
    expect(draft?.name).toBe('meeting-notes-to-docx')
    expect(draft?.description).toBe(
      'Use when the user asks to turn meeting notes into a formatted Word document.',
    )
    expect(draft?.body).toBe('1. Read the notes.\n2. Write the report.\n')
  })

  it('reads a bare document when the model skipped the fence', () => {
    const draft = skillDocFromReply(SKILL_MD())
    expect(draft?.name).toBe('meeting-notes-to-docx')
    expect(draft?.body).toContain('Read the notes')
  })

  it('trims whitespace around a bare document', () => {
    expect(skillDocFromReply(`\n${SKILL_MD()}\n`)?.name).toBe('meeting-notes-to-docx')
  })

  it('returns null when the description is missing', () => {
    const fenced = '```markdown\n---\nname: x\n---\n\nbody\n```'
    expect(skillDocFromReply(fenced)).toBeNull()
    expect(parseSkillDoc('---\nname: x\n---\n\nbody\n')).toBeNull()
    expect(parseSkillDoc('no frontmatter here\n')).toBeNull()
  })

  it('returns null for prose with no skill in it', () => {
    expect(skillDocFromReply('Sorry, I could not write the skill.')).toBeNull()
  })

  it('ignores chatter that precedes the document even when the reply starts like one', () => {
    const reply = `Sure — here it is.\n\n${'```markdown'}\n${SKILL_MD()}${'```'}`
    expect(skillDocFromReply(reply)?.name).toBe('meeting-notes-to-docx')
  })

  it('loads an installed skill back into fields', () => {
    const draft = parseSkillDoc(SKILL_MD('weekly-report', 'Turn progress into a report'))
    expect(draft).toEqual({
      name: 'weekly-report',
      description: 'Turn progress into a report',
      body: '1. Read the notes.\n2. Write the report.\n',
    })
  })
})

describe('draftFromInterview', () => {
  it('assembles a deterministic skill from the answers', () => {
    const draft = draftFromInterview({
      purpose: 'Turn meeting notes into a report',
      trigger: 'when the user asks for a report from notes',
      steps: '1. Read the notes\n- ask for missing figures\nKeep it under two pages',
      example: 'Make a report from these notes',
    })
    expect(draft.name).toBe('turn-meeting-notes-into-a-report')
    expect(draft.description).toBe('when the user asks for a report from notes')
    expect(draft.body).toContain('# Turn meeting notes into a report')
    // the step lines become numbered imperatives, list punctuation stripped
    expect(draft.body).toContain('1. Read the notes')
    expect(draft.body).toContain('2. ask for missing figures')
    expect(draft.body).toContain('3. Keep it under two pages')
    expect(draft.body).toContain('## Constraints')
    expect(draft.body).toContain('## Output')
    expect(draft.body).toContain('## Example')
  })

  it('falls back to a generic workflow when no steps were given', () => {
    const draft = draftFromInterview({ purpose: '生成周报', trigger: 't', steps: '', example: '' })
    expect(draft.body).toContain('## Workflow')
    expect(draft.name).toBe('my-skill')
    expect(draft.description).toBe('t')
    expect(draft.body).not.toContain('## Example')
  })

  it('suggests my-skill when the phrase has no latin words', () => {
    expect(suggestSkillName('把会议记录变成报告')).toBe('my-skill')
    expect(suggestSkillName('Weekly Report')).toBe('weekly-report')
    expect(suggestSkillName('a'.repeat(60))).toHaveLength(40)
  })
})

describe('legalSkillName / draftFieldErrors', () => {
  it('accepts folder-safe names and refuses the rest', () => {
    expect(legalSkillName('meeting-notes')).toBe(true)
    expect(legalSkillName('meeting_notes.v2')).toBe(true)
    expect(legalSkillName('周报生成')).toBe(true)
    for (const bad of [
      '',
      'a/b',
      'a\\b',
      'with:colon',
      'ends?',
      '"quoted"',
      '<tag>',
      '.hidden',
      '..',
      'ends.',
      'ends ',
      ' pad',
    ]) {
      expect(legalSkillName(bad)).toBe(false)
    }
  })

  it('flags the fields the save guard will refuse', () => {
    expect(draftFieldErrors({ name: 'ok', description: 'd', body: 'b' })).toEqual({
      name: false,
      description: false,
    })
    expect(draftFieldErrors({ name: 'a/b', description: '   ', body: 'b' })).toEqual({
      name: true,
      description: true,
    })
  })
})

describe('SKILL_PARAMS / insertSkillParam', () => {
  it('lists the vocabulary expandSkillBody replaces, each with its gloss key', () => {
    expect(SKILL_PARAMS.map((p) => p.token)).toEqual([
      '$ARGUMENTS',
      '$1',
      '${FILE_NAME}',
      '${SELECTION}',
      '${TODAY}',
      '${LANGUAGE}',
      '${APP}',
    ])
    for (const param of SKILL_PARAMS) {
      expect(param.desc.startsWith('intgSkillParam')).toBe(true)
    }
  })

  it('splices the token in at the caret and reports the caret after it', () => {
    expect(insertSkillParam('Step one.', '$ARGUMENTS', 5)).toEqual({
      body: 'Step $ARGUMENTSone.',
      caret: 5 + '$ARGUMENTS'.length,
    })
    expect(insertSkillParam('Report for ', '${TODAY}', 11).body).toBe('Report for ${TODAY}')
  })

  it('clamps a caret outside the body to the nearest end', () => {
    expect(insertSkillParam('ab', '$1', 99).body).toBe('ab$1')
    expect(insertSkillParam('ab', '$1', -3).body).toBe('$1ab')
  })
})

describe('interviewToUserMessage / SKILL_CREATOR_SYSTEM', () => {
  it('labels the answers and drops the empty optional one', () => {
    const message = interviewToUserMessage({
      purpose: 'a purpose',
      trigger: 'a trigger',
      steps: 'a step',
      example: '',
    })
    expect(message).toContain('## Purpose — what the skill should do\na purpose')
    expect(message).toContain('## When to use — becomes the description\na trigger')
    expect(message).toContain('## Steps and constraints — one step per line\na step')
    expect(message).not.toContain('Example')
  })

  it('asks the model for one fenced block, a kebab-case name and a when-to-use description', () => {
    expect(SKILL_CREATOR_SYSTEM).toContain('exactly one fenced code block')
    expect(SKILL_CREATOR_SYSTEM).toContain('`name` is kebab-case')
    expect(SKILL_CREATOR_SYSTEM).toContain('relevance')
  })

  it('teaches the parameters the app replaces at send time', () => {
    expect(SKILL_CREATOR_SYSTEM).toContain('$ARGUMENTS')
    expect(SKILL_CREATOR_SYSTEM).toContain('$ARGUMENTS[N]')
    expect(SKILL_CREATOR_SYSTEM).toContain('${FILE_NAME}')
    expect(SKILL_CREATOR_SYSTEM).toContain('${SELECTION}')
    expect(SKILL_CREATOR_SYSTEM).toContain('${TODAY}')
    expect(SKILL_CREATOR_SYSTEM).toContain('${LANGUAGE}')
    expect(SKILL_CREATOR_SYSTEM).toContain('${APP}')
  })
})
