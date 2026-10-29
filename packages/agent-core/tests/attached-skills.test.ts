import { describe, expect, it } from 'vitest'
import { expandSkillBody, skillContextVars, withAttachedSkills } from '../src/attached-skills'

describe('withAttachedSkills', () => {
  it('leaves the instruction alone when nothing is attached', () => {
    expect(withAttachedSkills('Make a table', [])).toBe('Make a table')
  })

  it('puts every skill body in front of the request, each in its own tagged block', () => {
    const out = withAttachedSkills('Make a table', [
      { name: 'tables', body: 'Always add a header row.\n' },
      { name: 'tone', body: 'Write plainly.' },
    ])
    expect(out.indexOf('<skill name="tables">')).toBeLessThan(out.indexOf('<skill name="tone">'))
    expect(out).toContain('<skill name="tables">\nAlways add a header row.\n</skill>')
    expect(out).toContain('<skill name="tone">\nWrite plainly.\n</skill>')
    expect(out.endsWith('User request:\nMake a table')).toBe(true)
    expect(out).toContain('the following skills to this request')
  })

  it('cannot have its tag broken by a quote in the name', () => {
    const out = withAttachedSkills('x', [{ name: 'a"b', body: 'c' }])
    expect(out).toContain('<skill name="ab">')
  })

  it('substitutes $ARGUMENTS with the user request, everywhere it appears', () => {
    const out = withAttachedSkills('Q3 data', [
      {
        name: 'report',
        body: 'Summarize $ARGUMENTS into a table. Mention $ARGUMENTS in the title.',
      },
    ])
    expect(out).toContain('Summarize Q3 data into a table. Mention Q3 data in the title.')
    expect(out.endsWith('User request:\nQ3 data')).toBe(true)
  })

  it('leaves bodies without $ARGUMENTS verbatim', () => {
    const out = withAttachedSkills('x', [{ name: 'tone', body: 'Write plainly.' }])
    expect(out).toContain('<skill name="tone">\nWrite plainly.\n</skill>')
  })
})

describe('expandSkillBody', () => {
  it('expands positional arguments with shell-style quoting', () => {
    expect(expandSkillBody('Move $0 to $1', '"Q3 report" final')).toBe('Move Q3 report to final')
    expect(expandSkillBody('Use $ARGUMENTS[1]', 'a b')).toBe('Use b')
  })

  it('keeps a placeholder with no argument as literal text', () => {
    expect(expandSkillBody('Move $0 to $2', 'only')).toBe('Move only to $2')
  })

  it('expands ${VARS} and leaves unknown names literal', () => {
    const out = expandSkillBody('${FILE_NAME} on ${TODAY}; ${NOPE}', 'x', {
      FILE_NAME: 'Q3.docx',
      TODAY: '2026/10/29',
    })
    expect(out).toBe('Q3.docx on 2026/10/29; ${NOPE}')
  })

  it('escapes \\$ to a literal dollar sign before a placeholder', () => {
    expect(expandSkillBody('Price: \\$1 off, apply $0', 'now')).toBe('Price: $1 off, apply now')
  })

  it('never re-scans text an argument carries in', () => {
    expect(expandSkillBody('Echo $0', '"$ARGUMENTS and $1"')).toBe('Echo $ARGUMENTS and $1')
  })

  it('passes bodies without a dollar sign through untouched', () => {
    expect(expandSkillBody('plain', 'anything')).toBe('plain')
  })
})

describe('skillContextVars', () => {
  it('builds the standard set and formats TODAY in the UI language', () => {
    const vars = skillContextVars({
      lang: 'zh',
      app: '文档',
      fileName: 'report.docx',
      selection: 'some text',
      now: new Date(2026, 9, 29),
    })
    expect(vars.LANGUAGE).toBe('中文')
    expect(vars.APP).toBe('文档')
    expect(vars.FILE_NAME).toBe('report.docx')
    expect(vars.SELECTION).toBe('some text')
    expect(vars.TODAY).toBe(new Date(2026, 9, 29).toLocaleDateString('zh'))
  })

  it('omits absent context instead of substituting an empty string', () => {
    const vars = skillContextVars({ lang: 'en', app: 'Sheets' })
    expect('FILE_NAME' in vars).toBe(false)
    expect('SELECTION' in vars).toBe(false)
  })
})
