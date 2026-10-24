import { describe, expect, it } from 'vitest'
import { decideNaming, type NamingContext } from '../src/main/naming-policy'

const base: NamingContext = {
  enabled: true,
  trigger: 'first-save',
  filePath: null,
  content: '# 季度营收总结\n\n本季度总营收同比增长 12%。',
  hasProvider: true,
}

describe('decideNaming', () => {
  it('names an untitled document on its first save', () => {
    expect(decideNaming(base)).toBe('name')
  })

  it('does not rename a file that already has a path', () => {
    // the important one: a document already on disk keeps its name
    expect(decideNaming({ ...base, filePath: '/a/note.md' })).toBe('skip-already-named')
  })

  it('does nothing when the preference is off', () => {
    expect(decideNaming({ ...base, enabled: false })).toBe('skip-disabled')
  })

  it('still allows a manual rename while the preference is off', () => {
    // the reader pressed the button — that is the whole point of the button
    expect(
      decideNaming({ ...base, enabled: false, trigger: 'manual', filePath: '/a/note.md' }),
    ).toBe('name')
  })

  it('allows a manual rename of a file that already has a path', () => {
    expect(decideNaming({ ...base, trigger: 'manual', filePath: '/a/note.md' })).toBe('name')
  })

  it('refuses to name a blank document', () => {
    expect(decideNaming({ ...base, content: '' })).toBe('skip-empty')
    expect(decideNaming({ ...base, content: '   \n\n  ' })).toBe('skip-empty')
    // frontmatter fences and an empty task list carry no words to name from
    expect(decideNaming({ ...base, content: '---\n---\n\n- [ ] \n' })).toBe('skip-empty')
  })

  it('names a document that is mostly code or a diagram', () => {
    const code = '```js\nexport function main() {\n  return 42\n}\n```'
    expect(decideNaming({ ...base, content: code })).toBe('name')
    const svg = '<svg viewBox="0 0 24 24"><path d="M0 0h24v24H0z"/></svg>'
    expect(decideNaming({ ...base, content: svg })).toBe('name')
  })

  it('does nothing when no provider key is configured', () => {
    expect(decideNaming({ ...base, hasProvider: false })).toBe('skip-no-provider')
    // a manual click with no key still cannot reach a model
    expect(
      decideNaming({ ...base, hasProvider: false, trigger: 'manual', filePath: '/a/n.md' }),
    ).toBe('skip-no-provider')
  })

  it('checks the provider before the content, so a missing key is reported as such', () => {
    // an empty doc with no key: the actionable problem is the missing key
    expect(decideNaming({ ...base, content: '', hasProvider: false })).toBe('skip-no-provider')
  })
})
