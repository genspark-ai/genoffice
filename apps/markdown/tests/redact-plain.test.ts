import { describe, expect, it } from 'vitest'

import {
  JSON_REDACT_KEY,
  clearJsonPath,
  hasPlainMark,
  jsonCanHoldMarks,
  jsonLabels,
  jsonPathAtLine,
  jsonPointer,
  markJsonPath,
  markPlainRange,
  plainLabels,
  projectJson,
  projectPlain,
} from '../src/renderer/editor/redact-plain'

/**
 * The source-text surface, where the mark has to survive in a file format that
 * has no place to put one.
 *
 * The property that matters throughout: the reader's own value is never
 * replaced. A mark records what to hide; the projection is a separate view.
 * The second property is format validity — a `.json` that a parser rejects is
 * worse than an unhidden secret, because the file is now broken.
 */

const ENV = ['DB_HOST=localhost', 'DB_PASSWORD=hunter2', 'API_KEY=sk-live-abc123'].join('\n')

describe('plain-text marks', () => {
  it('puts the mark on the line above the value, not on the value line', () => {
    const marked = markPlainRange(
      ENV,
      ENV.indexOf('hunter2'),
      ENV.indexOf('hunter2') + 7,
      'db password',
    )
    const lines = marked.split('\n')
    expect(lines[1]).toBe('// gx:redact:db password')
    expect(lines[2]).toBe('DB_PASSWORD=hunter2')
  })

  it('leaves every original line in place', () => {
    const marked = markPlainRange(ENV, 0, 0, 'x')
    for (const line of ENV.split('\n')) {
      expect(marked).toContain(line)
    }
  })

  it('reads the labels back', () => {
    const marked = markPlainRange(ENV, ENV.indexOf('hunter2'), 0, 'db password')
    expect(plainLabels(marked)).toEqual(['db password'])
    expect(hasPlainMark(marked)).toBe(true)
    expect(hasPlainMark(ENV)).toBe(false)
  })

  it('shows the model the placeholder and not the value', () => {
    const marked = markPlainRange(ENV, ENV.indexOf('hunter2'), 0, 'db password')
    const seen = projectPlain(marked)
    expect(seen).not.toContain('hunter2')
    expect(seen).toContain('{{db password}}')
    // and the mark line itself is a note to the reader, not content
    expect(seen).not.toContain('gx:redact')
  })

  it('leaves the unmarked lines exactly as they were', () => {
    const marked = markPlainRange(ENV, ENV.indexOf('hunter2'), 0, 'db password')
    // The mark line is dropped, so the projection is one line shorter — assert
    // on what survived rather than on an index that shifts.
    const seen = projectPlain(marked).split('\n')
    expect(seen).toContain('DB_HOST=localhost')
    expect(seen).toContain('API_KEY=sk-live-abc123')
    expect(seen).not.toContain('DB_PASSWORD=hunter2')
  })

  it('keeps a file that has two marks projecting both', () => {
    let text = markPlainRange(ENV, ENV.indexOf('hunter2'), 0, 'db password')
    text = markPlainRange(text, text.indexOf('sk-live'), 0, 'api key')
    const seen = projectPlain(text)
    expect(seen).toContain('{{db password}}')
    expect(seen).toContain('{{api key}}')
    expect(seen).not.toContain('hunter2')
    expect(seen).not.toContain('sk-live')
  })

  it('does not invent a value for a mark with nothing under it', () => {
    // A trailing mark covers no line. Replacing "nothing" with a placeholder
    // would add a line the reader never wrote.
    const trailing = `${ENV}\n// gx:redact:dangling`
    const seen = projectPlain(trailing).split('\n')
    expect(seen[seen.length - 2]).toBe('API_KEY=sk-live-abc123')
  })
})

describe('JSON marks', () => {
  const CONFIG = { db: { host: 'localhost', password: 'hunter2' }, retries: 3 }

  it('records a pointer and leaves every value in place', () => {
    const result = markJsonPath(CONFIG, ['db', 'password'], 'db password')
    expect(result.ok).toBe(true)
    if (!result.ok) return
    expect(result.json).toMatchObject({ db: { password: 'hunter2' } })
  })

  it('writes marks under a root key, so the file stays valid JSON', () => {
    // A `//` comment here would make the document unparseable. The whole
    // reason JSON gets its own mark form.
    const result = markJsonPath(CONFIG, ['db', 'password'], 'db password')
    expect(result.ok).toBe(true)
    if (!result.ok) return
    const json = result.json as Record<string, unknown>
    expect(json[JSON_REDACT_KEY]).toEqual([{ path: '/db/password', label: 'db password' }])
    expect(() => JSON.stringify(result.json)).not.toThrow()
  })

  it('shows the model the placeholder and not the value', () => {
    const result = markJsonPath(CONFIG, ['db', 'password'], 'db password')
    if (!result.ok) return
    const seen = projectJson(result.json)
    expect(JSON.stringify(seen)).not.toContain('hunter2')
    expect(seen).toMatchObject({ db: { host: 'localhost', password: '{{db password}}' } })
  })

  it('does not leave the mark array in the model view', () => {
    // The model should see the document's shape, not notes about itself.
    const result = markJsonPath(CONFIG, ['db', 'password'], 'db password')
    if (!result.ok) return
    expect(JSON.stringify(projectJson(result.json))).not.toContain(JSON_REDACT_KEY)
  })

  it('reads the labels back', () => {
    const result = markJsonPath(CONFIG, ['db', 'password'], 'db password')
    if (!result.ok) return
    expect(jsonLabels(result.json)).toEqual(['db password'])
  })

  it('refuses a document with no root object to hold the mark', () => {
    // A top-level array has nowhere to put a key without changing its shape.
    const result = markJsonPath([1, 2, 3], ['0'], 'x')
    expect(result.ok).toBe(false)
    expect(jsonCanHoldMarks([1, 2, 3])).toBe(false)
  })

  it('refuses to mark the same value twice', () => {
    const first = markJsonPath(CONFIG, ['db', 'password'], 'a')
    if (!first.ok) return
    expect(markJsonPath(first.json, ['db', 'password'], 'b').ok).toBe(false)
  })

  it('clears one mark without touching the others', () => {
    let json = markJsonPath(CONFIG, ['db', 'password'], 'db password')
    if (!json.ok) return
    json = markJsonPath(json.json, ['db', 'host'], 'db host')
    if (!json.ok) return
    const cleared = clearJsonPath(json.json, ['db', 'password'])
    expect(cleared.ok).toBe(true)
    if (!cleared.ok) return
    expect(jsonLabels(cleared.json)).toEqual(['db host'])
    expect(projectJson(cleared.json)).toMatchObject({ db: { password: 'hunter2' } })
  })

  it('drops the root key when the last mark goes', () => {
    const marked = markJsonPath(CONFIG, ['db', 'password'], 'x')
    if (!marked.ok) return
    const cleared = clearJsonPath(marked.json, ['db', 'password'])
    if (!cleared.ok) return
    expect(JSON_REDACT_KEY in (cleared.json as object)).toBe(false)
  })

  it('round-trips keys that contain / and ~ in a pointer', () => {
    expect(jsonPointer(['a/b', 'c~d'])).toBe('/a~1b/c~0d')
    const odd = { 'a/b': { 'c~d': 'secret' } }
    const result = markJsonPath(odd, ['a/b', 'c~d'], 'odd key')
    if (!result.ok) return
    expect(projectJson(result.json)).toMatchObject({ 'a/b': { 'c~d': '{{odd key}}' } })
  })

  it('leaves a document with no marks byte-identical', () => {
    expect(projectJson(CONFIG)).toBe(CONFIG)
  })
})

describe('a mark must address a value that exists', () => {
  /**
   * The regression the demo surfaced: a mark recorded from the `"password"`
   * line of `{"db":{"password":…}}` used to record `/password`, which matches
   * nothing. The writer then created that key at the root, so the model saw the
   * real secret *and* a value the file never had. Both directions of that are
   * failures; this pins them.
   */
  const CONFIG = { db: { host: 'localhost', password: 'hunter2' }, retries: 3 }
  const TEXT = JSON.stringify(CONFIG, null, 2) + '\n'

  it('resolves the whole path from the root, not just the key on the line', () => {
    const at = TEXT.indexOf('"password"')
    expect(jsonPathAtLine(CONFIG, TEXT, at)).toEqual(['db', 'password'])
  })

  it('resolves a top-level value', () => {
    const at = TEXT.indexOf('"retries"')
    expect(jsonPathAtLine(CONFIG, TEXT, at)).toEqual(['retries'])
  })

  it('records the resolved path, not the bare key', () => {
    const at = TEXT.indexOf('"password"')
    const path = jsonPathAtLine(CONFIG, TEXT, at)!
    const result = markJsonPath(CONFIG, path, 'db password')
    expect(result.ok).toBe(true)
    if (!result.ok) return
    expect((result.json as Record<string, unknown>)[JSON_REDACT_KEY]).toEqual([
      { path: '/db/password', label: 'db password' },
    ])
  })

  it('refuses a pointer that addresses nothing', () => {
    const result = markJsonPath(CONFIG, ['password'], 'db password')
    expect(result.ok).toBe(false)
    if (result.ok) return
    expect(result.reason).toMatch(/not in the document/)
  })

  it('withholds the value the pointer names, and leaves the rest alone', () => {
    const result = markJsonPath(CONFIG, ['db', 'password'], 'db password')
    if (!result.ok) return
    const seen = projectJson(result.json)
    expect(JSON.stringify(seen)).not.toContain('hunter2')
    expect(seen).toMatchObject({ db: { host: 'localhost', password: '{{db password}}' } })
  })

  it('never invents a key the document did not have', () => {
    // A hand-edited or older part can carry a pointer that no longer resolves.
    const damaged = { ...CONFIG, [JSON_REDACT_KEY]: [{ path: '/password', label: 'db password' }] }
    const seen = projectJson(damaged) as Record<string, unknown>
    expect('password' in seen).toBe(false)
    // and the real secret is not handed over either: the whole document is
    expect(JSON.stringify(seen)).not.toContain('hunter2')
  })

  it('keeps the document readable when every string is withheld', () => {
    const damaged = { ...CONFIG, [JSON_REDACT_KEY]: [{ path: '/nope', label: 'x' }] }
    const seen = projectJson(damaged) as Record<string, unknown>
    expect(seen.retries).toBe(3)
    expect(JSON_REDACT_KEY in seen).toBe(false)
  })
})
