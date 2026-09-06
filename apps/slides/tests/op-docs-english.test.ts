/**
 * The op docs are model-facing prompt material and must stay English only:
 * the system prompt and tool descriptions are English, the reply language is
 * chosen by systemSuffix, and mixed-language docs make the model mix languages
 * inside tool arguments. Full-width punctuation is caught too (it usually
 * arrives with CJK text and breaks JSON examples).
 */
import { describe, it, expect } from 'vitest'
import { readdirSync, readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { dirname, join } from 'node:path'

const dir = join(dirname(fileURLToPath(import.meta.url)), '../src/shared/prompts/ops')
const files = readdirSync(dir).filter((f) => f.endsWith('.md'))

// CJK punctuation, Hiragana/Katakana, CJK unified ideographs, full-width forms
const NON_ENGLISH_RE = /[\u3000-\u303F\u3040-\u30FF\u4E00-\u9FFF\uFF00-\uFFEF]/

describe('op docs are English only', () => {
  it('has the six group files plus the format guide', () => {
    expect(files.sort()).toEqual(
      [
        '_format.md',
        'deck.md',
        'element.md',
        'insert.md',
        'slide.md',
        'table.md',
        'text.md',
      ].sort(),
    )
  })

  for (const file of files) {
    it(`${file} contains no CJK characters or full-width punctuation`, () => {
      const lines = readFileSync(join(dir, file), 'utf8').split('\n')
      const offending = lines
        .map((line, i) => (NON_ENGLISH_RE.test(line) ? `${i + 1}: ${line}` : null))
        .filter(Boolean)
      expect(offending).toEqual([])
    })
  }
})
