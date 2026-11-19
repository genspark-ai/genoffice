/**
 * Op-docs sync inside pptx-ops: the markdown prompt docs (src/prompts/ops/*.md,
 * parsed into OP_DOCS) must track the executor registry exactly. A new op
 * without a doc block (or a block for a removed op) fails here, not in
 * production. Mirrors apps/slides/tests/op-docs*.test.ts so the package is
 * self-checking without the app harness.
 */
import { describe, it, expect } from 'vitest'
import { readdirSync, readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { OP_DOCS, OP_GROUPS, opUsage, opVocabulary } from '../src/op-docs'
import { opNames } from '../src/ops/registry'
import '../src/ops/index'

const here = dirname(fileURLToPath(import.meta.url))
const opsDir = join(here, '../src/prompts/ops')
const groupFiles = readdirSync(opsDir).filter((f) => f.endsWith('.md'))
const pending = new Set(Object.keys(OP_DOCS).filter((n) => OP_DOCS[n]!.pending))

function headings(md: string): string[] {
  return md
    .split('\n')
    .map((line) => /^### ([A-Za-z][A-Za-z0-9]*)/.exec(line)?.[1])
    .filter((n): n is string => Boolean(n))
}

describe('op-docs sync', () => {
  it('has the six group files plus the format guide', () => {
    expect(groupFiles.sort()).toEqual(
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

  it('every registered op appears in OP_DOCS', () => {
    expect(opNames().filter((n) => !OP_DOCS[n])).toEqual([])
  })

  it('every doc block matches a registered op (or a pending in-flight op)', () => {
    const registered = new Set(opNames())
    const stale = Object.keys(OP_DOCS).filter((n) => !registered.has(n) && !pending.has(n))
    expect(stale).toEqual([])
  })

  it('every registered op appears as a ### heading in the prompt docs', () => {
    const seen = new Set<string>()
    for (const file of groupFiles) {
      if (file === '_format.md') continue
      for (const name of headings(readFileSync(join(opsDir, file), 'utf8'))) seen.add(name)
    }
    expect(opNames().filter((n) => !seen.has(n))).toEqual([])
  })

  it('every prompt-doc heading matches a registered op (or pending)', () => {
    const registered = new Set(opNames())
    const stale: string[] = []
    for (const file of groupFiles) {
      if (file === '_format.md') continue
      for (const name of headings(readFileSync(join(opsDir, file), 'utf8'))) {
        if (!registered.has(name) && !pending.has(name)) stale.push(`${file}: ${name}`)
      }
    }
    expect(stale).toEqual([])
  })

  it('every block carries a compact signature and a body', () => {
    for (const [name, doc] of Object.entries(OP_DOCS)) {
      expect(doc.sig, name).toMatch(/^\{/)
      expect(doc.sig, name).not.toContain('`')
      expect(doc.body.length, name).toBeGreaterThan(20)
      expect(OP_GROUPS).toContain(doc.group)
    }
  })

  it('usage lines resolve for registered ops and stay undefined for unknown ops', () => {
    expect(opUsage('setText')).toContain('Usage: setText')
    expect(opUsage('sparkle')).toBeUndefined()
    expect(opVocabulary()).toContain('setText')
  })
})

// CJK punctuation, Hiragana/Katakana, CJK unified ideographs, full-width forms
const NON_ENGLISH_RE = /[\u3000-\u303F\u3040-\u30FF\u4E00-\u9FFF\uFF00-\uFFEF]/

describe('op docs are English only', () => {
  for (const file of groupFiles) {
    it(`${file} contains no CJK characters or full-width punctuation`, () => {
      const lines = readFileSync(join(opsDir, file), 'utf8').split('\n')
      const offending = lines
        .map((line, i) => (NON_ENGLISH_RE.test(line) ? `${i + 1}: ${line}` : null))
        .filter(Boolean)
      expect(offending).toEqual([])
    })
  }
})
