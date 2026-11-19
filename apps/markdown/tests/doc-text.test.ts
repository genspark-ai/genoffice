import { describe, expect, it } from 'vitest'
import {
  parseDocText,
  serializeDocText,
  stripLegacyFencedDivs,
} from '../src/renderer/markdown/docText'

describe('stripLegacyFencedDivs', () => {
  it('removes callout fences and keeps the body', () => {
    const out = stripLegacyFencedDivs(':::callout {type="warning"}\nBe careful.\n:::')
    expect(out).toBe('Be careful.')
  })

  it('turns a toggle summary into a bold paragraph', () => {
    const out = stripLegacyFencedDivs(':::toggle {summary="More info"}\nHidden body.\n:::')
    expect(out).toBe('**More info**\n\nHidden body.')
  })

  it('handles a toggle summary containing a closing brace', () => {
    const out = stripLegacyFencedDivs(':::toggle {summary="a } b"}\nBody.\n:::')
    expect(out).toBe('**a } b**\n\nBody.')
  })

  it('handles nested legacy divs', () => {
    const out = stripLegacyFencedDivs(':::callout\n:::toggle {summary="S"}\nbody\n:::\n:::')
    expect(out).toBe('**S**\n\nbody')
  })

  it('leaves ::: lines inside fenced code blocks alone', () => {
    const md = '```\n:::callout\n:::\n```'
    expect(stripLegacyFencedDivs(md)).toBe(md)
  })

  it('leaves unrelated ::: text alone', () => {
    const md = 'a line with ::: in the middle\n:::\nplain'
    expect(stripLegacyFencedDivs(md)).toBe(md)
  })

  it('passes plain documents through unchanged', () => {
    const md = '# Title\n\nHello **world**\n\n- item'
    expect(stripLegacyFencedDivs(md)).toBe(md)
  })
})

describe('parseDocText', () => {
  it('plain document without frontmatter', () => {
    const doc = parseDocText('# Title\n\nHello\n')
    expect(doc.frontmatter).toBe('')
    expect(doc.body).toBe('# Title\n\nHello\n')
    expect(doc.eol).toBe('\n')
    expect(doc.trailingNewline).toBe(true)
  })

  it('extracts frontmatter including fences and following blank lines', () => {
    const doc = parseDocText('---\ntitle: Test\n---\n\n# Title\n')
    expect(doc.frontmatter).toBe('---\ntitle: Test\n---\n\n')
    expect(doc.body).toBe('# Title\n')
  })

  it('frontmatter closed at EOF without newline', () => {
    const doc = parseDocText('---\ntitle: x\n---')
    expect(doc.frontmatter).toBe('---\ntitle: x\n---')
    expect(doc.body).toBe('')
    expect(doc.trailingNewline).toBe(false)
  })

  it('a --- inside body is not frontmatter when the file does not start with a fence', () => {
    const doc = parseDocText('# Title\n\n---\n\nmore\n')
    expect(doc.frontmatter).toBe('')
  })

  it('an unclosed opening fence stays in the body', () => {
    const doc = parseDocText('---\ntitle: x\nno close here\n')
    expect(doc.frontmatter).toBe('')
    expect(doc.body).toBe('---\ntitle: x\nno close here\n')
  })

  it('a fence-prefixed line (---x) does not close the frontmatter', () => {
    const doc = parseDocText('---\na: 1\n---x\nb: 2\n---\nbody\n')
    expect(doc.frontmatter).toBe('---\na: 1\n---x\nb: 2\n---\n')
    expect(doc.body).toBe('body\n')
  })

  it('detects CRLF and missing trailing newline', () => {
    const doc = parseDocText('# Hi\r\n\r\nthere')
    expect(doc.eol).toBe('\r\n')
    expect(doc.body).toBe('# Hi\n\nthere')
    expect(doc.trailingNewline).toBe(false)
  })

  it('empty file', () => {
    const doc = parseDocText('')
    expect(doc.frontmatter).toBe('')
    expect(doc.body).toBe('')
    expect(doc.trailingNewline).toBe(true)
  })

  it('strips a UTF-8 BOM before frontmatter detection and remembers it', () => {
    const doc = parseDocText('﻿---\ntitle: x\n---\n\nBody\n')
    expect(doc.bom).toBe(true)
    expect(doc.frontmatter).toBe('---\ntitle: x\n---\n\n')
    expect(doc.body).toBe('Body\n')
  })
})

describe('serializeDocText', () => {
  it('round-trips an unchanged document byte-for-byte', () => {
    const samples = [
      '# Title\n\nHello\n',
      '---\ntitle: Test\ntags: [a, b]\n---\n\n# Title\n\nBody\n',
      '---\ntitle: x\n---',
      '# Hi\r\n\r\n- a\r\n- b\r\n',
      '﻿---\ntitle: bom\n---\n\nBody\n',
      '',
    ]
    for (const raw of samples) {
      const doc = parseDocText(raw)
      expect(serializeDocText(doc, doc.body)).toBe(raw)
    }
  })

  it('appends the trailing newline to a new body when the original had one', () => {
    const doc = parseDocText('old\n')
    expect(serializeDocText(doc, 'new body')).toBe('new body\n')
  })

  it('keeps the file newline-free at EOF when the original was', () => {
    const doc = parseDocText('old')
    expect(serializeDocText(doc, 'new body\n')).toBe('new body')
  })

  it('re-applies CRLF endings to the whole file', () => {
    const doc = parseDocText('---\na: 1\r\n---\r\nbody\r\n')
    const out = serializeDocText(doc, 'body\n')
    expect(out).toBe('---\r\na: 1\r\n---\r\nbody\r\n')
  })
})
