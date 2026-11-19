import { describe, expect, it } from 'vitest'
import {
  modelTextOf,
  redactLabelsOf,
  redactNode,
  redactionCount,
  type RedactableNode,
} from '../src/renderer/ai/redact-view'
import {
  MAX_LABEL_LENGTH,
  checkPlaceholders,
  collectPlaceholders,
  isWholePlaceholder,
  placeholderInstruction,
  placeholderSource,
  readPlaceholderLabel,
  sanitizeLabel,
} from '../src/renderer/ai/redact'

const SECRET = '13800138000'

const text = (t: string, marks?: RedactableNode['marks']): RedactableNode =>
  marks ? { type: 'text', text: t, marks } : { type: 'text', text: t }
const marked = (t: string, label: string): RedactableNode =>
  text(t, [{ type: 'redaction', attrs: { label } }])
/** TipTap serialises a mark's type as its schema object, not the registered name */
const markedAsObject = (t: string, label: string): RedactableNode =>
  text(t, [{ type: { name: 'redaction' }, attrs: { label } }])
const para = (content: RedactableNode[]): RedactableNode => ({ type: 'paragraph', content })
const doc = (content: RedactableNode[]): RedactableNode => ({ type: 'doc', content })

describe('modelTextOf — what the model is handed', () => {
  it('writes the marker where the withheld words were', () => {
    expect(
      modelTextOf(
        doc([para([text('please call '), marked(SECRET, 'client phone'), text(' to confirm')])]),
      ),
    ).toBe('please call {{client phone}} to confirm')
  })

  it('recognises a mark whose type was serialised as an object', () => {
    // a strict `=== 'redaction'` silently matches nothing and the span leaks
    const out = modelTextOf(doc([para([markedAsObject(SECRET, 'phone')])]))
    expect(out).toBe('{{phone}}')
    expect(out).not.toContain(SECRET)
  })

  it('never contains a withheld secret', () => {
    const out = modelTextOf(doc([para([text('see '), marked('310101199001011234', 'ID number')])]))
    expect(out).not.toContain('310101199001011234')
  })

  it('does not leave a blank gap where a span was', () => {
    // a blanked span reads to the model as a missing word
    expect(modelTextOf(doc([para([text('A '), marked(SECRET, 'x'), text(' B')])]))).toBe(
      'A {{x}} B',
    )
  })

  it('does not add a trailing newline for a single block', () => {
    expect(modelTextOf(doc([para([text('a'), text('b')])]))).toBe('ab')
  })

  it('separates two blocks with a line break', () => {
    expect(modelTextOf(doc([para([text('one')]), para([text('two')])]))).toBe('one\ntwo')
  })

  it('leaves an unmarked document byte-identical', () => {
    const plain = doc([para([text('nothing secret')])])
    expect(modelTextOf(plain)).toBe('nothing secret')
  })

  it('reaches spans nested in a list, a quote and a table cell', () => {
    const tree = doc([
      {
        type: 'bulletList',
        content: [{ type: 'listItem', content: [para([marked('a1', 'A')])] }],
      },
      { type: 'blockquote', content: [para([marked('a2', 'B')])] },
      {
        type: 'table',
        content: [
          {
            type: 'tableRow',
            content: [{ type: 'tableCell', content: [para([marked('a3', 'C')])] }],
          },
        ],
      },
    ])
    const out = modelTextOf(tree)
    expect(out).toContain('{{A}}')
    expect(out).toContain('{{B}}')
    expect(out).toContain('{{C}}')
    expect(out).not.toContain('a1')
  })

  it('reads nothing from a shape it does not understand', () => {
    expect(modelTextOf({} as RedactableNode)).toBe('')
    expect(modelTextOf({ type: 'paragraph' } as RedactableNode)).toBe('')
  })

  it('keeps a bold that the reader put inside the span', () => {
    const out = modelTextOf(
      doc([
        para([text(SECRET, [{ type: 'redaction', attrs: { label: 'phone' } }, { type: 'bold' }])]),
      ]),
    )
    expect(out).toContain('{{phone}}')
  })
})

describe('withheld pictures', () => {
  it('turns a picture into a marker, so no path reaches the model', () => {
    const pic: RedactableNode = {
      type: 'docInlineImage',
      attrs: { src: 'assets/id-front.png', dataUrl: 'data:image/png;base64,AAAA' },
      marks: [{ type: 'redaction', attrs: { label: 'ID photo' } }],
    }
    const out = modelTextOf(doc([para([text('attach '), pic])]))
    expect(out).toContain('{{ID photo}}')
    expect(out).not.toContain('id-front.png')
    expect(out).not.toContain('base64')
  })

  it('leaves an ordinary picture alone', () => {
    const pic: RedactableNode = { type: 'docInlineImage', attrs: { src: 'a.png' } }
    // no text either way, but it must not become a marker
    expect(modelTextOf(doc([para([pic])]))).not.toContain('{{')
  })
})

describe('redactNode — the node-level view', () => {
  it('replaces a withheld text run with a plain marked-free text node', () => {
    const out = redactNode(doc([para([text('A '), marked(SECRET, 'phone'), text(' B')])]))
    expect(JSON.stringify(out)).toContain('{{phone}}')
    expect(JSON.stringify(out)).not.toContain(SECRET)
  })

  it('blanks a withheld picture’s path but keeps the node', () => {
    const pic: RedactableNode = {
      type: 'docInlineImage',
      attrs: { src: 'secret.png' },
      marks: [{ type: 'redaction', attrs: { label: 'ID photo' } }],
    }
    const out = redactNode(doc([para([pic])]))
    const flat = JSON.stringify(out)
    expect(flat).not.toContain('secret.png')
    expect(flat).toContain('redact')
  })

  it('does not mutate what it was given', () => {
    const source = doc([para([marked(SECRET, 'phone')])])
    const before = JSON.stringify(source)
    redactNode(source)
    expect(JSON.stringify(source)).toBe(before)
  })
})

describe('redactLabelsOf / redactionCount', () => {
  it('lists one label per span, in order', () => {
    expect(redactLabelsOf(doc([para([marked('a', 'phone'), marked('b', 'address')])]))).toEqual([
      'phone',
      'address',
    ])
  })

  it('counts spans across the tree', () => {
    expect(redactionCount(doc([para([marked('a', 'x')]), para([marked('b', 'y')])]))).toBe(2)
    expect(redactionCount(doc([para([text('plain')])]))).toBe(0)
  })
})

describe('the write guard', () => {
  const before = 'call {{client phone}} to confirm the order'

  it('accepts a reply that keeps every marker', () => {
    expect(
      checkPlaceholders(before, 'please call {{client phone}} before the end of today.'),
    ).toEqual([])
  })

  it('rejects a split across a line break', () => {
    expect(checkPlaceholders(before, 'call {{client\nphone}} to confirm').length).toBeGreaterThan(0)
  })

  it('rejects an interleaving, which an atom could never have prevented', () => {
    expect(checkPlaceholders('{{a}}{{b}}', '{{{{a}}b}}').length).toBeGreaterThan(0)
  })

  it('rejects a drop, a rename, an invention and a duplicate', () => {
    expect(checkPlaceholders('{{a}} {{b}}', '{{a}}').length).toBeGreaterThan(0)
    expect(checkPlaceholders('{{a}}', '{{a2}}').length).toBeGreaterThan(0)
    expect(checkPlaceholders('plain', 'plain {{new}}').length).toBeGreaterThan(0)
    expect(checkPlaceholders('{{a}}', '{{a}} {{a}}').length).toBeGreaterThan(0)
  })

  it('reports a half-typed marker as a split, not a mere absence', () => {
    expect(checkPlaceholders('{{a}}', '{{a}').some((i) => i.reason === 'split')).toBe(true)
    expect(checkPlaceholders('{{a}}', '{a}').some((i) => i.reason === 'split')).toBe(true)
  })

  it('does not report leftovers for a well-formed marker', () => {
    expect(checkPlaceholders('{{a}}', 'Call {{a}} now').some((i) => i.reason === 'split')).toBe(
      false,
    )
  })

  it('never throws on anything a model might produce', () => {
    for (const after of ['', '{{', '}}', '{{{{', '}}}}', '😀{{a}}😀', '{{a}}\n\n{{b}}']) {
      expect(() => checkPlaceholders('{{a}}', after)).not.toThrow()
    }
  })
})

describe('labels', () => {
  it('strips characters that would make the marker ambiguous', () => {
    expect(sanitizeLabel('a{b}<c>"d"')).toBe('abcd')
  })

  it('falls back to a neutral word', () => {
    expect(placeholderSource('')).toBe('{{private}}')
  })

  it('caps the length', () => {
    expect(sanitizeLabel('x'.repeat(100)).length).toBeLessThanOrEqual(MAX_LABEL_LENGTH)
  })

  it('recognises a whole marker but not partial text', () => {
    expect(isWholePlaceholder('{{a}}')).toBe(true)
    expect(isWholePlaceholder('call {{a}}')).toBe(false)
    expect(readPlaceholderLabel('{{client phone}}')).toBe('client phone')
    expect(readPlaceholderLabel('call {{a}}')).toBeNull()
  })
})

describe('the model instruction', () => {
  it('lists each marker this document has, once', () => {
    const text2 = placeholderInstruction(['phone', 'address', 'phone'])
    expect(text2).toContain('- {{phone}}')
    expect(text2).toContain('- {{address}}')
    expect(text2.split('\n').filter((l) => l === '- {{phone}}')).toHaveLength(1)
  })

  it('forbids each way a marker gets damaged', () => {
    const text2 = placeholderInstruction(['a'])
    expect(text2).toMatch(/never split/i)
    expect(text2).toMatch(/never merge/i)
    expect(text2).toMatch(/never rename/i)
    expect(text2).toMatch(/never drop/i)
    expect(text2).toMatch(/never add a placeholder/i)
  })

  it('says a placeholder may stand in for a picture', () => {
    expect(placeholderInstruction(['a'])).toMatch(/picture/i)
  })

  it('keeps its own wording, worked example included, in a Latin script', () => {
    // The reader's labels may be in any script, but this instruction ships to
    // every model in every language, and a worked example written in one
    // script teaches that script's conventions rather than the rule. With no
    // labels passed, every character in here is ours to keep Latin.
    const text = placeholderInstruction([])
    expect(text).not.toMatch(/\p{Script=Han}/u)
    expect(text).toContain('call {{client phone}}')
  })
})

describe('collectPlaceholders', () => {
  it('keeps duplicates', () => {
    expect(collectPlaceholders('{{a}} {{b}} {{a}}')).toEqual(['{{a}}', '{{b}}', '{{a}}'])
  })
})
