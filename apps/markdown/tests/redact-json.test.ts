import { describe, expect, it } from 'vitest'
import {
  modelTextOf,
  redactJson,
  redactLabels,
  redactLabelsOf,
} from '../src/renderer/editor/redact'

const SECRET = '13800138000'

const doc = (content: unknown[]) => ({ type: 'doc', content })
const para = (content: unknown[]) => ({ type: 'paragraph', content })
const text = (t: string, marks?: unknown) => ({
  type: 'text',
  text: t,
  ...(marks ? { marks } : {}),
})
const marked = (t: string, label: string) => text(t, [{ type: 'redaction', attrs: { label } }])

function dump(value: unknown): string {
  return JSON.stringify(value)
}

describe('redactJson — what the model is shown', () => {
  it('replaces the withheld text with its marker', () => {
    const out = redactJson(
      doc([para([text('please call '), marked(SECRET, 'client phone'), text(' confirm')])]),
    )
    expect(dump(out)).not.toContain(SECRET)
    expect(dump(out)).toContain('{{client phone}}')
  })

  it('keeps the prose around it', () => {
    const out = redactJson(
      doc([para([text('please call '), marked(SECRET, 'client phone'), text(' confirm')])]),
    )
    const flat = dump(out)
    expect(flat).toContain('please call')
    expect(flat).toContain('confirm')
  })

  it('leaves an unmarked document byte-identical', () => {
    const source = doc([para([text('nothing secret here')])])
    expect(dump(redactJson(source))).toBe(dump(source))
  })

  it('handles several spans in one paragraph', () => {
    const out = redactJson(
      doc([
        para([
          text('call '),
          marked('13800138000', 'client phone'),
          text('  or ship to '),
          marked('Pudong New Area, Shanghai', 'shipping address'),
        ]),
      ]),
    )
    const flat = dump(out)
    expect(flat).toContain('{{client phone}}')
    expect(flat).toContain('{{shipping address}}')
    expect(flat).not.toContain('13800138000')
    expect(flat).not.toContain('Pudong New Area, Shanghai')
  })

  it('handles a span split across several text nodes', () => {
    // the mark can land on more than one node; every piece has to go
    const out = redactJson(
      doc([
        para([text('see '), marked('passport', 'doc'), text(' no. '), marked('310101', 'number')]),
      ]),
    )
    const flat = dump(out)
    expect(flat).toContain('{{doc}}')
    expect(flat).toContain('{{number}}')
    expect(flat).not.toContain('310101')
  })

  it('removes only the withholding mark, keeping the reader’s other marks', () => {
    const boldMarked = text(SECRET, [
      { type: 'redaction', attrs: { label: 'client phone' } },
      { type: 'bold' },
    ])
    const flat = dump(redactJson(doc([para([boldMarked])])))
    expect(flat).toContain('{{client phone}}')
    // the reader bolded their own placeholder, so the marker stays bold; what
    // must not survive is the mark that told us to withhold the text
    expect(flat).toContain('"bold"')
    expect(flat).not.toContain('"redaction"')
  })

  it('never emits an empty marker for a blank label', () => {
    const flat = dump(redactJson(doc([para([marked(SECRET, '')])])))
    expect(flat).toContain('{{private}}')
  })

  it('reaches spans nested in lists, quotes and tables', () => {
    const out = redactJson(
      doc([
        {
          type: 'bulletList',
          content: [{ type: 'listItem', content: [para([marked(SECRET, 'phone')])] }],
        },
        { type: 'blockquote', content: [para([marked('a@b.com', 'email')])] },
        {
          type: 'table',
          content: [
            {
              type: 'tableRow',
              content: [{ type: 'tableCell', content: [para([marked('x', 'pin')])] }],
            },
          ],
        },
      ]),
    )
    const flat = dump(out)
    expect(flat).not.toContain(SECRET)
    expect(flat).not.toContain('a@b.com')
    expect(flat).toContain('{{phone}}')
    expect(flat).toContain('{{email}}')
    expect(flat).toContain('{{pin}}')
  })

  it('passes through shapes it does not know rather than throwing', () => {
    for (const odd of [null, 42, 'text', [], { type: 'x' }]) {
      expect(() => redactJson(odd)).not.toThrow()
    }
    expect(redactJson(null)).toBeNull()
  })

  it('does not mutate the document it was given', () => {
    const source = doc([para([marked(SECRET, 'client phone')])])
    const before = dump(source)
    redactJson(source)
    expect(dump(source)).toBe(before)
  })
})

describe('redactLabels — what the instruction lists', () => {
  it('collects one label per span', () => {
    expect(
      redactLabels(doc([para([marked('a', 'phone'), text(' x '), marked('b', 'address')])])),
    ).toEqual(['phone', 'address'])
  })

  it('returns nothing for a document with no spans', () => {
    expect(redactLabels(doc([para([text('plain')])]))).toEqual([])
  })
})

describe('modelTextOf — the text the model is handed', () => {
  it('writes the marker where the withheld span was, not a gap', () => {
    // a blanked span reads to the model as a missing word, which is worse than
    // useless: it looks like the document was damaged
    // named `out`, not `text`: a local would shadow the text() builder above
    const out = modelTextOf(
      doc([
        para([text('please call '), marked(SECRET, 'client phone'), text(' to confirm')]),
      ]) as never,
    )
    expect(out).toBe('please call {{client phone}} to confirm')
  })

  it('never contains the secret', () => {
    const out = modelTextOf(doc([para([marked(SECRET, 'phone')])]) as never)
    expect(out).not.toContain(SECRET)
    expect(out).toContain('{{phone}}')
  })

  it('joins a block’s children without inventing characters', () => {
    expect(modelTextOf(doc([para([text('a'), text('b')])]) as never)).toBe('ab')
  })

  it('reads nothing from a shape it does not understand', () => {
    expect(modelTextOf({} as never)).toBe('')
    expect(modelTextOf({ type: 'paragraph' } as never)).toBe('')
  })
})

describe('redactLabelsOf', () => {
  it('lists each withheld label', () => {
    expect(
      redactLabelsOf(doc([para([marked('a', 'phone'), marked('b', 'address')])]) as never),
    ).toEqual(['phone', 'address'])
  })

  it('falls back to a neutral label for a blank one', () => {
    expect(redactLabelsOf(doc([para([marked('x', '')])]) as never)).toEqual(['private'])
  })
})
