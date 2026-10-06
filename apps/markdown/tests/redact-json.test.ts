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
      doc([para([text('请拨打 '), marked(SECRET, '客户电话'), text(' 确认')])]),
    )
    expect(dump(out)).not.toContain(SECRET)
    expect(dump(out)).toContain('{{客户电话}}')
  })

  it('keeps the prose around it', () => {
    const out = redactJson(
      doc([para([text('请拨打 '), marked(SECRET, '客户电话'), text(' 确认')])]),
    )
    const flat = dump(out)
    expect(flat).toContain('请拨打')
    expect(flat).toContain('确认')
  })

  it('leaves an unmarked document byte-identical', () => {
    const source = doc([para([text('nothing secret here')])])
    expect(dump(redactJson(source))).toBe(dump(source))
  })

  it('handles several spans in one paragraph', () => {
    const out = redactJson(
      doc([
        para([
          text('拨打 '),
          marked('13800138000', '客户电话'),
          text(' 或寄到 '),
          marked('上海市浦东新区', '收货地址'),
        ]),
      ]),
    )
    const flat = dump(out)
    expect(flat).toContain('{{客户电话}}')
    expect(flat).toContain('{{收货地址}}')
    expect(flat).not.toContain('13800138000')
    expect(flat).not.toContain('上海市浦东新区')
  })

  it('handles a span split across several text nodes', () => {
    // the mark can land on more than one node; every piece has to go
    const out = redactJson(
      doc([para([text('凭 '), marked('身份证', '证件'), text('号 '), marked('310101', '号码')])]),
    )
    const flat = dump(out)
    expect(flat).toContain('{{证件}}')
    expect(flat).toContain('{{号码}}')
    expect(flat).not.toContain('310101')
  })

  it('removes only the withholding mark, keeping the reader’s other marks', () => {
    const boldMarked = text(SECRET, [
      { type: 'redaction', attrs: { label: '客户电话' } },
      { type: 'bold' },
    ])
    const flat = dump(redactJson(doc([para([boldMarked])])))
    expect(flat).toContain('{{客户电话}}')
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
          content: [{ type: 'listItem', content: [para([marked(SECRET, '电话')])] }],
        },
        { type: 'blockquote', content: [para([marked('a@b.com', '邮箱')])] },
        {
          type: 'table',
          content: [
            {
              type: 'tableRow',
              content: [{ type: 'tableCell', content: [para([marked('x', '密')])] }],
            },
          ],
        },
      ]),
    )
    const flat = dump(out)
    expect(flat).not.toContain(SECRET)
    expect(flat).not.toContain('a@b.com')
    expect(flat).toContain('{{电话}}')
    expect(flat).toContain('{{邮箱}}')
    expect(flat).toContain('{{密}}')
  })

  it('passes through shapes it does not know rather than throwing', () => {
    for (const odd of [null, 42, 'text', [], { type: 'x' }]) {
      expect(() => redactJson(odd)).not.toThrow()
    }
    expect(redactJson(null)).toBeNull()
  })

  it('does not mutate the document it was given', () => {
    const source = doc([para([marked(SECRET, '客户电话')])])
    const before = dump(source)
    redactJson(source)
    expect(dump(source)).toBe(before)
  })
})

describe('redactLabels — what the instruction lists', () => {
  it('collects one label per span', () => {
    expect(
      redactLabels(doc([para([marked('a', '电话'), text(' x '), marked('b', '地址')])])),
    ).toEqual(['电话', '地址'])
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
      doc([para([text('请拨打 '), marked(SECRET, '客户电话'), text(' 确认')])]) as never,
    )
    expect(out).toBe('请拨打 {{客户电话}} 确认')
  })

  it('never contains the secret', () => {
    const out = modelTextOf(doc([para([marked(SECRET, '电话')])]) as never)
    expect(out).not.toContain(SECRET)
    expect(out).toContain('{{电话}}')
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
      redactLabelsOf(doc([para([marked('a', '电话'), marked('b', '地址')])]) as never),
    ).toEqual(['电话', '地址'])
  })

  it('falls back to a neutral label for a blank one', () => {
    expect(redactLabelsOf(doc([para([marked('x', '')])]) as never)).toEqual(['private'])
  })
})
