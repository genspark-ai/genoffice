import { describe, expect, it } from 'vitest'
import {
  REDACT_EXT_URI,
  readRedactLabel,
  redactExtXml,
  setRedactExt,
  stripRedactExt,
} from '../src/redaction-xml'
import { XMLParser } from 'fast-xml-parser'

/**
 * The mark is a <a:extLst> inside the run's <a:rPr> (or a picture's nvPr). It has
 * to survive a save, be readable on reopen, and be removable again — and it must
 * never take the words with it.
 */

const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  trimValues: false,
  parseTagValue: false,
})

const RPR = '<a:rPr lang="en-US" sz="1800"><a:latin typeface="Calibri"/></a:rPr>'
const RPR_SELF = '<a:rPr lang="en-US"/>'
const NVPR = '<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>'

describe('the withheld-span marker in OOXML', () => {
  it('writes into a paired rPr and leaves everything else alone', () => {
    const out = setRedactExt(RPR, 'client phone')
    expect(out).toContain('lang="en-US"')
    expect(out).toContain('sz="1800"')
    expect(out).toContain('<a:latin typeface="Calibri"/>')
    expect(out).toContain(`<a:ext uri="${REDACT_EXT_URI}">`)
    expect(out).toContain('w:label="client phone"')
  })

  it('gives a self-closing rPr the paired form the mark needs', () => {
    const out = setRedactExt(RPR_SELF, 'client phone')
    expect(out).toContain('lang="en-US"')
    expect(out).toContain('</a:rPr>')
    expect(out).toContain('go:redact')
    // exactly one open, one close
    expect(out.match(/<a:rPr\b/g)).toHaveLength(1)
    expect(out.match(/<\/a:rPr>/g)).toHaveLength(1)
  })

  it('works on a picture nvPr, which is where a media mark lives', () => {
    const out = setRedactExt(NVPR, 'training recording')
    expect(out).toContain('picLocks')
    expect(out).toContain('w:label="training recording"')
    expect(out).toContain('</p:cNvPicPr>')
  })

  it('replaces rather than duplicates when applied twice', () => {
    const once = setRedactExt(RPR, 'first label')
    const twice = setRedactExt(once, 'second label')
    expect(twice.match(/go:redact/g)).toHaveLength(1)
    expect(twice).toContain('w:label="second label"')
    expect(twice).not.toContain('first label')
  })

  it('joins an existing extLst rather than nesting a second one', () => {
    const withOther =
      '<a:rPr lang="en-US"><a:extLst><a:ext uri="{OTHER}"><x:thing/></a:ext></a:extLst></a:rPr>'
    const out = setRedactExt(withOther, 'client phone')
    expect(out.match(/<a:extLst>/g)).toHaveLength(1)
    expect(out).toContain('{OTHER}')
    expect(out).toContain('go:redact')
  })

  it('reads the label back off the parse tree', () => {
    const xml = `<a:r>${setRedactExt(RPR, '客户电话')}<a:t>13800138000</a:t></a:r>`
    const run = parser.parse(xml)['a:r']
    expect(readRedactLabel(run['a:rPr'])).toBe('客户电话')
  })

  it('reads a label written through a self-closing rPr too', () => {
    const xml = `<a:r>${setRedactExt(RPR_SELF, 'client phone')}<a:t>x</a:t></a:r>`
    expect(readRedactLabel(parser.parse(xml)['a:r']['a:rPr'])).toBe('client phone')
  })

  it('reports nothing for a node with no mark', () => {
    expect(readRedactLabel(parser.parse(RPR)['a:rPr'])).toBeUndefined()
    expect(readRedactLabel(undefined)).toBeUndefined()
  })

  it('ignores another extension in the same list', () => {
    const foreign =
      '<a:rPr><a:extLst><a:ext uri="{SOMEONE-ELSE}"><go:redact w:label="decoy"/></a:ext></a:extLst></a:rPr>'
    expect(readRedactLabel(parser.parse(foreign)['a:rPr'])).toBeUndefined()
  })

  it('round-trips: write, parse, re-write, and the label is stable', () => {
    const first = setRedactExt(RPR, '客户电话')
    const label = readRedactLabel(parser.parse(`<a:r>${first}</a:r>`)['a:r']['a:rPr'])
    const second = setRedactExt(first, label!)
    expect(second).toBe(first)
  })

  it('strips the mark and leaves the rest of the node intact', () => {
    const marked = setRedactExt(RPR, 'client phone')
    const clean = stripRedactExt(marked)
    expect(clean).not.toContain('go:redact')
    expect(clean).not.toContain('extLst')
    expect(clean).toContain('lang="en-US"')
    expect(clean).toContain('<a:latin typeface="Calibri"/>')
  })

  it('leaves a foreign extension alone when stripping', () => {
    const mixed = '<a:rPr><a:extLst><a:ext uri="{OTHER}"><x:thing/></a:ext></a:extLst></a:rPr>'
    const marked = setRedactExt(mixed, 'client phone')
    const clean = stripRedactExt(marked)
    expect(clean).toContain('{OTHER}')
    expect(clean).toContain('x:thing')
    expect(clean).not.toContain('go:redact')
  })

  it('escapes a label that carries markup', () => {
    const nasty = 'a<b>&"c"'
    const out = setRedactExt(RPR, nasty)
    expect(out).toContain('&lt;b&gt;')
    expect(out).toContain('&amp;')
    expect(out).toContain('&quot;c&quot;')
    expect(out).not.toMatch(/w:label="a<b>/)
  })

  it('the element is well-formed enough for a re-parse', () => {
    const out = setRedactExt(RPR, 'client phone')
    expect(() => parser.parse(out)).not.toThrow()
    const node = parser.parse(`<a:r>${out}</a:r>`)['a:r']
    expect(node['a:rPr']['a:extLst']['a:ext']['go:redact']['@_w:label']).toBe('client phone')
  })

  it('keeps a stable uri, since a round-tripped deck has to land on the same value', () => {
    expect(REDACT_EXT_URI).toBe('{9F1B2C3D-4E5F-4A6B-8C7D-9E0F1A2B3C4D}')
    expect(redactExtXml('x')).toContain(REDACT_EXT_URI)
  })
})

describe('repeated apply/unapply', () => {
  it('does not accumulate empty extension lists over many cycles', () => {
    let xml = '<a:rPr lang="en-US" sz="1800"/>'
    for (let i = 0; i < 20; i += 1) {
      xml = stripRedactExt(setRedactExt(xml, `label ${i}`))
    }
    expect(xml).toBe('<a:rPr lang="en-US" sz="1800"/>')
  })
})
