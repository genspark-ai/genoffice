/**
 * OOXML schema-order gatekeeper: in the output of mergeRPrModel/mergePPrFormat, known
 * children must appear in CT_RPr/CT_PPr sequence — wrong order makes Word show the
 * "repair" dialog.
 */
import { describe, expect, it } from 'vitest'
import {
  mergePPrFormat,
  mergeRPrModel,
  PPR_CHILD_ORDER,
  RPR_CHILD_ORDER,
  splitXmlChildren,
} from '../src/generate'

/** Assert the ranks of known children inside the fragment are non-decreasing (unknown elements skipped) */
function assertSchemaOrder(fragment: string, rootTag: string, order: readonly string[]) {
  const open = new RegExp(`^<${rootTag}(?: [^>]*)?>`).exec(fragment)?.[0]
  if (!open) return // empty or self-closing: no children to verify
  const inner = fragment.slice(open.length, fragment.length - `</${rootTag}>`.length)
  let prev = -1
  let prevName = ''
  for (const child of splitXmlChildren(inner)) {
    const rank = order.indexOf(child.name)
    if (rank === -1) continue
    expect(
      rank,
      `${child.name} appears after ${prevName}, violating the ${rootTag} schema order`,
    ).toBeGreaterThanOrEqual(prev)
    prev = rank
    prevName = child.name
  }
}

describe('rPr schema order (mergeRPrModel)', () => {
  const RAW =
    '<w:rPr><w:rFonts w:ascii="Georgia"/><w:caps/><w:strike/><w:color w:val="112233"/>' +
    '<w:sz w:val="24"/><w:u w:val="double"/><w:bdr w:val="single"/><w:lang w:val="en-US"/></w:rPr>'

  it.each([
    ['add bold', { text: 'x', bold: true }],
    ['add italic + highlight', { text: 'x', italic: true, highlight: 'yellow' }],
    ['change color + size', { text: 'x', color: 'FF0000', sizeHalfPoints: 36 }],
    [
      'add style + vertAlign',
      { text: 'x', styleId: 'Emphasis', vertAlign: 'superscript' as const },
    ],
    ['clear underline + add font', { text: 'x', underline: false, font: '微软雅黑' }],
  ])('%s keeps CT_RPr child order', (_label, run) => {
    const out = mergeRPrModel(RAW, { ...run }, false)
    assertSchemaOrder(out, 'w:rPr', RPR_CHILD_ORDER)
  })

  it('hyperlink rStyle lands first even with raw children present', () => {
    // font matches raw (a parsed run always carries font) → the rFonts group keeps its
    // original bytes
    const out = mergeRPrModel(RAW, { text: 'x', bold: true, font: 'Georgia' }, true)
    assertSchemaOrder(out, 'w:rPr', RPR_CHILD_ORDER)
    expect(out.indexOf('<w:rStyle')).toBeLessThan(out.indexOf('<w:rFonts'))
  })

  it('raw children in unusual-but-valid subsets still merge in order', () => {
    const sparse = '<w:rPr><w:vanish/><w:vertAlign w:val="subscript"/></w:rPr>'
    const out = mergeRPrModel(sparse, { text: 'x', bold: true, color: 'AA0000' }, false)
    assertSchemaOrder(out, 'w:rPr', RPR_CHILD_ORDER)
    expect(out.indexOf('<w:b/>')).toBeLessThan(out.indexOf('<w:vanish/>'))
    expect(out.indexOf('<w:vanish/>')).toBeLessThan(out.indexOf('<w:color'))
  })
})

describe('pPr schema order (mergePPrFormat)', () => {
  const RAW =
    '<w:pPr><w:pStyle w:val="Heading1"/><w:keepNext/><w:tabs><w:tab w:val="left" w:pos="420"/></w:tabs>' +
    '<w:rPr><w:b/></w:rPr></w:pPr>'

  it.each([
    ['spacing + align', { spaceBefore: 120, align: 'center' as const }],
    ['borders + shading', { borders: 'tb', shadingFill: 'EEEEEE' }],
    ['indent + pageBreak', { indentLeft: 720, pageBreakBefore: true }],
  ])('%s keeps CT_PPr child order', (_label, format) => {
    const out = mergePPrFormat(RAW, format)
    assertSchemaOrder(out, 'w:pPr', PPR_CHILD_ORDER)
    // The paragraph-mark rPr must still sit at the end (one of the highest-ranked known
    // elements inside pPr)
    expect(out).toContain('<w:rPr><w:b/></w:rPr>')
  })

  // The fresh-build branch ('<w:pPr/>', taken by the header/footer writer) has no raw
  // children to interleave against, so it has to sort on its own.
  it.each([
    ['bidi + align', { bidi: true, align: 'left' as const }],
    ['tabs + bidi', { bidi: true, tabStops: [{ pos: 720, val: 'left' as const }] }],
    [
      'framePr + tabs + bidi + align',
      {
        align: 'center' as const,
        bidi: true,
        tabStops: [{ pos: 1440, val: 'right' as const }],
        dropCap: { type: 'drop' as const, lines: 3 },
      },
    ],
  ])('%s keeps CT_PPr child order when built from the model alone', (_label, format) => {
    assertSchemaOrder(mergePPrFormat('<w:pPr/>', format), 'w:pPr', PPR_CHILD_ORDER)
  })
})
