/**
 * Font dialog effects and character spacing (caps, dstrike, vanish, w:spacing, w:w,
 * w:kern, w:position) save from the run model: untouched runs keep their bytes,
 * edited fields are rebuilt, explicit off wins over a style value.
 */
import { describe, expect, it } from 'vitest'
import { parseDocx, saveDocx } from '../src/index'
import { mergeRPrModel } from '../src/generate'
import { buildDocx } from './helpers/build-docx'

const RPR =
  '<w:rPr><w:smallCaps/><w:dstrike/><w:vanish/><w:spacing w:val="40"/><w:w w:val="150"/>' +
  '<w:kern w:val="28"/><w:position w:val="6"/><w:sz w:val="24"/></w:rPr>'
const BODY = `<w:p><w:r>${RPR}<w:t>Spaced</w:t></w:r></w:p>`

describe('character spacing and effects in the run model', () => {
  it('parses every field, including explicit off values', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: BODY }))
    expect(doc.blocks[0].runs?.[0]).toMatchObject({
      caps: 'small',
      dstrike: true,
      vanish: true,
      vanishOwn: true,
      charSpacingTwips: 40,
      charScalePct: 150,
      kernHalfPoints: 28,
      positionHalfPoints: 6,
    })
    const off = await parseDocx(
      await buildDocx({
        bodyXml:
          '<w:p><w:r><w:rPr><w:caps w:val="0"/><w:dstrike w:val="0"/><w:vanish w:val="0"/></w:rPr><w:t>x</w:t></w:r></w:p>',
      }),
    )
    expect(off.blocks[0].runs?.[0]).toMatchObject({
      caps: 'none',
      dstrike: false,
      vanishOwn: false,
    })
    expect(off.blocks[0].runs?.[0].vanish).toBeUndefined()
  })

  it('keeps an untouched run byte-identical', async () => {
    const bytes = await buildDocx({ bodyXml: BODY })
    const doc = await parseDocx(bytes)
    const blocks = doc.blocks
      .filter((b) => !b.hidden && b.docxIndex !== null)
      .map((b) => ({ kind: 'original' as const, docxIndex: b.docxIndex! }))
    expect(await saveDocx(doc, blocks)).toEqual(bytes)
    const run = doc.blocks[0].runs![0]
    expect(mergeRPrModel(run.rawRPr!, run, false)).toBe(RPR)
  })

  it('rebuilds only the edited groups', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: BODY }))
    const run = doc.blocks[0].runs![0]
    const out = mergeRPrModel(
      run.rawRPr!,
      {
        ...run,
        caps: 'all',
        dstrike: false,
        vanishOwn: false,
        charSpacingTwips: -10,
        charScalePct: undefined,
        kernHalfPoints: 0,
        positionHalfPoints: undefined,
      },
      false,
    )
    expect(out).toBe(
      '<w:rPr><w:caps/><w:dstrike w:val="0"/><w:vanish w:val="0"/><w:spacing w:val="-10"/>' +
        '<w:kern w:val="0"/><w:sz w:val="24"/></w:rPr>',
    )
  })

  it('turning small caps off writes both explicit off toggles', () => {
    const out = mergeRPrModel(
      '<w:rPr><w:smallCaps/><w:sz w:val="24"/></w:rPr>',
      { text: 'x', caps: 'none', sizeHalfPoints: 24 },
      false,
    )
    expect(out).toBe('<w:rPr><w:caps w:val="0"/><w:smallCaps w:val="0"/><w:sz w:val="24"/></w:rPr>')
  })

  it('a style-inherited vanish never materializes on the run', () => {
    const raw = '<w:rPr><w:rStyle w:val="Hidden"/><w:sz w:val="24"/></w:rPr>'
    const out = mergeRPrModel(
      raw,
      { text: 'x', styleId: 'Hidden', sizeHalfPoints: 24, vanish: true },
      false,
    )
    expect(out).toBe(raw)
  })

  it('emits fresh children in schema order', () => {
    const out = mergeRPrModel(
      '<w:rPr/>',
      {
        text: 'x',
        caps: 'small',
        dstrike: true,
        vanishOwn: true,
        color: 'FF0000',
        charSpacingTwips: 40,
        charScalePct: 150,
        kernHalfPoints: 28,
        positionHalfPoints: -4,
        sizeHalfPoints: 24,
      },
      false,
    )
    expect(out).toBe(
      '<w:rPr><w:smallCaps/><w:dstrike/><w:vanish/><w:color w:val="FF0000"/><w:spacing w:val="40"/>' +
        '<w:w w:val="150"/><w:kern w:val="28"/><w:position w:val="-4"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>',
    )
  })
})
