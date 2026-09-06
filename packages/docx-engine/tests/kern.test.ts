/**
 * w:kern (pair-kerning threshold, half-points) is display-only: Word kerns a run
 * only when it is on and the font size reaches the threshold. The value is
 * exposed at docDefaults, style and run level; raw bytes round-trip via rawRPr.
 */
import { describe, expect, it } from 'vitest'
import { parseDocx } from '../src/index'
import { buildDocx } from './helpers/build-docx'

const STYLES =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
  '<w:docDefaults><w:rPrDefault><w:rPr><w:kern w:val="2"/><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>' +
  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/>' +
  '<w:rPr><w:kern w:val="28"/><w:sz w:val="56"/></w:rPr></w:style>' +
  '<w:style w:type="table" w:styleId="Grid"><w:name w:val="Table Grid"/><w:rPr><w:kern w:val="0"/></w:rPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Plain"><w:name w:val="Plain"/><w:basedOn w:val="Normal"/>' +
  '<w:rPr><w:kern w:val="0"/></w:rPr></w:style>' +
  '</w:styles>'

const BODY =
  '<w:p><w:r><w:t>default</w:t></w:r></w:p>' +
  '<w:p><w:r><w:rPr><w:kern w:val="0"/></w:rPr><w:t>off</w:t></w:r></w:p>' +
  '<w:p><w:r><w:rPr><w:kern w:val="32"/><w:sz w:val="40"/></w:rPr><w:t>big</w:t></w:r></w:p>'

describe('w:kern parsing', () => {
  it('exposes the threshold at docDefaults, style and run level', async () => {
    const parsed = await parseDocx(await buildDocx({ bodyXml: BODY, stylesXml: STYLES }))
    expect(parsed.docDefaults?.kernHalfPoints).toBe(2)
    expect(parsed.styles.get('Title')?.display?.kernHalfPoints).toBe(28)
    expect(parsed.styles.get('Plain')?.display?.kernHalfPoints).toBe(0)
    expect(parsed.styles.get('Normal')?.display?.kernHalfPoints).toBeUndefined()
    expect(parsed.styles.get('Grid')?.tableDisplay?.wholeTable?.kernHalfPoints).toBe(0)
    const runs = parsed.blocks
      .filter((b) => b.type === 'paragraph')
      .map((b) => (b as { runs: Array<{ kernHalfPoints?: number }> }).runs[0])
    expect(runs.map((r) => r.kernHalfPoints)).toEqual([undefined, 0, 32])
  })
})
