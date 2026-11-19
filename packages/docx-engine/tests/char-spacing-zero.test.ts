/**
 * An explicit `w:spacing w:val="0"` on a run or style is a real value that
 * cancels inherited letter spacing (a Titre1 heading spaced 2pt by its style
 * wrapped to two lines when the run's 0 was dropped as unset).
 */
import { describe, expect, it } from 'vitest'
import { parseDocx } from '../src/index'
import { buildDocx } from './helpers/build-docx'

const STYLES =
  '<w:style w:type="paragraph" w:styleId="Spaced"><w:name w:val="Spaced"/>' +
  '<w:rPr><w:spacing w:val="40"/></w:rPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Tight"><w:name w:val="Tight"/>' +
  '<w:basedOn w:val="Spaced"/><w:rPr><w:spacing w:val="0"/></w:rPr></w:style>'

describe('w:spacing w:val="0"', () => {
  it('keeps the explicit 0 on the run', async () => {
    const doc = await parseDocx(
      await buildDocx({
        extraStylesXml: STYLES,
        bodyXml:
          '<w:p><w:pPr><w:pStyle w:val="Spaced"/></w:pPr>' +
          '<w:r><w:rPr><w:spacing w:val="0"/></w:rPr><w:t>HEAD</w:t></w:r>' +
          '<w:r><w:t>ING</w:t></w:r></w:p>',
      }),
    )
    const runs = doc.blocks[0].runs!
    expect(runs[0].charSpacingTwips).toBe(0)
    expect(runs[1].charSpacingTwips).toBeUndefined()
  })

  it('lets a derived style cancel its parent letter spacing', async () => {
    const doc = await parseDocx(
      await buildDocx({ extraStylesXml: STYLES, bodyXml: '<w:p><w:r><w:t>x</w:t></w:r></w:p>' }),
    )
    expect(doc.styles.get('Spaced')!.display?.charSpacingTwips).toBe(40)
    expect(doc.styles.get('Tight')!.display?.charSpacingTwips).toBe(0)
  })
})
