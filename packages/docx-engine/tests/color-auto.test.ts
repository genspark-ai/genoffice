/**
 * `<w:color w:val="auto"/>` is Word's automatic colour: it overrides a colour
 * inherited from the style chain (the run renders in the default ink, not the
 * style's colour), so it is modeled as the literal 'auto' instead of being dropped.
 */
import { describe, expect, it } from 'vitest'
import { parseDocx, saveDocx, type SaveBlock } from '../src/index'
import { buildDocx } from './helpers/build-docx'

const STYLES =
  '<w:style w:type="paragraph" w:styleId="RedBody"><w:name w:val="Red Body"/>' +
  '<w:rPr><w:color w:val="FF0000"/></w:rPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="ResetBody"><w:name w:val="Reset Body"/>' +
  '<w:basedOn w:val="RedBody"/><w:rPr><w:color w:val="auto"/></w:rPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="PlainBody"><w:name w:val="Plain Body"/>' +
  '<w:basedOn w:val="RedBody"/><w:rPr><w:sz w:val="20"/></w:rPr></w:style>'

const BODY =
  '<w:p><w:pPr><w:pStyle w:val="RedBody"/></w:pPr>' +
  '<w:r><w:rPr><w:color w:val="auto"/></w:rPr><w:t>automatic</w:t></w:r>' +
  '<w:r><w:t xml:space="preserve"> inherited</w:t></w:r>' +
  '<w:r><w:rPr><w:color w:val="0000FF"/></w:rPr><w:t>blue</w:t></w:r></w:p>'

describe('w:color auto', () => {
  it('models an explicit auto run colour instead of dropping it', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: BODY, extraStylesXml: STYLES }))
    const runs = doc.blocks[0].runs!
    expect(runs.map((r) => r.color)).toEqual(['auto', undefined, '0000FF'])
  })

  it('lets a style-level auto reset the basedOn colour while plain children inherit it', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: BODY, extraStylesXml: STYLES }))
    expect(doc.styles.get('RedBody')?.display?.color).toBe('FF0000')
    expect(doc.styles.get('ResetBody')?.display?.color).toBe('auto')
    expect(doc.styles.get('PlainBody')?.display?.color).toBe('FF0000')
  })

  it('saves an untouched auto run byte-identically', async () => {
    const source = await buildDocx({ bodyXml: BODY, extraStylesXml: STYLES })
    const doc = await parseDocx(source)
    const originals: SaveBlock[] = doc.blocks
      .filter((b) => !b.hidden)
      .map((b) => ({ kind: 'original', docxIndex: b.docxIndex! }))
    expect(await saveDocx(doc, originals)).toBe(source)
  })

  it('keeps w:val="auto" when the run is rebuilt', async () => {
    const source = await buildDocx({ bodyXml: BODY, extraStylesXml: STYLES })
    const doc = await parseDocx(source)
    const block = doc.blocks[0]
    const runs = block.runs!.map((r) => ({ ...r }))
    runs[0].text = 'edited'
    const bytes = await saveDocx(doc, [
      { kind: 'generated', block: { type: 'paragraph', styleId: block.styleId, runs } },
    ])
    const reparsed = await parseDocx(bytes)
    expect(reparsed.blocks[0].runs!.map((r) => [r.text, r.color])).toEqual([
      ['edited', 'auto'],
      [' inherited', undefined],
      ['blue', '0000FF'],
    ])
  })
})
