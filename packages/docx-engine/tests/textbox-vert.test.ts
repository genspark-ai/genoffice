import { describe, expect, it } from 'vitest'
import { buildTextboxParagraphXml, parseDocx } from '../src/index'
import { buildDocx } from './helpers/build-docx'

const boxesOf = async (bodyXml: string) =>
  (await parseDocx(await buildDocx({ bodyXml }))).blocks.find((b) => b.textboxes?.length)
    ?.textboxes ?? []

describe('vertical text boxes (wps:bodyPr vert)', () => {
  it('writes the direction on a new text box and reads it back', async () => {
    const xml = buildTextboxParagraphXml({ widthEmu: 1080000, heightEmu: 1800000, vert: 'eaVert' })
    expect(xml).toContain('<wps:bodyPr vert="eaVert"/>')
    const [box] = await boxesOf(xml)
    expect(box.vert).toBe('eaVert')
    expect([box.widthPx, box.heightPx]).toEqual([113, 189])
  })

  it('leaves a horizontal text box without a direction', async () => {
    const xml = buildTextboxParagraphXml()
    expect(xml).toContain('<wps:bodyPr/>')
    expect((await boxesOf(xml))[0].vert).toBeUndefined()
  })

  it('reads the turned and stacked directions and ignores what is not a direction', async () => {
    const withVert = (vert: string) =>
      buildTextboxParagraphXml().replace('<wps:bodyPr/>', `<wps:bodyPr vert="${vert}"/>`)
    expect((await boxesOf(withVert('vert')))[0].vert).toBe('vert')
    expect((await boxesOf(withVert('vert270')))[0].vert).toBe('vert270')
    expect((await boxesOf(withVert('horz')))[0].vert).toBeUndefined()
    expect((await boxesOf(withVert('wordArtVert')))[0].vert).toBe('wordArtVert')
    expect((await boxesOf(withVert('mongolianVert')))[0].vert).toBe('mongolianVert')
    expect((await boxesOf(withVert('sideways')))[0].vert).toBeUndefined()
  })
})
