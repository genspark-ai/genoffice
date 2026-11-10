import { describe, expect, it } from 'vitest'
import { parseDocx, saveDocx, symbolGlyph } from '../src/index'
import { buildDocx } from './helpers/build-docx'

/**
 * Textless legacy VML shapes in the body (w:pict without a textbox): inline
 * ones ride as SVG run images, floating ones become geometry boxes, and
 * symbol runs sharing a run with empty picts keep their glyphs.
 */

const NS =
  'xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"'

const DIAMOND_TYPE =
  `<v:shapetype ${NS} id="_x0000_t110" coordsize="21600,21600" o:spt="110" ` +
  'path="m10800,l,10800,10800,21600,21600,10800xe"><v:stroke joinstyle="miter"/></v:shapetype>'
const diamond = (id: string, type = '') =>
  `<w:pict>${type}<v:shape ${NS} id="${id}" type="#_x0000_t110" ` +
  'style="width:1in;height:48pt;mso-position-horizontal-relative:char;mso-position-vertical-relative:line" ' +
  'fillcolor="white [3201]" strokecolor="#c2d69b [1942]" strokeweight="1pt">' +
  '<v:fill color2="#d6e3bc [1302]" focus="100%" type="gradient"/><w10:wrap type="none"/></v:shape></w:pict>'

const INLINE_PAIR =
  '<w:p><w:r><w:rPr><w:noProof/></w:rPr>' +
  diamond('_x0000_s1029', DIAMOND_TYPE) +
  diamond('_x0000_s1028') +
  '</w:r></w:p>'

const FLOATING_OVAL =
  `<w:p><w:r><w:pict><v:oval ${NS} id="_x0000_s1028" ` +
  'style="position:absolute;margin-left:11.25pt;margin-top:6.8pt;width:62.25pt;height:46.5pt;z-index:251658240"/>' +
  '</w:pict></w:r></w:p>'

const ROUNDRECT_TEXTBOX =
  `<w:p><w:r><w:pict><v:roundrect ${NS} id="_x0000_s1026" ` +
  'style="position:absolute;margin-left:48.75pt;margin-top:24.75pt;width:138pt;height:1in;z-index:251658240" arcsize="10923f">' +
  '<v:textbox><w:txbxContent><w:p><w:r><w:t>Top</w:t></w:r></w:p></w:txbxContent></v:textbox>' +
  '</v:roundrect></w:pict></w:r></w:p>'

const SPACED_FLOATING_OVAL = FLOATING_OVAL.replace(
  '<w:p>',
  '<w:p><w:pPr><w:spacing w:after="200" w:line="276" w:lineRule="auto"/></w:pPr>',
)

const emptyTextbox = (attrs: string) =>
  `<w:p><w:r><w:pict><v:shape ${NS} id="_x0000_s1026" type="#_x0000_t202" ` +
  'style="width:468.55pt;height:185.1pt;mso-position-horizontal:absolute;mso-position-horizontal-relative:char;mso-position-vertical:absolute;mso-position-vertical-relative:line" ' +
  `${attrs}><v:textbox inset="0,0,0,0"><w:txbxContent><w:p/></w:txbxContent></v:textbox>` +
  '<w10:wrap type="none"/></v:shape></w:pict></w:r></w:p>'

const SYM_WITH_EMPTY_PICTS =
  '<w:p><w:r><w:t>doc</w:t><w:sym w:font="Webdings" w:char="F045"/><w:t>ument</w:t>' +
  '<w:pict></w:pict><w:pict></w:pict></w:r></w:p>'

const unchangedSave = async (bytes: Uint8Array) => {
  const doc = await parseDocx(bytes)
  return saveDocx(
    doc,
    doc.blocks
      .filter((b) => !b.hidden)
      .map((b) => ({ kind: 'original' as const, docxIndex: b.docxIndex! })),
  )
}

describe('body VML shapes without text', () => {
  it('two inline shapes in one run become two SVG run images, the second resolving the shapetype of the first', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: INLINE_PAIR }))
    const block = doc.blocks[0]
    expect(block.type).toBe('paragraph')
    const images = block.runs!.filter((r) => r.image)
    expect(images).toHaveLength(2)
    for (const run of images) {
      expect(run.image!.widthPx).toBe(96)
      expect(run.image!.heightPx).toBe(64)
      const svg = decodeURIComponent(run.image!.dataUrl)
      expect(svg).toContain('<path d="M 48 0.67 L 0.67 32 L 48 63.33 L 95.33 32 Z"')
      expect(svg).toContain('stop-color="#FFFFFF"/><stop offset="1" stop-color="#d6e3bc"')
      expect(svg).toContain('stroke="#c2d69b"')
    }
    expect(images[1].image!.xml).not.toContain('v:shapetype')
  })

  it('a bare floating oval becomes a geometry box with the default black stroke at its pt margins', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: FLOATING_OVAL }))
    const block = doc.blocks[0]
    expect(block.type).toBe('passthrough')
    const box = block.textboxes?.[0]
    expect(box?.floating).toBe(true)
    expect(box?.widthPx).toBe(83)
    expect(box?.heightPx).toBe(62)
    expect(box?.offsetXEmu).toBe(Math.round(11.25 * 12700))
    expect(box?.offsetYEmu).toBe(Math.round(6.8 * 12700))
    expect(decodeURIComponent(box!.fillImageDataUrl!)).toContain(
      '<ellipse cx="41.5" cy="31" rx="41" ry="30.5" fill="#FFFFFF" stroke="#000000" stroke-width="1"/>',
    )
  })

  it('the floating shape paragraph keeps its own empty anchor line with the paragraph spacing', async () => {
    const bare = (await parseDocx(await buildDocx({ bodyXml: FLOATING_OVAL }))).blocks[0]
    expect(bare.textboxes?.[0]?.floating).toBe(true)
    expect(bare.anchorLine).toBeDefined()
    const spaced = (await parseDocx(await buildDocx({ bodyXml: SPACED_FLOATING_OVAL }))).blocks[0]
    expect(spaced.anchorLine?.format?.spaceAfter).toBe(200)
    expect(spaced.anchorLine?.format?.lineSpacing).toBe(1.15)
  })

  it('a text-empty inline textbox with a fill is drawn as an inline box; a white unstroked one is not', async () => {
    const red = (
      await parseDocx(await buildDocx({ bodyXml: emptyTextbox('fillcolor="red" stroked="f"') }))
    ).blocks[0]
    expect(red.label).toBe('Text box')
    const box = red.textboxes?.[0]
    expect(box?.floating).toBeUndefined()
    expect(box?.fill).toBe('FF0000')
    expect(box?.widthPx).toBe(625)
    expect(box?.heightPx).toBe(247)
    const white = (
      await parseDocx(await buildDocx({ bodyXml: emptyTextbox('fillcolor="white" stroked="f"') }))
    ).blocks[0]
    expect(white.textboxes ?? []).toHaveLength(0)
  })

  it('style rotation reaches both the floating box and the inline run image', async () => {
    const rotated = (pos: string) =>
      `<w:p><w:r><w:pict><v:rect ${NS} style="${pos}width:20pt;height:10pt;rotation:45"/></w:pict></w:r></w:p>`
    const box = (await parseDocx(await buildDocx({ bodyXml: rotated('position:absolute;') })))
      .blocks[0].textboxes?.[0]
    expect(box?.rotDeg).toBe(45)
    const run = (await parseDocx(await buildDocx({ bodyXml: rotated('') }))).blocks[0].runs!.find(
      (r) => r.image,
    )
    expect(run?.image?.rotDeg).toBe(45)
  })

  it('a roundrect textbox keeps its rounded frame geometry', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: ROUNDRECT_TEXTBOX }))
    const box = doc.blocks[0].textboxes?.[0]
    expect(box?.prst).toBe('roundRect')
    expect(box?.borderColor).toBe('000000')
    expect(box?.paras[0]?.runs[0]?.text).toBe('Top')
  })

  it('empty picts leave the paragraph editable and each w:sym keeps its symbol font', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: SYM_WITH_EMPTY_PICTS }))
    const block = doc.blocks[0]
    expect(block.type).toBe('paragraph')
    expect(block.runs!.map((r) => [r.text, r.sym?.font])).toEqual([
      ['doc', undefined],
      [symbolGlyph('Webdings', 'F045'), 'Webdings'],
      ['ument', undefined],
    ])
  })

  it('no edits -> byte-identical save for every shape form', async () => {
    for (const bodyXml of [
      INLINE_PAIR,
      FLOATING_OVAL,
      ROUNDRECT_TEXTBOX,
      SYM_WITH_EMPTY_PICTS,
      emptyTextbox('fillcolor="red" stroked="f"'),
    ]) {
      const bytes = await buildDocx({ bodyXml })
      expect(await unchangedSave(bytes)).toBe(bytes)
    }
  })
})
