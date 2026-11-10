import { describe, expect, it } from 'vitest'
import { createBlankPptx, openPptx } from '../src/index'

describe('resolveSlideChain', () => {
  it('follows the slide rels when present', async () => {
    const { archive } = await openPptx(await createBlankPptx())
    const chain = archive.resolveSlideChain('ppt/slides/slide1.xml')
    expect(chain.layoutPath).toBe('ppt/slideLayouts/slideLayout1.xml')
    expect(chain.masterPath).toBe('ppt/slideMasters/slideMaster1.xml')
    expect(chain.themePath).toBe('ppt/theme/theme1.xml')
  })

  it('a slide with a ctrTitle placeholder falls back to the title layout', async () => {
    const { archive } = await openPptx(await createBlankPptx())
    archive.entries.delete('ppt/slides/_rels/slide1.xml.rels')
    // Add a second layout typed "title" behind the existing one
    const master = archive.readText('ppt/slideMasters/slideMaster1.xml')!
    archive.entries.set(
      'ppt/slideMasters/slideMaster1.xml',
      Buffer.from(
        master.replace(
          '</p:sldLayoutIdLst>',
          '<p:sldLayoutId id="2147483650" r:id="rIdT"/></p:sldLayoutIdLst>',
        ),
        'utf8',
      ),
    )
    const rels = archive.readText('ppt/slideMasters/_rels/slideMaster1.xml.rels')!
    archive.entries.set(
      'ppt/slideMasters/_rels/slideMaster1.xml.rels',
      Buffer.from(
        rels.replace(
          '</Relationships>',
          '<Relationship Id="rIdT" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayoutT.xml"/></Relationships>',
        ),
        'utf8',
      ),
    )
    archive.entries.set(
      'ppt/slideLayouts/slideLayoutT.xml',
      Buffer.from('<p:sldLayout type="title" xmlns:p="p"><p:cSld/></p:sldLayout>', 'utf8'),
    )
    const slide = archive.readText('ppt/slides/slide1.xml')!
    archive.entries.set(
      'ppt/slides/slide1.xml',
      Buffer.from(
        slide.replace(
          '</p:spTree>',
          '<p:sp><p:nvSpPr><p:cNvPr id="9" name="t"/><p:cNvSpPr/><p:nvPr><p:ph type="ctrTitle"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp></p:spTree>',
        ),
        'utf8',
      ),
    )
    expect(archive.resolveSlideChain('ppt/slides/slide1.xml').layoutPath).toBe(
      'ppt/slideLayouts/slideLayoutT.xml',
    )
    // Without the ctrTitle placeholder the first layout wins
    archive.entries.set('ppt/slides/slide1.xml', Buffer.from(slide, 'utf8'))
    expect(archive.resolveSlideChain('ppt/slides/slide1.xml').layoutPath).toBe(
      'ppt/slideLayouts/slideLayout1.xml',
    )
  })

  it('falls back to the first layout of the first master when the slide has no rels part', async () => {
    const { archive } = await openPptx(await createBlankPptx())
    archive.entries.delete('ppt/slides/_rels/slide1.xml.rels')
    const chain = archive.resolveSlideChain('ppt/slides/slide1.xml')
    expect(chain.layoutPath).toBe('ppt/slideLayouts/slideLayout1.xml')
    expect(chain.masterPath).toBe('ppt/slideMasters/slideMaster1.xml')
    expect(chain.themePath).toBe('ppt/theme/theme1.xml')
  })
})
