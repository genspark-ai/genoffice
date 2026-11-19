import { describe, expect, it } from 'vitest'
import { parseDocx, saveDocx } from '../src/index'
import { buildDocx } from './helpers/build-docx'

const FRAME_P =
  '<w:p><w:pPr>' +
  '<w:framePr w:w="2851" w:h="3241" w:hSpace="180" w:wrap="around" w:vAnchor="text" w:hAnchor="page" w:x="1816" w:y="299"/>' +
  '<w:textDirection w:val="tbRl"/>' +
  '</w:pPr><w:r><w:t>sideways frame</w:t></w:r></w:p>'

const ALIGNED_FRAME_P =
  '<w:p><w:pPr>' +
  '<w:framePr w:w="1200" w:wrap="notBeside" w:vAnchor="margin" w:hAnchor="margin" w:xAlign="right" w:yAlign="top" w:anchorLock="1"/>' +
  '</w:pPr><w:r><w:t>margin note</w:t></w:r></w:p>'

const DROP_CAP_P =
  '<w:p><w:pPr><w:framePr w:dropCap="drop" w:lines="3" w:wrap="around"/></w:pPr>' +
  '<w:r><w:t>Lorem</w:t></w:r></w:p>'

describe('positioned frame and text flow parsing', () => {
  it('reads w:framePr geometry and the paragraph text direction', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: FRAME_P }))
    const f = doc.blocks[0].format
    expect(f?.frame).toEqual({
      wTwips: 2851,
      hTwips: 3241,
      hRule: 'atLeast',
      xTwips: 1816,
      yTwips: 299,
      vAnchor: 'text',
      wrap: 'around',
      hSpaceTwips: 180,
    })
    expect(f?.textDirection).toBe('tbRl')
    expect(f?.dropCap).toBeUndefined()
  })

  it('keeps alignment anchors and the lock flag', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: ALIGNED_FRAME_P }))
    expect(doc.blocks[0].format?.frame).toMatchObject({
      wTwips: 1200,
      hAnchor: 'margin',
      vAnchor: 'margin',
      xAlign: 'right',
      yAlign: 'top',
      anchorLock: true,
      wrap: 'notBeside',
    })
    expect(doc.blocks[0].format?.textDirection).toBeUndefined()
  })

  it('w:wrap="tight" is kept as a wrapping frame', async () => {
    const doc = await parseDocx(
      await buildDocx({ bodyXml: ALIGNED_FRAME_P.replace('notBeside', 'tight') }),
    )
    expect(doc.blocks[0].format?.frame?.wrap).toBe('tight')
  })

  it('a drop-cap framePr stays a drop cap, not a frame', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: DROP_CAP_P }))
    expect(doc.blocks[0].format?.dropCap?.type).toBe('drop')
    expect(doc.blocks[0].format?.frame).toBeUndefined()
  })

  it('a rebuilt pPr emits the frame and its text direction', async () => {
    const { generateParagraphXml } = await import('../src/generate')
    const xml = generateParagraphXml(
      {
        type: 'paragraph' as const,
        runs: [{ text: 'sideways' }],
        format: {
          frame: { wTwips: 2851, hTwips: 3241, xTwips: 1816, yTwips: 299, wrap: 'around' as const },
          textDirection: 'tbRl' as const,
        },
      },
      { headingStyleIds: new Map(), allocateHyperlinkRel: () => 'rId1' },
    )
    expect(xml).toContain('<w:textDirection w:val="tbRl"/>')
    expect(xml).toContain('w:framePr w:w="2851" w:h="3241" w:hRule="atLeast" w:wrap="around"')
  })

  it('round-trips the frame attributes through save', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: FRAME_P }))
    const saved = await saveDocx(doc, [{ kind: 'original', docxIndex: 0 }])
    const again = await parseDocx(saved)
    expect(again.blocks[0].format?.frame).toEqual(doc.blocks[0].format?.frame)
    expect(again.blocks[0].format?.textDirection).toBe('tbRl')
    expect(again.internal.documentXml).toContain('w:hSpace="180"')
  })
})
