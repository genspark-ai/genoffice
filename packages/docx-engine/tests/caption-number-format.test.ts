import { describe, expect, it } from 'vitest'
import { formatSeqNumber, generateCaptionXml, generateParagraphXml } from '../src/generate'

describe('caption number formats', () => {
  it('formats a SEQ number the way its switch shows it', () => {
    expect(formatSeqNumber(4, 'ARABIC')).toBe('4')
    expect(formatSeqNumber(3, 'alphabetic')).toBe('c')
    expect(formatSeqNumber(27, 'ALPHABETIC')).toBe('AA')
    expect(formatSeqNumber(14, 'roman')).toBe('xiv')
    expect(formatSeqNumber(1994, 'ROMAN')).toBe('MCMXCIV')
  })

  it('writes the format switch and the formatted number', () => {
    const xml = generateCaptionXml('Figure', 2, 'Layout', { format: 'ROMAN' })
    expect(xml).toContain(' SEQ Figure \\* ROMAN ')
    expect(xml).toContain('<w:t>II</w:t>')
    expect(xml).toContain('>Figure </w:t>')
  })

  it('leaves the label out of the text but keeps it as the sequence name', () => {
    const xml = generateCaptionXml('Figure', 1, '', { excludeLabel: true })
    expect(xml).not.toContain('>Figure </w:t>')
    expect(xml).toContain(' SEQ Figure \\* ARABIC ')
  })

  it('quotes a label that has a space', () => {
    expect(generateCaptionXml('Photo plate', 1, '')).toContain(' SEQ "Photo plate" \\* ARABIC ')
  })
})

describe('paragraph bookmarks', () => {
  it('span the paragraph text so a REF field to one has something to show', () => {
    const xml = generateParagraphXml(
      {
        type: 'paragraph',
        bookmarks: ['Intro'],
        hiddenBookmarks: ['_Ref1'],
        runs: [{ text: 'x' }],
      },
      { headingStyleIds: new Map(), allocateHyperlinkRel: () => 'rId1' },
    )
    expect(xml).toMatch(
      /<w:bookmarkStart w:id="(\d+)" w:name="_Ref1"\/><w:bookmarkStart w:id="(\d+)" w:name="Intro"\/><w:r>.*<\/w:r><w:bookmarkEnd w:id="\1"\/><w:bookmarkEnd w:id="\2"\/><\/w:p>$/,
    )
  })
})
