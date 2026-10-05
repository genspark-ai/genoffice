import { describe, expect, it } from 'vitest'

import {
  addElement,
  copyElementData,
  createBlankPptx,
  openPptx,
  pasteElements,
  type ElementClipboardItem,
} from '../src/index'
import { relsPathFor } from '../src/zip'

const OFF = { x: 914400, y: 914400, cx: 1828800, cy: 914400 }
const SHIFT = { dx: 152400, dy: 152400 }
const HYPERLINK_REL =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'

/** A real clipboard item, plus one extra EXTERNAL relationship carrying `url`. */
async function itemWithExternalUrl(url: string): Promise<ElementClipboardItem> {
  const opened = await openPptx(await createBlankPptx())
  const slide = opened.deck.slides[0]!
  const el = addElement(slide, {
    kind: 'textbox',
    offset: { ...OFF },
    paragraphs: [{ runs: [{ text: 'copy me' }] }],
  })
  const clip = copyElementData(opened, slide, el)
  return {
    ...clip,
    rels: [...clip.rels, { rid: 'rId1', type: HYPERLINK_REL, target: url, external: true }],
  }
}

const lastTarget = (xml: string): string | undefined =>
  [...xml.matchAll(/Target="([^"]*)"/g)].at(-1)?.[1]

/** every relationship must remain a single self-contained element */
const isWellFormed = (xml: string): boolean => {
  const open = (xml.match(/<Relationship\b/g) ?? []).length
  const selfClosing = (xml.match(/\/>/g) ?? []).length
  return open === selfClosing
}

describe('relationship writers must not expand $ patterns in a replacement', () => {
  // String.prototype.replace treats the REPLACEMENT as a pattern: $& is the
  // matched text, $` the text before it, $' the text after, $$ a literal $.
  // escapeXmlAttr escapes the XML metacharacters but not '$', and an external
  // target is a user-supplied URL, so a $& in it substituted the matched
  // </Relationships> into the middle of the Target attribute and left the rels
  // part malformed - PowerPoint then offers to repair the file.
  // A function replacer inserts the text verbatim, so every $-sequence is
  // preserved as written; a string replacer would rewrite them.
  const urls: [label: string, url: string, escaped: string][] = [
    ['$&', 'https://example.com/?a=1$&b=2', 'https://example.com/?a=1$&amp;b=2'],
    ['$$', 'https://example.com/?a=$$b', 'https://example.com/?a=$$b'],
    [
      '$& twice',
      'https://example.com/?a=1$&b=2$&c=3',
      'https://example.com/?a=1$&amp;b=2$&amp;c=3',
    ],
  ]

  for (const [label, url, escaped] of urls) {
    it(`keeps ${label} in an external hyperlink target verbatim`, async () => {
      const opened = await openPptx(await createBlankPptx())
      const item = await itemWithExternalUrl(url)
      const r = pasteElements(opened, 0, [item], SHIFT)
      expect(r).not.toBeNull()

      const rels = opened.archive.readText(relsPathFor(opened.deck.slides[0]!.path))!
      expect(lastTarget(rels)).toBe(escaped)
      expect(rels).toContain('TargetMode="External"')
      expect(isWellFormed(rels)).toBe(true)
    })
  }

  it('leaves an ordinary URL byte-identical to the previous behaviour', async () => {
    const url = 'https://example.com/q4/report.pdf?v=2'
    const opened = await openPptx(await createBlankPptx())
    pasteElements(opened, 0, [await itemWithExternalUrl(url)], SHIFT)
    const rels = opened.archive.readText(relsPathFor(opened.deck.slides[0]!.path))!
    expect(lastTarget(rels)).toBe(url)
  })
})
