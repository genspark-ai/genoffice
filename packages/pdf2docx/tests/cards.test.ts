/** P20 card regions: backdrop plate + text group → one anchored text box. */
import { describe, expect, it } from 'vitest'
import { detectCardRegions, type CardCandidate } from '../src/analyze'
import type { ImageBlock, IrPage, Line, PageSection, Span, TableBlock, TextBlock } from '../src/ir'
import { rebuildDocx } from '../src/rebuild'

function span(text: string, over: Partial<Span> = {}): Span {
  return {
    text,
    box: { x0: 72, y0: 690, x1: 300, y1: 700 },
    fontSize: 12,
    fontFamily: 'Helvetica',
    bold: false,
    italic: false,
    color: '000000',
    dir: 'ltr',
    script: 'latin',
    ...over,
  }
}

function blockAt(
  text: string,
  x0: number,
  top: number,
  widthPt: number,
  fontSize = 12,
  over: Partial<TextBlock> = {},
  spanOver: Partial<Span> = {},
): TextBlock {
  const s = span(text, {
    fontSize,
    box: { x0, y0: top - fontSize, x1: x0 + widthPt, y1: top },
    ...spanOver,
  })
  const line: Line = {
    spans: [s],
    box: s.box,
    baseline: top - fontSize * 0.8,
    endsWithHyphen: false,
  }
  return {
    kind: 'text',
    lines: [line],
    box: s.box,
    align: 'left',
    firstLineIndentPt: 0,
    dir: 'ltr',
    ...over,
  }
}

const plateImage = (): ImageBlock => ({
  kind: 'image',
  box: { x0: 60, y0: 480, x1: 540, y1: 600 },
  data: new Uint8Array([1]),
  mime: 'image/png',
  pixelWidth: 2,
  pixelHeight: 2,
  float: { wrap: 'behind', xOffsetPt: 60 },
})

/** portrait page: dark plate 60..540 × y 480..600, two light text members on it */
function cardPage(over: Partial<IrPage> = {}): {
  page: IrPage
  candidates: CardCandidate[]
  members: TextBlock[]
} {
  const head = blockAt(
    'FINAL TAKEAWAY',
    84,
    588,
    200,
    14,
    { spacingBeforePt: 30 },
    { color: 'FFFFFF' },
  )
  const body = blockAt(
    'the conclusion line',
    84,
    550,
    380,
    12,
    { spacingBeforePt: 12 },
    { color: 'EEEEEE' },
  )
  const after = blockAt('after the card', 72, 440, 300, 12, { spacingBeforePt: 24 })
  const img = plateImage()
  const blocks: IrPage['blocks'] = [head, body, after]
  const sections: PageSection[] = [
    {
      box: { x0: 60, y0: 60, x1: 540, y1: 720 },
      columns: [{ box: { x0: 60, y0: 60, x1: 540, y1: 720 }, blocks }],
      gutterWidthsPt: [],
      dir: 'ltr',
    },
  ]
  const page: IrPage = {
    index: 0,
    widthPt: 612,
    heightPt: 792,
    rotation: 0,
    blocks,
    degraded: false,
    scanned: false,
    hasStructTree: false,
    bgPanels: [img],
    sections,
    ...over,
  }
  const candidates: CardCandidate[] = [{ img, box: img.box, color: '1A1A2E', rounded: false }]
  return { page, candidates, members: [head, body] }
}

const docXml = async (docx: Uint8Array): Promise<string> => {
  const { default: JSZip } = await import('jszip')
  const zip = await JSZip.loadAsync(docx)
  return zip.file('word/document.xml')!.async('string')
}

describe('detectCardRegions (P20)', () => {
  it('tags a contiguous text group fully on a contrasting plate', () => {
    const { page, candidates, members } = cardPage()
    detectCardRegions(page, candidates)
    expect(page.cards).toHaveLength(1)
    expect(page.cards![0]!.color).toBe('1A1A2E')
    expect(members.every((m) => m.cardId === 0)).toBe(true)
    expect(candidates[0]!.img.cardId).toBe(0)
  })

  it('keeps the panel when a table overlaps the plate (form banner lesson)', () => {
    const { page, candidates } = cardPage()
    const table: TableBlock = {
      kind: 'table',
      box: { x0: 60, y0: 470, x1: 540, y1: 560 },
      rows: [],
      colWidthsPt: [480],
    } as unknown as TableBlock
    page.sections![0]!.columns[0]!.blocks.push(table)
    page.blocks.push(table)
    detectCardRegions(page, candidates)
    expect(page.cards).toBeUndefined()
  })

  it('keeps the panel when a text block straddles the plate edge', () => {
    const { page, candidates } = cardPage()
    const straddler = blockAt('half on the plate', 60, 486, 300, 12)
    straddler.box = { x0: 60, y0: 420, x1: 360, y1: 500 } // ~17% inside
    page.sections![0]!.columns[0]!.blocks.push(straddler)
    detectCardRegions(page, candidates)
    expect(page.cards).toBeUndefined()
  })

  it('keeps the panel when the text is not light (P10 B bar)', () => {
    const { page, candidates } = cardPage()
    for (const b of page.blocks) {
      if (b.kind === 'text') for (const l of b.lines) for (const s of l.spans) s.color = '333344'
    }
    detectCardRegions(page, candidates)
    expect(page.cards).toBeUndefined()
  })

  it('keeps the panel for a light plate under dark text (form fill-in areas)', () => {
    const { page, candidates } = cardPage()
    candidates[0]!.color = 'E7FFF2'
    for (const b of page.blocks) {
      if (b.kind === 'text') for (const l of b.lines) for (const s of l.spans) s.color = '111111'
    }
    detectCardRegions(page, candidates)
    expect(page.cards).toBeUndefined()
  })

  it('keeps the panel for a single-line label banner (form-gov green strips)', () => {
    const { page, candidates } = cardPage()
    // strip the body member: one single-line block on the plate is a banner
    const col = page.sections![0]!.columns[0]!
    col.blocks = col.blocks.filter(
      (b) => b.kind !== 'text' || b.lines[0]!.spans[0]!.text !== 'the conclusion line',
    )
    page.blocks = col.blocks
    detectCardRegions(page, candidates)
    expect(page.cards).toBeUndefined()
  })

  it('never fires on translucent scrims (they are not candidates upstream)', () => {
    // upstream filter: alpha panels are excluded from the candidate list —
    // here we just document that an empty candidate list is a no-op
    const { page } = cardPage()
    detectCardRegions(page, [])
    expect(page.cards).toBeUndefined()
  })
})

describe('card rebuild (P20)', () => {
  const detected = () => {
    const { page, candidates } = cardPage()
    detectCardRegions(page, candidates)
    return page
  }

  it('emits ONE anchored text box carrying plate fill, size and both members', async () => {
    const xml = await docXml(await rebuildDocx([detected()]))
    expect(xml.match(/<wps:txbx>/g)).toHaveLength(1)
    expect(xml).toContain('<a:solidFill><a:srgbClr val="1A1A2E"/></a:solidFill>')
    // plate: 480pt wide, 120pt tall
    expect(xml).toContain(
      `<wp:extent cx="${Math.round(480 * 12700)}" cy="${Math.round(120 * 12700)}"/>`,
    )
    expect(xml).toContain('<wp:wrapTopAndBottom/>')
    expect(xml).toContain('FINAL TAKEAWAY')
    expect(xml).toContain('the conclusion line')
    // members are inside the box, not in the body flow
    const txbx = /<w:txbxContent>([\s\S]*?)<\/w:txbxContent>/.exec(xml)?.[1] ?? ''
    expect(txbx).toContain('FINAL TAKEAWAY')
    const body = xml.replace(/<w:txbxContent>[\s\S]*?<\/w:txbxContent>/g, '')
    expect(body).not.toContain('FINAL TAKEAWAY')
    // the plate's behindDoc pin is gone (the box paints it)
    expect(xml).not.toContain('behindDoc="1"')
  })

  it('writes measured insets (plate left edge 60 → text left 84 = 24pt)', async () => {
    const xml = await docXml(await rebuildDocx([detected()]))
    const bodyPr = /<wps:bodyPr[^>]*>/.exec(xml)?.[0] ?? ''
    expect(bodyPr).toContain(`lIns="${Math.round(24 * 12700)}"`)
    // top inset: plate top 600 − head top 588 = 12pt
    expect(bodyPr).toContain(`tIns="${Math.round(12 * 12700)}"`)
  })

  it('flow text after the card stays in the body', async () => {
    const xml = await docXml(await rebuildDocx([detected()]))
    const body = xml.replace(/<w:txbxContent>[\s\S]*?<\/w:txbxContent>/g, '')
    expect(body).toContain('after the card')
  })

  it('canvas pages ignore card tags (plate pins, text frames as usual)', async () => {
    const page = detected()
    page.canvas = true
    page.widthPt = 960
    page.heightPt = 540
    const xml = await docXml(await rebuildDocx([page]))
    expect(xml).not.toContain('<wps:txbx>')
    expect(xml).toContain('behindDoc="1"')
  })
})
