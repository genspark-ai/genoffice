import { describe, expect, it } from 'vitest'
import { createSlidesSkill, type DeckAccess } from '../src/renderer/ai/slides-skill'
import type { PictureRenderNode, RenderSlide, ShapeRenderNode } from '@genoffice/pptx-render'

/**
 * The first line of defence. The write guard stops a model from *damaging* a
 * span; this stops it ever *seeing* one, and tells it what the markers mean so
 * it can write around them instead of through them.
 *
 * Both halves are live: the instruction is rebuilt per request from the deck as
 * it is now, so a mark applied a moment ago is in the next prompt and a removed
 * one is gone from it.
 */

const LABEL = '客户电话'
const SECRET = '13800138000'
/** the section heading placeholderInstruction emits; the base slides prompt
 *  already says "placeholder" about layout slots, so assertions anchor here */
const HEADER = '## Private placeholders'

const glyph = (text: string, redact?: string) => ({
  text,
  x: 0,
  baselineY: 20,
  fontFamily: 'Arial',
  fontSizePx: 24,
  color: '#000',
  bold: false,
  italic: false,
  underline: !!redact,
  widthPx: text.length * 12,
  ...(redact ? { redact } : {}),
})

const shape = (runs: Array<ReturnType<typeof glyph>>, redact?: string): ShapeRenderNode =>
  ({
    id: 'r_1',
    sourceId: 'e_1',
    type: 'shape',
    box: { x: 0, y: 0, w: 400, h: 60 },
    fill: { kind: 'none' },
    text: {
      lines: [{ runs, top: 0, height: 28 }],
      insets: { l: 0, t: 0, r: 0, b: 0 },
      anchor: 'top',
      fontScale: 1,
      contentHeight: 28,
    },
    ...(redact ? { redact } : {}),
  }) as unknown as ShapeRenderNode

const picture = (redact?: string): PictureRenderNode =>
  ({
    id: 'r_2',
    sourceId: 'e_2',
    type: 'picture',
    box: { x: 0, y: 80, w: 100, h: 100 },
    dataUrl: 'data:image/png;base64,AAAA',
    ...(redact ? { redact } : {}),
  }) as unknown as PictureRenderNode

const deck = (...nodes: ShapeRenderNode[]): RenderSlide =>
  ({
    index: 0,
    widthPx: 1280,
    heightPx: 720,
    scale: 1,
    nodes,
    background: null,
  }) as unknown as RenderSlide

const access = (slides: RenderSlide[]): DeckAccess =>
  ({
    getSlides: () => slides,
    getCurrent: () => 0,
    getSelectedIds: () => [],
    applySlide: () => {},
    applyDeck: () => {},
  }) as unknown as DeckAccess

const promptFor = (slides: RenderSlide[]) => createSlidesSkill(access(slides)).systemPrompt

describe('the model is told what the placeholders mean', () => {
  it('names a span withheld in the text', () => {
    const p = promptFor([deck(shape([glyph(`Call ${SECRET}`, LABEL)]))])
    expect(p).toContain(HEADER)
    expect(p).toContain('{{客户电话}}')
    expect(p).toContain('placeholder')
  })

  it('names a withheld picture', () => {
    const p = promptFor([deck(picture('公司 logo') as unknown as ShapeRenderNode)])
    expect(p).toContain('{{公司 logo}}')
  })

  it('gathers spans from every page, not just the current one', () => {
    const p = promptFor([
      deck(shape([glyph('nothing here')])),
      deck(shape([glyph('x', 'second page label')])),
    ])
    expect(p).toContain('{{second page label}}')
  })

  it('does not repeat a label that appears on several pages', () => {
    const p = promptFor([
      deck(shape([glyph('a', LABEL)]), picture(LABEL) as unknown as ShapeRenderNode),
      deck(shape([glyph('b', LABEL)])),
    ])
    // count the list entry, not the marker: the instruction also shows one in
    // its worked example
    expect(p.split('\n').filter((l) => l === `- {{${LABEL}}}`)).toHaveLength(1)
  })

  it('says nothing at all about markers when there are none', () => {
    // one page holding both a plain shape and a plain picture
    const plain = promptFor([
      deck(shape([glyph('nothing secret')]), picture() as unknown as ShapeRenderNode),
    ])
    // the section header, not the word: the base slides prompt already uses
    // "placeholder" to mean a layout title/body slot, so a word-level assertion
    // here would be a false signal
    expect(plain).not.toContain(HEADER)
    expect(plain).not.toContain('{{')
  })

  it('never leaks the words it is protecting', () => {
    const p = promptFor([deck(shape([glyph(`Call ${SECRET}`, LABEL)]))])
    expect(p).not.toContain(SECRET)
  })

  it('keeps the media-tool note that was already there', () => {
    // the two injections compose; losing either would be a regression the
    // other test would not catch
    const p = promptFor([deck(shape([glyph('x', LABEL)]))])
    expect(p.length).toBeGreaterThan(0)
    expect(p).toContain(HEADER)
    expect(p).toContain('{{客户电话}}')
  })

  it('is rebuilt per request, so a mark applied now is in the next prompt', () => {
    const slides = [deck(shape([glyph('plain')]))]
    const skill = createSlidesSkill(access(slides))
    expect(skill.systemPrompt).not.toContain(HEADER)
    // the reader applies a mark; the same skill object now reports it
    ;(slides[0]!.nodes[0] as ShapeRenderNode).text!.lines[0]!.runs[0]!.redact = 'added later'
    expect(skill.systemPrompt).toContain('{{added later}}')
  })
})
