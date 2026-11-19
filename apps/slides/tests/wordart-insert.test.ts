import { beforeEach, describe, expect, it, vi } from 'vitest'
import {
  WORDART_BODY_PR,
  WORDART_FONT_PT,
  WORDART_H_PX,
  WORDART_W_PX,
  wordArtInsertSpec,
} from '../src/renderer/wordart-insert'
import { insertWordArt } from '../src/renderer/insert-actions'
import { t } from '../src/renderer/i18n/locale'
import type { ActionCtx } from '../src/renderer/action-context'

const PX_PER_PT = 96 / 72

describe('WordArt insert geometry (PowerPoint for Mac measurements)', () => {
  it('drops a 354.6 x 72.7 pt box, 54 pt text, wrap="none" + spAutoFit', () => {
    expect(WORDART_W_PX).toBe(Math.round(354.6 * PX_PER_PT))
    expect(WORDART_H_PX).toBe(Math.round(72.7 * PX_PER_PT))
    expect(WORDART_W_PX).toBe(473)
    expect(WORDART_H_PX).toBe(97)
    expect(WORDART_FONT_PT).toBe(54)
    expect(WORDART_BODY_PR).toEqual({ wrap: 'none', autoFit: 'resize' })
  })

  it('centers the box on the slide (left 302.7 pt, top 233.6 pt on a 960 x 540 slide)', () => {
    const spec = wordArtInsertSpec({ widthPx: 1280, heightPx: 720 })
    expect(spec).toEqual({
      x: Math.round((1280 - 473) / 2),
      y: Math.round((720 - 97) / 2),
      w: 473,
      h: 97,
      bodyPr: { wrap: 'none', autoFit: 'resize' },
    })
    expect(Math.abs(spec.x - 302.7 * PX_PER_PT)).toBeLessThan(1)
    expect(Math.abs(spec.y - 233.6 * PX_PER_PT)).toBeLessThan(1)
  })
})

describe('insertWordArt', () => {
  const addElement = vi.fn()
  const applySlide = vi.fn()
  const setSelectedIds = vi.fn()
  const setEditing = vi.fn()
  const setStatus = vi.fn()
  const slide = { widthPx: 1280, heightPx: 720 }
  const ctx = {
    slide,
    current: 2,
    applySlide,
    setSelectedIds,
    setEditing,
    setStatus,
  } as unknown as ActionCtx
  const preset = {
    id: 'p',
    nameKey: 'n',
    fill: '#FF0000',
    outline: { color: '#000000', widthEmu: 12700 },
    bold: true,
    italic: true,
  }

  beforeEach(() => {
    vi.clearAllMocks()
    ;(window as unknown as { slidesApi: unknown }).slidesApi = { addElement }
    addElement.mockResolvedValue({ slide, sourceId: 'sp7' })
  })

  it('adds a centered auto-fit text box carrying the preset styles at 54 pt', async () => {
    await insertWordArt(ctx, preset)
    expect(addElement).toHaveBeenCalledWith({
      slideIndex: 2,
      kind: 'textbox',
      xPx: 404,
      yPx: 312,
      wPx: 473,
      hPx: 97,
      fitWidthPx: 1280,
      bodyPr: { wrap: 'none', autoFit: 'resize' },
      paragraphs: [
        {
          align: 'center',
          runs: [
            {
              text: t('appWordArtPlaceholder'),
              fontSize: 54,
              bold: true,
              italic: true,
              color: '#FF0000',
              outline: { color: '#000000', widthEmu: 12700 },
            },
          ],
        },
      ],
    })
  })

  it('selects the new shape and enters editing with the whole placeholder selected', async () => {
    await insertWordArt(ctx, preset)
    expect(applySlide).toHaveBeenCalledWith(2, slide)
    expect(setSelectedIds).toHaveBeenCalledWith(['sp7'])
    expect(setEditing).toHaveBeenCalledWith({ sourceId: 'sp7', selectAll: true })
    expect(setStatus).toHaveBeenCalledWith(t('appStatusWordArtInserted'))
  })

  it('leaves the state alone when the add fails', async () => {
    addElement.mockResolvedValueOnce(null)
    await insertWordArt(ctx, preset)
    expect(setEditing).not.toHaveBeenCalled()
    expect(setSelectedIds).not.toHaveBeenCalled()
  })
})
