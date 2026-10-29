import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { parseSlide, setRedactExt } from '@genoffice/pptx-engine'
import { buildRenderSlide, type RenderSlide } from '@genoffice/pptx-render'

// file-actions pulls export-render → pptx-render → konva → native `canvas`
// (missing in CI). Mock the UI-only modules so this unit test exercises only
// the save path and the naming decision.
vi.mock('../src/renderer/export-render', () => ({ renderSlidesToPngBase64: vi.fn() }))
vi.mock('../src/renderer/components/toast-bus', () => ({ showToast: vi.fn() }))
vi.mock('../src/renderer/i18n/locale', () => ({ t: (k: string) => k }))

import { nameWithAi, save } from '../src/renderer/file-actions'
import type { ActionCtx } from '../src/renderer/action-context'

/**
 * The other half of `file-naming.test.ts`: that file pins what the model is
 * asked, this one pins where the answer goes — and, just as importantly, where
 * the model is *not* asked.
 *
 * A save never calls the model: a deck that has never been on disk keeps the
 * app's own untitled name. Only the reader's explicit "Name with AI…" action
 * asks, and that answer reaches the Save As dialog, not the file on disk.
 */

const LABEL = '客户电话'
const SECRET = '13800138000'
const MODEL_NAME = '客户回访记录'

const SP = (runs: string, cx = 3000000) =>
  '<p:sp><p:nvSpPr><p:cNvPr id="2" name="s"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
  `<p:spPr><a:xfrm><a:off x="100000" y="100000"/><a:ext cx="${cx}" cy="1200000"/></a:xfrm>` +
  '<a:prstGeom prst="rect"/></p:spPr>' +
  `<p:txBody><a:bodyPr wrap="square"/><a:lstStyle/><a:p>${runs}</a:p></p:txBody></p:sp>`

const RPR = '<a:rPr lang="en-US"/>'
const MARKED = setRedactExt(RPR, LABEL)
const run = (rPr: string, text: string) => `<a:r>${rPr}<a:t>${text}</a:t></a:r>`

const render = (body: string): RenderSlide =>
  buildRenderSlide(
    parseSlide({
      path: 'ppt/slides/slide1.xml',
      slideXml:
        '<?xml version="1.0"?><p:sld xmlns:p="p" xmlns:a="a" xmlns:r="r"><p:cSld>' +
        `<p:spTree><p:nvGrpSpPr/><p:grpSpPr/>${body}</p:spTree></p:cSld></p:sld>`,
      ctx: {},
    }),
    { cx: 9144000, cy: 6858000 },
    { fitWidthPx: 1280 },
  )

/** A deck whose phone number the reader has hidden from the model. */
function redactedDeck(): RenderSlide[] {
  return [render(SP(run(RPR, '客户回访') + run(MARKED, SECRET)))]
}

/** A deck with nothing withheld. */
function plainDeck(): RenderSlide[] {
  return [render(SP(run(RPR, '客户回访')))]
}

function ctxWith(over: {
  path: string | null
  slides?: RenderSlide[]
  editing?: boolean
}): ActionCtx {
  return {
    slides: over.slides ?? plainDeck(),
    current: 0,
    path: over.path,
    editingActiveRef: { current: over.editing ?? false },
    flushNotes: () => Promise.resolve(),
    setSlides: () => {},
    setSelectedIds: () => {},
    setEnteredGroupId: () => {},
    setEditing: () => {},
    setEditingCell: () => {},
    setDirty: () => {},
    setPath: () => {},
    setStatus: () => {},
  } as unknown as ActionCtx
}

interface NamingInput {
  content: string
}

let suggestFileName: ReturnType<typeof vi.fn>
let slidesApi: { save: ReturnType<typeof vi.fn>; saveAs: ReturnType<typeof vi.fn> }

beforeEach(() => {
  suggestFileName = vi.fn(async () => ({ ok: true, name: MODEL_NAME }))
  slidesApi = {
    save: vi.fn(async () => ({ ok: true })),
    saveAs: vi.fn(async () => ({ ok: true })),
  }
  vi.stubGlobal('window', {
    slidesApi,
    aiNamingApi: {
      suggestFileName,
    },
  })
})

afterEach(() => {
  vi.unstubAllGlobals()
  vi.restoreAllMocks()
})

describe('saving never asks the model', () => {
  it('does not call the model for a deck that has never been saved', async () => {
    await save(() => ctxWith({ path: null }), true)

    // A first save must not wait on a model round-trip.
    expect(suggestFileName).not.toHaveBeenCalled()
  })

  it('leaves the untitled name to the main process', async () => {
    await save(() => ctxWith({ path: null }), true)

    // No proposed name is passed: the deck keeps the localized "Untitled
    // presentation" the main process produces in pickDraftPath.
    expect(slidesApi.save).toHaveBeenCalledWith()
  })

  it('does not call the model for a deck that already has a name', async () => {
    await save(() => ctxWith({ path: '/tmp/deck.pptx' }), true)

    expect(suggestFileName).not.toHaveBeenCalled()
    expect(slidesApi.save).toHaveBeenCalledWith()
  })

  it('does not call the model for a deck with nothing to name', async () => {
    await save(() => ctxWith({ path: null, slides: [] }), true)

    expect(suggestFileName).not.toHaveBeenCalled()
    expect(slidesApi.save).toHaveBeenCalledWith()
  })
})

describe('the reader asks for a name, and the answer goes to the Save As dialog', () => {
  it('asks the model once and prefills the dialog', async () => {
    await nameWithAi(() => ctxWith({ path: '/tmp/deck.pptx' }))

    expect(suggestFileName).toHaveBeenCalledTimes(1)
    // the answer goes into the Save As dialog, not onto the file behind the
    // reader's back
    expect(slidesApi.saveAs).toHaveBeenCalledWith(`${MODEL_NAME}.pptx`)
    expect(slidesApi.save).not.toHaveBeenCalled()
  })

  it('asks from the projected text, so a withheld value cannot reach the name', async () => {
    await nameWithAi(() => ctxWith({ path: '/tmp/deck.pptx', slides: redactedDeck() }))

    const input = suggestFileName.mock.calls[0]![0] as NamingInput
    expect(input.content).not.toContain(SECRET)
    expect(input.content).toContain(LABEL)
  })

  it('leaves the dialog on the current name when the model declines', async () => {
    suggestFileName.mockResolvedValue({ ok: false, reason: 'skip-no-provider' })

    await nameWithAi(() => ctxWith({ path: '/tmp/deck.pptx' }))

    expect(slidesApi.saveAs).not.toHaveBeenCalled()
  })

  it('leaves the dialog on the current name when the model returns an empty name', async () => {
    suggestFileName.mockResolvedValue({ ok: true, name: '' })

    await nameWithAi(() => ctxWith({ path: '/tmp/deck.pptx' }))

    expect(slidesApi.saveAs).not.toHaveBeenCalled()
  })

  it('leaves the dialog on the current name when the model errors', async () => {
    suggestFileName.mockRejectedValue(new Error('provider unreachable'))

    await nameWithAi(() => ctxWith({ path: '/tmp/deck.pptx' }))

    expect(slidesApi.saveAs).not.toHaveBeenCalled()
  })
})
