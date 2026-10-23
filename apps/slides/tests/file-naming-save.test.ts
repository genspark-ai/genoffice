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
import { noteDocumentSwapped } from '../src/renderer/file-naming'
import type { ActionCtx } from '../src/renderer/action-context'

/**
 * The other half of `file-naming.test.ts`: that file pins what the model is
 * asked, this one pins where the answer goes.
 *
 * A deck is named by the model only on the save of a deck that has never been
 * on disk; once it has a name, the save path leaves it alone and only the
 * ribbon button may ask again — and that answer reaches the Save As dialog,
 * not the file on disk.
 */

const LABEL = '客户电话'
const SECRET = '13800138000'
const MODEL_NAME = '客户回访记录'

const SP = (runs: string, cx = 3000000) =>
  '<p:sp><p:nvSpPr><p:cNvPr id="2" name="s"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
  `<p:spPr><a:xfrm><a:off x="100000" y="100000"/><a:ext cx="${cx}" cy="1200000"/></a:xfrm>` +
  '<a:prstGeom prst="rect"/></a:spPr>' +
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

/** The same deck before the reader's mark is committed: the words are still raw. */
function unmarkedDeck(): RenderSlide[] {
  return [render(SP(run(RPR, '客户回访') + run(RPR, SECRET)))]
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
  trigger: string
  filePath?: string | null
}

let suggestFileName: ReturnType<typeof vi.fn>
let slidesApi: { save: ReturnType<typeof vi.fn>; saveAs: ReturnType<typeof vi.fn> }

beforeEach(() => {
  // The one-attempt-per-deck latch is module state, so every case starts from a
  // freshly swapped document.
  noteDocumentSwapped()
  suggestFileName = vi.fn(async () => ({ ok: true, name: MODEL_NAME }))
  slidesApi = {
    save: vi.fn(async () => ({ ok: true })),
    saveAs: vi.fn(async () => ({ ok: true })),
  }
  vi.stubGlobal('window', {
    slidesApi,
    aiOffice: {
      suggestFileName,
      getFileNamingEnabled: vi.fn(async () => true),
      setFileNamingEnabled: vi.fn(async () => true),
    },
  })
})

afterEach(() => {
  vi.unstubAllGlobals()
  vi.restoreAllMocks()
})

describe('the first save of a never-saved deck asks the model', () => {
  it('proposes the model name to the main process', async () => {
    await save(() => ctxWith({ path: null }), true)

    expect(suggestFileName).toHaveBeenCalledTimes(1)
    expect(slidesApi.save).toHaveBeenCalledWith(MODEL_NAME)
  })

  it('asks from the projected text, so a withheld value cannot reach the name', async () => {
    await save(() => ctxWith({ path: null, slides: redactedDeck() }), true)

    const input = suggestFileName.mock.calls[0]![0] as NamingInput
    expect(input.content).not.toContain(SECRET)
    expect(input.content).toContain(LABEL)
    expect(slidesApi.save).toHaveBeenCalledWith(MODEL_NAME)
  })

  it('keeps the default name when the model declines', async () => {
    // The main process answers ok:false when no provider is configured, when
    // the preference is off, or when the model returned nothing usable.
    suggestFileName.mockResolvedValue({ ok: false, reason: 'skip-no-provider' })

    await save(() => ctxWith({ path: null }), true)

    expect(slidesApi.save).toHaveBeenCalledWith(undefined)
  })

  it('keeps the default name for a deck with nothing to name', async () => {
    await save(() => ctxWith({ path: null, slides: [] }), true)

    expect(suggestFileName).not.toHaveBeenCalled()
    expect(slidesApi.save).toHaveBeenCalledWith(undefined)
  })

  it('names the deck as it is after an open edit commits, not as it was before', async () => {
    // The reader withholds a span inside the text box they are editing: the
    // mark is carried by the edit the save triggers, so until that commit the
    // tree still holds the raw words. Reading the deck from the snapshot taken
    // before the commit would hand the model the very words the reader hid.
    let call = 0
    const getCtx = () => {
      const editing = call === 0
      call += 1
      return ctxWith({
        path: null,
        slides: editing ? unmarkedDeck() : redactedDeck(),
        editing,
      })
    }

    await save(getCtx, true)

    const input = suggestFileName.mock.calls[0]![0] as NamingInput
    expect(input.content).not.toContain(SECRET)
    // the marker is there instead, so the model still sees that a span is
    // withheld and can write around it
    expect(input.content).toContain(LABEL)
  })
})

describe('a deck that already has a name is left alone', () => {
  it('is not renamed by the save path', async () => {
    await save(() => ctxWith({ path: '/tmp/deck.pptx' }), true)

    expect(suggestFileName).not.toHaveBeenCalled()
    expect(slidesApi.save).toHaveBeenCalledWith(undefined)
  })

  it('is renamed only when the reader asks for it', async () => {
    await nameWithAi(() => ctxWith({ path: '/tmp/deck.pptx' }))

    expect(suggestFileName).toHaveBeenCalledTimes(1)
    const input = suggestFileName.mock.calls[0]![0] as NamingInput
    expect(input.trigger).toBe('manual')
    // the answer goes into the Save As dialog, not onto the file behind the
    // reader's back
    expect(slidesApi.saveAs).toHaveBeenCalledWith(`${MODEL_NAME}.pptx`)
  })

  it('leaves the dialog on the current name when the model declines', async () => {
    suggestFileName.mockResolvedValue({ ok: false, reason: 'skip-no-provider' })

    await nameWithAi(() => ctxWith({ path: '/tmp/deck.pptx' }))

    expect(slidesApi.saveAs).not.toHaveBeenCalled()
  })
})
