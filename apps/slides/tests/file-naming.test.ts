import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { parseSlide } from '@genoffice/pptx-engine'
import { setRedactExt } from '@genoffice/pptx-engine'
import { buildRenderSlide, type RenderSlide } from '@genoffice/pptx-render'
import { firstSaveName, modelSaveAsName, noteDocumentSwapped } from '../src/renderer/file-naming'

/**
 * These go through the naming functions the save path and the ribbon actually
 * call, not through `nameForSave` directly. A test of the shared client would
 * still pass if this app stopped using the projection — which is the failure
 * that leaks a secret, so the app's own call site is what needs covering.
 *
 * The slide XML and the render-tree construction are the ones
 * `redact-view.test.ts` uses, deliberately: that file already pins how a
 * withheld run reads through the projection, and reusing its shape here keeps
 * this one about the naming decision rather than about the layout engine.
 */

const LABEL = '客户电话'
const SECRET = '13800138000'

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
  return [render(SP(run(RPR, 'Quarterly review')))]
}

/** Calls the naming bridge received, in order. */
let sent: Array<{ content: string; trigger: string; filePath: string | null }> = []

/** What the model answers with; '' is a model that returned nothing usable. */
let answer = '客户回访记录'
/** What the preference reads as. First-save naming consults it; manual does not. */
let namingEnabled = true

/**
 * Install the naming bridge the preload exposes. Assigned onto the real window
 * rather than replacing it, so the jsdom DOM other code reads stays intact.
 */
function installBridge(api: {
  suggestFileName: (
    input: { content: string; trigger: string; filePath?: string | null },
  ) => Promise<unknown>
  getFileNamingEnabled: () => Promise<boolean>
}): void {
  ;(window as unknown as { aiOffice: unknown }).aiOffice = api
}

beforeEach(() => {
  sent = []
  answer = '客户回访记录'
  namingEnabled = true
  // The latch is module state, deliberately: one attempt per document. Every
  // test starts from a freshly swapped document so the cases stay independent.
  noteDocumentSwapped()
  installBridge({
    suggestFileName: async (input) => {
      sent.push({
        content: input.content,
        trigger: input.trigger,
        filePath: input.filePath ?? null,
      })
      return answer ? { ok: true, name: answer } : { ok: false, reason: 'empty-name' }
    },
    getFileNamingEnabled: async () => namingEnabled,
  })
})

afterEach(() => {
  delete (window as unknown as { aiOffice?: unknown }).aiOffice
  vi.restoreAllMocks()
})

describe('the first-save name is derived from the projected text', () => {
  it('never hands the model a value the reader withheld', async () => {
    await firstSaveName(redactedDeck())
    // The one assertion that matters. The file name is written to disk, synced
    // and emailed, so a secret that reached it would outlive the redaction
    // meant to contain it.
    expect(sent).toHaveLength(1)
    expect(sent[0]!.content).not.toContain(SECRET)
    expect(JSON.stringify(sent[0]!.content)).not.toContain('13800138')
  })

  it('sends the marker the reader chose, so the model can still read the slide', async () => {
    await firstSaveName(redactedDeck())
    expect(sent[0]!.content).toContain(LABEL)
    expect(sent[0]!.content).toContain('客户回访')
  })

  it('sends the whole deck for one with nothing withheld', async () => {
    await firstSaveName(plainDeck())
    expect(sent[0]!.content).toContain('Quarterly review')
  })

  it('asks as a first-save attempt, which is the one the preference governs', async () => {
    await firstSaveName(plainDeck())
    expect(sent[0]!.trigger).toBe('first-save')
  })
})

describe('the first-save name falls back to the main-process default', () => {
  it('proposes nothing when naming is declined', async () => {
    // Slides' own name for an untitled deck is the localized default inside
    // pickDraftPath, so the fallback here is an empty proposal: the save lands
    // under exactly the name it would have had anyway.
    namingEnabled = false
    expect(await firstSaveName(plainDeck())).toBe('')
    expect(sent).toHaveLength(0)
  })

  it('proposes nothing when the model errors', async () => {
    installBridge({
      suggestFileName: async () => {
        throw new Error('provider unreachable')
      },
      getFileNamingEnabled: async () => true,
    })
    expect(await firstSaveName(plainDeck())).toBe('')
  })

  it('proposes nothing when the model times out with nothing to show', async () => {
    // The main process aborts a naming turn at 30s and answers ok:false, so a
    // timeout reaches the app as a decline rather than as a throw.
    installBridge({
      suggestFileName: async () => ({ ok: false, reason: 'timeout' }),
      getFileNamingEnabled: async () => true,
    })
    expect(await firstSaveName(plainDeck())).toBe('')
  })

  it('proposes nothing when the model returns an empty name', async () => {
    answer = ''
    expect(await firstSaveName(plainDeck())).toBe('')
  })

  it('proposes nothing when there is no naming bridge at all', async () => {
    delete (window as unknown as { aiOffice?: unknown }).aiOffice
    expect(await firstSaveName(plainDeck())).toBe('')
  })

  it('proposes nothing for a deck with no text to work from', async () => {
    // An empty deck has no name to derive, so there is nothing worth sending.
    await firstSaveName([])
    expect(sent).toHaveLength(0)
  })

  it('proposes the model name when one does arrive', async () => {
    expect(await firstSaveName(plainDeck())).toBe('客户回访记录')
  })
})

describe('one naming turn per document', () => {
  it('does not ask again on a second save', async () => {
    expect(await firstSaveName(plainDeck())).toBe('客户回访记录')
    expect(await firstSaveName(plainDeck())).toBe('')
    expect(sent).toHaveLength(1)
  })

  it('does not start a second call while the first is still in flight', async () => {
    // Two saves overlapping is what a second ⌘S during a slow turn looks like.
    const [first, second] = await Promise.all([
      firstSaveName(plainDeck()),
      firstSaveName(plainDeck()),
    ])
    expect(sent).toHaveLength(1)
    expect(first).toBe('客户回访记录')
    expect(second).toBe('')
  })

  it('asks again once a different deck is loaded', async () => {
    expect(await firstSaveName(plainDeck())).toBe('客户回访记录')
    noteDocumentSwapped()
    expect(await firstSaveName([...plainDeck(), ...plainDeck()])).toBe('客户回访记录')
    expect(sent).toHaveLength(2)
  })
})

describe('the manual name goes to the Save As dialog, projected', () => {
  it('never hands the model a value the reader withheld', async () => {
    await modelSaveAsName(redactedDeck(), '/tmp/deck.pptx')
    expect(sent).toHaveLength(1)
    expect(sent[0]!.content).not.toContain(SECRET)
    expect(sent[0]!.content).toContain(LABEL)
  })

  it('tells the main process which file is being renamed, so it can refuse a real one', async () => {
    await modelSaveAsName(plainDeck(), '/tmp/deck.pptx')
    expect(sent[0]!.filePath).toBe('/tmp/deck.pptx')
    expect(sent[0]!.trigger).toBe('manual')
  })

  it('goes through even with first-save naming switched off', async () => {
    // The reader asked for it by clicking, so the preference does not apply.
    namingEnabled = false
    expect(await modelSaveAsName(plainDeck(), '/tmp/deck.pptx')).toBe('客户回访记录')
    expect(sent).toHaveLength(1)
  })

  it('returns nothing when the model declines, leaving the dialog on the current name', async () => {
    answer = ''
    expect(await modelSaveAsName(plainDeck(), '/tmp/deck.pptx')).toBe('')
  })

  it('is repeatable: a second request asks again', async () => {
    // The latch is a first-save promise, not a rate limit — a reader who
    // changed their mind and clicked again gets a fresh answer.
    await modelSaveAsName(plainDeck(), '/tmp/deck.pptx')
    await modelSaveAsName(plainDeck(), '/tmp/deck.pptx')
    expect(sent).toHaveLength(2)
  })
})
