import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { parseSlide } from '@genoffice/pptx-engine'
import { setRedactExt } from '@genoffice/pptx-engine'
import { buildRenderSlide, type RenderSlide } from '@genoffice/pptx-render'
import { modelSaveAsName } from '../src/renderer/file-naming'

/**
 * These go through the naming function the ribbon actually calls, not through
 * `suggestFileName` directly. A test of the shared client would still pass if
 * this app stopped using the projection — which is the failure that leaks a
 * secret, so the app's own call site is what needs covering.
 *
 * The slide XML and the render-tree construction are the ones
 * `redact-view.test.ts` uses, deliberately: that file already pins how a
 * withheld run reads through the projection, and reusing its shape here keeps
 * this one about the naming decision rather than about the layout engine.
 */

const LABEL = '\u5ba2\u6237\u7535\u8bdd'
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
  return [render(SP(run(RPR, '\u5ba2\u6237\u56de\u8bbf') + run(MARKED, SECRET)))]
}

/** A deck with nothing withheld. */
function plainDeck(): RenderSlide[] {
  return [render(SP(run(RPR, 'Quarterly review')))]
}

/** The inputs the naming bridge received, in order. */
let sent: Array<{ content: string }> = []

/** What the model answers with; '' is a model that returned nothing usable. */
let answer = '\u5ba2\u6237\u56de\u8bbf\u8bb0\u5f55'

/**
 * Install the naming bridge the preload exposes. Assigned onto the real window
 * rather than replacing it, so the jsdom DOM other code reads stays intact.
 */
function installBridge(api: {
  suggestFileName: (input: { content: string }) => Promise<unknown>
}): void {
  ;(window as unknown as { aiNamingApi: unknown }).aiNamingApi = api
}

beforeEach(() => {
  sent = []
  answer = '\u5ba2\u6237\u56de\u8bbf\u8bb0\u5f55'
  installBridge({
    suggestFileName: async (input) => {
      sent.push({ content: input.content })
      return answer ? { ok: true, name: answer } : { ok: false, reason: 'empty-name' }
    },
  })
})

afterEach(() => {
  delete (window as unknown as { aiNamingApi?: unknown }).aiNamingApi
  vi.restoreAllMocks()
})

describe('the name is derived from the projected text', () => {
  it('never hands the model a value the reader withheld', async () => {
    await modelSaveAsName(redactedDeck())
    // The one assertion that matters. The file name is written to disk, synced
    // and emailed, so a secret that reached it would outlive the redaction
    // meant to contain it.
    expect(sent).toHaveLength(1)
    expect(sent[0]!.content).not.toContain(SECRET)
    expect(JSON.stringify(sent[0]!.content)).not.toContain('13800138')
  })

  it('sends the marker the reader chose, so the model can still read the slide', async () => {
    await modelSaveAsName(redactedDeck())
    expect(sent[0]!.content).toContain(LABEL)
    expect(sent[0]!.content).toContain('\u5ba2\u6237\u56de\u8bbf')
  })

  it('sends the whole deck for one with nothing withheld', async () => {
    await modelSaveAsName(plainDeck())
    expect(sent[0]!.content).toContain('Quarterly review')
  })

  it('hands the bridge the projected content and nothing else', async () => {
    // The bridge carries one field: the text the model may see. A leftover
    // trigger or file path would be a channel the model does not need.
    const received: Array<Record<string, unknown>> = []
    installBridge({
      suggestFileName: async (input) => {
        received.push(input as unknown as Record<string, unknown>)
        return { ok: true, name: '\u5b63\u5ea6\u56de\u987e' }
      },
    })
    await modelSaveAsName(plainDeck())
    expect(Object.keys(received[0]!)).toEqual(['content'])
  })
})

describe('a name is offered only when the model has one', () => {
  it('offers the model name when one does arrive', async () => {
    expect(await modelSaveAsName(plainDeck())).toBe('\u5ba2\u6237\u56de\u8bbf\u8bb0\u5f55')
  })

  it('offers nothing when the model declines', async () => {
    installBridge({
      suggestFileName: async () => ({ ok: false, reason: 'timeout' }),
    })
    expect(await modelSaveAsName(plainDeck())).toBe('')
  })

  it('offers nothing when the model errors', async () => {
    installBridge({
      suggestFileName: async () => {
        throw new Error('provider unreachable')
      },
    })
    expect(await modelSaveAsName(plainDeck())).toBe('')
  })

  it('offers nothing when the model returns an empty name', async () => {
    answer = ''
    expect(await modelSaveAsName(plainDeck())).toBe('')
  })

  it('offers nothing when there is no naming bridge at all', async () => {
    delete (window as unknown as { aiNamingApi?: unknown }).aiNamingApi
    expect(await modelSaveAsName(plainDeck())).toBe('')
  })

  it('offers nothing for a deck with no text to work from', async () => {
    // An empty deck has no name to derive, so there is nothing worth sending.
    await modelSaveAsName([])
    expect(sent).toHaveLength(0)
  })
})

describe('the reader asks each time', () => {
  it('asks again on a second request', async () => {
    // The name is a proposal the reader may reject and ask for again, not a
    // once-per-document decision.
    await modelSaveAsName(plainDeck())
    await modelSaveAsName(plainDeck())
    expect(sent).toHaveLength(2)
  })
})
