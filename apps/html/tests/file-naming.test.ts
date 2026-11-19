import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'

import { buildParseMap } from '../src/renderer/document/parse-map'
import { localFileName, modelSaveAsName } from '../src/renderer/file-naming'

/**
 * These go through the naming functions the ribbon actually calls, not through
 * the shared client directly. A test of the shared client would still pass if
 * this app stopped using the projection — which is the failure that leaks a
 * secret, so the app's own call site is what needs covering.
 */

const LABEL = 'API key'
const SECRET = 'sk-LIVE-9999'

const PAGE = [
  '<!doctype html><html><head>',
  '<title>Checkout rollout</title>',
  '</head><body>',
  `<h1>Checkout rollout</h1>`,
  `<p>Rotate the <span data-gx-redact="${LABEL}">${SECRET}</span> today</p>`,
  '</body></html>',
].join('\n')

/** The payloads the naming bridge received, in order. */
let sent: Array<{ content: string }> = []

/** What the model answers with; '' is a model that returned nothing usable. */
let answer = 'Checkout rollout notes'
/** Whether the bridge reports the naming as a success. */
let ok = true

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
  answer = 'Checkout rollout notes'
  ok = true
  installBridge({
    suggestFileName: async (input) => {
      sent.push({ content: input.content })
      // A failed naming carries no name of its own. The stub sends one anyway:
      // a bridge that reported a failure and attached a name must not have
      // that name used, and only a test that carries one can say so.
      if (!ok) return { ok: false, reason: 'no-provider', name: 'A name the failure did not earn' }
      return { ok: true, name: answer }
    },
  })
})

afterEach(() => {
  delete (window as unknown as { aiNamingApi?: unknown }).aiNamingApi
  vi.restoreAllMocks()
})

describe('the name the reader asked for', () => {
  it('is the name the model proposed', async () => {
    expect(await modelSaveAsName(PAGE, buildParseMap(PAGE, 1))).toBe('Checkout rollout notes')
    expect(sent).toHaveLength(1)
  })

  it('never hands the model a value the reader withheld', async () => {
    await modelSaveAsName(PAGE, buildParseMap(PAGE, 1))
    // The one assertion that matters. The file name is written to disk, synced
    // and emailed, so a secret that reached it would outlive the redaction
    // meant to contain it.
    expect(sent[0]!.content).not.toContain(SECRET)
    expect(JSON.stringify(sent[0]!.content)).not.toContain('sk-LIVE')
  })

  it('sends the marker the reader chose, so the model can still read the line', async () => {
    await modelSaveAsName(PAGE, buildParseMap(PAGE, 1))
    expect(sent[0]!.content).toContain(`{{${LABEL}}}`)
    expect(sent[0]!.content).toContain('Checkout rollout')
  })

  it('sends the whole page for one with nothing withheld', async () => {
    const plain =
      '<html><head><title>Notes</title></head><body><p>the build is green</p></body></html>'
    await modelSaveAsName(plain, buildParseMap(plain, 1))
    expect(sent[0]!.content).toContain('the build is green')
    expect(sent[0]!.content).toContain('Notes')
  })

  it('sends the text and nothing else', async () => {
    // The one channel takes the text and the text only: no trigger to branch
    // on, no path to refuse. A name is asked for, not switched on.
    await modelSaveAsName(PAGE, buildParseMap(PAGE, 1))
    expect(Object.keys(sent[0]!)).toEqual(['content'])
  })

  it('withholds everything rather than sending raw source when the map is unusable', async () => {
    // A missing or damaged map cannot say which element a mark belongs to, so
    // the projection answers "all withheld". Handing over the raw source
    // instead would be the exact leak this feature exists to prevent.
    await modelSaveAsName(PAGE, null as never)
    expect(sent[0]!.content).not.toContain(SECRET)
    expect(sent[0]!.content).not.toContain('Rotate the')
  })

  it('is repeatable: a second request asks again', async () => {
    // Nothing is latched: the reader clicked, so the reader gets an answer.
    await modelSaveAsName(PAGE, buildParseMap(PAGE, 1))
    await modelSaveAsName(PAGE, buildParseMap(PAGE, 1))
    expect(sent).toHaveLength(2)
  })
})

describe('a naming that could not happen offers no name', () => {
  it('offers nothing when the bridge reports a failure', async () => {
    ok = false
    expect(await modelSaveAsName(PAGE, buildParseMap(PAGE, 1))).toBe('')
  })

  it('offers nothing when the model returns an empty name', async () => {
    answer = ''
    expect(await modelSaveAsName(PAGE, buildParseMap(PAGE, 1))).toBe('')
  })

  it('offers nothing when the bridge throws', async () => {
    installBridge({
      suggestFileName: async () => {
        throw new Error('provider unreachable')
      },
    })
    expect(await modelSaveAsName(PAGE, buildParseMap(PAGE, 1))).toBe('')
  })

  it('offers nothing when there is no naming bridge at all', async () => {
    delete (window as unknown as { aiNamingApi?: unknown }).aiNamingApi
    expect(await modelSaveAsName(PAGE, buildParseMap(PAGE, 1))).toBe('')
  })

  it('offers nothing for a page with no text to name', async () => {
    expect(await modelSaveAsName('', buildParseMap('', 1))).toBe('')
    expect(sent).toHaveLength(0)
  })
})

describe("the page's own name", () => {
  it('is the page title', () => {
    expect(localFileName(PAGE, null)).toBe('Checkout rollout')
  })

  it('falls back to the request name for a page with no title of its own', () => {
    const untitled = '<html><body><p>body words only</p></body></html>'
    expect(localFileName(untitled, 'From the prompt')).toBe('From the prompt')
  })

  it('falls back to the first words of the page when there is neither', () => {
    const untitled = '<html><body><p>body words only</p></body></html>'
    expect(localFileName(untitled, null)).toBe('body words only')
  })
})

/**
 * The renderer source, read rather than rendered: the call sites are the claim.
 *
 * Read through Vite's `import.meta.glob` rather than from a path: the test
 * environment is a DOM, where `import.meta.url` is the page's URL and not a
 * file's, and the working directory a test runs in is not the one it names.
 */
const APP_SOURCE = import.meta.glob('../src/renderer/App.tsx', {
  query: '?raw',
  import: 'default',
  eager: true,
})['../src/renderer/App.tsx'] as string

function appSource(): string {
  return APP_SOURCE
}

/**
 * The save path's own source: the `doSave` callback, up to its dependency list.
 *
 * Read from the file rather than exercised through the component: the claim is
 * about which calls the save path makes, and a component that renders is a
 * slow and indirect way to ask that.
 */
function savePathSource(): string {
  const src = appSource()
  const start = src.indexOf('const doSave = useCallback(')
  expect(start, 'the save path was not found').toBeGreaterThan(-1)
  const end = src.indexOf('\n  )\n', start)
  expect(end, 'the end of the save path was not found').toBeGreaterThan(start)
  return src.slice(start, end)
}

describe('the save path never asks the model', () => {
  it('names a first save from the page itself', () => {
    const savePath = savePathSource()
    // The region is the save path: it has to be, or the assertions below are
    // about nothing.
    expect(savePath).toContain('window.htmlApi.save(')
    expect(savePath).toContain('localFileName(')
    expect(savePath).not.toMatch(/modelSaveAsName|suggestFileName|aiNamingApi/)
  })

  it('reaches the model through one call site, the one the reader triggers', () => {
    const src = appSource()
    expect(src.match(/modelSaveAsName\(/g) ?? []).toHaveLength(1)
    expect(src).not.toContain('suggestFileName')
    expect(src).not.toContain('aiNamingApi')
  })
})
