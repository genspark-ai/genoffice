import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'

import { buildParseMap } from '../src/renderer/document/parse-map'
import { firstSaveName, modelSaveAsName, noteDocumentSwapped } from '../src/renderer/file-naming'

/**
 * These go through the naming functions the save path and the ribbon actually
 * call, not through `nameForSave` directly. A test of the shared client would
 * still pass if this app stopped using the projection — which is the failure
 * that leaks a secret, so the app's own call site is what needs covering.
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

/** Calls the naming bridge received, in order. */
let sent: Array<{ content: string; trigger: string; filePath: string | null }> = []

/** What the model answers with; '' is a model that returned nothing usable. */
let answer = 'Checkout rollout notes'
/** What the preference reads as. First-save naming consults it; manual does not. */
let namingEnabled = true

/**
 * Install the naming bridge the preload exposes. Assigned onto the real window
 * rather than replacing it, so the jsdom DOM other code reads stays intact.
 */
function installBridge(api: {
  suggestFileName: (input: {
    content: string
    trigger: string
    filePath?: string | null
  }) => Promise<unknown>
  getFileNamingEnabled: () => Promise<boolean>
}): void {
  ;(window as unknown as { aiOffice: unknown }).aiOffice = api
}

beforeEach(() => {
  sent = []
  answer = 'Checkout rollout notes'
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
    await firstSaveName(PAGE, buildParseMap(PAGE, 1), null)
    // The one assertion that matters. The file name is written to disk, synced
    // and emailed, so a secret that reached it would outlive the redaction
    // meant to contain it.
    expect(sent).toHaveLength(1)
    expect(sent[0]!.content).not.toContain(SECRET)
    expect(JSON.stringify(sent[0]!.content)).not.toContain('sk-LIVE')
  })

  it('sends the marker the reader chose, so the model can still read the line', async () => {
    await firstSaveName(PAGE, buildParseMap(PAGE, 1), null)
    expect(sent[0]!.content).toContain(`{{${LABEL}}}`)
    expect(sent[0]!.content).toContain('Checkout rollout')
  })

  it('sends the whole page for one with nothing withheld', async () => {
    const plain =
      '<html><head><title>Notes</title></head><body><p>the build is green</p></body></html>'
    await firstSaveName(plain, buildParseMap(plain, 1), null)
    expect(sent[0]!.content).toContain('the build is green')
    expect(sent[0]!.content).toContain('Notes')
  })

  it('asks as a first-save attempt, which is the one the preference governs', async () => {
    await firstSaveName(PAGE, buildParseMap(PAGE, 1), null)
    expect(sent[0]!.trigger).toBe('first-save')
  })

  it('withholds everything rather than sending raw source when the map is unusable', async () => {
    // A missing or damaged map cannot say which element a mark belongs to, so
    // the projection answers "all withheld". Handing over the raw source
    // instead would be the exact leak this feature exists to prevent.
    await firstSaveName(PAGE, null as never, null)
    expect(sent[0]!.content).not.toContain(SECRET)
    expect(sent[0]!.content).not.toContain('Rotate the')
  })
})

describe('the first-save name falls back to the page own title', () => {
  it('keeps the title when naming is declined', async () => {
    namingEnabled = false
    expect(await firstSaveName(PAGE, buildParseMap(PAGE, 1), null)).toBe('Checkout rollout')
  })

  it('keeps the title when the model errors', async () => {
    installBridge({
      suggestFileName: async () => {
        throw new Error('provider unreachable')
      },
      getFileNamingEnabled: async () => true,
    })
    expect(await firstSaveName(PAGE, buildParseMap(PAGE, 1), null)).toBe('Checkout rollout')
  })

  it('keeps the title when the model times out with nothing to show', async () => {
    // The main process aborts a naming turn at 30s and answers ok:false, so a
    // timeout reaches the app as a decline rather than as a throw.
    installBridge({
      suggestFileName: async () => ({ ok: false, reason: 'timeout' }),
      getFileNamingEnabled: async () => true,
    })
    expect(await firstSaveName(PAGE, buildParseMap(PAGE, 1), null)).toBe('Checkout rollout')
  })

  it('keeps the title when the model returns an empty name', async () => {
    answer = ''
    expect(await firstSaveName(PAGE, buildParseMap(PAGE, 1), null)).toBe('Checkout rollout')
  })

  it('keeps the title when there is no naming bridge at all', async () => {
    delete (window as unknown as { aiOffice?: unknown }).aiOffice
    expect(await firstSaveName(PAGE, buildParseMap(PAGE, 1), null)).toBe('Checkout rollout')
  })

  it('falls back to the request name for a page with no title of its own', async () => {
    namingEnabled = false
    const untitled = '<html><body><p>body words only</p></body></html>'
    expect(await firstSaveName(untitled, buildParseMap(untitled, 1), 'From the prompt')).toBe(
      'From the prompt',
    )
  })

  it('prefers the model name when one does arrive', async () => {
    expect(await firstSaveName(PAGE, buildParseMap(PAGE, 1), null)).toBe('Checkout rollout notes')
  })
})

describe('one naming turn per document', () => {
  it('does not ask again on a second save', async () => {
    expect(await firstSaveName(PAGE, buildParseMap(PAGE, 1), null)).toBe('Checkout rollout notes')
    // A save that re-asked would put a model call behind every keystroke the
    // reader made after the first one.
    expect(await firstSaveName(PAGE, buildParseMap(PAGE, 1), null)).toBe('Checkout rollout')
    expect(sent).toHaveLength(1)
  })

  it('does not start a second call while the first is still in flight', async () => {
    // Two saves overlapping is what a second ⌘S during a slow turn looks like.
    const [first, second] = await Promise.all([
      firstSaveName(PAGE, buildParseMap(PAGE, 1), null),
      firstSaveName(PAGE, buildParseMap(PAGE, 1), null),
    ])
    expect(sent).toHaveLength(1)
    expect(first).toBe('Checkout rollout notes')
    expect(second).toBe('Checkout rollout')
  })

  it('asks again once a different document is loaded', async () => {
    expect(await firstSaveName(PAGE, buildParseMap(PAGE, 1), null)).toBe('Checkout rollout notes')
    noteDocumentSwapped()
    const other = '<html><head><title>Other</title></head><body></body></html>'
    expect(await firstSaveName(other, buildParseMap(other, 1), null)).toBe('Checkout rollout notes')
    expect(sent).toHaveLength(2)
  })
})

describe('the manual name goes to the Save As dialog, projected', () => {
  it('never hands the model a value the reader withheld', async () => {
    await modelSaveAsName(PAGE, buildParseMap(PAGE, 1), '/tmp/page.html')
    expect(sent).toHaveLength(1)
    expect(sent[0]!.content).not.toContain(SECRET)
    expect(sent[0]!.content).toContain(`{{${LABEL}}}`)
  })

  it('tells the main process which file is being renamed, so it can refuse a real one', async () => {
    await modelSaveAsName(PAGE, buildParseMap(PAGE, 1), '/tmp/page.html')
    expect(sent[0]!.filePath).toBe('/tmp/page.html')
    expect(sent[0]!.trigger).toBe('manual')
  })

  it('goes through even with first-save naming switched off', async () => {
    // The reader asked for it by clicking, so the preference does not apply.
    namingEnabled = false
    expect(await modelSaveAsName(PAGE, buildParseMap(PAGE, 1), '/tmp/page.html')).toBe(
      'Checkout rollout notes',
    )
    expect(sent).toHaveLength(1)
  })

  it('returns nothing when the model declines, leaving the dialog on the current name', async () => {
    answer = ''
    expect(await modelSaveAsName(PAGE, buildParseMap(PAGE, 1), '/tmp/page.html')).toBe('')
  })

  it('is repeatable: a second request asks again', async () => {
    // The latch is a first-save promise, not a rate limit — a reader who
    // changed their mind and clicked again gets a fresh answer.
    await modelSaveAsName(PAGE, buildParseMap(PAGE, 1), '/tmp/page.html')
    await modelSaveAsName(PAGE, buildParseMap(PAGE, 1), '/tmp/page.html')
    expect(sent).toHaveLength(2)
  })
})
