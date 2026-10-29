import { describe, expect, it } from 'vitest'

import { buildParseMap } from '../src/renderer/document/parse-map'
import { modelTextOf } from '../src/renderer/document/redact-view'

const LABEL = 'API key'
const TEXT_SECRET = 'sk-TEXT-0000'
const ATTR_SECRET = 'sk-ATTR-1111'
const SCRIPT_SECRET = 'sk-SCRIPT-2222'

/** one page carrying all three kinds of withheld span */
const PAGE = [
  '<!doctype html><html><head>',
  `<meta data-gx-redact-content="${LABEL}" name="api-key" content="${ATTR_SECRET}">`,
  '<script>',
  `  const KEY = /*gx:redact:${LABEL}*/ "${SCRIPT_SECRET}";`,
  '</script>',
  '</head><body>',
  `<p>Call <span data-gx-redact="${LABEL}">${TEXT_SECRET}</span> now</p>`,
  '<p>Ordinary paragraph.</p>',
  '</body></html>',
].join('\n')

const SECRETS = [TEXT_SECRET, ATTR_SECRET, SCRIPT_SECRET]

const map = buildParseMap(PAGE, 1)

describe('the page as the model may read it', () => {
  it('never carries the value of a withheld span', () => {
    // The one assertion that matters. If this function can return the raw
    // source on any path, the guard downstream still protects the write but the
    // secret is already in the prompt.
    const view = modelTextOf(PAGE, map)
    for (const secret of SECRETS) expect(view).not.toContain(secret)
  })

  it('shows the label the reader chose, in place of the value', () => {
    const view = modelTextOf(PAGE, map)
    expect(view).toContain(`{{${LABEL}}}`)
    // one marker per withheld span — a duplicate reads as two separate things
    // to the write guard, which refuses the edit.
    expect(view.match(/\{\{API key\}\}/g)).toHaveLength(3)
  })

  it('leaves the rest of the page byte for byte what it was', () => {
    // The projection replaces spans in place; everything outside a mark has to
    // come back unchanged, or the model is reasoning about a page that does not
    // exist.
    const view = modelTextOf(PAGE, map)
    expect(view).toContain('<p>Ordinary paragraph.</p>')
    expect(view).toContain('<!doctype html><html><head>')
    expect(view).toContain('data-gx-redact-content="API key"')
  })

  it('returns the source unchanged when nothing is withheld', () => {
    const plain = '<!doctype html><html><body><p>Hello.</p></body></html>'
    expect(modelTextOf(plain, buildParseMap(plain, 1))).toBe(plain)
  })

  it('withholds the whole page when the parse map is unavailable', () => {
    // Without the map the marks cannot be attributed to elements, so the
    // withheld spans are unknown — and unknown is not "none". Answering with
    // the source here would hand over exactly what the reader hid.
    for (const missing of [null, undefined]) {
      const view = modelTextOf(PAGE, missing)
      for (const secret of SECRETS) expect(view).not.toContain(secret)
      expect(view).toBe('{{private}}')
    }
  })

  it('withholds the whole page when the map is not one it can read', () => {
    const broken = { version: 1, elements: 'not an array', bySid: new Map(), errorCount: 0 }
    const view = modelTextOf(PAGE, broken as never)
    for (const secret of SECRETS) expect(view).not.toContain(secret)
    expect(view).toBe('{{private}}')
  })

  it('answers with something rather than throwing on input it does not recognise', () => {
    // A throw is a leak with extra steps: the caller's fallback is the raw page.
    for (const junk of [null, undefined, 42, {}]) {
      expect(() => modelTextOf(junk as never, map)).not.toThrow()
      expect(modelTextOf(junk as never, map)).not.toContain('Ordinary paragraph')
    }
  })

  it('says nothing about a page that is not there', () => {
    expect(modelTextOf('', map)).toBe('')
  })
})
