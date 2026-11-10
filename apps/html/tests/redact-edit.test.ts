import { describe, expect, it } from 'vitest'
import { buildParseMap, type ParseMap } from '../src/renderer/document/parse-map'
import { compileOps } from '../src/renderer/document/ops'
import { applyPatches } from '../src/renderer/document/patch'
import {
  redactClearPlan,
  redactMarkPlan,
  type RedactPlan,
} from '../src/renderer/document/redact-edit'
import { buildProjection, collectWithheld, sanitizeLabel } from '../src/renderer/ai/redact'

/**
 * The reader-facing half of withholding a span: which of the three marks a source
 * selection turns into, and what is the smallest edit that writes it.
 *
 * The property under test is never "a plan was produced" but "after the plan runs,
 * `ai/redact.ts` — the projection the model actually reads — withholds exactly the
 * selected characters, and the rest of the file is untouched". A mark that looks
 * right and withholds nothing is the failure this file exists to catch.
 */

const LABEL = 'client phone'
const SECRET = '13800138000'

const mapOf = (html: string): ParseMap => buildParseMap(html, 1, null)

/** run a plan the way the app does: ops compile, patches are spliced, then committed */
function run(html: string, plan: RedactPlan | null): string {
  if (!plan) throw new Error('expected a plan')
  if (plan.ops.length > 0) {
    const compiled = compileOps(html, mapOf(html), plan.ops)
    expect(compiled.errors).toEqual([])
    return applyPatches(html, compiled.patches)
  }
  return applyPatches(html, plan.patches)
}

/** select the first occurrence of `text` in `html` */
function spanOf(html: string, text: string): [number, number] {
  const from = html.indexOf(text)
  if (from < 0) throw new Error(`fixture does not contain ${text}`)
  return [from, from + text.length]
}

describe('withholding visible text', () => {
  const html = `<p>Call ${SECRET} now</p>`

  it('wraps the selection in a span carrying the mark, byte for byte', () => {
    const [from, to] = spanOf(html, SECRET)
    const out = run(html, redactMarkPlan(html, mapOf(html), from, to, LABEL))
    expect(out).toBe(`<p>Call <span data-gx-redact="${LABEL}">${SECRET}</span> now</p>`)
  })

  it('leaves the projection holding the marker instead of the words', () => {
    const [from, to] = spanOf(html, SECRET)
    const out = run(html, redactMarkPlan(html, mapOf(html), from, to, LABEL))
    const view = buildProjection(out, mapOf(out)).view
    expect(view).toContain(`<span data-gx-redact="${LABEL}">{{${LABEL}}}</span>`)
    expect(view).not.toContain(SECRET)
  })

  it('addresses the range through the source, not through the decoded text', () => {
    // `&amp;` is six characters in the file and one in the document: reading the
    // range on the wrong side of that gap wraps the wrong words
    const withEntity = `<p>Tom &amp; Jerry call ${SECRET}</p>`
    const [from, to] = spanOf(withEntity, SECRET)
    const out = run(withEntity, redactMarkPlan(withEntity, mapOf(withEntity), from, to, LABEL))
    expect(out).toBe(`<p>Tom &amp; Jerry call <span data-gx-redact="${LABEL}">${SECRET}</span></p>`)
  })
})

describe('withholding an attribute value', () => {
  const html = `<img src="https://cdn.example.com/logo.png" alt="Logo">`

  it('marks the attribute the reader dragged over, name and quotes included', () => {
    // in the source, selecting `src="…"` whole is the gesture, not just the value
    const [from, to] = spanOf(html, 'src="https://cdn.example.com/logo.png"')
    const out = run(html, redactMarkPlan(html, mapOf(html), from, to, 'logo url'))
    expect(out).toBe(
      `<img src="https://cdn.example.com/logo.png" alt="Logo" data-gx-redact-src="logo url">`,
    )
    const view = buildProjection(out, mapOf(out)).view
    expect(view).toContain('src="{{logo url}}"')
    // the other attribute is none of the reader's business
    expect(view).toContain('alt="Logo"')
  })

  it('marks the element content when the selection is a tag that has one', () => {
    const para = '<p title="note">hi</p>'
    const [from, to] = spanOf(para, '<p')
    const out = run(para, redactMarkPlan(para, mapOf(para), from, to, 'greeting'))
    expect(out).toBe('<p title="note" data-gx-redact="greeting">hi</p>')
    expect(collectWithheld(out, mapOf(out))).toEqual([
      expect.objectContaining({ kind: 'content', label: 'greeting' }),
    ])
  })

  it('writes nothing only when a void element carries nothing at all', () => {
    // `data-gx-redact` on an <img> is still inert — the projection drops an
    // empty range — so a void element's payload is its attributes, and a whole
    // tag selection marks every value on it. With no values there is genuinely
    // nothing to withhold and the command reports that rather than pretending.
    const bare = '<br>'
    expect(redactMarkPlan(bare, mapOf(bare), 0, bare.length, 'logo')).toBeNull()
  })
})

describe('withholding a value inside a <script>', () => {
  const html = `<script>const apiKey = "${SECRET}";</script>`

  it('puts the mark in front of the literal, not inside the quotes', () => {
    // the selection starts at the `=`, well before the opening quote: inserting
    // where the selection starts would land between the quotes, where the engine's
    // own scan finds the closing quote first and the span comes back inverted
    const [from, to] = spanOf(html, `= "${SECRET}"`)
    const out = run(html, redactMarkPlan(html, mapOf(html), from, to, 'api key'))
    expect(out).toBe(`<script>const apiKey = /*gx:redact:api key*/"${SECRET}";</script>`)
    expect(collectWithheld(out, mapOf(out))).toEqual([
      expect.objectContaining({ kind: 'literal', label: 'api key' }),
    ])
  })

  it('finds the same literal when only the value is selected', () => {
    const [from, to] = spanOf(html, SECRET)
    const out = run(html, redactMarkPlan(html, mapOf(html), from, to, 'api key'))
    expect(out).toBe(`<script>const apiKey = /*gx:redact:api key*/"${SECRET}";</script>`)
    expect(buildProjection(out, mapOf(out)).view).not.toContain(SECRET)
  })

  it('keeps a label that could close the comment from becoming code', () => {
    const [from, to] = spanOf(html, SECRET)
    const out = run(html, redactMarkPlan(html, mapOf(html), from, to, 'a*/b'))
    // sanitizeLabel is the one rule for both carriers, so the two characters are
    // dropped rather than blanked — a label is a name, and keeping one rule
    // means there is no second sanitizer to drift out of step
    expect(out).toBe(`<script>const apiKey = /*gx:redact:ab*/"${SECRET}";</script>`)
    // the script is still one comment followed by one string
    expect(collectWithheld(out, mapOf(out))).toEqual([
      expect.objectContaining({ kind: 'literal', label: 'ab' }),
    ])
  })

  it('strips the same characters from an attribute label, so both carriers agree', () => {
    // the same label reaching the same reader by the other route must read back
    // the same way, or the prompt and the file would disagree
    const attr = sanitizeLabel('a*/b')
    expect(attr).toBe('ab')
  })
})

describe('a void element', () => {
  it('withholds every value on a whole-tag selection rather than doing nothing', () => {
    // `<img src="…">` has no content, so a content mark on it is inert and the
    // projection would drop it. Silently writing nothing is the worst outcome:
    // the reader would think the command failed.
    const html = '<img src="https://cdn.example.com/logo.png" alt="company logo">'
    const from = html.indexOf('<img')
    const out = run(html, redactMarkPlan(html, mapOf(html), from, html.length, 'logo'))
    expect(out).toContain('data-gx-redact-src="logo"')
    expect(out).toContain('data-gx-redact-alt="logo"')
    const view = buildProjection(out, mapOf(out)).view
    expect(view).not.toContain('cdn.example.com')
    expect(view).not.toContain('company logo')
  })

  it('leaves a void element with no values alone', () => {
    const html = '<br>'
    expect(redactMarkPlan(html, mapOf(html), 0, html.length, 'x')).toBeNull()
  })
})

describe('a selection no mark describes', () => {
  it('marks a <style> element rather than wrapping a span inside its text', () => {
    const html = '<style>.a { color: red }</style>'
    const [from, to] = spanOf(html, 'color: red')
    const out = run(html, redactMarkPlan(html, mapOf(html), from, to, 'theme css'))
    expect(out).toBe('<style data-gx-redact="theme css">.a { color: red }</style>')
  })

  it('falls back to the element when the selection crosses markup', () => {
    const html = '<div><p>one</p><p>two</p></div>'
    const [from, to] = spanOf(html, 'one</p><p>two')
    const out = run(html, redactMarkPlan(html, mapOf(html), from, to, 'body copy'))
    expect(out).toBe('<div data-gx-redact="body copy"><p>one</p><p>two</p></div>')
  })

  it('returns nothing for a label that sanitizes away', () => {
    const html = `<p>Call ${SECRET} now</p>`
    const [from, to] = spanOf(html, SECRET)
    expect(redactMarkPlan(html, mapOf(html), from, to, '  <>  ')).toBeNull()
  })
})

describe('clearing a mark', () => {
  it('drops a whole wrapper the reader selected, text and all but the wrapper', () => {
    const html = `<p>Call <span data-gx-redact="${LABEL}">${SECRET}</span> now</p>`
    const [from, to] = spanOf(html, `<span data-gx-redact="${LABEL}">${SECRET}</span>`)
    const out = run(html, redactClearPlan(html, mapOf(html), from, to))
    expect(out).toBe(`<p>Call ${SECRET} now</p>`)
  })

  it('only removes the attribute when the selection is the marked text itself', () => {
    // deliberately conservative: unwrapping here would strip tags off a span the
    // author wrote, so the mark comes off and the element stays
    const html = `<p>Call <span data-gx-redact="${LABEL}">${SECRET}</span> now</p>`
    const [from, to] = spanOf(html, SECRET)
    const out = run(html, redactClearPlan(html, mapOf(html), from, to))
    expect(out).toBe(`<p>Call <span>${SECRET}</span> now</p>`)
  })

  it('keeps an author element and takes only the attribute off it', () => {
    const html = `<p data-gx-redact="body copy">Call ${SECRET}</p>`
    const [from, to] = spanOf(html, SECRET)
    const out = run(html, redactClearPlan(html, mapOf(html), from, to))
    expect(out).toBe(`<p>Call ${SECRET}</p>`)
  })

  it('leaves the attribute alone and drops only its mark', () => {
    const html = `<img src="https://cdn.example.com/logo.png" data-gx-redact-src="logo url">`
    const [from, to] = spanOf(html, 'https://cdn.example.com/logo.png')
    const out = run(html, redactClearPlan(html, mapOf(html), from, to))
    expect(out).toBe(`<img src="https://cdn.example.com/logo.png">`)
  })

  it('removes the comment and nothing else from a <script>', () => {
    const html = `<script>const apiKey = /*gx:redact:api key*/"${SECRET}";</script>`
    const [from, to] = spanOf(html, `"${SECRET}"`)
    const out = run(html, redactClearPlan(html, mapOf(html), from, to))
    expect(out).toBe(`<script>const apiKey = "${SECRET}";</script>`)
    expect(collectWithheld(out, mapOf(out))).toEqual([])
  })

  it('finds nothing to clear on a plain selection', () => {
    const html = `<p>Call ${SECRET} now</p>`
    const [from, to] = spanOf(html, SECRET)
    expect(redactClearPlan(html, mapOf(html), from, to)).toBeNull()
  })
})
