import { describe, expect, it } from 'vitest'
import { buildParseMap, type ParseMap } from '../src/renderer/document/parse-map'
import { compileOps, type HtmlOp } from '../src/renderer/document/ops'
import { applyPatches } from '../src/renderer/document/patch'
import { createHtmlSkillCore, type HtmlDocAccess } from '../src/renderer/ai/tools'
import { createDocumentSkill } from '../src/renderer/ai/html-skill'

/**
 * The one test that matters: **no surface a model can read may hand it a withheld
 * value.**
 *
 * There are four of them and they were found separately, which is exactly why
 * they are all asserted here. `buildContext` used to embed the whole file,
 * `get_outline` prints a text and an attribute preview per element, and
 * `read_source` addresses by line — and the line numbers it prints belong to
 * whichever text it was shown, so a projection that shortened a value has to
 * renumber with it.
 *
 * `apply_ops` is the other direction and is asserted separately below: the model
 * not seeing something is no use if it can still delete it.
 */

const LABEL = 'API key'
const TEXT_SECRET = 'sk-TEXT-0000'
const ATTR_SECRET = 'sk-ATTR-1111'
const SCRIPT_SECRET = 'sk-SCRIPT-2222'

/** one page carrying all three kinds, on separate lines so line reads are distinct */
const PAGE = [
  '<!doctype html><html><head>',
  `<meta name="api-key" content="${ATTR_SECRET}" ${'data-gx-redact-meta'}:none>`,
  '<script>',
  `  const KEY = /*gx:redact:${LABEL}*/ "${SCRIPT_SECRET}";`,
  '  render(KEY);',
  '</script>',
  '</head><body>',
  `<p>Call <span data-gx-redact="${LABEL}">${TEXT_SECRET}</span> now</p>`,
  '<p>Ordinary paragraph.</p>',
  '</body></html>',
].join('\n')

/** the meta mark is added here so the fixture reads as one page a reader wrote */
const FIXTURE = PAGE.replace(' data-gx-redact-meta:none', '').replace(
  '<meta name="api-key"',
  `<meta data-gx-redact-content="${LABEL}" name="api-key"`,
)

const SECRETS = [TEXT_SECRET, ATTR_SECRET, SCRIPT_SECRET]

function fakeAccess(initial: string) {
  let text = initial
  let version = 1
  const lastManual = 0
  let map: ParseMap | null = null
  const getMap = () => {
    if (!map || map.version !== version) map = buildParseMap(text, version, map)
    return map
  }
  const access: HtmlDocAccess = {
    getText: () => text,
    getVersion: () => version,
    getMap,
    getLastManualVersion: () => lastManual,
    getFilePath: () => '/tmp/page.html',
    getSelectedSid: () => null,
    applyOps: (ops: HtmlOp[]) => {
      const compiled = compileOps(text, getMap(), ops)
      if (compiled.errors.length) return { ok: false, errors: compiled.errors }
      text = applyPatches(text, compiled.patches)
      version++
      return { ok: true, ranges: [] }
    },
    replaceAll: (html) => {
      text = html
      version++
    },
  } as unknown as HtmlDocAccess
  return {
    access,
    getText: () => text,
    setText: (t: string) => {
      text = t
      version++
    },
  }
}

const call = async (core: ReturnType<typeof createHtmlSkillCore>, tool: string, input: unknown) => {
  const r = await core.executeTool({ name: tool, input } as never)
  return typeof r === 'string' ? r : r.output
}

describe('no read surface hands the model a withheld value', () => {
  it('buildContext — the whole-file preview, the outline and the selection', async () => {
    const { access } = fakeAccess(FIXTURE)
    const out = createHtmlSkillCore(access).buildContext()
    for (const secret of SECRETS) expect(out, secret).not.toContain(secret)
    expect(out).toContain(`{{${LABEL}}}`)
  })

  it('get_outline — the per-element text preview and the src/href preview', async () => {
    const { access } = fakeAccess(FIXTURE)
    const out = await call(createHtmlSkillCore(access), 'get_outline', { depth: 4 })
    for (const secret of SECRETS) expect(out, secret).not.toContain(secret)
  })

  it('read_source by line range', async () => {
    const { access } = fakeAccess(FIXTURE)
    const out = await call(createHtmlSkillCore(access), 'read_source', {
      start_line: 1,
      end_line: 9,
    })
    for (const secret of SECRETS) expect(out, secret).not.toContain(secret)
  })

  it('read_source whole file', async () => {
    const { access } = fakeAccess(FIXTURE)
    const out = await call(createHtmlSkillCore(access), 'read_source', {})
    for (const secret of SECRETS) expect(out, secret).not.toContain(secret)
  })

  it('read_source by element id', async () => {
    const { access } = fakeAccess(FIXTURE)
    const map = access.getMap()
    const para = map.elements.find((e) => e.tag === 'p' && e.inner[1] - e.inner[0] > 40)
    expect(para, 'the paragraph with the withheld span').toBeTruthy()
    const out = await call(createHtmlSkillCore(access), 'read_source', { sid: para!.sid })
    expect(out).not.toContain(TEXT_SECRET)
  })

  it('the system prompt names the placeholder and not the value', () => {
    const { access } = fakeAccess(FIXTURE)
    const prompt = createDocumentSkill(access).systemPrompt
    for (const secret of SECRETS) expect(prompt, secret).not.toContain(secret)
    expect(prompt).toContain(`{{${LABEL}}}`)
  })

  it('a page with nothing withheld costs nothing and says nothing', () => {
    const plain = '<html><body><p>Nothing secret here</p></body></html>'
    const { access } = fakeAccess(plain)
    const prompt = createDocumentSkill(access).systemPrompt
    expect(prompt).not.toContain('{{')
    expect(createHtmlSkillCore(access).buildContext()).not.toContain('{{')
  })
})

describe('line numbering follows the text the model was shown', () => {
  it('a range that ends after a shortening substitution is not truncated', async () => {
    const { access } = fakeAccess(FIXTURE)
    const core = createHtmlSkillCore(access)
    // A marker is always single-line, so a substitution never changes the line
    // COUNT — it changes the width of the lines it sits on. The end offset of
    // any later line therefore moves, and a range that ends past a marked line
    // comes back short if the view's offset is used as the file's.
    const lastLine = access.getText().split('\n').length
    const out = await call(core, 'read_source', { start_line: 1, end_line: lastLine })
    expect(out).toContain('Ordinary paragraph.')
    expect(out.trimEnd().endsWith('</html>')).toBe(true)
    expect(out).toContain(String(lastLine).padStart(5) + '| </body></html>')
  })

  it('the numbers it prints match the view, not the file', async () => {
    const { access } = fakeAccess(FIXTURE)
    const core = createHtmlSkillCore(access)
    const all = await call(core, 'read_source', {})
    // the marker is shorter than the value it replaced, so the file's and the
    // view's line counts can only agree here if the offsets were translated
    expect(all.split('\n')[0]).toMatch(/^\s*1\|/)
  })
})

describe('the model cannot destroy a span either', () => {
  const opsCore = async () => {
    const { access } = fakeAccess(FIXTURE)
    const core = createHtmlSkillCore(access)
    // the staleness check compares against the last version the model SAW, so a
    // write has to be preceded by a read — as it is in any real session
    await call(core, 'get_outline', { depth: 1 })
    return { access, core }
  }

  it('refuses a str_replace whose needle is the withheld value', async () => {
    const { core } = await opsCore()
    const out = await call(core, 'apply_ops', {
      ops: [{ op: 'str_replace', old: SCRIPT_SECRET, new: 'x' }],
    })
    expect(out).toContain('0 of 1 ops applied')
    expect(out).toContain('withheld')
  })

  it('refuses an op on the element that carries the mark', async () => {
    const { access, core } = await opsCore()
    const map = access.getMap()
    const span = map.elements.find((e) => e.tag === 'span' && e.startTag[1] - e.startTag[0] > 20)!
    const out = await call(core, 'apply_ops', {
      ops: [{ op: 'set_text_node', sid: span.sid, index: 0, text: 'gone' }],
    })
    expect(out).toContain('0 of 1 ops applied')
  })

  it('refuses removing the paragraph that holds the mark', async () => {
    const { access, core } = await opsCore()
    const map = access.getMap()
    const para = map.elements.find((e) => e.tag === 'p' && e.inner[1] - e.inner[0] > 40)!
    const out = await call(core, 'apply_ops', { ops: [{ op: 'remove', sid: para.sid }] })
    expect(out).toContain('0 of 1 ops applied')
  })

  it('refuses an op on the <head> that contains a marked script', async () => {
    const { access, core } = await opsCore()
    const map = access.getMap()
    const head = map.elements.find((e) => e.tag === 'head')!
    const out = await call(core, 'apply_ops', {
      ops: [{ op: 'set_inner_html', sid: head.sid, html: '' }],
    })
    expect(out).toContain('0 of 1 ops applied')
  })

  it('allows a str_replace that rewrites the sentence AROUND the marker', async () => {
    const { core } = await opsCore()
    const out = await call(core, 'apply_ops', {
      ops: [{ op: 'str_replace', old: ' now</p>', new: ' before Friday.</p>' }],
    })
    expect(out).not.toContain('0 of 1 ops applied')
  })

  it('allows an op on an element that merely sits next to a mark', async () => {
    const { access, core } = await opsCore()
    const map = access.getMap()
    // the paragraph that merely sits NEXT to a mark: the other <p> on the page
    const plain = map.elements.find((e) => {
      if (e.tag !== 'p') return false
      return access.getText().slice(e.range[0], e.range[1]).includes('Ordinary')
    })!
    const out = await call(core, 'apply_ops', {
      ops: [{ op: 'set_style', sid: plain.sid, styles: { color: '#333333' } }],
    })
    expect(out).not.toContain('0 of 1 ops applied')
  })
})
