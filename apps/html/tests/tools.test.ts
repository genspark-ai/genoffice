import { describe, expect, it } from 'vitest'
import { buildParseMap, type ParseMap } from '../src/renderer/document/parse-map'
import { compileOps, type HtmlOp } from '../src/renderer/document/ops'
import { applyPatches } from '../src/renderer/document/patch'
import {
  createHtmlSkillCore,
  buildOutline,
  expandAttachmentRefs,
  findTruncatedDataUrl,
  type HtmlDocAccess,
} from '../src/renderer/ai/tools'
import type { Brief } from '../src/renderer/document/brief'
import type { PageWriteSpec } from '../src/renderer/ai/page-writer'
import type { BriefPlanSpec } from '../src/renderer/ai/brief-writer'

/** in-memory stand-in for the App: same compile → apply → version bookkeeping */
function fakeAccess(initial: string) {
  let text = initial
  let version = 1
  let lastManual = 0
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
      const ranges: Array<[number, number]> = []
      let delta = 0
      for (const p of [...compiled.patches].sort((a, b) => a.from - b.from)) {
        ranges.push([p.from + delta, p.from + delta + p.text.length])
        delta += p.text.length - (p.to - p.from)
      }
      text = applyPatches(text, compiled.patches)
      version++
      return { ok: true, ranges }
    },
    replaceAll: (html) => {
      text = html
      version++
    },
  }
  return {
    access,
    userTypes(next: string) {
      text = next
      version++
      lastManual = version
    },
    get text() {
      return text
    },
  }
}

const DOC = `<!doctype html>
<html>
<head><title>Report</title></head>
<body>
<h1 id="t">Quarterly</h1>
<section class="kpis">
  <div class="card">Revenue</div>
  <div class="card">Margin</div>
</section>
<p>Closing note.</p>
</body>
</html>`

const call = (name: string, input: Record<string, unknown>) => ({ id: 'c1', name, input })
/** the brief writer stub: plan_page hands the drafted JSON to the card */
const drafts = (access: HtmlDocAccess, raw: Record<string, unknown>) => {
  access.planBrief = async () => ({ ok: true, raw })
}

describe('html skill tools', () => {
  it('context and outline list sids with line ranges and previews', () => {
    const doc = fakeAccess(DOC)
    const core = createHtmlSkillCore(doc.access)
    const ctx = core.buildContext()
    expect(ctx).toContain('title: Report')
    expect(ctx).toMatch(/sid=\d+\s+h1#t\s+"Quarterly"/)
    const outline = buildOutline(DOC, buildParseMap(DOC, 1), { depth: 4 })
    expect(outline).toMatch(/L6-L9\s+sid=\d+\s+section\.kpis/)
    expect(outline).toContain('div.card  "Revenue"')
  })

  it('read_source addresses by sid and by lines with line numbers', () => {
    const doc = fakeAccess(DOC)
    const core = createHtmlSkillCore(doc.access)
    const sid = doc.access.getMap().elements.find((e) => e.tag === 'p')!.sid
    const bySid = core.executeTool(call('read_source', { sid }))
    expect(bySid.output).toBe('   10| <p>Closing note.</p>')
    const byLines = core.executeTool(call('read_source', { start_line: 5, end_line: 6 }))
    expect(byLines.output.split('\n')).toEqual([
      '    5| <h1 id="t">Quarterly</h1>',
      '    6| <section class="kpis">',
    ])
  })

  it('apply_ops applies a mixed batch atomically and reports locations', () => {
    const doc = fakeAccess(DOC)
    const core = createHtmlSkillCore(doc.access)
    core.buildContext()
    const map = doc.access.getMap()
    const h1 = map.elements.find((e) => e.tag === 'h1')!.sid
    const cards = map.elements.filter((e) => e.tag === 'div')
    const r = core.executeTool(
      call('apply_ops', {
        ops: [
          { op: 'set_text', sid: h1, text: 'Q3 Results' },
          { op: 'set_attr', sid: cards[1]!.sid, name: 'data-metric', value: 'margin' },
          { op: 'str_replace', old: 'Closing note.', new: 'Closing remarks.' },
        ],
        summary: 'tidy',
      }),
    )
    expect(r.isError).toBeFalsy()
    expect(r.mutated).toBe(true)
    expect(r.output).toMatch(/Applied 3 op\(s\) at L5, L8, L10/)
    expect(doc.text).toContain('<h1 id="t">Q3 Results</h1>')
    expect(doc.text).toContain('<div class="card" data-metric="margin">Margin</div>')
    expect(doc.text).toContain('Closing remarks.')
  })

  it('a failing op rejects the whole batch with per-op reasons', () => {
    const doc = fakeAccess(DOC)
    const core = createHtmlSkillCore(doc.access)
    core.buildContext()
    const before = doc.text
    const r = core.executeTool(
      call('apply_ops', {
        ops: [
          { op: 'str_replace', old: 'Closing note.', new: 'x' },
          { op: 'str_replace', old: '<div class="card">', new: '<div class="card big">' },
        ],
      }),
    )
    expect(r.isError).toBe(true)
    expect(r.output).toContain('0 of 2 ops applied')
    expect(r.output).toMatch(
      /Op 2\/2 FAILED \(AMBIGUOUS\): old text matches 2 locations \(lines 7, 8\)/,
    )
    expect(doc.text).toBe(before)
  })

  it('refuses to edit after the user typed since the model last looked', () => {
    const doc = fakeAccess(DOC)
    const core = createHtmlSkillCore(doc.access)
    core.buildContext()
    doc.userTypes(DOC.replace('Closing note.', 'Closing NOTE.'))
    const r = core.executeTool(
      call('apply_ops', { ops: [{ op: 'str_replace', old: 'Quarterly', new: 'Q' }] }),
    )
    expect(r.isError).toBe(true)
    expect(r.output).toMatch(/changed since you last looked/)
    // re-reading clears the staleness
    core.executeTool(call('get_outline', {}))
    const ok = core.executeTool(
      call('apply_ops', { ops: [{ op: 'str_replace', old: 'Quarterly', new: 'Q' }] }),
    )
    expect(ok.isError).toBeFalsy()
  })

  it('context tells an empty document how it gets created', () => {
    const core = createHtmlSkillCore(fakeAccess('').access)
    expect(core.buildContext()).toContain('The document is empty')
    expect(core.buildContext()).toContain('plan_page')
  })
})

describe('generation flow tools', () => {
  it('ask_clarification hands the card answers (or a skip) back to the model', async () => {
    const doc = fakeAccess('')
    let seen: unknown = null
    doc.access.askClarification = async (questions) => {
      seen = questions
      return { answers: 'Audience: board\nTone: executive' }
    }
    const core = createHtmlSkillCore(doc.access)
    const r = await core.executeTool(
      call('ask_clarification', {
        questions: [
          { id: 'aud', label: 'Audience?', options: ['board', 'team', 'customers'] },
          { label: 'no id', options: ['a', 'b', 'c', 'd', 'e', 'f', 'g'] },
          { id: 'empty', label: 'no options', options: [] },
        ],
      }),
    )
    expect(r.isError).toBeFalsy()
    expect(r.output).toContain('Audience: board')
    const qs = seen as Array<{ id: string; options: string[] }>
    expect(qs).toHaveLength(2)
    expect(qs[1]!.id).toBe('q2')
    expect(qs[1]!.options).toHaveLength(5)
    doc.access.askClarification = async () => ({ answers: '', cancelled: true })
    const skipped = await createHtmlSkillCore(doc.access).executeTool(
      call('ask_clarification', { questions: [{ id: 'a', label: 'x', options: ['1'] }] }),
    )
    expect(skipped.output).toMatch(/skipped/)
  })

  it('plan_page drafts the brief through the writer and relays confirm / redo / dismiss decisions', async () => {
    const doc = fakeAccess('')
    const core = createHtmlSkillCore(doc.access)
    core.buildContext()
    doc.access.confirmBrief = async (brief) => ({ kind: 'confirmed', brief })
    doc.access.planBrief = async () => ({ ok: false, error: 'connection dropped' })
    const nothing = await core.executeTool(call('plan_page', { mode: 'extract' }))
    expect(nothing.isError).toBe(true)
    expect(nothing.output).toContain('connection dropped')
    drafts(doc.access, { core_hook: '', style: { tone: 'x' }, sections: [] })
    const bad = await core.executeTool(call('plan_page', { mode: 'extract' }))
    expect(bad.isError).toBe(true)
    expect(bad.output).toContain('core_hook is required')
    const proposal = {
      core_hook: 'Costs fell 20% while output doubled',
      style: {
        tone: 'executive',
        palette: { primary: '#123456' },
        typography: { heading: 'Inter' },
      },
      alternatives: [
        { tone: 'playful', name: 'Candy grid', tokens: { radius: '16px' } },
        { tone: '' },
        { tone: 'dark', name: 'Night' },
        { tone: 'fourth', name: 'Dropped' },
      ],
      sections: [
        { title: 'Summary', brief: 'three numbers' },
        { title: 'Detail', brief: 'table' },
      ],
      meta: { audience: 'board' },
    }
    let spec: BriefPlanSpec | undefined
    doc.access.planBrief = async (s) => {
      spec = s
      return { ok: true, raw: proposal }
    }
    let shown: Brief | undefined
    doc.access.confirmBrief = async (brief) => {
      shown = brief
      return {
        kind: 'confirmed',
        brief: { ...brief, core_hook: 'Edited hook', user_edited: ['core_hook'] },
      }
    }
    const ok = await core.executeTool(
      call('plan_page', { mode: 'extract', notes: 'board wants numbers' }),
    )
    expect(ok.isError).toBeFalsy()
    // the writer gets the mode, the notes and the page to read
    expect(spec?.mode).toBe('extract')
    expect(spec?.notes).toBe('board wants numbers')
    expect(spec?.page).toBe('')
    // invalid directions are dropped and at most two alternatives survive; docx_friendly defaults off
    expect(shown?.alternatives?.map((s) => s.name)).toEqual(['Candy grid', 'Night'])
    expect(shown?.alternatives?.[0]?.tokens?.radius).toBe('16px')
    expect(shown?.style.docx_friendly).toBe(false)
    expect(ok.output).toContain('Fields the user edited (keep them exactly): core_hook')
    expect(ok.output).toContain('Edited hook')
    expect(ok.output).toContain('pinned')
    doc.access.confirmBrief = async () => ({ kind: 'redo', note: 'warmer colors' })
    const redo = await core.executeTool(call('plan_page', { mode: 'extract' }))
    expect(redo.output).toContain('different brief: warmer colors')
    // the rejected proposal travels back so the next notes can say what must change
    expect(redo.output).toContain('Costs fell 20%')
    doc.access.confirmBrief = async () => ({ kind: 'cancelled' })
    const cancelled = await core.executeTool(call('plan_page', { mode: 'extract' }))
    expect(cancelled.output).toMatch(/dismissed/)
  })

  it('plan_page in new mode hands the confirmed brief to the page writer and lands the page', async () => {
    const doc = fakeAccess('')
    let spec: PageWriteSpec | undefined
    doc.access.confirmBrief = async (brief) => ({ kind: 'confirmed', brief })
    doc.access.getInstruction = () => 'landing page for Acme'
    doc.access.writePage = async (s) => {
      spec = s
      return {
        ok: true,
        html: '<!doctype html><html><body><section data-section="hero">Hi</section></body></html>',
      }
    }
    drafts(doc.access, {
      core_hook: 'Ship in a day',
      style: { tone: 'technical' },
      sections: [{ title: 'Hero', brief: 'x' }],
      context: 'Acme sells widgets since 1999',
    })
    const core = createHtmlSkillCore(doc.access)
    const r = await core.executeTool(call('plan_page', { mode: 'new' }))
    expect(r.isError).toBeFalsy()
    expect(r.mutated).toBe(true)
    expect(doc.text).toContain('data-section="hero"')
    expect(r.output).toContain('Page written by the system')
    expect(r.output).toContain('sections: hero')
    expect(r.output).not.toContain('<section')
    expect(spec?.kind).toBe('design')
    expect(spec?.instruction).toBe('landing page for Acme')
    expect(spec?.context).toBe('Acme sells widgets since 1999')
    if (spec?.kind === 'design') expect(spec.brief.core_hook).toBe('Ship in a day')
  })

  it('plan_page passes structured reference material to the page writer as text', async () => {
    const doc = fakeAccess('')
    let spec: PageWriteSpec | undefined
    doc.access.confirmBrief = async (brief) => ({ kind: 'confirmed', brief })
    doc.access.writePage = async (s) => {
      spec = s
      return { ok: true, html: '<html></html>' }
    }
    drafts(doc.access, {
      core_hook: 'h',
      style: { tone: 't' },
      sections: [{ title: 'A', brief: 'x' }],
      context: { facts: ['founded 1999'], sources: [{ url: 'https://acme.test' }] },
    })
    await createHtmlSkillCore(doc.access).executeTool(call('plan_page', { mode: 'new' }))
    expect(spec?.context).toContain('founded 1999')
    expect(spec?.context).toContain('https://acme.test')
    expect(spec?.context).not.toContain('[object Object]')
  })

  it('plan_page reports a failed or partial page write', async () => {
    const doc = fakeAccess('')
    doc.access.confirmBrief = async (brief) => ({ kind: 'confirmed', brief })
    drafts(doc.access, {
      core_hook: 'h',
      style: { tone: 't' },
      sections: [{ title: 'A', brief: 'x' }],
    })
    const proposal = { mode: 'new' }
    doc.access.writePage = async () => ({ ok: false, error: 'connection dropped' })
    const failed = await createHtmlSkillCore(doc.access).executeTool(call('plan_page', proposal))
    expect(failed.isError).toBe(true)
    expect(failed.output).toContain('connection dropped')
    expect(doc.text).toBe('')
    doc.access.writePage = async () => ({ ok: true, html: '<html><body>half', truncated: true })
    const partial = await createHtmlSkillCore(doc.access).executeTool(call('plan_page', proposal))
    expect(partial.mutated).toBe(true)
    expect(partial.output).toContain('INCOMPLETE')
    expect(doc.text).toBe('<html><body>half')
  })

  it('plan_page without a mode writes an empty document and only pins on an existing one', async () => {
    const proposal = {
      core_hook: 'h',
      style: { tone: 't' },
      sections: [{ title: 'A', brief: 'x' }],
    }
    const empty = fakeAccess('')
    let writes = 0
    drafts(empty.access, proposal)
    empty.access.confirmBrief = async (brief) => ({ kind: 'confirmed', brief })
    empty.access.writePage = async () => {
      writes++
      return { ok: true, html: '<html></html>' }
    }
    await createHtmlSkillCore(empty.access).executeTool(call('plan_page', {}))
    expect(writes).toBe(1)
    const existing = fakeAccess(DOC)
    let spec: BriefPlanSpec | undefined
    existing.access.planBrief = async (s) => {
      spec = s
      return { ok: true, raw: proposal }
    }
    existing.access.confirmBrief = async (brief) => ({ kind: 'confirmed', brief })
    existing.access.writePage = async () => {
      writes++
      return { ok: true, html: '<html></html>' }
    }
    const existingCore = createHtmlSkillCore(existing.access)
    existingCore.buildContext()
    const r = await existingCore.executeTool(call('plan_page', {}))
    expect(writes).toBe(1)
    expect(r.output).toContain('extract')
    expect(spec?.page).toContain('<title>Report</title>')
    expect(existing.text.replace(/\n<meta name="genoffice:brief"[^>]*>/, '')).toBe(DOC)
  })

  it('write_document routes a content plan to the page writer', async () => {
    const doc = fakeAccess('')
    let spec: PageWriteSpec | undefined
    doc.access.writePage = async (s) => {
      spec = s
      return { ok: true, html: '<html><body><h1>Guide</h1></body></html>' }
    }
    const core = createHtmlSkillCore(doc.access)
    const bad = await core.executeTool(call('write_document', { plan: '' }))
    expect(bad.isError).toBe(true)
    const r = await core.executeTool(
      call('write_document', { title: 'Guide', plan: '1. intro 2. steps', context: 'facts' }),
    )
    expect(r.mutated).toBe(true)
    expect(spec?.kind).toBe('content')
    if (spec?.kind === 'content') {
      expect(spec.title).toBe('Guide')
      expect(spec.plan).toBe('1. intro 2. steps')
    }
    expect(doc.text).toContain('<h1>Guide</h1>')
  })

  it('plan_page in extract mode pins the brief into the open document with one edit', async () => {
    const doc = fakeAccess(DOC)
    doc.access.confirmBrief = async (brief) => ({ kind: 'confirmed', brief })
    const core = createHtmlSkillCore(doc.access)
    core.buildContext()
    drafts(doc.access, {
      core_hook: 'Quarterly numbers',
      style: { tone: 'plain' },
      sections: [{ title: 'Body', brief: 'x' }],
    })
    const proposal = { mode: 'extract' }
    const r = await core.executeTool(call('plan_page', proposal))
    expect(r.mutated).toBe(true)
    expect(r.output).toContain('now pinned')
    expect(doc.text).toMatch(/<head>\n<meta name="genoffice:brief"[^>]*><title>Report<\/title>/)
    expect(doc.text.replace(/\n<meta name="genoffice:brief"[^>]*>/, '')).toBe(DOC)
    const again = await core.executeTool(call('plan_page', proposal))
    expect(again.mutated).toBeFalsy()
    expect(again.output).toContain('already current')
  })

  it('pinning falls back to a whole-document write that keeps edits flushed by applyOps', async () => {
    const doc = fakeAccess(DOC)
    doc.access.confirmBrief = async (brief) => ({ kind: 'confirmed', brief })
    const realApply = doc.access.applyOps
    // the app wrapper lands pending live edits before compiling; simulate that landing plus a failed op
    doc.access.applyOps = (ops) => {
      realApply([{ op: 'str_replace', old: 'Quarterly', new: 'Quarterly (live)' }])
      return {
        ok: false,
        errors: [{ index: 0, kind: 'not_found', message: `gone: ${ops.length} op` }],
      }
    }
    drafts(doc.access, {
      core_hook: 'h',
      style: { tone: 't' },
      sections: [{ title: 'A', brief: 'x' }],
    })
    const core = createHtmlSkillCore(doc.access)
    core.buildContext()
    const r = await core.executeTool(call('plan_page', { mode: 'extract' }))
    expect(r.mutated).toBe(true)
    expect(doc.text).toContain('Quarterly (live)')
    expect(doc.text).toMatch(/<meta name="genoffice:brief"/)
  })

  it('plan_page refuses to overwrite text the user typed while the brief card was open', async () => {
    const doc = fakeAccess('')
    let writes = 0
    doc.access.writePage = async () => {
      writes++
      return { ok: true, html: '<html></html>' }
    }
    doc.access.confirmBrief = async (brief) => {
      doc.userTypes('<p>my draft</p>')
      return { kind: 'confirmed', brief }
    }
    drafts(doc.access, {
      core_hook: 'h',
      style: { tone: 't' },
      sections: [{ title: 'A', brief: 'x' }],
    })
    const r = await createHtmlSkillCore(doc.access).executeTool(call('plan_page', { mode: 'new' }))
    expect(r.isError).toBe(true)
    expect(writes).toBe(0)
    expect(doc.text).toBe('<p>my draft</p>')
  })

  it('context surfaces a pinned brief', () => {
    const html =
      '<html><head><meta charset="utf-8"><meta name="genoffice:brief" content=\'{"core_hook":"Ship faster","style":{"tone":"technical","palette":{},"typography":{}},"sections":[{"title":"Why","brief":"x"}],"version":1}\'></head><body><p>x</p></body></html>'
    const core = createHtmlSkillCore(fakeAccess(html).access)
    const ctx = core.buildContext()
    expect(ctx).toContain('## Brief')
    expect(ctx).toContain('core hook: Ship faster')
    expect(ctx).toContain('sections: Why')
  })
})

describe('attachment:// references in apply_ops', () => {
  const resolver = async (name: string) =>
    name === 'logo.png'
      ? ({ ok: true, src: 'assets/logo.png' } as const)
      : ({ ok: false, error: `no attachment named "${name}"` } as const)

  it('expands refs in html, value and style urls before applying', async () => {
    const doc = fakeAccess(DOC)
    doc.access.resolveAttachmentSrc = resolver
    const core = createHtmlSkillCore(doc.access)
    core.buildContext()
    const map = doc.access.getMap()
    const h1 = map.elements.find((e) => e.tag === 'h1')!.sid
    const section = map.elements.find((e) => e.tag === 'section')!.sid
    const r = await core.executeTool(
      call('apply_ops', {
        ops: [
          {
            op: 'insert_html',
            sid: h1,
            position: 'before',
            html: '<img src="attachment://logo.png" alt="logo" width="120">',
          },
          {
            op: 'set_style',
            sid: section,
            styles: { 'background-image': 'url(attachment://logo.png)' },
          },
        ],
        summary: 'embed the logo',
      }),
    )
    expect(r.isError).toBeFalsy()
    expect(r.mutated).toBe(true)
    expect(doc.text).toContain('<img src="assets/logo.png" alt="logo" width="120">')
    expect(doc.text).toContain('url(assets/logo.png)')
    expect(doc.text).not.toContain('attachment://')
  })

  it('an unresolved ref fails the whole batch with the resolver message', async () => {
    const doc = fakeAccess(DOC)
    doc.access.resolveAttachmentSrc = resolver
    const core = createHtmlSkillCore(doc.access)
    core.buildContext()
    const h1 = doc.access.getMap().elements.find((e) => e.tag === 'h1')!.sid
    const r = await core.executeTool(
      call('apply_ops', {
        ops: [
          { op: 'set_text', sid: h1, text: 'New title' },
          { op: 'set_attr', sid: h1, name: 'data-logo', value: 'attachment://missing.png' },
        ],
      }),
    )
    expect(r.isError).toBe(true)
    expect(r.output).toContain('no attachment named "missing.png"')
    expect(doc.text).toContain('Quarterly')
  })

  it('without a resolver the reference stays literal text (no crash)', () => {
    const doc = fakeAccess(DOC)
    const core = createHtmlSkillCore(doc.access)
    core.buildContext()
    const h1 = doc.access.getMap().elements.find((e) => e.tag === 'h1')!.sid
    const r = core.executeTool(
      call('apply_ops', {
        ops: [{ op: 'set_attr', sid: h1, name: 'data-x', value: 'attachment://logo.png' }],
      }),
    ) as { isError?: boolean }
    expect(r.isError).toBeFalsy()
    expect(doc.text).toContain('data-x="attachment://logo.png"')
  })

  it('a manual edit while the image is being read makes the batch stale', async () => {
    const doc = fakeAccess(DOC)
    doc.access.resolveAttachmentSrc = async () => {
      doc.userTypes(DOC.replace('Quarterly', 'Typed meanwhile'))
      return { ok: true, src: 'assets/logo.png' }
    }
    const core = createHtmlSkillCore(doc.access)
    core.buildContext()
    const h1 = doc.access.getMap().elements.find((e) => e.tag === 'h1')!.sid
    const r = await core.executeTool(
      call('apply_ops', {
        ops: [{ op: 'set_attr', sid: h1, name: 'data-logo', value: 'attachment://logo.png' }],
      }),
    )
    expect(r.isError).toBe(true)
    expect(r.output).toContain('The document changed')
    expect(doc.text).not.toContain('assets/logo.png')
  })
})

describe('expandAttachmentRefs', () => {
  it('never rewrites str_replace old (it must match the document)', async () => {
    const { ops, errors } = await expandAttachmentRefs(
      [
        {
          op: 'str_replace',
          old: 'src="attachment://logo.png"',
          new: 'a attachment://logo.png b',
        },
      ],
      async () => ({ ok: true, src: 'assets/logo.png' }),
    )
    expect(errors).toEqual([])
    expect(ops[0]).toMatchObject({
      old: 'src="attachment://logo.png"',
      new: 'a assets/logo.png b',
    })
  })

  it('URL-encoded names decode before resolution and resolve once per ref', async () => {
    const seen: string[] = []
    const { ops } = await expandAttachmentRefs(
      [
        { op: 'set_attr', sid: 1, name: 'src', value: 'attachment://my%20logo.png' },
        {
          op: 'insert_html',
          sid: 2,
          position: 'after',
          html: '<img src="attachment://my%20logo.png">',
        },
      ],
      async (name) => {
        seen.push(name)
        return { ok: true, src: 'assets/my logo.png' }
      },
    )
    expect(seen).toEqual(['my logo.png'])
    expect(ops[0]).toMatchObject({ value: 'assets/my logo.png' })
    expect(ops[1]).toMatchObject({ html: '<img src="assets/my logo.png">' })
  })

  it('known names with spaces and parentheses match whole, in any casing or encoding', async () => {
    const seen: string[] = []
    const names = ['logo (1).png', 'Screenshot 2024.png']
    const { ops, errors } = await expandAttachmentRefs(
      [
        {
          op: 'insert_html',
          sid: 1,
          position: 'after',
          html: '<img src="attachment://logo (1).png"> <img src="attachment://screenshot 2024.png">',
        },
        {
          op: 'set_style',
          sid: 2,
          styles: { 'background-image': 'url(attachment://logo (1).png) no-repeat' },
        },
        { op: 'set_attr', sid: 3, name: 'src', value: 'attachment://logo%20(1).png' },
      ],
      async (name) => {
        seen.push(name)
        return { ok: true, src: `assets/${name}` }
      },
      names,
    )
    expect(errors).toEqual([])
    expect(seen.sort()).toEqual(['Screenshot 2024.png', 'logo (1).png'])
    expect(ops[0]).toMatchObject({
      html: '<img src="assets/logo (1).png"> <img src="assets/Screenshot 2024.png">',
    })
    expect(ops[1]).toMatchObject({
      styles: { 'background-image': 'url(assets/logo (1).png) no-repeat' },
    })
    expect(ops[2]).toMatchObject({ value: 'assets/logo (1).png' })
  })

  it('a known name is not matched as a prefix of a longer token', async () => {
    const seen: string[] = []
    const { errors } = await expandAttachmentRefs(
      [
        { op: 'set_attr', sid: 1, name: 'src', value: 'attachment://logo.png.bak' },
        { op: 'set_attr', sid: 2, name: 'data-x', value: 'attachment://logo.png-dark.png' },
      ],
      async (name) => {
        seen.push(name)
        return { ok: false, error: 'no attachment' }
      },
      ['logo.png'],
    )
    expect(seen.sort()).toEqual(['logo.png-dark.png', 'logo.png.bak'])
    expect(errors).toHaveLength(2)
  })

  it('a percent sign in an unknown name is passed through instead of throwing', async () => {
    const seen: string[] = []
    const { errors } = await expandAttachmentRefs(
      [{ op: 'set_attr', sid: 1, name: 'src', value: 'attachment://100%.png' }],
      async (name) => {
        seen.push(name)
        return { ok: false, error: 'no attachment' }
      },
    )
    expect(seen).toEqual(['100%.png'])
    expect(errors).toEqual(['attachment://100%.png: no attachment'])
  })
})

describe('truncated data: URL guard in apply_ops', () => {
  it('rejects the batch when an op re-types an ellipsized data: src', () => {
    const doc = fakeAccess(DOC)
    const core = createHtmlSkillCore(doc.access)
    core.buildContext()
    const h1 = doc.access.getMap().elements.find((e) => e.tag === 'h1')!.sid
    const before = doc.text
    const r = core.executeTool(
      call('apply_ops', {
        ops: [
          { op: 'set_text', sid: h1, text: 'New title' },
          {
            op: 'insert_html',
            sid: h1,
            position: 'before',
            html: '<img src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUg\u2026" alt="logo">',
          },
        ],
      }),
    ) as { output: string; isError?: boolean }
    expect(r.isError).toBe(true)
    expect(r.output).toContain('op 2')
    expect(r.output).toContain('truncated data: URL')
    expect(doc.text).toBe(before)
  })

  it('catches three-dot truncation in str_replace old and in style urls', () => {
    const ellipsized: HtmlOp[] = [
      { op: 'str_replace', old: 'src="data:image/png;base64,iVBOR..."', new: 'src="x.png"' },
    ]
    expect(findTruncatedDataUrl(ellipsized)).toBe(0)
    const styled = [
      {
        op: 'set_style',
        sid: 1,
        styles: { 'background-image': 'url(data:image/png;base64,AAA\u2026)' },
      },
    ] as HtmlOp[]
    expect(findTruncatedDataUrl(styled)).toBe(0)
  })

  it('leaves complete data: URLs and prose ellipses alone', () => {
    const doc = fakeAccess(DOC)
    const core = createHtmlSkillCore(doc.access)
    core.buildContext()
    const h1 = doc.access.getMap().elements.find((e) => e.tag === 'h1')!.sid
    const r = core.executeTool(
      call('apply_ops', {
        ops: [
          {
            op: 'insert_html',
            sid: h1,
            position: 'before',
            html: '<img src="data:image/gif;base64,R0lGODlhAQABAAAAACw=" alt="dot">',
          },
          { op: 'set_text', sid: h1, text: 'To be continued\u2026 (data: not a url)' },
        ],
      }),
    ) as { output: string; isError?: boolean; mutated?: boolean }
    expect(r.isError).toBeFalsy()
    expect(r.mutated).toBe(true)
    expect(doc.text).toContain('R0lGODlhAQABAAAAACw=')
    expect(doc.text).toContain('To be continued\u2026')
  })
})
