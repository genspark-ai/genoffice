import { describe, expect, it } from 'vitest'
import { buildParseMap, type ParseMap } from '../src/renderer/document/parse-map'
import { compileOps, type HtmlOp } from '../src/renderer/document/ops'
import { applyPatches } from '../src/renderer/document/patch'
import { createHtmlSkillCore, buildOutline, type HtmlDocAccess } from '../src/renderer/ai/tools'
import type { Brief } from '../src/renderer/document/brief'
import type { PageWriteSpec } from '../src/renderer/ai/page-writer'

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

  it('plan_page validates the brief and relays confirm / redo / dismiss decisions', async () => {
    const doc = fakeAccess('')
    const core = createHtmlSkillCore(doc.access)
    core.buildContext()
    const bad = await core.executeTool(
      call('plan_page', { core_hook: '', style: { tone: 'x' }, sections: [] }),
    )
    expect(bad.isError).toBe(true)
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
      mode: 'extract',
    }
    let shown: Brief | undefined
    doc.access.confirmBrief = async (brief) => {
      shown = brief
      return {
        kind: 'confirmed',
        brief: { ...brief, core_hook: 'Edited hook', user_edited: ['core_hook'] },
      }
    }
    const ok = await core.executeTool(call('plan_page', proposal))
    expect(ok.isError).toBeFalsy()
    // invalid directions are dropped and at most two alternatives survive; docx_friendly defaults off
    expect(shown?.alternatives?.map((s) => s.name)).toEqual(['Candy grid', 'Night'])
    expect(shown?.alternatives?.[0]?.tokens?.radius).toBe('16px')
    expect(shown?.style.docx_friendly).toBe(false)
    expect(ok.output).toContain('Fields the user edited (keep them exactly): core_hook')
    expect(ok.output).toContain('Edited hook')
    expect(ok.output).toContain('pinned')
    doc.access.confirmBrief = async () => ({ kind: 'redo', note: 'warmer colors' })
    const redo = await core.executeTool(call('plan_page', proposal))
    expect(redo.output).toContain('different brief: warmer colors')
    doc.access.confirmBrief = async () => ({ kind: 'cancelled' })
    const cancelled = await core.executeTool(call('plan_page', proposal))
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
    const core = createHtmlSkillCore(doc.access)
    const r = await core.executeTool(
      call('plan_page', {
        mode: 'new',
        core_hook: 'Ship in a day',
        style: { tone: 'technical' },
        sections: [{ title: 'Hero', brief: 'x' }],
        context: 'Acme sells widgets since 1999',
      }),
    )
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

  it('plan_page reports a failed or partial page write', async () => {
    const doc = fakeAccess('')
    doc.access.confirmBrief = async (brief) => ({ kind: 'confirmed', brief })
    const proposal = {
      mode: 'new',
      core_hook: 'h',
      style: { tone: 't' },
      sections: [{ title: 'A', brief: 'x' }],
    }
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
    empty.access.confirmBrief = async (brief) => ({ kind: 'confirmed', brief })
    empty.access.writePage = async () => {
      writes++
      return { ok: true, html: '<html></html>' }
    }
    await createHtmlSkillCore(empty.access).executeTool(call('plan_page', proposal))
    expect(writes).toBe(1)
    const existing = fakeAccess(DOC)
    existing.access.confirmBrief = async (brief) => ({ kind: 'confirmed', brief })
    existing.access.writePage = async () => {
      writes++
      return { ok: true, html: '<html></html>' }
    }
    const existingCore = createHtmlSkillCore(existing.access)
    existingCore.buildContext()
    const r = await existingCore.executeTool(call('plan_page', proposal))
    expect(writes).toBe(1)
    expect(r.output).toContain('extract')
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
    const proposal = {
      mode: 'extract',
      core_hook: 'Quarterly numbers',
      style: { tone: 'plain' },
      sections: [{ title: 'Body', brief: 'x' }],
    }
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
    const core = createHtmlSkillCore(doc.access)
    core.buildContext()
    const r = await core.executeTool(
      call('plan_page', {
        mode: 'extract',
        core_hook: 'h',
        style: { tone: 't' },
        sections: [{ title: 'A', brief: 'x' }],
      }),
    )
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
    const r = await createHtmlSkillCore(doc.access).executeTool(
      call('plan_page', {
        mode: 'new',
        core_hook: 'h',
        style: { tone: 't' },
        sections: [{ title: 'A', brief: 'x' }],
      }),
    )
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
