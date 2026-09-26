import { copyFileSync, existsSync, readFileSync } from 'node:fs'
import { join, resolve } from 'node:path'
import { Client } from '@modelcontextprotocol/sdk/client/index.js'
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js'
import { afterAll, beforeAll, describe, expect, it } from 'vitest'
import { defaultRegistry } from '../src/cli'
import { DECK_TOOLS } from '../src/mcp/deck'
import { loadCatalog } from '../src/commands/guide'
import type { GuideDomain } from '../src/op-catalog'
import { forbiddenConstructs } from '../src/mcp/op-schemas'
import { createContext, disposeContext, runJson, type McpContext } from '../src/mcp/run'
import { createMcpServer } from '../src/mcp/server'
import {
  buildArgv,
  resolveTool,
  resolveTools,
  stripVerbPrefix,
  toolShape,
  TOOLS,
} from '../src/mcp/tools'
import { tempDir, writeMinimalPdf } from './helpers'

const REPO = resolve(__dirname, '../../..')
const DOCX = join(REPO, 'apps/docs/tests/pagination-corpus/docx/01-simple-english.docx')

const registry = defaultRegistry()

describe('mcp tool table', () => {
  it('names every command option it exposes and builds a schema for each tool', () => {
    const tools = resolveTools(registry)
    expect(tools.length).toBe(TOOLS.length)
    const names = [...tools.map((t) => t.name), ...DECK_TOOLS.map((t) => t.name)]
    expect(new Set(names).size).toBe(names.length)
    for (const name of names) expect(name).toMatch(/^[a-z][a-z0-9_]*$/)
    for (const tool of tools) {
      const shape = toolShape(tool)
      for (const p of tool.params) expect(shape[p.key]).toBeDefined()
      for (const p of tool.positionals ?? []) expect(shape[p.key]).toBeDefined()
    }
  })

  it('rejects a tool that names an option the command does not have', () => {
    expect(() =>
      resolveTool({ name: 'x', command: 'info', description: '', options: ['nope'] }, registry),
    ).toThrow('has no --nope')
  })

  it('drops the verb prefix from option descriptions', () => {
    expect(stripVerbPrefix('read: block index range (default: all)')).toBe(
      'block index range (default: all)',
    )
    expect(stripVerbPrefix('pptx --spec <dir>: the deck outline')).toBe('the deck outline')
    expect(stripVerbPrefix('worksheet (default: the active one)')).toBe(
      'worksheet (default: the active one)',
    )
    expect(stripVerbPrefix('auto | 0.5k | 1k')).toBe('auto | 0.5k | 1k')
  })

  it('turns tool arguments into the command line, with inline JSON and text as files', () => {
    const tools = new Map(resolveTools(registry).map((t) => [t.name, t]))
    const apply = buildArgv(tools.get('docs_apply')!, {
      file: 'a.docx',
      ops: [{ op: 'findReplace', find: 'a', replace: 'b' }],
      dry_run: true,
      force: false,
      author: 'me',
    })
    expect(apply.argv).toEqual(['docs', 'apply', 'a.docx', '--author', 'me', '--dry-run'])
    expect(apply.inline).toEqual([
      { option: 'ops', ext: '.json', text: '[{"op":"findReplace","find":"a","replace":"b"}]' },
    ])

    const docx = buildArgv(tools.get('create_docx')!, { markdown: '# Hi', out: 'x.docx' })
    expect(docx.argv).toEqual(['create', '--type', 'docx', '--out', 'x.docx'])
    expect(docx.inline).toEqual([{ option: 'from', ext: '.md', text: '# Hi' }])

    const render = buildArgv(tools.get('slides_render')!, {
      file: 'd.pptx',
      out: 'shots',
      slide: 2,
      scale: 1.5,
    })
    expect(render.argv).toEqual([
      'slides',
      'render',
      'd.pptx',
      '--out',
      'shots',
      '--slide',
      '2',
      '--scale',
      '1.5',
    ])

    const guide = buildArgv(tools.get('guide')!, { domain: 'slides' })
    expect(guide.argv).toEqual(['guide', 'slides'])
    expect(() => buildArgv(tools.get('info')!, {})).toThrow('missing file')
  })

  it('exposes the batch modes on slides_apply as on docs_apply and sheet_apply', () => {
    const tools = new Map(resolveTools(registry).map((t) => [t.name, t]))
    for (const name of ['docs_apply', 'sheet_apply', 'slides_apply']) {
      const keys = tools.get(name)!.params.map((p) => p.key)
      expect(keys).toEqual(expect.arrayContaining(['best_effort', 'stop_on_error']))
    }
    const best = buildArgv(tools.get('slides_apply')!, {
      file: 'd.pptx',
      ops: [],
      best_effort: true,
    })
    expect(best.argv).toEqual(['slides', 'apply', 'd.pptx', '--best-effort'])
    const stop = buildArgv(tools.get('slides_apply')!, {
      file: 'd.pptx',
      ops: [],
      stop_on_error: true,
    })
    expect(stop.argv).toEqual(['slides', 'apply', 'd.pptx', '--stop-on-error'])
  })

  it('marks open and selection as stdio-only', () => {
    const local = TOOLS.filter((t) => t.localOnly).map((t) => t.name)
    expect(local.sort()).toEqual(['open', 'selection'])
  })
})

describe('mcp server', () => {
  let ctx: McpContext
  let client: Client
  let dir: string

  beforeAll(async () => {
    dir = tempDir()
    ctx = createContext({
      cwd: dir,
      env: { ...process.env, GENOFFICE_AUDIT_LOG: 'off' },
      log: () => {},
    })
    const server = await createMcpServer(ctx, { registry })
    const [a, b] = InMemoryTransport.createLinkedPair()
    await server.connect(a)
    client = new Client({ name: 'test', version: '0' })
    await client.connect(b)
  })

  afterAll(async () => {
    await client.close()
    disposeContext(ctx)
  })

  const call = async (name: string, args: Record<string, unknown>) => {
    const r = await client.callTool({ name, arguments: args })
    const content = r.content as { type: string; text?: string }[]
    const text = content.find((c) => c.type === 'text')?.text ?? ''
    return { isError: r.isError === true, content, text, json: () => JSON.parse(text) }
  }

  it('lists the command tools and the deck tools with schemas', async () => {
    const { tools } = await client.listTools()
    const names = tools.map((t) => t.name)
    expect(names).toEqual(
      expect.arrayContaining([
        'info',
        'docs_read',
        'docs_apply',
        'pdf_read',
        'sheet_apply',
        'slides_render',
        'guide',
        'deck_start',
        'deck_page',
        'deck_build',
        'deck_replace',
      ]),
    )
    expect(names).not.toContain('mcp')
    expect(names).not.toContain('install')
    expect(names).toEqual(expect.arrayContaining(['open', 'selection']))
    const selection = tools.find((t) => t.name === 'selection')!
    expect(selection.annotations?.readOnlyHint).toBe(true)
    expect(selection.inputSchema.required).toEqual(['file'])
    const slides = tools.find((t) => t.name === 'slides_apply')!
    const slidesProps = slides.inputSchema.properties as Record<string, { type?: unknown }>
    expect(slidesProps.best_effort?.type).toBe('boolean')
    expect(slidesProps.stop_on_error?.type).toBe('boolean')
    const apply = tools.find((t) => t.name === 'docs_apply')!
    const props = apply.inputSchema.properties as Record<
      string,
      { type?: unknown; anyOf?: unknown }
    >
    expect(apply.inputSchema.required).toEqual(expect.arrayContaining(['file', 'ops']))
    expect(props.dry_run?.type).toBe('boolean')
    expect(props.ops?.type).toBe('array')
    expect(apply.annotations?.readOnlyHint).toBe(false)
    expect(apply.annotations?.openWorldHint).toBe(false)
    expect(tools.find((t) => t.name === 'search')!.annotations?.openWorldHint).toBe(true)
    expect(tools.find((t) => t.name === 'docs_read')!.annotations?.readOnlyHint).toBe(true)
    expect(tools.find((t) => t.name === 'pdf_read')!.annotations?.readOnlyHint).toBe(true)
  })

  type Schema = {
    type?: unknown
    enum?: unknown[]
    items?: Schema
    anyOf?: Schema[]
    properties?: Record<string, Schema>
    required?: string[]
    description?: string
  }
  const OPS_TOOLS: Record<string, GuideDomain> = {
    docs_apply: 'docs',
    sheet_apply: 'sheets',
    slides_apply: 'slides',
  }
  const opsParam = async (tool: string): Promise<Schema> => {
    const { tools } = await client.listTools()
    const props = tools.find((t) => t.name === tool)!.inputSchema.properties as Record<
      string,
      Schema
    >
    return props.ops!
  }

  it('advertises ops as one object variant per callable op of guide <domain> --json', async () => {
    for (const [tool, domain] of Object.entries(OPS_TOOLS)) {
      const ops = await opsParam(tool)
      expect(ops.type).toBe('array')
      const variants = ops.items!.anyOf!
      const guide = await runJson(['guide', domain, '--json'], ctx)
      const callable = (guide.ok!.detail!.ops as { op: string; available?: false }[])
        .filter((o) => o.available !== false)
        .map((o) => o.op)
      expect(variants.map((v) => v.properties!.op!.enum![0])).toEqual(callable)
      expect(callable).toEqual(
        (await loadCatalog(domain)).ops.filter((o) => o.available !== false).map((o) => o.op),
      )
      for (const v of variants) {
        expect(v.type).toBe('object')
        expect(v.required).toContain('op')
        expect(v.properties!.op).toEqual({ enum: [v.properties!.op!.enum![0]] })
        if (v.description) expect(v.description.length).toBeLessThanOrEqual(60)
      }
    }
    const docs = (await opsParam('docs_apply')).items!.anyOf!
    const setHeading = docs.find((v) => v.properties!.op!.enum![0] === 'setHeadingLevel')!
    expect(setHeading.required).toEqual(['op', 'target', 'level'])
    expect(setHeading.properties!.level).toEqual({ type: 'number' })
    const sheets = (await opsParam('sheet_apply')).items!.anyOf!
    const setCell = sheets.find((v) => v.properties!.op!.enum![0] === 'set_cell')!
    expect(setCell.required).toEqual(['op', 'address', 'value'])
    expect(setCell.properties!.sheet).toEqual({ type: 'string' })
    const slides = (await opsParam('slides_apply')).items!.anyOf!
    const transform = slides.find((v) => v.properties!.op!.enum![0] === 'setTransform')!
    expect(transform.required).toEqual(['op', 'target', 'box'])
    expect(transform.properties!.box!.required).toEqual(['x', 'y', 'cx', 'cy'])
  })

  it('types cells, data and create_pptx ops from the same catalogs', async () => {
    const { tools } = await client.listTools()
    const props = (name: string) =>
      tools.find((t) => t.name === name)!.inputSchema.properties as Record<string, Schema>
    expect(props('sheet_apply').cells!.items!.required).toEqual(['cell'])
    expect(props('create_xlsx').data!.anyOf).toHaveLength(2)
    expect(props('merge').data!.type).toBe('object')
    const names = props('create_pptx').ops!.items!.properties!.op!.enum!
    const slides = (await opsParam('slides_apply')).items!.anyOf!
    expect(names).toEqual(slides.map((v) => v.properties!.op!.enum![0]))
  })

  it('emits no schema construct that Gemini or strict clients reject, within the size budget', async () => {
    const list = await client.listTools()
    for (const tool of list.tools) {
      const { $schema: _draft, ...schema } = tool.inputSchema as Record<string, unknown>
      expect(forbiddenConstructs(schema), tool.name).toEqual([])
    }
    const bytes = JSON.stringify(list).length
    expect(bytes).toBeLessThanOrEqual(90 * 1024)
    expect(bytes).toBeGreaterThan(60 * 1024)
  })

  it('leaves op validation to the CLI: a typo or a wrong field type gets the structured op error', async () => {
    const pptx = join(dir, 'typed.pptx')
    const created = await call('create_pptx', {
      ops: [
        {
          op: 'addElement',
          target: { slide: 0 },
          kind: 'textbox',
          offset: { x: 914400, y: 914400, cx: 3657600, cy: 914400 },
          paragraphs: [{ runs: [{ text: 'Hello' }] }],
        },
      ],
      out: pptx,
    })
    expect(created.isError).toBe(false)
    const read = await call('slides_read', { file: pptx })
    const el = read.json().detail.pages[0].elements[0].id as string

    const typo = await call('slides_apply', {
      file: pptx,
      ops: [{ op: 'setTxt', target: { slide: 0, el }, paragraphs: [] }],
    })
    expect(typo.isError).toBe(true)
    expect(typo.text).not.toContain('Input validation error')
    expect(typo.json().detail.failures[0]).toMatchObject({
      index: 0,
      op: 'setTxt',
      reason: 'unknown_op',
      did_you_mean: 'setText',
    })

    const wrongType = await call('slides_apply', {
      file: pptx,
      ops: [{ op: 'setText', target: { slide: 0, el }, paragraphs: 'nope' }],
    })
    expect(wrongType.isError).toBe(true)
    expect(wrongType.json().detail.failures[0]).toMatchObject({
      index: 0,
      op: 'setText',
      reason: 'op_rejected',
      usage: expect.stringContaining('setText {paragraphs:'),
    })

    const docx = join(dir, 'typed.docx')
    copyFileSync(DOCX, docx)
    const docsTypo = await call('docs_apply', {
      file: docx,
      ops: [{ op: 'findReplac', find: 'a', replace: 'b' }],
    })
    expect(docsTypo.isError).toBe(true)
    expect(docsTypo.json().detail.failures[0]).toMatchObject({
      reason: 'unknown_op',
      did_you_mean: 'findReplace',
    })
  })

  it('falls back to untyped arrays with compactSchemas', async () => {
    const ctx2 = createContext({
      cwd: dir,
      env: { ...process.env, GENOFFICE_AUDIT_LOG: 'off' },
      log: () => {},
    })
    const server = await createMcpServer(ctx2, { registry, compactSchemas: true })
    const [a, b] = InMemoryTransport.createLinkedPair()
    await server.connect(a)
    const client2 = new Client({ name: 'test3', version: '0' })
    await client2.connect(b)
    try {
      const { tools } = await client2.listTools()
      const props = tools.find((t) => t.name === 'slides_apply')!.inputSchema.properties as Record<
        string,
        Schema
      >
      expect(props.ops!.anyOf!.map((v) => v.type)).toEqual(['array', 'object'])
      expect(props.ops!.items).toBeUndefined()
      expect(JSON.stringify(tools).length).toBeLessThan(40 * 1024)
    } finally {
      await client2.close()
      disposeContext(ctx2)
    }
  })

  it('reads one PDF page headless through pdf_read', async () => {
    const pdf = writeMinimalPdf(join(dir, 'two.pdf'), ['First page', 'Second page'])
    const r = await call('pdf_read', { file: pdf, page: 2, max_chars: 6 })
    expect(r.isError).toBe(false)
    expect(r.json().detail).toMatchObject({
      pages: 2,
      range: '2-2',
      pages_read: [{ page: 2, text: 'Second…(+5 chars)', truncated: true }],
    })
  })

  it('runs a read-only command and returns the JSON envelope', async () => {
    const r = await call('info', { file: DOCX })
    expect(r.isError).toBe(false)
    expect(r.json()).toMatchObject({ status: 'ok', command: 'info', detail: { format: 'docx' } })
  })

  it('returns command errors as isError with the machine-readable reason', async () => {
    const r = await call('info', { file: join(dir, 'missing.docx') })
    expect(r.isError).toBe(true)
    expect(r.json()).toMatchObject({ status: 'error', code: 2, error: 'file_not_found' })
  })

  it('applies inline ops to a document without any file from the client', async () => {
    const copy = join(dir, 'edit.docx')
    copyFileSync(DOCX, copy)
    const before = await call('docs_read', { file: copy, range: '0', full: true })
    const first = (before.json().detail.items[0].text as string).split(' ')[0]!
    const r = await call('docs_apply', {
      file: copy,
      ops: [{ op: 'findReplace', find: first, replace: 'GENOFFICE' }],
    })
    expect(r.isError).toBe(false)
    expect(r.json()).toMatchObject({ status: 'ok', command: 'docs', output_path: copy })
    const after = await call('docs_read', { file: copy, range: '0', full: true })
    expect(after.json().detail.items[0].text).toContain('GENOFFICE')
  })

  it('spells out the lost ops ahead of a partial batch envelope, without isError', async () => {
    const textbox = (slide: number, text: string) => ({
      op: 'addElement',
      target: { slide },
      kind: 'textbox',
      offset: { x: 914400, y: 914400, cx: 3657600, cy: 914400 },
      paragraphs: [{ runs: [{ text }] }],
    })
    const pptx = join(dir, 'partial.pptx')
    const created = await call('create_pptx', { ops: [textbox(0, 'Hello')], out: pptx })
    expect(created.isError).toBe(false)
    const r = await call('slides_apply', {
      file: pptx,
      ops: [textbox(0, 'one'), textbox(7, 'nope'), textbox(0, 'three')],
      best_effort: true,
    })
    expect(r.isError).toBe(false)
    expect(r.content).toHaveLength(2)
    const [first, second] = r.content as { text: string }[]
    expect(first!.text.split('\n')[0]).toBe('partial: 2 of 3 ops applied, 1 failed')
    expect(first!.text).toMatch(/\n {2}op 1 \(addElement\): .+/)
    const envelope = JSON.parse(second!.text)
    expect(envelope).toMatchObject({ status: 'partial', command: 'slides' })
    expect(envelope.detail.batch).toEqual({ total: 3, applied: 2, failed: 1, skipped: 0 })
    expect(envelope.detail.failures[0].index).toBe(1)

    const stopped = await call('slides_apply', {
      file: pptx,
      ops: [textbox(0, 'four'), textbox(7, 'nope'), textbox(0, 'six')],
      stop_on_error: true,
    })
    expect(stopped.isError).toBe(false)
    expect((stopped.content[0] as { text: string }).text.split('\n')[0]).toBe(
      'partial: 1 of 3 ops applied, 1 failed, 1 skipped',
    )

    const clean = await call('slides_apply', { file: pptx, ops: [textbox(0, 'seven')] })
    expect(clean.isError).toBe(false)
    expect(clean.content).toHaveLength(1)
  })

  it('serves the guides as plain text tools and as resources', async () => {
    const r = await call('guide', { domain: 'docs' })
    expect(r.isError).toBe(false)
    expect(r.text).toContain('Word ops')
    expect(() => r.json()).toThrow()
    const { resources } = await client.listResources()
    expect(resources.map((x) => x.uri)).toContain('genoffice://guide/slides/spec')
    const spec = await client.readResource({ uri: 'genoffice://guide/slides/spec' })
    expect((spec.contents[0] as { text: string }).text.length).toBeGreaterThan(200)
  })

  it('walks the deck stages in order and refuses to skip one', async () => {
    const deck = join(dir, 'deck')
    const text = (t: string) => ({
      type: 'text',
      x: 80,
      y: 80,
      w: 800,
      h: 80,
      paragraphs: [{ runs: [{ text: t, sizePt: 28, bold: true, color: '#112233' }] }],
    })
    const early = await call('deck_page', { dir: deck, index: 0, page: { elements: [text('x')] } })
    expect(early.isError).toBe(true)
    expect(early.json().suggestion).toContain('deck_start')

    const outlinePage = (title: string, type: string, layout: string) => ({
      title,
      type,
      layout,
      brief: `${title}: three cards with the real 2024 figures from the brief, one source line under each card.`,
      image_queries: [],
    })
    const start = await call('deck_start', {
      dir: deck,
      style: '# Style\n\nPalette: #112233 on white.',
      outline: {
        core_hook: 'Two pages, both built',
        pages: [
          outlinePage('Cover', 'cover', 'cover_typography_hero'),
          outlinePage('Close', 'closing', 'closing_cta'),
        ],
      },
    })
    expect(start.isError).toBe(false)
    const started = start.json()
    expect(started.detail.pages.map((p: { title: string }) => p.title)).toEqual(['Cover', 'Close'])
    expect(started.detail.guides.design).toContain('style')
    expect(existsSync(join(deck, 'style.md'))).toBe(true)

    const blocked = await call('deck_build', { dir: deck })
    expect(blocked.isError).toBe(true)
    expect(blocked.json().detail.missing_pages).toEqual([0, 1])

    const skipped = await call('deck_page', {
      dir: deck,
      index: 1,
      page: { title: 'Close', type: 'closing', layout: 'closing_cta', elements: [text('Close')] },
    })
    expect(skipped.isError).toBe(true)
    expect(skipped.json().suggestion).toContain('page 0')

    const wrong = await call('deck_page', {
      dir: deck,
      index: 0,
      page: { title: 'Cover', type: 'closing', layout: 'closing_cta', elements: [text('Cover')] },
    })
    expect(wrong.isError).toBe(true)
    expect(wrong.json().detail.kept).toBe(false)
    expect(existsSync(join(deck, 'pages', '01.json'))).toBe(false)
    expect(wrong.json().detail.missing_pages).toEqual([0, 1])

    const p0 = await call('deck_page', {
      dir: deck,
      index: 0,
      page: {
        title: 'Cover',
        type: 'cover',
        layout: 'cover_typography_hero',
        elements: [text('Cover')],
      },
    })
    expect(p0.isError).toBe(false)
    expect(p0.json().detail.missing_pages).toEqual([1])
    expect(p0.json().detail.next).toContain('page 1')

    const bad = await call('deck_page', { dir: deck, index: 5, page: { elements: [text('x')] } })
    expect(bad.isError).toBe(true)
    expect(bad.json().error).toBe('out_of_range')

    const p1 = await call('deck_page', {
      dir: deck,
      index: 1,
      page: { title: 'Close', type: 'closing', layout: 'closing_cta', elements: [text('Close')] },
    })
    expect(p1.isError).toBe(false)
    expect(p1.json().detail.next).toContain('deck_build')

    const built = await call('deck_build', { dir: deck })
    expect(built.isError).toBe(false)
    expect(built.json().output_path).toBe(join(deck, 'deck.pptx'))
    expect(readFileSync(join(deck, 'deck.pptx')).byteLength).toBeGreaterThan(1000)

    const replaced = await call('deck_replace', {
      dir: deck,
      index: 1,
      page: { title: 'Close', type: 'closing', layout: 'closing_cta', elements: [text('Thanks')] },
    })
    expect(replaced.isError).toBe(false)
    expect(replaced.json().detail.slide).toBe(1)
    const read = await call('slides_read', { file: join(deck, 'deck.pptx'), slide: 1, full: true })
    expect(read.text).toContain('Thanks')

    const broken = await call('deck_replace', {
      dir: deck,
      index: 1,
      page: {
        title: 'Close',
        type: 'cover',
        layout: 'cover_typography_hero',
        elements: [text('Nope')],
      },
    })
    expect(broken.isError).toBe(true)
    expect(readFileSync(join(deck, 'pages', '02.json'), 'utf-8')).toContain('Thanks')
  })

  it('keeps the deck tools inside GENOFFICE_ALLOWED_ROOTS', async () => {
    const inside = join(dir, 'roots')
    const ctx2 = createContext({
      cwd: dir,
      env: { ...process.env, GENOFFICE_AUDIT_LOG: 'off', GENOFFICE_ALLOWED_ROOTS: inside },
      log: () => {},
    })
    const server = await createMcpServer(ctx2, { registry })
    const [a, b] = InMemoryTransport.createLinkedPair()
    await server.connect(a)
    const client2 = new Client({ name: 'test2', version: '0' })
    await client2.connect(b)
    try {
      const r = await client2.callTool({
        name: 'deck_start',
        arguments: {
          dir: join(dir, 'outside'),
          style: '# s',
          outline: { core_hook: 'x', pages: [] },
        },
      })
      expect(r.isError).toBe(true)
      const text = (r.content as { text: string }[])[0]!.text
      expect(JSON.parse(text).error).toBe('outside_allowed_roots')
      expect(existsSync(join(dir, 'outside'))).toBe(false)
    } finally {
      await client2.close()
      disposeContext(ctx2)
    }
  })
})
