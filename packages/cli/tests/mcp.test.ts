import { copyFileSync, existsSync, readFileSync } from 'node:fs'
import { join, resolve } from 'node:path'
import { Client } from '@modelcontextprotocol/sdk/client/index.js'
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js'
import { afterAll, beforeAll, describe, expect, it } from 'vitest'
import { defaultRegistry } from '../src/cli'
import { DECK_TOOLS } from '../src/mcp/deck'
import { createContext, disposeContext, type McpContext } from '../src/mcp/run'
import { createMcpServer } from '../src/mcp/server'
import {
  buildArgv,
  resolveTool,
  resolveTools,
  stripVerbPrefix,
  toolShape,
  TOOLS,
} from '../src/mcp/tools'
import { tempDir } from './helpers'

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
    const server = createMcpServer(ctx, { registry })
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
    const apply = tools.find((t) => t.name === 'docs_apply')!
    const props = apply.inputSchema.properties as Record<
      string,
      { type?: unknown; anyOf?: unknown }
    >
    expect(apply.inputSchema.required).toEqual(expect.arrayContaining(['file', 'ops']))
    expect(props.dry_run?.type).toBe('boolean')
    expect(props.ops?.anyOf).toBeDefined()
    expect(apply.annotations?.readOnlyHint).toBe(false)
    expect(apply.annotations?.openWorldHint).toBe(false)
    expect(tools.find((t) => t.name === 'search')!.annotations?.openWorldHint).toBe(true)
    expect(tools.find((t) => t.name === 'docs_read')!.annotations?.readOnlyHint).toBe(true)
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
    const server = createMcpServer(ctx2, { registry })
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
