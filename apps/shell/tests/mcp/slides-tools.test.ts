import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import { mkdtemp, rm } from 'node:fs/promises'
import { existsSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { Client } from '@modelcontextprotocol/sdk/client/index.js'
import { StreamableHTTPClientTransport } from '@modelcontextprotocol/sdk/client/streamableHttp.js'
import { McpServerService } from '../../src/main/mcp/mcp-server'
import { outlineToOps, parsePptxOutline } from '../../src/main/mcp/pptx-outline'
import {
  createSlidesTools,
  slidesDriver,
  type SlidesControl,
} from '../../src/main/mcp/tools/slides-tools'
import { createSessionHost, createSessionTools } from '../../src/main/mcp/tools/session-tools'
import { fakeCli } from './fake-cli'

/**
 * Slides tool surface over a real MCP session: headless create_pptx (markdown
 * and JSON outlines, path policy, background gating, CLI delegation) and the
 * visible deck session tools driven through a fake SlidesControl.
 */

let service: McpServerService | undefined
let dir: string
let client: Client | undefined

async function freePort(): Promise<number> {
  const { createServer } = await import('node:http')
  return new Promise((resolve, reject) => {
    const server = createServer()
    server.on('error', reject)
    server.listen(0, '127.0.0.1', () => {
      const address = server.address()
      const port = typeof address === 'object' && address ? address.port : 0
      server.close(() => resolve(port))
    })
  })
}

beforeEach(async () => {
  dir = await mkdtemp(join(tmpdir(), 'genoffice-mcp-pptx-'))
})

afterEach(async () => {
  await client?.close()
  client = undefined
  await service?.stop()
  service = undefined
  await rm(dir, { recursive: true, force: true })
})

async function startService(tools: ReturnType<typeof createSlidesTools>): Promise<void> {
  const port = await freePort()
  service = new McpServerService({ port, tools })
  await service.start()
  client = new Client({ name: 'slides-test', version: '0.0.0' })
  await client.connect(new StreamableHTTPClientTransport(new URL(`http://127.0.0.1:${port}/mcp`)))
}

/** content tools + the shared session tools wired against the same host, as the shell does */
function sessionSurface(control: SlidesControl): ReturnType<typeof createSlidesTools> {
  const host = createSessionHost()
  return [
    ...createSessionTools([slidesDriver(control)], host),
    ...createSlidesTools({ defaultSaveDir: () => dir, slides: control }, host),
  ]
}

function baseDeps(background?: boolean): {
  defaultSaveDir: () => string
  background?: boolean
} {
  return { defaultSaveDir: () => dir, ...(background === undefined ? {} : { background }) }
}

function text(content: unknown): string {
  const arr = content as Array<{ type: string; text?: string }>
  return arr.map((c) => c.text ?? '').join('')
}

describe('pptx outline mapping', () => {
  it('maps markdown headings, bullets, nesting and numbered lists', () => {
    const slides = parsePptxOutline(
      'markdown',
      '# First\n\n- one\n  - nested\n2. wait this is markdown\n\n# Second\n\n## Bold line\nplain',
    )
    expect(slides).toHaveLength(2)
    expect(slides[0]!.title).toBe('First')
    expect(slides[0]!.paragraphs).toEqual([
      { text: 'one', bullet: 'char' },
      { text: 'nested', bullet: 'char', level: 1 },
      { text: 'wait this is markdown', bullet: 'number' },
    ])
    expect(slides[1]!.title).toBe('Second')
    expect(slides[1]!.paragraphs).toEqual([{ text: 'Bold line', bold: true }, { text: 'plain' }])
  })

  it('maps the JSON shape including string slide shorthand', () => {
    const slides = parsePptxOutline(
      'json',
      JSON.stringify({
        slides: [
          'Just a body',
          { title: 'T', bullets: ['a', { text: 'b', level: 2, bullet: 'number' }] },
        ],
      }),
    )
    expect(slides[0]).toEqual({ paragraphs: [{ text: 'Just a body' }] })
    expect(slides[1]!.title).toBe('T')
    // entries under `bullets` are bullets: a bare string defaults to a character
    // bullet, and an explicit `bullet` is honored
    expect(slides[1]!.paragraphs).toEqual([
      { text: 'a', bullet: 'char' },
      { text: 'b', bullet: 'number', level: 2 },
    ])
  })

  it('keeps `paragraphs` as plain text while `bullets` renders bullets', () => {
    const slides = parsePptxOutline(
      'json',
      JSON.stringify({
        slides: [
          { title: 'A', bullets: ['b1', { text: 'b2', bullet: 'char' }] },
          { title: 'B', paragraphs: ['p1', { text: 'p2' }] },
        ],
      }),
    )
    expect(slides[0]!.paragraphs).toEqual([
      { text: 'b1', bullet: 'char' },
      { text: 'b2', bullet: 'char' },
    ])
    expect(slides[1]!.paragraphs).toEqual([{ text: 'p1' }, { text: 'p2' }])
  })

  it('rejects malformed JSON and empty outlines', () => {
    expect(() => parsePptxOutline('json', '{nope')).toThrow(/valid JSON/)
    expect(() => parsePptxOutline('json', JSON.stringify({ slides: 3 }))).toThrow(/slides/)
    expect(() => outlineToOps(parsePptxOutline('markdown', '   \n  \n'))).toThrow(/no slides/)
    expect(() => outlineToOps([])).toThrow(/no slides/)
    // a heading-less outline still yields one slide from its body lines
    const slides = parsePptxOutline('markdown', 'just some words')
    expect(slides).toEqual([{ paragraphs: [{ text: 'just some words' }] }])
  })

  it('maps an outline to one flat op array: create every page, then fill them', () => {
    const ops = outlineToOps(parsePptxOutline('markdown', '# A\n- x\n\n# B\n- y\n\n# C'))
    const slidesOf = (op: { op: string; target?: { slide: number } }): string =>
      `${op.op}@${op.target?.slide}`
    // pages are created before any fill op, so the CLI's per-op application can
    // target a page a previous op added
    expect(ops.slice(0, 2).map(slidesOf)).toEqual(['addBlankSlide@0', 'addBlankSlide@1'])
    const adds = ops.filter((o) => o.op === 'addElement')
    // A: title+body, B: title+body, C: title only
    expect(adds).toHaveLength(5)
    for (const op of adds) expect(op.target?.slide ?? 0).toBeLessThan(3)
  })
})

describe('headless create_pptx', () => {
  it('delegates to the CLI with the mapped ops and reports its output', async () => {
    const cli = fakeCli()
    await startService(createSlidesTools({ ...baseDeps(), cli: cli.runner }, createSessionHost()))
    const result = await client!.callTool({
      name: 'create_pptx',
      arguments: { title: 'Deck Plan', outline: '# Intro\n- point one\n- point two\n\n# Next' },
    })
    expect(result.isError).toBeFalsy()
    const payload = JSON.parse(text(result.content)) as { path: string }
    expect(existsSync(payload.path)).toBe(true)
    expect(payload.path.startsWith(dir)).toBe(true)
    expect(payload.path.endsWith('.pptx')).toBe(true)

    // `create --type pptx --ops <staged file> --out <path>`
    const args = cli.last()
    expect(args.slice(0, 4)).toEqual(['create', '--type', 'pptx', '--ops'])
    expect(args).toContain('--out')
    const ops = JSON.parse(cli.lastOps() ?? '[]') as Array<{
      op: string
      paragraphs?: Array<{ runs: Array<{ text: string }> }>
    }>
    const texts = ops.flatMap((o) => (o.paragraphs ?? []).flatMap((p) => p.runs.map((r) => r.text)))
    expect(texts).toContain('Intro')
    expect(texts).toContain('point one')
  })

  it('accepts a JSON outline and appends the extension to extensionless paths', async () => {
    const cli = fakeCli()
    await startService(createSlidesTools({ ...baseDeps(), cli: cli.runner }, createSessionHost()))
    const result = await client!.callTool({
      name: 'create_pptx',
      arguments: {
        title: 'Json Deck',
        format: 'json',
        outline: JSON.stringify({
          slides: [{ title: 'Hello', bullets: ['a', { text: 'b', level: 1 }] }],
        }),
        path: join(dir, 'json-deck'),
      },
    })
    expect(result.isError).toBeFalsy()
    const payload = JSON.parse(text(result.content)) as { path: string }
    expect(payload.path).toBe(join(dir, 'json-deck.pptx'))
    expect(cli.last()).toContain(join(dir, 'json-deck.pptx'))
  })

  it('enforces the shared path policy: absolute paths, clobber guard', async () => {
    const cli = fakeCli()
    await startService(createSlidesTools({ ...baseDeps(), cli: cli.runner }, createSessionHost()))
    const absolute = await client!.callTool({
      name: 'create_pptx',
      arguments: { title: 'X', outline: '# A', path: 'relative.pptx' },
    })
    expect(absolute.isError).toBe(true)
    expect(text(absolute.content)).toMatch(/path must be absolute/)

    // Both calls name the same explicit path: the clobber guard is about the
    // file on disk, so the second run must not depend on how the first run
    // derived its name (a title-derived path differs in case, and a
    // case-insensitive filesystem would hide a missing guard).
    const target = join(dir, 'taken.pptx')
    const first = await client!.callTool({
      name: 'create_pptx',
      arguments: { title: 'Taken', outline: '# A', path: target },
    })
    expect(first.isError).toBeFalsy()
    const second = await client!.callTool({
      name: 'create_pptx',
      arguments: { title: 'Taken', outline: '# A', path: target },
    })
    expect(second.isError).toBe(true)
    expect(text(second.content)).toMatch(/file already exists/)
    const third = await client!.callTool({
      name: 'create_pptx',
      arguments: { title: 'Taken', outline: '# A', path: target, overwrite: true },
    })
    expect(third.isError).toBeFalsy()
    // overwrite maps to the CLI's --force
    expect(cli.last()).toContain('--force')
  })

  it('is hidden when background is off, present when on (session tools unaffected)', async () => {
    const fake = fakeSlidesControl()
    const cli = fakeCli()
    await startService(
      createSlidesTools(
        { ...baseDeps(false), slides: fake.control, cli: cli.runner },
        createSessionHost(),
      ),
    )
    let names = (await client!.listTools()).tools.map((t) => t.name)
    expect(names).not.toContain('create_pptx')
    expect(names).toContain('read_deck')

    await client!.close()
    await service!.stop()
    service = undefined
    await startService(
      createSlidesTools(
        { ...baseDeps(true), slides: fake.control, cli: cli.runner },
        createSessionHost(),
      ),
    )
    names = (await client!.listTools()).tools.map((t) => t.name)
    expect(names).toContain('create_pptx')
    expect(names).toContain('read_deck')
  })
})

/** In-memory SlidesControl standing in for the shell's slides-bridge */
function fakeSlidesControl(): { control: SlidesControl; saved: Array<{ path: string }> } {
  let nextId = 1
  const saved: Array<{ path: string }> = []
  return {
    saved,
    control: {
      openBlankTab: async () => nextId++,
      runTxn: async (wcId, req) => {
        if (wcId <= 0) throw new Error('no deck')
        const ops = req.ops as Array<{ op: string }>
        if (ops.some((o) => o.op === 'explode')) {
          return { applied: false, failures: [{ index: 0, error: 'boom' }] }
        }
        return { applied: true, records: ops.map((o) => ({ op: o.op })) }
      },
      readDeck: async (wcId) => ({ slides: [{ index: 0, wc: wcId, elements: [] }] }),
      saveDeck: async (wcId, path) => {
        if (wcId <= 0) throw new Error('no deck')
        saved.push({ path })
        return { path }
      },
    },
  }
}

describe('visible deck session tools', () => {
  it('expose session lifecycle: create, read, apply, save-closes', async () => {
    const fake = fakeSlidesControl()
    await startService(sessionSurface(fake.control))

    const before = await client!.callTool({ name: 'read_deck', arguments: {} })
    expect(before.isError).toBe(true)
    expect(text(before.content)).toMatch(/no session is open — call create_session first/)

    const created = await client!.callTool({
      name: 'create_session',
      arguments: { family: 'pptx' },
    })
    expect(created.isError).toBeFalsy()

    const read = await client!.callTool({ name: 'read_deck', arguments: {} })
    expect(read.isError).toBeFalsy()
    expect(text(read.content)).toContain('"wc": 1')

    const apply = await client!.callTool({
      name: 'apply_slide_ops',
      arguments: {
        ops: [
          {
            op: 'addElement',
            target: { slide: 0 },
            kind: 'textbox',
            offset: { x: 0, y: 0, cx: 1, cy: 1 },
          },
        ],
      },
    })
    expect(apply.isError).toBeFalsy()
    expect(text(apply.content)).toContain('"applied": true')

    const saved = await client!.callTool({
      name: 'save_session',
      arguments: { path: join(dir, 'out.pptx') },
    })
    expect(saved.isError).toBeFalsy()
    expect(fake.saved).toEqual([{ path: join(dir, 'out.pptx') }])

    // the session ended with the save: further edits must ask for create_session
    const after = await client!.callTool({
      name: 'apply_slide_ops',
      arguments: { ops: [{ op: 'x' }] },
    })
    expect(after.isError).toBe(true)
    expect(text(after.content)).toMatch(/no session is open/)
  })

  it('surfaces failed transactions as errors, keeps dry runs successful, caps batch size', async () => {
    const fake = fakeSlidesControl()
    await startService(sessionSurface(fake.control))
    await client!.callTool({ name: 'create_session', arguments: { family: 'pptx' } })

    const failed = await client!.callTool({
      name: 'apply_slide_ops',
      arguments: { ops: [{ op: 'explode' }] },
    })
    expect(failed.isError).toBe(true)
    expect(text(failed.content)).toContain('boom')

    const dry = await client!.callTool({
      name: 'apply_slide_ops',
      arguments: { ops: [{ op: 'explode' }], dryRun: true },
    })
    expect(dry.isError).toBeFalsy()

    const tooBig = await client!.callTool({
      name: 'apply_slide_ops',
      arguments: { ops: Array.from({ length: 51 }, (_, i) => ({ op: `op${i}` })) },
    })
    expect(tooBig.isError).toBe(true)
  })

  it('disappear entirely without a control (headless runs keep the file-only surface)', async () => {
    await startService(createSlidesTools(baseDeps(), createSessionHost()))
    const names = (await client!.listTools()).tools.map((t) => t.name)
    for (const tool of ['create_session', 'read_deck', 'apply_slide_ops', 'save_session']) {
      expect(names).not.toContain(tool)
    }
  })
})
