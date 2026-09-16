import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { mkdtemp, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import { Client } from '@modelcontextprotocol/sdk/client/index.js'
import { StreamableHTTPClientTransport } from '@modelcontextprotocol/sdk/client/streamableHttp.js'
import type { OpenDocumentTab } from '../../src/shared/tabs-api'
import { McpServerService } from '../../src/main/mcp/mcp-server'
import {
  closeSavePath,
  createOpenDocumentTools,
  resolveOpenDocument,
  type OpenDocumentsControl,
} from '../../src/main/mcp/tools/open-documents-tools'

/**
 * `open_documents` over a real MCP session: the list/read/close actions against
 * a fake control, plus the pure helpers (target resolution, the default save
 * path an untitled dirty document is written to before closing).
 */

let service: McpServerService | undefined
let client: Client | undefined
let dir: string

async function freePort(): Promise<number> {
  const { createServer } = await import('node:http')
  return new Promise((resolvePort, reject) => {
    const server = createServer()
    server.on('error', reject)
    server.listen(0, '127.0.0.1', () => {
      const address = server.address()
      const port = typeof address === 'object' && address ? address.port : 0
      server.close(() => resolvePort(port))
    })
  })
}

beforeEach(async () => {
  dir = await mkdtemp(join(tmpdir(), 'genoffice-mcp-open-docs-'))
})

afterEach(async () => {
  await client?.close()
  client = undefined
  await service?.stop()
  service = undefined
  await rm(dir, { recursive: true, force: true })
})

async function startService(control?: OpenDocumentsControl): Promise<void> {
  const port = await freePort()
  service = new McpServerService({
    port,
    tools: createOpenDocumentTools({ defaultSaveDir: () => dir, ...(control ? { control } : {}) }),
  })
  await service.start()
  client = new Client({ name: 'open-documents-test', version: '0.0.0' })
  await client.connect(new StreamableHTTPClientTransport(new URL(`http://127.0.0.1:${port}/mcp`)))
}

/** callTool's result carries the payload on `.content`; accept either shape */
function textOf(result: unknown): string {
  const content = (result as { content?: unknown })?.content ?? result
  const parts = Array.isArray(content) ? content : [content]
  return parts.map((part) => (part as { text?: string } | undefined)?.text ?? '').join('')
}

/** the handler serializes its return value as JSON text */
function jsonOf(result: unknown): Record<string, unknown> {
  return JSON.parse(textOf(result)) as Record<string, unknown>
}

/** a document as the tab manager reports it */
function doc(overrides: Partial<OpenDocumentTab> & { id: string }): OpenDocumentTab {
  return {
    kind: 'markdown',
    title: 'notes.md',
    active: false,
    dirty: false,
    ...overrides,
  } as OpenDocumentTab
}

/** the six-open-document scenario: three the user opened, one per family edge case */
const DOCUMENTS: OpenDocumentTab[] = [
  doc({ id: 't1', kind: 'docs', title: 'one.docx', filePath: 'C:/docs/one.docx', dirty: true }),
  doc({ id: 't2', kind: 'sheets', title: 'two.xlsx', filePath: 'C:/docs/two.xlsx' }),
  doc({
    id: 't3',
    kind: 'slides',
    title: 'three.pptx',
    filePath: 'C:/docs/three.pptx',
    dirty: true,
  }),
  doc({ id: 't4', kind: 'markdown', title: 'four.md', filePath: 'C:/docs/four.md' }),
  doc({ id: 't5', kind: 'html', title: 'five.html', filePath: 'C:/docs/five.html' }),
  doc({ id: 't6', kind: 'pdf', title: 'six.pdf', filePath: 'C:/docs/six.pdf' }),
  doc({ id: 't7', kind: 'docs', title: 'Untitled', dirty: true, active: true }),
]

function controlOf(overrides: Partial<OpenDocumentsControl> = {}): OpenDocumentsControl {
  return {
    list: vi.fn(async () => DOCUMENTS),
    read: vi.fn(async (tab: OpenDocumentTab) => ({ content: `live:${tab.id}` })),
    close: vi.fn(async () => ({})),
    ...overrides,
  }
}

describe('resolveOpenDocument', () => {
  it('matches by exact tab id first', () => {
    expect(resolveOpenDocument(DOCUMENTS, 't3')?.id).toBe('t3')
  })

  it('matches by path, including a path spelled differently', () => {
    const other = resolve(DOCUMENTS[0]!.filePath!)
    expect(resolveOpenDocument(DOCUMENTS, other)?.id).toBe('t1')
  })

  it('returns undefined when nothing matches', () => {
    expect(resolveOpenDocument(DOCUMENTS, 'C:/docs/nope.docx')).toBeUndefined()
  })
})

describe('closeSavePath', () => {
  it('reuses the document path when it has one', () => {
    expect(closeSavePath(DOCUMENTS[0]!, () => dir)).toBe('C:/docs/one.docx')
  })

  it('falls back to the default folder for an untitled document, with its extension', () => {
    const target = closeSavePath(DOCUMENTS[6]!, () => dir)
    expect(target.startsWith(resolve(dir))).toBe(true)
    expect(target.endsWith('.docx')).toBe(true)
  })

  it('never hands back a path that already exists', async () => {
    const { writeFile } = await import('node:fs/promises')
    const first = closeSavePath(DOCUMENTS[6]!, () => dir)
    await writeFile(first, 'x')
    // a second untitled close must not overwrite the first
    expect(closeSavePath(DOCUMENTS[6]!, () => dir)).not.toBe(first)
  })
})

describe('open_documents tool', () => {
  it('lists every open document with type, path and unsaved state', async () => {
    await startService(controlOf())
    const result = jsonOf(
      await client!.callTool({ name: 'open_documents', arguments: { action: 'list' } }),
    )
    const documents = result.documents as Array<Record<string, unknown>>
    expect(documents).toHaveLength(DOCUMENTS.length)

    const first = documents[0]!
    expect(first.id).toBe('t1')
    expect(first.type).toBe('Word document')
    expect(first.path).toBe('C:/docs/one.docx')
    expect(first.saved).toBe(true)
    expect(first.dirty).toBe(true)

    // an untitled document reports no path and saved:false
    const untitled = documents.find((d) => d.id === 't7')!
    expect(untitled.path).toBeNull()
    expect(untitled.saved).toBe(false)
    expect(untitled.active).toBe(true)

    // the family vocabulary carries through for every kind
    const types = Object.fromEntries(documents.map((d) => [d.id, d.type]))
    expect(types).toMatchObject({
      t2: 'spreadsheet',
      t3: 'presentation',
      t4: 'Markdown document',
      t5: 'HTML page',
      t6: 'PDF document',
    })
  })

  it('reads one document by path and reports its live content', async () => {
    const read = vi.fn(async (tab: OpenDocumentTab) => ({ content: `live:${tab.id}` }))
    await startService(controlOf({ read }))
    const result = jsonOf(
      await client!.callTool({
        name: 'open_documents',
        arguments: { action: 'read', target: 'C:/docs/four.md' },
      }),
    )
    expect(result.id).toBe('t4')
    expect(result.dirty).toBe(false)
    expect(result.content).toEqual({ content: 'live:t4' })
    expect(read).toHaveBeenCalledTimes(1)
  })

  it('reads an untitled document by its tab id', async () => {
    await startService(controlOf())
    const result = jsonOf(
      await client!.callTool({
        name: 'open_documents',
        arguments: { action: 'read', target: 't7' },
      }),
    )
    expect(result.id).toBe('t7')
    expect(result.path).toBeNull()
  })

  it('closes with save by default, reporting where it was saved', async () => {
    const close = vi.fn(async () => ({ savedPath: 'C:/docs/three.pptx' }))
    await startService(controlOf({ close }))
    const result = jsonOf(
      await client!.callTool({
        name: 'open_documents',
        arguments: { action: 'close', target: 'C:/docs/three.pptx' },
      }),
    )
    expect(result.closed).toBe(true)
    expect(result.savedTo).toBe('C:/docs/three.pptx')
    expect(close).toHaveBeenCalledWith(expect.objectContaining({ id: 't3' }), { unsaved: 'save' })
  })

  it('passes discard through and flags that changes were dropped', async () => {
    const close = vi.fn(async () => ({}))
    await startService(controlOf({ close }))
    const result = jsonOf(
      await client!.callTool({
        name: 'open_documents',
        arguments: { action: 'close', target: 't3', unsaved: 'discard' },
      }),
    )
    expect(close).toHaveBeenCalledWith(expect.objectContaining({ id: 't3' }), {
      unsaved: 'discard',
    })
    expect(result.discardedUnsavedChanges).toBe(true)
  })

  it('refuses to read a PDF, naming the reason', async () => {
    await startService(controlOf())
    const result = await client!.callTool({
      name: 'open_documents',
      arguments: { action: 'read', target: 'C:/docs/six.pdf' },
    })
    expect(result.isError).toBe(true)
    expect(textOf(result.content)).toContain('viewer')
  })

  it('lists an unknown target as an error that names what is open', async () => {
    await startService(controlOf())
    const result = await client!.callTool({
      name: 'open_documents',
      arguments: { action: 'close', target: 'C:/docs/missing.docx' },
    })
    expect(result.isError).toBe(true)
    const message = textOf(result.content)
    expect(message).toContain('no open document matches')
    expect(message).toContain('C:/docs/one.docx')
  })

  it('requires a target for read and close', async () => {
    await startService(controlOf())
    for (const action of ['read', 'close']) {
      const result = await client!.callTool({ name: 'open_documents', arguments: { action } })
      expect(result.isError).toBe(true)
      expect(textOf(result.content)).toContain('needs a target')
    }
  })

  it('explains an empty workspace instead of failing silently', async () => {
    await startService(controlOf({ list: vi.fn(async () => []) }))
    const listed = jsonOf(
      await client!.callTool({ name: 'open_documents', arguments: { action: 'list' } }),
    )
    expect(listed.documents).toEqual([])

    const closed = await client!.callTool({
      name: 'open_documents',
      arguments: { action: 'close', target: 'x.md' },
    })
    expect(closed.isError).toBe(true)
    expect(textOf(closed.content)).toContain('no document is open')
  })

  it('reports a clear error when the app is not running', async () => {
    await startService(undefined)
    const result = await client!.callTool({ name: 'open_documents', arguments: { action: 'list' } })
    expect(result.isError).toBe(true)
    expect(textOf(result.content)).toContain('GenOffice is not running')
  })
})
