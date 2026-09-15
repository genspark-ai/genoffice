/**
 * Tests for @genoffice/agent-skills/extensions/docs-skill.
 *
 * W4 deliverable: read_blocks migrated, first dialog (replace_document) working.
 *
 * The tests use a mock DocsEditor (no Tiptap dependency) and a real
 * ReactUIAdapter from @genoffice/agent-runtime so we exercise the full
 * tool registration → execute → UI dialog path.
 */

import { describe, expect, it, beforeEach, afterEach, vi } from 'vitest'
import { ReactUIAdapter } from '@genoffice/agent-runtime'
import {
  createDocsSkillExtension,
  createReadBlocksTool,
  createReplaceDocumentTool,
  type DocsEditor,
  type DocsBlock,
} from '../src/extensions/docs-skill'

// ---------------------------------------------------------------------------
// Mock editor
// ---------------------------------------------------------------------------

class MockDocsEditor implements DocsEditor {
  private blocks: DocsBlock[]
  replacedHtml: string | undefined = undefined
  selectionHtml: string | undefined = undefined
  selectionReplaced: string | undefined = undefined
  lastOps: { ops: ReadonlyArray<unknown>; dryRun: boolean } | null = null
  appliedCount = 0

  constructor(blocks: DocsBlock[]) {
    this.blocks = blocks
  }

  // Legacy: also supports replaceAll for the replace_document tool
  replaceAll(html: string): void {
    this.replacedHtml = html
  }

  getBlockCount(): number {
    return this.blocks.length
  }

  getBlock(index: number): DocsBlock {
    return this.blocks[index]!
  }

  getRangeHtml(start: number, end: number): string {
    return this.blocks
      .slice(start, end + 1)
      .map((b) => b.html)
      .join('\n')
  }

  clampRange(start: number, end: number): { start: number; end: number } | null {
    if (start < 0 || end < start) return null
    const last = this.blocks.length - 1
    if (start > last) return null
    return { start, end: Math.min(end, last) }
  }

  // ---- Mutation API ----
  insertBlocks(afterIndex: number, blocksHtml: string): { inserted: number } {
    const inserted = (blocksHtml.match(/<(p|h[1-6]|ul|ol|table|blockquote)[^>]*>/gi) ?? []).length || 1
    const at = afterIndex < 0 ? 0 : Math.min(afterIndex + 1, this.blocks.length)
    const newBlocks: DocsBlock[] = []
    for (let i = 0; i < inserted; i++) {
      newBlocks.push({ index: at + i, kind: 'paragraph', html: blocksHtml })
    }
    this.blocks.splice(at, 0, ...newBlocks)
    this.reindex()
    return { inserted }
  }

  replaceBlockRange(start: number, end: number, blocksHtml: string): { inserted: number; removed: number } {
    const removed = end - start + 1
    const inserted = (blocksHtml.match(/<(p|h[1-6]|ul|ol|table|blockquote)[^>]*>/gi) ?? []).length || 1
    this.blocks.splice(start, removed, { index: start, kind: 'paragraph', html: blocksHtml })
    this.reindex()
    return { inserted, removed }
  }

  replaceSelection(inlineHtml: string): { replaced: boolean } {
    if (this.selectionHtml === undefined) return { replaced: false }
    this.selectionReplaced = inlineHtml
    return { replaced: true }
  }

  applyOps(ops: ReadonlyArray<unknown>, dryRun: boolean): { applied: number; dryRun: boolean } {
    this.lastOps = { ops, dryRun }
    if (!dryRun) this.appliedCount += ops.length
    return { applied: ops.length, dryRun }
  }

  markDocSeen(): void {
    // no-op in mock
  }

  private reindex(): void {
    for (let i = 0; i < this.blocks.length; i++) {
      this.blocks[i]!.index = i
    }
  }
}

function makeDoc(): MockDocsEditor {
  return new MockDocsEditor([
    { index: 0, kind: 'heading', html: '<h1>Hello</h1>' },
    { index: 1, kind: 'paragraph', html: '<p>First paragraph with <em>emphasis</em>.</p>' },
    { index: 2, kind: 'paragraph', html: '<p>Second paragraph.</p>' },
    { index: 3, kind: 'list', html: '<ul><li>Item A</li><li>Item B</li></ul>' },
  ])
}

// ---------------------------------------------------------------------------
// read_blocks tests
// ---------------------------------------------------------------------------

describe('read_blocks tool', () => {
  let adapter: ReactUIAdapter
  let editor: MockDocsEditor

  beforeEach(() => {
    adapter = new ReactUIAdapter()
    editor = makeDoc()
    adapter.setEditorInstance(editor)
  })

  it('returns "no editor" when none is set', async () => {
    const a = new ReactUIAdapter()
    const tool = createReadBlocksTool({ uiAdapter: a })
    const result = await tool.execute('id-1', { startBlockIndex: 0, endBlockIndex: 1 }, undefined, undefined, {} as never)
    expect(((result.content[0] as { type: 'text'; text: string }).text)).toContain('No editor available')
  })

  it('reads a full range and returns concatenated HTML', async () => {
    const tool = createReadBlocksTool({ uiAdapter: adapter })
    const result = await tool.execute('id-2', { startBlockIndex: 1, endBlockIndex: 2 }, undefined, undefined, {} as never)
    const text = ((result.content[0] as { type: 'text'; text: string }).text)
    expect(text).toContain('First paragraph')
    expect(text).toContain('Second paragraph')
    expect(result.details).toEqual({ blockCount: 2, truncated: false, offset: text.length })
  })

  it('clamps out-of-range end to last block', async () => {
    const tool = createReadBlocksTool({ uiAdapter: adapter })
    const result = await tool.execute('id-3', { startBlockIndex: 1, endBlockIndex: 999 }, undefined, undefined, {} as never)
    expect(result.details.blockCount).toBe(3) // blocks 1,2,3
    expect(((result.content[0] as { type: 'text'; text: string }).text)).toContain('Item B')
  })

  it('returns error text for invalid range (start > blockCount)', async () => {
    const tool = createReadBlocksTool({ uiAdapter: adapter })
    const result = await tool.execute('id-4', { startBlockIndex: 50, endBlockIndex: 60 }, undefined, undefined, {} as never)
    expect(((result.content[0] as { type: 'text'; text: string }).text)).toContain('Invalid range')
    expect(((result.content[0] as { type: 'text'; text: string }).text)).toContain('4 block(s)')
  })

  it('returns error text for negative range', async () => {
    const tool = createReadBlocksTool({ uiAdapter: adapter })
    const result = await tool.execute('id-5', { startBlockIndex: 2, endBlockIndex: 1 }, undefined, undefined, {} as never)
    expect(((result.content[0] as { type: 'text'; text: string }).text)).toContain('Invalid range')
  })

  it('paginates with offset (simulate large doc)', async () => {
    // Build a single big block to force truncation
    const big = '<p>' + 'x'.repeat(250_000) + '</p>'
    const bigDoc = new MockDocsEditor([{ index: 0, kind: 'paragraph', html: big }])
    adapter.setEditorInstance(bigDoc)
    const tool = createReadBlocksTool({ uiAdapter: adapter })

    const page1 = await tool.execute('id-6', { startBlockIndex: 0, endBlockIndex: 0 }, undefined, undefined, {} as never)
    const text1 = ((page1.content[0] as { type: 'text'; text: string }).text)
    expect(page1.details.truncated).toBe(true)
    const match = text1.match(/offset=(\d+)/)
    expect(match).not.toBeNull()
    const nextOffset = Number(match![1])

    const page2 = await tool.execute('id-7', { startBlockIndex: 0, endBlockIndex: 0, offset: nextOffset }, undefined, undefined, {} as never)
    expect(page2.details.truncated).toBe(false)
    expect(((page2.content[0] as { type: 'text'; text: string }).text)).toContain('(end of range')
  })

  it('rejects offset beyond content length', async () => {
    const tool = createReadBlocksTool({ uiAdapter: adapter })
    const result = await tool.execute('id-8', { startBlockIndex: 0, endBlockIndex: 0, offset: 99_999 }, undefined, undefined, {} as never)
    expect(((result.content[0] as { type: 'text'; text: string }).text)).toContain('offset 99999 is beyond the content')
  })
})

// ---------------------------------------------------------------------------
// replace_document dialog tests
// ---------------------------------------------------------------------------

describe('replace_document tool (first dialog)', () => {
  let adapter: ReactUIAdapter
  let editor: MockDocsEditor

  beforeEach(() => {
    adapter = new ReactUIAdapter()
    editor = makeDoc()
    adapter.setEditorInstance(editor)
  })

  it('shows a confirm dialog before replacing', async () => {
    const tool = createReplaceDocumentTool({ uiAdapter: adapter })
    const pendingDialogs = vi.fn()
    adapter.onDialogs(pendingDialogs)

    const promise = tool.execute('id-9', { html: '<p>New content</p>', reason: 'Reset to clean draft' }, undefined, undefined, {} as never)

    // Yield so pushDialog can run
    await new Promise((r) => setTimeout(r, 5))
    expect(adapter.dialogs.length).toBe(1)
    expect(adapter.dialogs[0]!.kind).toBe('confirm')
    expect(adapter.dialogs[0]!.title).toBe('Replace entire document?')
    const dlg0 = adapter.dialogs[0] as { kind: 'confirm'; message: string }
    expect(dlg0.message).toContain('Reset to clean draft')
    expect(pendingDialogs).toHaveBeenCalled()

    // User clicks OK
    adapter.resolveDialog(adapter.dialogs[0]!.id, true)
    const result = await promise
    expect(result.details.replaced).toBe(true)
    expect(editor.replacedHtml).toBe('<p>New content</p>')
  })

  it('respects user cancellation (does NOT replace)', async () => {
    const tool = createReplaceDocumentTool({ uiAdapter: adapter })
    const promise = tool.execute('id-10', { html: '<p>New</p>', reason: 'Try' }, undefined, undefined, {} as never)

    await new Promise((r) => setTimeout(r, 5))
    expect(adapter.dialogs.length).toBe(1)
    adapter.resolveDialog(adapter.dialogs[0]!.id, false)
    const result = await promise
    expect(result.details.replaced).toBe(false)
    expect(result.details.reason).toBe('user_cancelled')
    expect(editor.replacedHtml).toBeUndefined()
    expect(((result.content[0] as { type: 'text'; text: string }).text)).toContain('User cancelled')
  })

  it('auto-cancels on dialog timeout', async () => {
    const tool = createReplaceDocumentTool({ uiAdapter: adapter, confirmTimeoutMs: 50 })
    const result = await tool.execute('id-11', { html: '<p>New</p>', reason: 'Try' }, undefined, undefined, {} as never)
    expect(result.details.replaced).toBe(false)
    expect(editor.replacedHtml).toBeUndefined()
  })
})

// ---------------------------------------------------------------------------
// Extension factory end-to-end (real session, real LLM call)
// ---------------------------------------------------------------------------

describe('docs-skill extension factory', () => {
  let adapter: ReactUIAdapter
  let editor: MockDocsEditor

  beforeEach(() => {
    adapter = new ReactUIAdapter()
    editor = makeDoc()
    adapter.setEditorInstance(editor)
  })

  afterEach(() => {
    vi.restoreAllMocks()
  })

  it('registers read_blocks and replace_document as pi tools', () => {
    // We don't need a full session for this; just verify the factory shape
    // matches what createAgentSession expects.
    const factory = createDocsSkillExtension({ uiAdapter: adapter })
    expect(typeof factory).toBe('function')

    // And calling the factory with a fake pi API should register two tools
    const registered: Array<{ name: string }> = []
    const pi = {
      registerTool: (tool: { name: string }) => registered.push({ name: tool.name }),
      on: () => {},
    } as never
    factory(pi)
    const names = registered.map((t) => t.name)
    expect(names).toContain('read_blocks')
    expect(names).toContain('replace_document')
  })

  it('can be loaded into a real pi session via createOfficeSession + factory', async () => {
    const { createOfficeSession } = await import('@genoffice/agent-runtime')
    const { session, dispose } = await createOfficeSession({
      cwd: process.cwd(),
      extensionFactories: [createDocsSkillExtension({ uiAdapter: adapter, enabledTools: ['read_blocks', 'get_document_context', 'insert_content', 'replace_blocks', 'replace_selection', 'apply_ops', 'create_document', 'read_comments', 'reply_comment', 'resolve_comment'] })],
    })

    // getAllTools returns ToolInfo[] (name + metadata), getToolDefinition returns
    // the full ToolDefinition (with execute). We use both to verify wiring.
    const toolInfos = session.getAllTools()
    const names = toolInfos.map((t) => t.name)
    expect(names).toContain('read_blocks')
    expect(names).not.toContain('replace_document') // not in enabledTools

    const readBlocks = session.getToolDefinition('read_blocks')
    expect(readBlocks).toBeDefined()
    expect(typeof readBlocks!.execute).toBe('function')

    // Execute the tool — this proves the wire works end-to-end (extension
    // factory → pi session → registered tool → ui adapter → mock editor)
    // without needing a real LLM call.
    const result = await readBlocks!.execute(
      'id-12',
      { startBlockIndex: 0, endBlockIndex: 1 },
      undefined,
      undefined,
      // minimal ctx — the tool doesn't use it
      {} as never,
    )
    expect(((result.content[0] as { type: 'text'; text: string }).text)).toContain('<h1>Hello</h1>')
    expect(((result.content[0] as { type: 'text'; text: string }).text)).toContain('First paragraph')

    dispose()
  }, 30_000)
})

// ============================================================================
// W6: additional migrated tools
// ============================================================================

describe('get_document_context tool', () => {
  let adapter: ReactUIAdapter
  let editor: MockDocsEditor

  beforeEach(() => {
    adapter = new ReactUIAdapter()
    editor = makeDoc()
    adapter.setEditorInstance(editor)
  })

  it('returns block list with index + kind + preview', async () => {
    const { createGetDocumentContextTool } = await import('../src/extensions/docs-skill')
    const tool = createGetDocumentContextTool({ uiAdapter: adapter })
    const result = await tool.execute('id', {}, undefined, undefined, {} as never)
    const text = (result.content[0] as { type: 'text'; text: string }).text
    expect(text).toContain('4 block(s)')
    expect(text).toContain('[0] heading: <h1>Hello</h1>')
    expect(text).toContain('[1] paragraph:')
    expect(text).toContain('[3] list:')
    expect(result.details.blockCount).toBe(4)
  })

  it('truncates very long block lists', async () => {
    const many: DocsBlock[] = Array.from({ length: 500 }, (_, i) => ({
      index: i,
      kind: 'paragraph',
      html: `<p>Block ${i} ${'x'.repeat(200)}</p>`,
    }))
    const bigDoc = new MockDocsEditor(many)
    adapter.setEditorInstance(bigDoc)
    const { createGetDocumentContextTool } = await import('../src/extensions/docs-skill')
    const tool = createGetDocumentContextTool({ uiAdapter: adapter })
    const result = await tool.execute('id', {}, undefined, undefined, {} as never)
    const text = (result.content[0] as { type: 'text'; text: string }).text
    expect(text).toContain('…(truncated)')
    expect(result.details.blockCount).toBe(500)
  })
})

describe('insert_content tool', () => {
  let adapter: ReactUIAdapter
  let editor: MockDocsEditor

  beforeEach(() => {
    adapter = new ReactUIAdapter()
    editor = makeDoc()
    adapter.setEditorInstance(editor)
  })

  it('inserts after a given index and updates block count', async () => {
    const { createInsertContentTool } = await import('../src/extensions/docs-skill')
    const tool = createInsertContentTool({ uiAdapter: adapter })
    const before = editor.getBlockCount()
    const result = await tool.execute('id', { html: '<p>Inserted</p>', afterBlockIndex: 1 }, undefined, undefined, {} as never)
    expect(editor.getBlockCount()).toBeGreaterThan(before)
    const text = (result.content[0] as { type: 'text'; text: string }).text
    expect(text).toContain('Inserted')
    expect(result.details.afterIndex).toBe(1)
  })

  it('defaults to appending at the end when afterBlockIndex is omitted', async () => {
    const { createInsertContentTool } = await import('../src/extensions/docs-skill')
    const tool = createInsertContentTool({ uiAdapter: adapter })
    const before = editor.getBlockCount()
    const result = await tool.execute('id', { html: '<p>Tail</p>' }, undefined, undefined, {} as never)
    expect(editor.getBlockCount()).toBe(before + 1)
    expect(result.details.afterIndex).toBe(before - 1)
  })

  it('returns "no editor" when editor missing', async () => {
    const a = new ReactUIAdapter()
    const { createInsertContentTool } = await import('../src/extensions/docs-skill')
    const tool = createInsertContentTool({ uiAdapter: a })
    const result = await tool.execute('id', { html: '<p>x</p>' }, undefined, undefined, {} as never)
    expect((result.content[0] as { type: 'text'; text: string }).text).toContain('No editor')
  })
})

describe('replace_blocks tool', () => {
  let adapter: ReactUIAdapter
  let editor: MockDocsEditor

  beforeEach(() => {
    adapter = new ReactUIAdapter()
    editor = makeDoc()
    adapter.setEditorInstance(editor)
  })

  it('replaces a range and reports the counts', async () => {
    const { createReplaceBlocksTool } = await import('../src/extensions/docs-skill')
    const tool = createReplaceBlocksTool({ uiAdapter: adapter })
    const result = await tool.execute('id', { startBlockIndex: 1, endBlockIndex: 2, html: '<p>Replaced</p>' }, undefined, undefined, {} as never)
    const text = (result.content[0] as { type: 'text'; text: string }).text
    expect(text).toContain('Replaced 2 block(s)')
    expect(result.details).toEqual({ removed: 2, inserted: 1 })
  })

  it('rejects invalid range', async () => {
    const { createReplaceBlocksTool } = await import('../src/extensions/docs-skill')
    const tool = createReplaceBlocksTool({ uiAdapter: adapter })
    const result = await tool.execute('id', { startBlockIndex: 100, endBlockIndex: 200, html: '<p>x</p>' }, undefined, undefined, {} as never)
    expect((result.content[0] as { type: 'text'; text: string }).text).toContain('Invalid range')
  })
})

describe('replace_selection tool', () => {
  let adapter: ReactUIAdapter
  let editor: MockDocsEditor

  beforeEach(() => {
    adapter = new ReactUIAdapter()
    editor = makeDoc()
    editor.selectionHtml = '<p>some selected text</p>'
    adapter.setEditorInstance(editor)
  })

  it('replaces when a selection is active', async () => {
    const { createReplaceSelectionTool } = await import('../src/extensions/docs-skill')
    const tool = createReplaceSelectionTool({ uiAdapter: adapter })
    const result = await tool.execute('id', { html: 'New text' }, undefined, undefined, {} as never)
    expect(result.details.replaced).toBe(true)
    expect(editor.selectionReplaced).toBe('New text')
  })

  it('returns helpful error when no selection', async () => {
    editor.selectionHtml = undefined
    const { createReplaceSelectionTool } = await import('../src/extensions/docs-skill')
    const tool = createReplaceSelectionTool({ uiAdapter: adapter })
    const result = await tool.execute('id', { html: 'New' }, undefined, undefined, {} as never)
    const text = (result.content[0] as { type: 'text'; text: string }).text
    expect(text).toContain('No active text selection')
    expect(text).toContain('replace_blocks')
  })
})

describe('apply_ops tool', () => {
  let adapter: ReactUIAdapter
  let editor: MockDocsEditor

  beforeEach(() => {
    adapter = new ReactUIAdapter()
    editor = makeDoc()
    adapter.setEditorInstance(editor)
  })

  it('passes ops to the editor and reports count', async () => {
    const { createApplyOpsTool } = await import('../src/extensions/docs-skill')
    const tool = createApplyOpsTool({ uiAdapter: adapter })
    const ops = [{ op: 'setFont', target: { nodeType: 'docHeading' }, color: '#FF0000' }]
    const result = await tool.execute('id', { ops }, undefined, undefined, {} as never)
    expect(editor.lastOps?.ops).toEqual(ops)
    expect(result.details.applied).toBe(1)
    expect(result.details.dryRun).toBe(false)
  })

  it('respects dryRun flag', async () => {
    const { createApplyOpsTool } = await import('../src/extensions/docs-skill')
    const tool = createApplyOpsTool({ uiAdapter: adapter })
    const beforeApplied = editor.appliedCount
    const result = await tool.execute('id', { ops: [{ op: 'x' }], dryRun: true }, undefined, undefined, {} as never)
    expect(editor.appliedCount).toBe(beforeApplied) // unchanged
    expect(result.details.dryRun).toBe(true)
  })

  it('rejects empty ops array', async () => {
    const { createApplyOpsTool } = await import('../src/extensions/docs-skill')
    const tool = createApplyOpsTool({ uiAdapter: adapter })
    const result = await tool.execute('id', { ops: [] }, undefined, undefined, {} as never)
    expect((result.content[0] as { type: 'text'; text: string }).text).toContain('empty')
  })
})

describe('create_document tool', () => {
  let adapter: ReactUIAdapter

  beforeEach(() => {
    adapter = new ReactUIAdapter()
  })

  it('requires createNewDocument on the editor', async () => {
    const editor = makeDoc()
    adapter.setEditorInstance(editor)
    const { createCreateDocumentTool } = await import('../src/extensions/docs-skill')
    const tool = createCreateDocumentTool({ uiAdapter: adapter })
    const result = await tool.execute('id', { html: '<p>New doc</p>', title: 'Untitled' }, undefined, undefined, {} as never)
    expect((result.content[0] as { type: 'text'; text: string }).text).toContain('createNewDocument')
  })

  it('invokes createNewDocument when present', async () => {
    const editor = Object.assign(makeDoc(), {
      createNewDocument: vi.fn(),
    })
    adapter.setEditorInstance(editor)
    const { createCreateDocumentTool } = await import('../src/extensions/docs-skill')
    const tool = createCreateDocumentTool({ uiAdapter: adapter })
    const result = await tool.execute('id', { html: '<p>New</p>', title: 'Doc' }, undefined, undefined, {} as never)
    expect(editor.createNewDocument).toHaveBeenCalledWith('<p>New</p>', 'Doc')
    expect(result.details.created).toBe(true)
    expect(result.details.title).toBe('Doc')
  })
})

describe('comment tools (read_comments, reply_comment, resolve_comment)', () => {
  let adapter: ReactUIAdapter
  let editor: MockDocsEditor

  beforeEach(() => {
    adapter = new ReactUIAdapter()
    editor = makeDoc()
    adapter.setEditorInstance(editor)
  })

  it('read_comments returns empty when no comments stored', async () => {
    const { createReadCommentsTool } = await import('../src/extensions/docs-skill')
    const tool = createReadCommentsTool({ uiAdapter: adapter })
    const result = await tool.execute('id', {}, undefined, undefined, {} as never)
    expect((result.content[0] as { type: 'text'; text: string }).text).toContain('No comments')
    expect(result.details.count).toBe(0)
  })

  it('read_comments lists stored threads', async () => {
    adapter.setCustomData('comments', [
      { id: 'c1', author: 'Alice', date: '2026-09-15T10:00:00Z', blockIndex: 2, anchorText: 'revenue', resolved: false, replies: [] },
      { id: 'c2', author: 'Bob', date: '2026-09-14T10:00:00Z', blockIndex: 4, anchorText: 'enterprise', resolved: true, replies: [{ author: 'AI', date: '2026-09-14T11:00:00Z', text: 'fixed' }] },
    ])
    const { createReadCommentsTool } = await import('../src/extensions/docs-skill')
    const tool = createReadCommentsTool({ uiAdapter: adapter })
    const result = await tool.execute('id', {}, undefined, undefined, {} as never)
    const text = (result.content[0] as { type: 'text'; text: string }).text
    expect(text).toContain('[c1]')
    expect(text).toContain('Alice')
    expect(text).toContain('[c2]')
    expect(text).toContain('(resolved)')
    expect(result.details.count).toBe(2)
  })

  it('reply_comment appends to the thread', async () => {
    adapter.setCustomData('comments', [
      { id: 'c1', author: 'Alice', date: '2026-09-15T10:00:00Z', blockIndex: 2, anchorText: 'revenue', resolved: false, replies: [] },
    ])
    const { createReplyCommentTool } = await import('../src/extensions/docs-skill')
    const tool = createReplyCommentTool({ uiAdapter: adapter })
    const result = await tool.execute('id', { parentId: 'c1', text: 'Done!' }, undefined, undefined, {} as never)
    expect(result.details.replied).toBe(true)
    const comments = adapter.getCustomData<Array<{ replies: unknown[] }>>('comments')!
    expect(comments[0]!.replies.length).toBe(1)
  })

  it('reply_comment returns not-found for missing parent', async () => {
    const { createReplyCommentTool } = await import('../src/extensions/docs-skill')
    const tool = createReplyCommentTool({ uiAdapter: adapter })
    const result = await tool.execute('id', { parentId: 'missing', text: 'x' }, undefined, undefined, {} as never)
    expect((result.content[0] as { type: 'text'; text: string }).text).toContain('not found')
  })

  it('resolve_comment marks the thread as resolved', async () => {
    adapter.setCustomData('comments', [
      { id: 'c1', author: 'Alice', date: '2026-09-15T10:00:00Z', blockIndex: 2, anchorText: 'revenue', resolved: false, replies: [] },
    ])
    const { createResolveCommentTool } = await import('../src/extensions/docs-skill')
    const tool = createResolveCommentTool({ uiAdapter: adapter })
    const result = await tool.execute('id', { id: 'c1' }, undefined, undefined, {} as never)
    expect(result.details.resolved).toBe(true)
    const comments = adapter.getCustomData<Array<{ resolved: boolean }>>('comments')!
    expect(comments[0]!.resolved).toBe(true)
  })
})

// ============================================================================
// W6: extension factory with all tools
// ============================================================================

describe('W6: full docs-skill extension', () => {
  let adapter: ReactUIAdapter
  let editor: MockDocsEditor

  beforeEach(() => {
    adapter = new ReactUIAdapter()
    editor = makeDoc()
    adapter.setEditorInstance(editor)
  })

  it('registers all 11 docs tools by default', () => {
    const registered: string[] = []
    const pi = {
      registerTool: (tool: { name: string }) => registered.push(tool.name),
      on: () => {},
    } as never
    createDocsSkillExtension({ uiAdapter: adapter })(pi)
    const expected = [
      'read_blocks', 'get_document_context', 'insert_content', 'replace_blocks',
      'replace_selection', 'apply_ops', 'create_document', 'replace_document',
      'read_comments', 'reply_comment', 'resolve_comment',
    ]
    for (const name of expected) expect(registered).toContain(name)
    expect(registered.length).toBe(expected.length)
  })

  it('enabledTools restricts the registered set', () => {
    const registered: string[] = []
    const pi = {
      registerTool: (tool: { name: string }) => registered.push(tool.name),
      on: () => {},
    } as never
    createDocsSkillExtension({ uiAdapter: adapter, enabledTools: ['read_blocks', 'get_document_context'] })(pi)
    expect(registered).toEqual(['read_blocks', 'get_document_context'])
  })

  it('loads all tools into a real pi session', async () => {
    const { createOfficeSession } = await import('@genoffice/agent-runtime')
    const { session, dispose } = await createOfficeSession({
      cwd: process.cwd(),
      extensionFactories: [createDocsSkillExtension({ uiAdapter: adapter })],
    })
    const toolNames = session.getAllTools().map((t) => t.name)
    expect(toolNames).toContain('read_blocks')
    expect(toolNames).toContain('get_document_context')
    expect(toolNames).toContain('insert_content')
    expect(toolNames).toContain('replace_blocks')
    expect(toolNames).toContain('replace_selection')
    expect(toolNames).toContain('apply_ops')
    expect(toolNames).toContain('create_document')
    expect(toolNames).toContain('replace_document')
    expect(toolNames).toContain('read_comments')
    expect(toolNames).toContain('reply_comment')
    expect(toolNames).toContain('resolve_comment')
    dispose()
  }, 30_000)
})
