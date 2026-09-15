/**
 * End-to-end test for read_blocks (W5 deliverable).
 *
 * Unlike the unit tests in `docs-skill.test.ts` that call the tool directly,
 * this e2e exercises the *full* pi agent pipeline:
 *   1. Real createOfficeSession() wires ReactUIAdapter as UI context
 *   2. Docs-skill extension is loaded via extensionFactories
 *   3. A mock editor is attached via uiAdapter.setEditorInstance()
 *   4. We invoke the read_blocks tool *through the session's tool dispatcher*
 *      (the same code path the LLM uses when it emits a tool_call)
 *   5. The result flows back as a real AgentToolResult
 *
 * This proves the W4 migration is wired correctly end-to-end without needing
 * a full Electron build or a real LLM call. The W8 UI integration is a
 * separate deliverable that will replace step 4 with an actual model call.
 */

import { describe, expect, it, beforeEach } from 'vitest'
import { createOfficeSession, ReactUIAdapter } from '@genoffice/agent-runtime'
import { createDocsSkillExtension, type DocsEditor, type DocsBlock } from '../src/extensions/docs-skill'

// ---------------------------------------------------------------------------
// Mock editor with realistic content
// ---------------------------------------------------------------------------

function makeRealisticDoc(): DocsEditor & { replacedHtml?: string } {
  const blocks: DocsBlock[] = [
    { index: 0, kind: 'heading', html: '<h1>Quarterly Report</h1>' },
    { index: 1, kind: 'paragraph', html: '<p>This report covers Q3 2026 financial highlights.</p>' },
    { index: 2, kind: 'heading', html: '<h2>Revenue</h2>' },
    { index: 3, kind: 'paragraph', html: '<p>Revenue grew by <strong>23%</strong> year-over-year.</p>' },
    { index: 4, kind: 'list', html: '<ul><li>Enterprise: +35%</li><li>SMB: +18%</li><li>Consumer: +12%</li></ul>' },
    { index: 5, kind: 'heading', html: '<h2>Risks</h2>' },
    { index: 6, kind: 'paragraph', html: '<p>Main risks include FX volatility and supply chain delays.</p>' },
  ]
  let replacedHtml: string | undefined
  let selectionHtml: string | undefined = '<p>placeholder selection</p>'
  let lastOps: { ops: ReadonlyArray<unknown>; dryRun: boolean } | null = null
  return {
    getBlockCount: () => blocks.length,
    getBlock: (i) => blocks[i]!,
    getRangeHtml: (s, e) => blocks.slice(s, e + 1).map((b) => b.html).join('\n'),
    clampRange: (s, e) => {
      if (s < 0 || e < s) return null
      if (s > blocks.length - 1) return null
      return { start: s, end: Math.min(e, blocks.length - 1) }
    },
    insertBlocks: (afterIndex, html) => {
      const inserted = (html.match(/<(p|h[1-6]|ul|ol|table|blockquote)[^>]*>/gi) ?? []).length || 1
      blocks.splice(afterIndex + 1, 0, ...Array.from({ length: inserted }, () => ({ index: 0, kind: 'paragraph', html })))
      return { inserted }
    },
    replaceBlockRange: (s, e, html) => {
      const removed = e - s + 1
      const inserted = (html.match(/<(p|h[1-6]|ul|ol|table|blockquote)[^>]*>/gi) ?? []).length || 1
      blocks.splice(s, removed, { index: s, kind: 'paragraph', html })
      return { inserted, removed }
    },
    replaceSelection: (html) => {
      if (selectionHtml === undefined) return { replaced: false }
      selectionHtml = html
      return { replaced: true }
    },
    applyOps: (ops, dryRun) => {
      lastOps = { ops, dryRun }
      return { applied: ops.length, dryRun }
    },
    markDocSeen: () => {},
    get replacedHtml() {
      return replacedHtml
    },
    set replacedHtml(v) {
      replacedHtml = v
    },
  }
}

describe('e2e: read_blocks through the real pi session', () => {
  let adapter: ReactUIAdapter
  let editor: ReturnType<typeof makeRealisticDoc>

  beforeEach(() => {
    adapter = new ReactUIAdapter()
    editor = makeRealisticDoc()
    adapter.setEditorInstance(editor)
  })

  it('loads the extension into a real session and exposes the tool', async () => {
    const { session, dispose } = await createOfficeSession({
      cwd: process.cwd(),
      extensionFactories: [createDocsSkillExtension({ uiAdapter: adapter, enabledTools: ['read_blocks', 'get_document_context', 'insert_content', 'replace_blocks', 'replace_selection', 'apply_ops', 'create_document', 'read_comments', 'reply_comment', 'resolve_comment'] })],
    })

    const toolNames = session.getAllTools().map((t) => t.name)
    expect(toolNames).toContain('read_blocks')
    expect(toolNames).not.toContain('replace_document') // not in the enabledTools list

    dispose()
  }, 30_000)

  it('executes read_blocks through the real session dispatcher and returns HTML', async () => {
    const { session, dispose } = await createOfficeSession({
      cwd: process.cwd(),
      extensionFactories: [createDocsSkillExtension({ uiAdapter: adapter, enabledTools: ['read_blocks', 'get_document_context', 'insert_content', 'replace_blocks', 'replace_selection', 'apply_ops', 'create_document', 'read_comments', 'reply_comment', 'resolve_comment'] })],
    })

    // This is the exact code path the LLM tool call would trigger.
    const tool = session.getToolDefinition('read_blocks')!
    const result = await tool.execute(
      'e2e-call-1',
      { startBlockIndex: 0, endBlockIndex: 4 },
      undefined,
      undefined,
      // minimal ExtensionContext — the tool does not use it
      {} as never,
    )

    const text = (result.content[0] as { type: 'text'; text: string }).text
    expect(text).toContain('<h1>Quarterly Report</h1>')
    expect(text).toContain('Q3 2026')
    expect(text).toContain('Enterprise: +35%')
    expect(text).toContain('Consumer: +12%')
    expect(result.details).toEqual({ blockCount: 5, truncated: false, offset: text.length })

    dispose()
  }, 30_000)

  it('handles pagination with a large document', async () => {
    // Build a doc with a single huge block to force pagination
    const big = '<p>' + 'A'.repeat(300_000) + '</p>'
    const bigDoc: DocsEditor = {
      getBlockCount: () => 1,
      getBlock: () => ({ index: 0, kind: 'paragraph', html: big }),
      getRangeHtml: () => big,
      clampRange: (s, e) => (s === 0 ? { start: 0, end: Math.min(e, 0) } : null),
      insertBlocks: () => ({ inserted: 0 }),
      replaceBlockRange: () => ({ inserted: 0, removed: 0 }),
      replaceSelection: () => ({ replaced: false }),
      applyOps: () => ({ applied: 0, dryRun: false }),
      markDocSeen: () => {},
    }
    adapter.setEditorInstance(bigDoc)

    const { session, dispose } = await createOfficeSession({
      cwd: process.cwd(),
      extensionFactories: [createDocsSkillExtension({ uiAdapter: adapter, enabledTools: ['read_blocks', 'get_document_context', 'insert_content', 'replace_blocks', 'replace_selection', 'apply_ops', 'create_document', 'read_comments', 'reply_comment', 'resolve_comment'] })],
    })

    const tool = session.getToolDefinition('read_blocks')!

    // First page — should be truncated and tell us where to continue
    const page1 = await tool.execute('e2e-page1', { startBlockIndex: 0, endBlockIndex: 0 }, undefined, undefined, {} as never)
    const page1Details = page1.details as { blockCount: number; truncated: boolean; offset: number }
    expect(page1Details.truncated).toBe(true)
    const firstText = (page1.content[0] as { type: 'text'; text: string }).text
    const m = firstText.match(/offset=(\d+)/)
    expect(m).not.toBeNull()
    const offset = Number(m![1])
    expect(offset).toBeGreaterThan(0)

    // Second page — should reach the end
    const page2 = await tool.execute('e2e-page2', { startBlockIndex: 0, endBlockIndex: 0, offset }, undefined, undefined, {} as never)
    const page2Details = page2.details as { blockCount: number; truncated: boolean; offset: number }
    expect(page2Details.truncated).toBe(false)
    const secondText = (page2.content[0] as { type: 'text'; text: string }).text
    expect(secondText).toContain('(end of range')

    // Concatenate the slices to recover the full content
    const reconstructed = firstText.replace(/\n…\(truncated.*$/, '') + secondText.replace(/\n\(end of range.*$/, '')
    expect(reconstructed.length).toBe(big.length)

    dispose()
  }, 30_000)

  it('handles out-of-range gracefully (does not throw)', async () => {
    const { session, dispose } = await createOfficeSession({
      cwd: process.cwd(),
      extensionFactories: [createDocsSkillExtension({ uiAdapter: adapter, enabledTools: ['read_blocks', 'get_document_context', 'insert_content', 'replace_blocks', 'replace_selection', 'apply_ops', 'create_document', 'read_comments', 'reply_comment', 'resolve_comment'] })],
    })

    const tool = session.getToolDefinition('read_blocks')!
    const result = await tool.execute('e2e-oor', { startBlockIndex: 100, endBlockIndex: 200 }, undefined, undefined, {} as never)
    const text = (result.content[0] as { type: 'text'; text: string }).text
    expect(text).toContain('Invalid range')
    expect(text).toContain('7 block(s)')

    dispose()
  }, 30_000)

  it('returns "no editor" when the adapter has no editor attached', async () => {
    const emptyAdapter = new ReactUIAdapter() // no setEditorInstance call
    const { session, dispose } = await createOfficeSession({
      cwd: process.cwd(),
      extensionFactories: [createDocsSkillExtension({ uiAdapter: emptyAdapter, enableReplaceDocument: false })],
    })

    const tool = session.getToolDefinition('read_blocks')!
    const result = await tool.execute('e2e-noeditor', { startBlockIndex: 0, endBlockIndex: 1 }, undefined, undefined, {} as never)
    const text = (result.content[0] as { type: 'text'; text: string }).text
    expect(text).toContain('No editor available')

    dispose()
  }, 30_000)
})
