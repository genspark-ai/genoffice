import type { Editor } from '@tiptap/react'
import { BLANK_BULLET_NUM_ID, BLANK_ORDERED_NUM_ID } from '@genoffice/docx-engine'
import type { McpCommandMessage, McpEditorCommand } from '../shared/ipc'
import { executeTool, markDocSeen } from './ai/tools'
import { findNumId, type NumIds } from './ai/protocol'
import { save, type FileActionContext } from './file-actions'

/**
 * MCP bridge (renderer half).
 *
 * The shell main process pushes editor commands over `docs:mcp-command`; this
 * module runs them against the *live* Tiptap editor so an external agent drives
 * the same visible editor the built-in agent does, then reports the outcome
 * back on `docs:mcp-result`. Command execution is serialized so a burst cannot
 * interleave two edits into one document.
 *
 * The executors are the built-in agent's own (`executeTool`), so external edits
 * inherit the same parsing, atomicity, formatting rules and stale-index guard.
 */

export interface McpBridgeDeps {
  /** live file-action context (refreshed every render by App) */
  getCtx: () => FileActionContext
}

function numIdsFor(ctx: FileActionContext): NumIds {
  const blocks = ctx.doc?.parsed.blocks ?? []
  const isBlank = ctx.doc?.isBlank === true
  return {
    bullet: findNumId(blocks, 'bullet') ?? (isBlank ? BLANK_BULLET_NUM_ID : null),
    ordered: findNumId(blocks, 'ordered') ?? (isBlank ? BLANK_ORDERED_NUM_ID : null),
  }
}

/** the payload shapes this bridge accepts per command (validated by the executors) */
interface InsertContentInput {
  html: string
  afterBlockIndex?: number
}
interface ReplaceBlocksInput {
  startBlockIndex: number
  endBlockIndex: number
  html: string
}
interface ApplyOpsInput {
  ops: unknown[]
  dryRun?: boolean
}

/** mirror of AiPanel's clearAiHighlights: auto-accept the external edit, keeping undo history */
function clearAiChangedFlags(editor: Editor): void {
  const view = editor.view
  let tr = view.state.tr
  let touched = false
  view.state.doc.forEach((node, offset) => {
    if (node.attrs.aiChanged) {
      tr = tr.setNodeMarkup(offset, undefined, { ...node.attrs, aiChanged: false })
      touched = true
    }
  })
  if (!touched) return
  tr = tr.setMeta('addToHistory', false)
  view.dispatch(tr)
  markDocSeen(editor)
}

async function runCommand(
  deps: McpBridgeDeps,
  command: McpEditorCommand,
  payload: unknown,
): Promise<unknown> {
  const ctx = deps.getCtx()
  const editor = ctx.editor
  if (!editor || !ctx.doc) throw new Error('the document is not ready')

  switch (command) {
    case 'insert_content': {
      const input = (payload ?? {}) as InsertContentInput
      if (typeof input.html !== 'string') throw new Error('insert_content requires "html"')
      const outcome = await executeTool(
        editor,
        { id: 'mcp', name: 'insert_content', input: { ...input } },
        numIdsFor(ctx),
      )
      if (outcome.isError) throw new Error(outcome.output)
      clearAiChangedFlags(editor)
      return { summary: outcome.summary, mutated: outcome.mutated }
    }

    case 'replace_blocks': {
      const input = (payload ?? {}) as ReplaceBlocksInput
      if (typeof input.html !== 'string') throw new Error('replace_blocks requires "html"')
      const outcome = await executeTool(
        editor,
        { id: 'mcp', name: 'replace_blocks', input: { ...input } },
        numIdsFor(ctx),
      )
      if (outcome.isError) throw new Error(outcome.output)
      clearAiChangedFlags(editor)
      return { summary: outcome.summary, mutated: outcome.mutated }
    }

    case 'apply_ops': {
      const input = (payload ?? {}) as ApplyOpsInput
      // Route through executeTool rather than executeOps: the wrapper is what
      // refuses a block-indexed batch after the user edited the document (the
      // stale-index guard) and supplies numbering ids, so an external batch
      // behaves exactly like the built-in agent's.
      const outcome = await executeTool(
        editor,
        {
          id: 'mcp',
          name: 'apply_ops',
          input: { ops: input.ops, ...(input.dryRun === true ? { dryRun: true } : {}) },
        },
        numIdsFor(ctx),
      )
      if (outcome.isError) throw new Error(outcome.output)
      if (input.dryRun !== true) clearAiChangedFlags(editor)
      return { summary: outcome.summary, output: outcome.output, mutated: outcome.mutated }
    }

    case 'read_document': {
      const outcome = await executeTool(
        editor,
        { id: 'mcp', name: 'get_document_context', input: {} },
        numIdsFor(ctx),
      )
      if (outcome.isError) throw new Error(outcome.output)
      return { text: outcome.output }
    }

    case 'save_document': {
      const input = (payload ?? {}) as { path?: string; overwrite?: boolean }
      if (typeof input.path !== 'string' || !input.path) {
        throw new Error('save_document requires an absolute "path"')
      }
      const failures: string[] = []
      const ok = await save(ctx, false, true, undefined, {
        path: input.path,
        overwrite: input.overwrite === true,
        onError: (message) => failures.push(message),
      })
      // the genuine reason matters here: "already exists, pass overwrite:true"
      // is actionable, while a generic failure sends the agent guessing
      if (!ok) throw new Error(failures[0] ?? 'the document could not be saved')
      return { ok: true, path: input.path }
    }

    default: {
      const unreachable: never = command
      throw new Error(`unknown MCP command: ${String(unreachable)}`)
    }
  }
}

/**
 * A fresh tab boots asynchronously (`newFile()` calls setContent, then setDoc),
 * so a command that lands before `doc` exists would be wiped by that blank
 * reset. Wait for the loaded document before announcing readiness.
 */
const READY_POLL_MS = 50
const READY_POLL_LIMIT = 400

async function announceWhenLoaded(
  deps: McpBridgeDeps,
  signal: () => void,
  isCancelled: () => boolean,
): Promise<void> {
  for (let i = 0; i < READY_POLL_LIMIT; i++) {
    if (isCancelled()) return
    const ctx = deps.getCtx()
    if (ctx?.editor && ctx.doc) {
      signal()
      return
    }
    await new Promise((resolve) => setTimeout(resolve, READY_POLL_MS))
  }
}

/** Subscribe the live editor to MCP commands. Returns the unsubscribe function. */
export function installMcpBridge(deps: McpBridgeDeps): () => void {
  const desktop = window.desktop
  if (!desktop?.onMcpCommand || !desktop.reportMcpResult) return () => {}
  let cancelled = false
  let queue: Promise<void> = Promise.resolve()
  const unsubscribe = desktop.onMcpCommand((message: McpCommandMessage) => {
    if (!message || typeof message.requestId !== 'string') return
    queue = queue.then(async () => {
      try {
        const result = await runCommand(deps, message.command, message.payload)
        desktop.reportMcpResult({ requestId: message.requestId, ok: true, result })
      } catch (error) {
        desktop.reportMcpResult({
          requestId: message.requestId,
          ok: false,
          error: error instanceof Error ? error.message : String(error),
        })
      }
    })
  })
  // Let the shell know this tab can accept commands (device for targeted routing).
  void announceWhenLoaded(
    deps,
    () => desktop.signalMcpReady?.(),
    () => cancelled,
  )
  return () => {
    cancelled = true
    unsubscribe()
  }
}
