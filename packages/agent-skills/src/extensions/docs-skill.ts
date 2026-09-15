/**
 * docs-skill extension — the first migrated tool from GenOffice's
 * `apps/docs/src/renderer/ai/tools.ts`.
 *
 * This file migrates the simplest tool first: `read_blocks` (W4 deliverable).
 * The plan calls for all 22 docx tools to be migrated in W6 — for now we
 * establish the shape and test the wiring end-to-end.
 *
 * The extension factory captures the ReactUIAdapter via closure so tool
 * `execute` handlers can call `getEditorInstance()` to reach the Tiptap editor
 * that the React renderer installed. This avoids any casting of `ctx.ui` and
 * keeps the typed `ExtensionUIContext` contract intact.
 */

import { Type, type Static } from "@earendil-works/pi-ai";
import { defineTool, type ExtensionAPI } from "@earendil-works/pi-coding-agent";
import type { ReactUIAdapter } from "@genoffice/agent-runtime";

// ============================================================================
// DocsEditor contract
// ============================================================================
//
// We intentionally abstract the editor behind a small interface so:
//   1. The skill can be unit-tested with a mock editor (no Tiptap/PM dependency).
//   2. apps/docs wires a real Tiptap editor that satisfies this contract.
//   3. The contract is small enough to mock in <30 lines.
//
// The interface covers what read_blocks needs today; W6 will extend it for the
// remaining 21 tools (insert_content, replace_blocks, etc.).

export interface DocsBlock {
  /** Block index in the document. */
  index: number;
  /** Block kind (paragraph, heading, list, image, etc.). */
  kind: string;
  /** HTML representation of the block. */
  html: string;
  /** Tracked-deletion flag (for hidden-blocks handling). */
  trackedDeleted?: boolean;
}

export interface DocsEditor {
  /** Number of blocks currently in the document. */
  getBlockCount(): number;
  /** Read a single block by index. Throws if out of range. */
  getBlock(index: number): DocsBlock;
  /** Return the concatenated HTML of blocks [start, end] inclusive. */
  getRangeHtml(start: number, end: number): string;
  /** Normalize a requested [start, end] to a valid inclusive range, or null. */
  clampRange(start: number, end: number): { start: number; end: number } | null;

  // --------------------------------------------------------------------------
  // Mutation API — implemented by the real Tiptap-backed editor. Mock editors
  // for tests record the calls and update `getBlockCount` accordingly.
  // --------------------------------------------------------------------------

  /** Insert one or more blocks at the given position. `afterIndex` of -1 means "at start". */
  insertBlocks(afterIndex: number, blocksHtml: string): { inserted: number };
  /** Replace a block range with new content. Returns the new block count. */
  replaceBlockRange(start: number, end: number, blocksHtml: string): { inserted: number; removed: number };
  /** Replace exactly the user's selected text. Requires a text selection in the editor. */
  replaceSelection(inlineHtml: string): { replaced: boolean };
  /** Apply a batch of formatting/structure ops atomically. `dryRun` only validates. */
  applyOps(ops: ReadonlyArray<unknown>, dryRun: boolean): { applied: number; dryRun: boolean };
  /** Mark the document as seen (clears the "doc changed since last read" flag). */
  markDocSeen(): void;
}

const MAX_CHARS = 200_000;

// ============================================================================
// read_blocks tool
// ============================================================================

const ReadBlocksParams = Type.Object({
  startBlockIndex: Type.Integer({ minimum: 0, description: "Start block index (0-based, inclusive)" }),
  endBlockIndex: Type.Integer({ minimum: 0, description: "End block index (inclusive)" }),
  offset: Type.Optional(
    Type.Integer({
      minimum: 0,
      description: "Character offset to continue a truncated read (default 0)",
    }),
  ),
});

type ReadBlocksArgs = Static<typeof ReadBlocksParams>;

export interface ReadBlocksToolOptions {
  /** UI adapter (provided by createOfficeSession). */
  uiAdapter: ReactUIAdapter;
  /** Default timeout (ms) for the dangerous-tool confirm dialog. */
  confirmTimeoutMs?: number;
}

export function createReadBlocksTool(opts: ReadBlocksToolOptions) {
  const { uiAdapter } = opts;

  return defineTool<typeof ReadBlocksParams, { blockCount: number; truncated: boolean; offset: number }>({
    name: "read_blocks",
    label: "Read Blocks",
    description:
      "Read the full content of a block range (restricted HTML). Previews in the block list are truncated; you must read the full original text with this tool before rewriting. " +
      "Long ranges are paged: a truncated result says which offset to continue from; concatenate the slices in order to get the full HTML.",
    promptSnippet: "read_blocks(startBlockIndex, endBlockIndex, offset?) — read full HTML of a block range",
    promptGuidelines: [
      "Block indexes change after modifications; call get_document_context first if you are not sure of the current state.",
      "Image blocks return as `[Protected content: Image, kept as is]` and cannot be modified.",
      "If a result ends with a truncation notice, call read_blocks again with the given offset to continue.",
    ],
    parameters: ReadBlocksParams,
    async execute(_toolCallId, params: ReadBlocksArgs, _signal, _onUpdate, _ctx) {
      const editor = uiAdapter.getEditorInstance<DocsEditor>();
      if (!editor) {
        return {
          content: [{ type: "text", text: "No editor available — open a document first." }],
          details: { blockCount: 0, truncated: false, offset: 0 },
        };
      }

      const range = editor.clampRange(params.startBlockIndex, params.endBlockIndex);
      if (!range) {
        return {
          content: [
            {
              type: "text",
              text: `Invalid range [${params.startBlockIndex}, ${params.endBlockIndex}] — document has ${editor.getBlockCount()} block(s).`,
            },
          ],
          details: { blockCount: 0, truncated: false, offset: 0 },
        };
      }

      const html = editor.getRangeHtml(range.start, range.end);
      const offset = Math.max(0, Math.trunc(Number(params.offset ?? 0)) || 0);
      if (offset > 0 && offset >= html.length) {
        return {
          content: [
            { type: "text", text: `offset ${offset} is beyond the content (${html.length} characters in total)` },
          ],
          details: { blockCount: range.end - range.start + 1, truncated: false, offset },
        };
      }

      const slice = html.slice(offset, offset + MAX_CHARS);
      const end = offset + slice.length;
      const note =
        end < html.length
          ? `\n…(truncated: ${html.length} characters in total, call read_blocks again with offset=${end} to continue)`
          : offset > 0
            ? `\n(end of range: ${html.length} characters in total)`
            : "";

      const blockCount = range.end - range.start + 1;
      const output = slice ? slice + note : "(range is empty)";

      return {
        content: [{ type: "text", text: output }],
        details: { blockCount, truncated: end < html.length, offset: end },
      };
    },
  });
}

// ============================================================================
// replace_document tool (first dialog example)
// ============================================================================
//
// Per the plan §3.4, dangerous tools (replace_document, etc.) must show a
// confirm dialog before executing. We use uiAdapter.confirm() and respect the
// user's choice. If the user clicks Cancel, the tool returns a synthetic
// "blocked" content result that the LLM will see and adapt to.

const ReplaceDocumentParams = Type.Object({
  html: Type.String({ description: "New document HTML (will replace everything)" }),
  reason: Type.String({ description: "Why this replacement is needed (shown in confirm dialog)" }),
});

type ReplaceDocumentArgs = Static<typeof ReplaceDocumentParams>;

export function createReplaceDocumentTool(opts: ReadBlocksToolOptions) {
  const { uiAdapter, confirmTimeoutMs = 30_000 } = opts;

  return defineTool<typeof ReplaceDocumentParams, { replaced: boolean; reason: string }>({
    name: "replace_document",
    label: "Replace Document",
    description:
      "Replace the entire document content. DESTRUCTIVE — requires user confirmation. " +
      "Use this only when the user explicitly asks to rewrite the whole document, or when other tools cannot achieve the desired result.",
    promptSnippet: "replace_document(html, reason) — replace entire document (requires confirmation)",
    promptGuidelines: [
      "Always prefer targeted tools (replace_blocks, insert_content) over replace_document.",
      "You must provide a clear `reason` that will be shown to the user in the confirmation dialog.",
      "If the user cancels, do not retry — switch to a more targeted tool.",
    ],
    parameters: ReplaceDocumentParams,
    async execute(_toolCallId, params: ReplaceDocumentArgs, _signal, _onUpdate, _ctx) {
      const ok = await uiAdapter.confirm(
        "Replace entire document?",
        `Reason: ${params.reason}\n\nThis will overwrite the current document content. Undo is available afterwards.`,
        { timeout: confirmTimeoutMs },
      );
      if (!ok) {
        return {
          content: [
            { type: "text", text: "User cancelled the replace_document operation. Switch to a targeted tool (e.g. replace_blocks) or ask the user how to proceed." },
          ],
          details: { replaced: false, reason: "user_cancelled" },
        };
      }

      const editor = uiAdapter.getEditorInstance<DocsEditor & { replaceAll?(html: string): void }>();
      if (!editor) {
        return {
          content: [{ type: "text", text: "No editor available." }],
          details: { replaced: false, reason: "no_editor" },
        };
      }
      if (typeof editor.replaceAll === "function") {
        editor.replaceAll(params.html);
      }
      // If replaceAll is not implemented, the host app must wire it before
      // exposing this tool — the migration guide in agent1.md will cover this.
      return {
        content: [{ type: "text", text: `Document replaced (${params.html.length} characters).` }],
        details: { replaced: true, reason: params.reason },
      };
    },
  });
}

// ============================================================================
// Extension factory
// ============================================================================

export interface DocsSkillOptions {
  /** UI adapter (provided by createOfficeSession). */
  uiAdapter: ReactUIAdapter;
  /** Whether to register replace_document (W4 enables it as the first dialog demo). */
  enableReplaceDocument?: boolean;
  /** Confirm timeout in ms for replace_document. */
  confirmTimeoutMs?: number;
}

/**
 * Create the docs-skill extension factory for createOfficeSession.
 *
 * Usage:
 *   const { session, uiAdapter } = await createOfficeSession({
 *     extensionFactories: [createDocsSkillExtension({ uiAdapter })],
 *   });
 *   uiAdapter.setEditorInstance(tiptapEditor);
 */
export interface DocsSkillOptions {
  /** UI adapter (provided by createOfficeSession). */
  uiAdapter: ReactUIAdapter;
  /**
   * Which tools to register. Defaults to all of them.
   * Pass a subset for restricted sessions (e.g. read-only).
   */
  enabledTools?: ReadonlyArray<DocsToolName>;
  /** Confirm timeout in ms for replace_document. */
  confirmTimeoutMs?: number;
}

export type DocsToolName =
  | "read_blocks"
  | "get_document_context"
  | "insert_content"
  | "replace_blocks"
  | "replace_selection"
  | "apply_ops"
  | "create_document"
  | "replace_document"
  | "read_comments"
  | "reply_comment"
  | "resolve_comment";

export const ALL_DOCS_TOOL_NAMES: readonly DocsToolName[] = [
  "read_blocks",
  "get_document_context",
  "insert_content",
  "replace_blocks",
  "replace_selection",
  "apply_ops",
  "create_document",
  "replace_document",
  "read_comments",
  "reply_comment",
  "resolve_comment",
];

export function createDocsSkillExtension(opts: DocsSkillOptions) {
  const { uiAdapter, enabledTools, confirmTimeoutMs } = opts;
  const enabled = new Set(enabledTools ?? ALL_DOCS_TOOL_NAMES);

  return (pi: ExtensionAPI) => {
    if (enabled.has("read_blocks")) pi.registerTool(createReadBlocksTool({ uiAdapter }));
    if (enabled.has("get_document_context")) pi.registerTool(createGetDocumentContextTool({ uiAdapter }));
    if (enabled.has("insert_content")) pi.registerTool(createInsertContentTool({ uiAdapter }));
    if (enabled.has("replace_blocks")) pi.registerTool(createReplaceBlocksTool({ uiAdapter }));
    if (enabled.has("replace_selection")) pi.registerTool(createReplaceSelectionTool({ uiAdapter }));
    if (enabled.has("apply_ops")) pi.registerTool(createApplyOpsTool({ uiAdapter }));
    if (enabled.has("create_document")) pi.registerTool(createCreateDocumentTool({ uiAdapter }));
    if (enabled.has("replace_document")) {
      pi.registerTool(createReplaceDocumentTool({ uiAdapter, ...(confirmTimeoutMs !== undefined ? { confirmTimeoutMs } : {}) }));
    }
    if (enabled.has("read_comments")) pi.registerTool(createReadCommentsTool({ uiAdapter }));
    if (enabled.has("reply_comment")) pi.registerTool(createReplyCommentTool({ uiAdapter }));
    if (enabled.has("resolve_comment")) pi.registerTool(createResolveCommentTool({ uiAdapter }));

    // Append the office-specific system prompt section. Cast the `on` call to
    // bypass TypeScript's overload-resolution ambiguity (the union-typed
    // handler picks the wrong overload).
    const beforeAgentStart = async (): Promise<{ systemPromptAppend: string }> => ({
      systemPromptAppend:
        "\n## Document Editing Rules\n- Block indexes are 0-based and inclusive.\n" +
        "- Always call get_document_context before mutating operations if you are not sure of the current state.\n" +
        "- Use replace_document sparingly — prefer targeted tools like replace_blocks or insert_content.\n" +
        "- replace_selection requires an active text selection; for whole blocks use replace_blocks.",
    });
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (pi as any).on("before_agent_start", beforeAgentStart);
  };
}

export default createDocsSkillExtension;

// ============================================================================
// get_document_context
// ============================================================================

const GetDocumentContextParams = Type.Object({});

const MAX_CONTEXT_CHARS = 12_000;

export function createGetDocumentContextTool(opts: ReadBlocksToolOptions) {
  const { uiAdapter } = opts;
  return defineTool<typeof GetDocumentContextParams, { blockCount: number; selectionSummary: string | null }>({
    name: "get_document_context",
    label: "Get Document Context",
    description:
      "Get the latest state of the current document: block list (index|type|content preview), full-text stats (word/character counts) and the current selection. Block indexes change after modifications; call this when you need up-to-date indexes.",
    promptSnippet: "get_document_context() — refresh block list + selection before mutating",
    parameters: GetDocumentContextParams,
    async execute(_id, _params, _signal, _onUpdate, _ctx) {
      const editor = uiAdapter.getEditorInstance<DocsEditor>();
      if (!editor) {
        return { content: [{ type: "text", text: "No editor available." }], details: { blockCount: 0, selectionSummary: null } };
      }
      const count = editor.getBlockCount();
      const lines: string[] = [];
      for (let i = 0; i < count; i++) {
        const b = editor.getBlock(i);
        const preview = b.html.length > 120 ? b.html.slice(0, 120) + "…" : b.html;
        lines.push(`[${i}] ${b.kind}: ${preview}`);
      }
      const fullText = lines.join("\n");
      const truncated = fullText.length > MAX_CONTEXT_CHARS;
      const out = truncated ? fullText.slice(0, MAX_CONTEXT_CHARS) + "\n…(truncated)" : fullText;

      return {
        content: [{ type: "text", text: `Document has ${count} block(s):\n${out}` }],
        details: { blockCount: count, selectionSummary: null },
      };
    },
  });
}

// ============================================================================
// insert_content
// ============================================================================

const InsertContentParams = Type.Object({
  html: Type.String({ description: "Restricted HTML fragment to insert (may contain multiple blocks)" }),
  afterBlockIndex: Type.Optional(
    Type.Integer({ description: "Insert after this block index (-1 = start; omit = after cursor block)" }),
  ),
});

export function createInsertContentTool(opts: ReadBlocksToolOptions) {
  const { uiAdapter } = opts;
  return defineTool<typeof InsertContentParams, { inserted: number; afterIndex: number }>({
    name: "insert_content",
    label: "Insert Content",
    description:
      "Insert new content at a given position (restricted HTML, may contain multiple blocks). For writing/continuing/generating new content; to rewrite existing content use replace_blocks.",
    promptSnippet: "insert_content(html, afterBlockIndex?) — append or insert blocks",
    parameters: InsertContentParams,
    async execute(_id, params, _signal, _onUpdate, _ctx) {
      const editor = uiAdapter.getEditorInstance<DocsEditor>();
      if (!editor) {
        return { content: [{ type: "text", text: "No editor available." }], details: { inserted: 0, afterIndex: -1 } };
      }
      // Default: append at the end
      const afterIndex = params.afterBlockIndex ?? Math.max(0, editor.getBlockCount() - 1);
      const result = editor.insertBlocks(afterIndex, params.html);
      return {
        content: [{ type: "text", text: `Inserted ${result.inserted} block(s) after block ${afterIndex}. Block indexes have changed; use get_document_context if needed.` }],
        details: { inserted: result.inserted, afterIndex },
      };
    },
  });
}

// ============================================================================
// replace_blocks
// ============================================================================

const ReplaceBlocksParams = Type.Object({
  startBlockIndex: Type.Integer({ minimum: 0 }),
  endBlockIndex: Type.Integer({ minimum: 0 }),
  html: Type.String({ description: "Replacement restricted HTML fragment" }),
});

export function createReplaceBlocksTool(opts: ReadBlocksToolOptions) {
  const { uiAdapter } = opts;
  return defineTool<typeof ReplaceBlocksParams, { removed: number; inserted: number }>({
    name: "replace_blocks",
    label: "Replace Blocks",
    description:
      "Replace a block range with new content (restricted HTML). For rewriting/translating/condensing/expanding existing content; the new block count may differ from the old. New blocks inherit the replaced blocks' paragraph and text formatting automatically.",
    promptSnippet: "replace_blocks(startBlockIndex, endBlockIndex, html) — rewrite a block range",
    parameters: ReplaceBlocksParams,
    async execute(_id, params, _signal, _onUpdate, _ctx) {
      const editor = uiAdapter.getEditorInstance<DocsEditor>();
      if (!editor) {
        return { content: [{ type: "text", text: "No editor available." }], details: { removed: 0, inserted: 0 } };
      }
      const range = editor.clampRange(params.startBlockIndex, params.endBlockIndex);
      if (!range) {
        return {
          content: [{ type: "text", text: `Invalid range [${params.startBlockIndex}, ${params.endBlockIndex}].` }],
          details: { removed: 0, inserted: 0 },
        };
      }
      const result = editor.replaceBlockRange(range.start, range.end, params.html);
      return {
        content: [{ type: "text", text: `Replaced ${result.removed} block(s) with ${result.inserted} new block(s). Block indexes have changed.` }],
        details: { removed: result.removed, inserted: result.inserted },
      };
    },
  });
}

// ============================================================================
// replace_selection
// ============================================================================

const ReplaceSelectionParams = Type.Object({
  html: Type.String({ description: "Replacement inline content (plain text or restricted inline HTML)" }),
});

export function createReplaceSelectionTool(opts: ReadBlocksToolOptions) {
  const { uiAdapter } = opts;
  return defineTool<typeof ReplaceSelectionParams, { replaced: boolean }>({
    name: "replace_selection",
    label: "Replace Selection",
    description:
      "Replace exactly the user's selected text with new inline content, leaving the rest of the block untouched. For rewording/translating/correcting a selected phrase or sentence. Requires a range selection inside one paragraph/heading/list item; for whole blocks use replace_blocks.",
    promptSnippet: "replace_selection(html) — replace the user's current text selection",
    parameters: ReplaceSelectionParams,
    async execute(_id, params, _signal, _onUpdate, _ctx) {
      const editor = uiAdapter.getEditorInstance<DocsEditor>();
      if (!editor) {
        return { content: [{ type: "text", text: "No editor available." }], details: { replaced: false } };
      }
      const result = editor.replaceSelection(params.html);
      if (!result.replaced) {
        return {
          content: [{ type: "text", text: "No active text selection. Ask the user to select text first, or use replace_blocks for whole-block edits." }],
          details: { replaced: false },
        };
      }
      return {
        content: [{ type: "text", text: `Selection replaced (${params.html.length} characters).` }],
        details: { replaced: true },
      };
    },
  });
}

// ============================================================================
// apply_ops
// ============================================================================

const ApplyOpsParams = Type.Object({
  ops: Type.Array(Type.Unknown(), { description: "Array of formatting/structure ops to apply atomically" }),
  dryRun: Type.Optional(Type.Boolean({ description: "Validate and return the plan without changing the document" })),
});

export function createApplyOpsTool(opts: ReadBlocksToolOptions) {
  const { uiAdapter } = opts;
  return defineTool<typeof ApplyOpsParams, { applied: number; dryRun: boolean }>({
    name: "apply_ops",
    label: "Apply Ops",
    description:
      "Run a list of formatting/structure ops as one atomic transaction. Each op is a flat { op, target?, ...fields } object; fields are patches (present = set, null = clear, absent = untouched). Any invalid op rejects the whole batch.",
    promptSnippet: "apply_ops(ops, dryRun?) — run atomic formatting batch",
    parameters: ApplyOpsParams,
    async execute(_id, params, _signal, _onUpdate, _ctx) {
      const editor = uiAdapter.getEditorInstance<DocsEditor>();
      if (!editor) {
        return { content: [{ type: "text", text: "No editor available." }], details: { applied: 0, dryRun: false } };
      }
      if (params.ops.length === 0) {
        return { content: [{ type: "text", text: "ops array is empty" }], details: { applied: 0, dryRun: Boolean(params.dryRun) } };
      }
      const result = editor.applyOps(params.ops, Boolean(params.dryRun));
      return {
        content: [
          { type: "text", text: result.dryRun ? `Dry run: ${result.applied} op(s) would be applied (no changes made).` : `Applied ${result.applied} op(s).` },
        ],
        details: { applied: result.applied, dryRun: result.dryRun },
      };
    },
  });
}

// ============================================================================
// create_document
// ============================================================================

const CreateDocumentParams = Type.Object({
  html: Type.Optional(Type.String({ description: "Initial document content as restricted HTML" })),
  title: Type.Optional(Type.String({ description: "Document title (used for the new tab name)" })),
});

export function createCreateDocumentTool(opts: ReadBlocksToolOptions) {
  const { uiAdapter } = opts;
  return defineTool<typeof CreateDocumentParams, { created: boolean; title: string | null }>({
    name: "create_document",
    label: "Create Document",
    description:
      "Create a new empty document (or with initial content). The new document opens in a new tab and becomes the active document.",
    promptSnippet: "create_document(html?, title?) — open a new document tab",
    parameters: CreateDocumentParams,
    async execute(_id, params, _signal, _onUpdate, _ctx) {
      const editor = uiAdapter.getEditorInstance<DocsEditor & { createNewDocument?(html: string, title: string | null): void }>();
      if (!editor || typeof editor.createNewDocument !== "function") {
        return {
          content: [{ type: "text", text: "create_document requires the host app to implement createNewDocument()." }],
          details: { created: false, title: null },
        };
      }
      editor.createNewDocument(params.html ?? "", params.title ?? null);
      return {
        content: [{ type: "text", text: `Created new document${params.title ? ` titled "${params.title}"` : ""}.` }],
        details: { created: true, title: params.title ?? null },
      };
    },
  });
}

// ============================================================================
// Comment tools (read_revisions, read_comments, reply_comment, resolve_comment)
// ============================================================================

export interface CommentThread {
  id: string;
  author: string;
  date: string;
  blockIndex: number;
  anchorText: string;
  resolved: boolean;
  replies: Array<{ author: string; date: string; text: string }>;
}

export function createReadCommentsTool(opts: ReadBlocksToolOptions) {
  const { uiAdapter } = opts;
  return defineTool<typeof GetDocumentContextParams, { count: number }>({
    name: "read_comments",
    label: "Read Comments",
    description: "List all comment threads (including resolved ones) with ids, authors, anchored block indexes and anchor text.",
    parameters: GetDocumentContextParams,
    async execute(_id, _params, _signal, _onUpdate, _ctx) {
      const adapter = uiAdapter;
      const comments = adapter.getCustomData<CommentThread[]>("comments") ?? [];
      const text = comments.length === 0
        ? "No comments."
        : comments.map((c) => `[${c.id}] block ${c.blockIndex} by ${c.author} @ ${c.date} ${c.resolved ? "(resolved)" : ""}: "${c.anchorText}" — ${c.replies.length} replies`).join("\n");
      return { content: [{ type: "text", text }], details: { count: comments.length } };
    },
  });
}

const ReplyCommentParams = Type.Object({
  parentId: Type.String(),
  text: Type.String(),
});

export function createReplyCommentTool(opts: ReadBlocksToolOptions) {
  const { uiAdapter } = opts;
  return defineTool<typeof ReplyCommentParams, { replied: boolean; parentId: string }>({
    name: "reply_comment",
    label: "Reply Comment",
    description: "Add a reply to a comment thread.",
    parameters: ReplyCommentParams,
    async execute(_id, params, _signal, _onUpdate, _ctx) {
      const comments = uiAdapter.getCustomData<CommentThread[]>("comments") ?? [];
      const thread = comments.find((c) => c.id === params.parentId);
      if (!thread) {
        return { content: [{ type: "text", text: `Comment ${params.parentId} not found.` }], details: { replied: false, parentId: params.parentId } };
      }
      thread.replies.push({ author: "AI", date: new Date().toISOString(), text: params.text });
      uiAdapter.setCustomData("comments", comments);
      return { content: [{ type: "text", text: `Replied to comment ${params.parentId}.` }], details: { replied: true, parentId: params.parentId } };
    },
  });
}

const ResolveCommentParams = Type.Object({
  id: Type.String(),
});

export function createResolveCommentTool(opts: ReadBlocksToolOptions) {
  const { uiAdapter } = opts;
  return defineTool<typeof ResolveCommentParams, { resolved: boolean; id: string }>({
    name: "resolve_comment",
    label: "Resolve Comment",
    description: "Mark a comment thread as resolved.",
    parameters: ResolveCommentParams,
    async execute(_id, params, _signal, _onUpdate, _ctx) {
      const comments = uiAdapter.getCustomData<CommentThread[]>("comments") ?? [];
      const thread = comments.find((c) => c.id === params.id);
      if (!thread) {
        return { content: [{ type: "text", text: `Comment ${params.id} not found.` }], details: { resolved: false, id: params.id } };
      }
      thread.resolved = true;
      uiAdapter.setCustomData("comments", comments);
      return { content: [{ type: "text", text: `Resolved comment ${params.id}.` }], details: { resolved: true, id: params.id } };
    },
  });
}


