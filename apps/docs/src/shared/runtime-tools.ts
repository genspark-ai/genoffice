// Modified for ForaOffice: reliable document targeting and portable inspection.
import type { AgentToolDef } from '@genoffice/agent-core'

/** Runtime lifecycle controls, separate from the pinned native AI tool catalog. */
export const DOCS_RUNTIME_TOOLS: AgentToolDef[] = [
  {
    name: 'query_document',
    description:
      'Read a bounded current-content window with revision-bound mutation targets. Partial reads refresh only the body baseline. Continue using nextCursor; stale continuation requires a fresh query.',
    inputSchema: {
      type: 'object',
      additionalProperties: false,
      properties: {
        startBlockIndex: { type: 'integer', minimum: 0 },
        endBlockIndex: { type: 'integer', minimum: 0 },
        type: { type: 'string' },
        containsText: { type: 'string', maxLength: 2000 },
        domain: {
          type: 'string',
          enum: ['body', 'cells', 'images', 'sections', 'headers', 'footers'],
        },
        limit: { type: 'integer', minimum: 1, maximum: 50 },
        maxItemChars: { type: 'integer', minimum: 1, maximum: 4000 },
        cursor: { type: 'string', maxLength: 6000 },
      },
    },
  },
  {
    name: 'read_document_media',
    description:
      'Read an authorized standalone, inline or table-cell image as standard MCP image content. Prefer the revision-bound mediaRef from fresh context/read_blocks; alternatively use blockIndex and imageIndex (required for multiple images). Missing resources, stale references and unsupported targets are distinguished. Internal asset addresses are not exposed.',
    inputSchema: {
      type: 'object',
      properties: {
        mediaRef: {
          type: 'string',
          description: 'Revision-bound reference returned by context/read_blocks.',
        },
        blockIndex: { type: 'integer', minimum: 0 },
        imageIndex: {
          type: 'integer',
          minimum: 0,
          description: 'Zero-based image ordinal within the block.',
        },
      },
      anyOf: [{ required: ['mediaRef'] }, { required: ['blockIndex'] }],
    },
  },
  {
    name: 'get_editor_capabilities',
    description:
      'Discover current document tool availability, restricted HTML rules, and apply_ops operation signatures. In external mode all inference workflows execute in the Harness.',
    inputSchema: { type: 'object', properties: {} },
  },
  {
    name: 'document_history',
    description:
      'Undo or redo one live-document edit. Refused if this caller has not read the latest document.',
    inputSchema: {
      type: 'object',
      properties: { action: { type: 'string', enum: ['undo', 'redo'] } },
      required: ['action'],
    },
  },
  {
    name: 'start_document_draft',
    description:
      'Start a visible, unaccepted draft and capture its runtime-owned referenceSnapshot. Read full context and references first. replaceDocument selects the landing position; structural deletion still requires exact deleteStructures on each cumulative update. The external Harness supplies HTML; the editor performs no inference.',
    inputSchema: {
      type: 'object',
      properties: {
        afterBlockIndex: {
          type: 'integer',
          description: 'Insert after this block; -1 means the beginning.',
        },
        replaceDocument: {
          type: 'boolean',
          description: 'Replace the complete document after explicit finalization.',
        },
      },
    },
  },
  {
    name: 'update_document_draft',
    description:
      'Update a draft with cumulative restricted HTML and an increasing sequence number. Replays must have identical content and deletion intent. Complete image references and structural effects are checked before preview; image/structure references bind to the draft referenceSnapshot while its authored source remains unchanged. Incomplete image tokens buffer without publishing; inspect validatedSequence/buffering. New images use insert_picture.',
    inputSchema: {
      type: 'object',
      properties: {
        draftId: { type: 'string' },
        sequence: { type: 'integer', minimum: 1 },
        html: { type: 'string', maxLength: 200_000 },
        deleteStructures: {
          type: 'array',
          items: { type: 'string' },
          description:
            'Exact removed structure references from the runtime-owned draft referenceSnapshot. Required for whole-document structural deletion; preserved images/compatible tables are excluded.',
        },
      },
      required: ['draftId', 'sequence', 'html'],
    },
  },
  {
    name: 'inspect_document_draft',
    description: 'Read the state and update sequence of a draft owned by this caller.',
    inputSchema: {
      type: 'object',
      properties: { draftId: { type: 'string' } },
      required: ['draftId'],
    },
  },
  {
    name: 'stop_document_draft',
    description:
      'Stop accepting updates and wait for the Harness to finalize, retain partial content, or discard. The draft remains unaccepted.',
    inputSchema: {
      type: 'object',
      properties: { draftId: { type: 'string' } },
      required: ['draftId'],
    },
  },
  {
    name: 'finish_document_draft',
    description:
      'Finalize, retain an identified validated prefix, or discard a draft. Reference/deletion drafts refuse after captured body or authored-side state changes; discard and read fresh context to restart. Incomplete HTML cannot finalize; retain_partial reports omittedTail and validatedSequence when retaining an earlier validated snapshot. Committed content uses normal tracking and undo.',
    inputSchema: {
      type: 'object',
      properties: {
        draftId: { type: 'string' },
        decision: { type: 'string', enum: ['finalize', 'retain_partial', 'discard'] },
      },
      required: ['draftId', 'decision'],
    },
  },
]

export const DOCS_DRAFT_CONTROLS = new Set([
  'update_document_draft',
  'inspect_document_draft',
  'stop_document_draft',
  'finish_document_draft',
])
