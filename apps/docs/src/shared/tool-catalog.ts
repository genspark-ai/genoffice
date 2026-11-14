// Modified for ForaOffice: reliable document targeting and portable inspection.
import type { AgentToolDef } from '@genoffice/agent-core'
import { MIND_MAP_TOOL } from '@genoffice/agent-core/mind-map'
import { PAGE_SETUP_TOOL, SECTION_BREAK_TOOL } from './page-tool-schema'
import { DEFINE_STYLE_TOOL, LIST_STYLES_TOOL } from './style-tool-schema'
import {
  SET_WATERMARK_TOOL,
  INSERT_PICTURE_TOOL,
  INSERT_TEXT_BOX_TOOL,
} from './floating-tool-schema'

/** Editing remains in the runtime; inference and orchestration use the Harness's own facilities. */

import { DOCS_IMAGE_SOURCE_SCHEMA } from './media-policy'

export const DOCS_READ_TOOLS = new Set([
  'query_document',
  'read_attachment',
  'search_knowledge_base',
  'read_knowledge_file',
  'read_document_media',
  'get_editor_capabilities',
  'get_document_context',
  'read_blocks',
  'read_revisions',
  'read_comments',
  'read_notes',
  'list_styles',
  'generate_mindmap',
  'inspect_document_draft',
])

export function docsToolCatalog(operationNames: readonly string[] = []): AgentToolDef[] {
  return [
    {
      name: 'get_document_context',
      description:
        'Get the latest state of the current document: block list (index|type|content preview), full-text stats (word/character counts) and the current selection. Refreshes the full caller baseline, including document-side state. Block indexes change after modifications; call this when you need up-to-date indexes or a stale-target error requires full context.',
      inputSchema: { type: 'object', properties: {}, required: [] },
    },
    {
      name: 'read_blocks',
      description:
        'Read the full content of a block range (restricted HTML). Previews in the block list are truncated; you must read the full original text with this tool before rewriting. ' +
        'Long ranges are paged: a truncated result says which offset to continue from; concatenate the slices in order to get the full HTML. ' +
        'Read-only data-doc-readonly elements describe protected objects or field results; they are not writable HTML or object references. This read refreshes the body baseline only, not document-side state; use read_document or get_document_context when full context is required.',
      inputSchema: {
        type: 'object',
        properties: {
          startBlockIndex: {
            type: 'integer',
            description: 'start block index (0-based, inclusive)',
          },
          endBlockIndex: { type: 'integer', description: 'end block index (inclusive)' },
          offset: {
            type: 'integer',
            description:
              'character offset to continue a truncated read (default 0); use the offset given in the previous truncation notice',
          },
        },
        required: ['startBlockIndex', 'endBlockIndex'],
      },
    },
    {
      name: 'insert_content',
      description:
        'Insert new content at a given position (restricted HTML, may contain multiple blocks). For writing/continuing/generating new content; to rewrite existing content use replace_blocks. Image tags must use exact current media:D:R:B.I references from read_blocks; every invalid or unsupported image rejects the whole write. New images use insert_picture.',
      inputSchema: {
        type: 'object',
        properties: {
          html: {
            type: 'string',
            description:
              'restricted HTML fragment to insert; <span data-tab></span> is a tab character (literal tabs collapse to spaces)',
          },
          afterBlockIndex: {
            type: 'integer',
            description:
              'insert after this block index; -1 = start of document; omitted = after the block containing the cursor',
          },
        },
        required: ['html'],
      },
    },
    {
      name: 'replace_blocks',
      description:
        'Replace a block range with new content (restricted HTML). For rewriting/translating/condensing/expanding existing content; the new block count may differ from the old. Compatible paragraph/text formatting (font, size, color, indent, spacing, alignment) is inherited by block role; this is not a lossless DOCX round trip and does not guarantee section boundaries, fields or protected objects. A rewritten <table> keeps column widths, borders, shading and cell formatting (unchanged cells keep their content). Read-only data-doc-readonly elements are refused before editing; exclude those ranges or edit surrounding text in place. Intentional replacement or explicit deleteBlocks may remove structures. Preserve writable inline pictures with the exact current media:D:R:B.I HTML references from read_blocks. Invalid, stale, unsupported or legacy unbound image references reject the whole write.',
      inputSchema: {
        type: 'object',
        properties: {
          deleteStructures: {
            type: 'array',
            items: { type: 'string' },
            description:
              'Exact revision-bound structure IDs from query_document that this replacement intentionally removes.',
          },
          startBlockIndex: {
            type: 'integer',
            description: 'start block index (0-based, inclusive)',
          },
          endBlockIndex: { type: 'integer', description: 'end block index (inclusive)' },
          html: {
            type: 'string',
            description:
              'replacement restricted HTML fragment; <span data-tab></span> is a tab character',
          },
        },
        required: ['startBlockIndex', 'endBlockIndex', 'html'],
      },
    },
    {
      name: 'replace_selection',
      description:
        "Replace exactly the user's selected text (the <sel>…</sel> span in the context) with new inline content, leaving the rest of the block untouched. For rewording/translating/correcting a selected phrase or sentence inside a paragraph, including a table cell. The new text inherits the selection's formatting unless the fragment styles it. Requires a range selection inside one paragraph/heading/list item; ranges across paragraphs or cells are refused. Pending deletion revisions are preserved and cannot be rewritten. Image tags must copy exact current media:D:R:B.I references; invalid or unsupported images reject the whole write. For whole blocks use replace_blocks.",
      inputSchema: {
        type: 'object',
        properties: {
          deleteStructures: {
            type: 'array',
            items: { type: 'string' },
            description:
              'Exact revision-bound structure IDs from query_document that this replacement intentionally removes.',
          },
          html: {
            type: 'string',
            description:
              'replacement inline content: plain text or restricted inline HTML (strong em u s a br formula)',
          },
        },
        required: ['html'],
      },
    },
    {
      name: 'apply_ops',
      description:
        (operationNames.length
          ? `Run a list of formatting/structure ops as one atomic transaction: ${operationNames.join(', ')}. `
          : 'Run formatting/structure ops as one atomic transaction. Call get_editor_capabilities for operation names and signatures. ') +
        'Each op is a flat { op, target?, ...fields } object; fields are patches (present = set, null = clear, absent = untouched). Any invalid op rejects the whole batch with its usage line. findReplace defaults to exactly one literal match across formatting runs within one paragraph; zero/multiple matches reject atomically. Intentional bulk replacement requires match:"all" and an explicit target scope.',
      inputSchema: {
        type: 'object',
        properties: {
          ops: {
            type: 'array',
            description:
              'ops executed in order, e.g. [{"op":"setFont","target":{"nodeType":"docHeading"},"color":"#FF0000"}]',
            items: { type: 'object' },
          },
          dryRun: {
            type: 'boolean',
            description:
              'Check parameters, current target types/ranges and schema executability by building the whole transaction, then discard it. No document, history or revision changes; later edits can invalidate this preview.',
          },
        },
        required: ['ops'],
      },
    },
    {
      name: 'read_revisions',
      description:
        'List every pending tracked revision (insertions, deletions, formatting/move/table changes) with its id, type, author, date, block index, the affected text and, for formatting changes, what changed. Ids are positional (r1 = first in document order) and renumber after every edit, so read again before accept_changes / reject_changes.',
      inputSchema: { type: 'object', properties: {}, required: [] },
    },
    {
      name: 'accept_changes',
      description:
        'Accept pending tracked changes: insertions become plain text, deleted text disappears, new formatting stays. Select with all: true, ids from read_revisions, or any combination of author / type / blockIndex / blockRange / before. Only when the user asks to accept changes.',
      inputSchema: {
        type: 'object',
        properties: {
          all: { type: 'boolean', description: 'true = every pending change' },
          ids: {
            type: 'array',
            items: { type: 'string' },
            description:
              'revision ids from read_revisions (r1, r2, …); positional, so re-read after any edit',
          },
          author: { type: 'string', description: 'only changes by this author (case-insensitive)' },
          type: {
            type: 'string',
            enum: ['insertion', 'deletion', 'formatting', 'move'],
            description: 'only changes of this type',
          },
          blockIndex: { type: 'integer', description: 'only changes inside this block' },
          blockRange: {
            type: 'array',
            items: { type: 'integer' },
            minItems: 2,
            maxItems: 2,
            description: '[from, to] block indexes, inclusive',
          },
          before: { type: 'string', description: 'ISO date; only changes recorded before it' },
        },
        required: [],
      },
    },
    {
      name: 'reject_changes',
      description:
        'Reject pending tracked changes: inserted text disappears, deleted text is restored, formatting goes back to what it was. Same selector as accept_changes. Only when the user asks to reject changes.',
      inputSchema: {
        type: 'object',
        properties: {
          all: { type: 'boolean', description: 'true = every pending change' },
          ids: {
            type: 'array',
            items: { type: 'string' },
            description:
              'revision ids from read_revisions (r1, r2, …); positional, so re-read after any edit',
          },
          author: { type: 'string', description: 'only changes by this author (case-insensitive)' },
          type: {
            type: 'string',
            enum: ['insertion', 'deletion', 'formatting', 'move'],
            description: 'only changes of this type',
          },
          blockIndex: { type: 'integer', description: 'only changes inside this block' },
          blockRange: {
            type: 'array',
            items: { type: 'integer' },
            minItems: 2,
            maxItems: 2,
            description: '[from, to] block indexes, inclusive',
          },
          before: { type: 'string', description: 'ISO date; only changes recorded before it' },
        },
        required: [],
      },
    },
    {
      name: 'read_comments',
      description:
        'List all comment threads (including resolved ones) with ids, authors, anchored block indexes and anchor text. Unresolved threads already ride along in the message context; use this for the full picture.',
      inputSchema: { type: 'object', properties: {}, required: [] },
    },
    {
      name: 'reply_comment',
      description:
        'Add a reply to a comment thread: after completing a requested change (summarize what changed), or to answer/ask back when the comment is a question or is ambiguous.',
      inputSchema: {
        type: 'object',
        properties: {
          parentId: { type: 'string', description: 'id of the comment thread to reply to' },
          text: { type: 'string', description: 'reply text' },
        },
        required: ['parentId', 'text'],
      },
    },
    {
      name: 'resolve_comment',
      description:
        'Mark a comment thread as resolved. Only after the requested change was applied (reply first), or when the user explicitly asked to resolve.',
      inputSchema: {
        type: 'object',
        properties: {
          id: { type: 'string', description: 'id of the comment thread to resolve' },
        },
        required: ['id'],
      },
    },
    {
      name: 'insert_footnote',
      description:
        'Add a footnote: a superscript reference mark goes into the block (right after afterText, else at its end) and the note text prints at the bottom of that page. Use for sources, asides and clarifications the user asks to footnote.',
      inputSchema: {
        type: 'object',
        properties: {
          blockIndex: { type: 'integer', description: 'block that gets the reference mark' },
          afterText: {
            type: 'string',
            description: 'exact text in that block the mark follows; omitted = end of the block',
          },
          text: { type: 'string', description: 'the note text (\\n separates paragraphs)' },
        },
        required: ['blockIndex', 'text'],
      },
    },
    {
      name: 'insert_endnote',
      description:
        'Add an endnote (same as insert_footnote, but the text collects at the end of the document). Use when the document already uses endnotes or the user asks for them.',
      inputSchema: {
        type: 'object',
        properties: {
          blockIndex: { type: 'integer', description: 'block that gets the reference mark' },
          afterText: {
            type: 'string',
            description: 'exact text in that block the mark follows; omitted = end of the block',
          },
          text: { type: 'string', description: 'the note text (\\n separates paragraphs)' },
        },
        required: ['blockIndex', 'text'],
      },
    },
    {
      name: 'delete_note',
      description:
        'Remove a footnote or endnote together with its reference mark; ids from read_notes.',
      inputSchema: {
        type: 'object',
        properties: {
          kind: { type: 'string', enum: ['footnote', 'endnote'] },
          id: { type: 'string', description: 'note id from read_notes' },
        },
        required: ['kind', 'id'],
      },
    },
    {
      name: 'edit_note',
      description:
        'Change the text of an existing footnote or endnote in place (findReplace inside the note): the note keeps its id, reference mark and formatting. Use for a typo or a citation fix instead of delete_note + insert_footnote.',
      inputSchema: {
        type: 'object',
        properties: {
          kind: {
            type: 'string',
            enum: ['footnote', 'endnote'],
            description: 'omit when the id is unambiguous across footnotes and endnotes',
          },
          id: { type: 'string', description: 'note id from read_notes' },
          find: { type: 'string', description: 'exact text inside the note to replace' },
          replace: { type: 'string', description: 'replacement text ("" deletes the match)' },
          matchCase: { type: 'boolean', description: 'defaults to true' },
        },
        required: ['id', 'find', 'replace'],
      },
    },
    {
      name: 'read_notes',
      description:
        'List every footnote and endnote with its id, number, the block holding its reference mark and its text.',
      inputSchema: { type: 'object', properties: {}, required: [] },
    },
    {
      name: 'add_comment',
      description:
        'Start a new comment thread on a block or on an exact text span inside it, without changing the document: review notes, questions, suggestions the user should decide on.',
      inputSchema: {
        type: 'object',
        properties: {
          blockIndex: { type: 'integer', description: 'block to annotate' },
          text: {
            type: 'string',
            description:
              'exact text inside the block to anchor the comment to; omitted = the whole block',
          },
          occurrence: {
            type: 'integer',
            description: 'which match to anchor when text occurs more than once (1 = first)',
          },
          comment: { type: 'string', description: 'comment body; \\n starts a new paragraph' },
          author: { type: 'string', description: 'author name; default: the AI author' },
          initials: { type: 'string', description: 'author initials shown in the margin' },
        },
        required: ['blockIndex', 'comment'],
      },
    },
    {
      name: 'delete_comment',
      description:
        'Delete a comment. A thread root with replies is refused unless withReplies is true; a reply id deletes just that reply.',
      inputSchema: {
        type: 'object',
        properties: {
          id: { type: 'string', description: 'id of the comment to delete' },
          withReplies: {
            type: 'boolean',
            description: 'also delete the replies of a thread root (default false)',
          },
        },
        required: ['id'],
      },
    },
    {
      name: 'insert_image',
      description:
        'Insert an image at the captured cursor/selection. Accepts an http(s)/data image URL or an authorized local path/host asset source. Sources are read by the host, then embedded using the normal insertion pipeline.',
      inputSchema: {
        type: 'object',
        properties: {
          url: {
            type: 'string',
            description: 'direct http(s) image link or data:image/png|jpeg|gif;base64,...',
          },
          maxWidthPx: { type: 'integer', description: 'maximum width (px), default 480' },
          source: DOCS_IMAGE_SOURCE_SCHEMA,
          label: {
            type: 'string',
            maxLength: 500,
            description: 'optional picture label, preserved in the editor document model',
          },
        },
        oneOf: [
          { type: 'object', properties: { url: { type: 'string' } }, required: ['url'] },
          {
            type: 'object',
            properties: { source: DOCS_IMAGE_SOURCE_SCHEMA },
            required: ['source'],
          },
        ],
      },
    },
    {
      name: 'insert_chart',
      description:
        'Insert a chart (saved as a native Word chart). Data must be real: from the document content or web_search results — do not make up numbers.',
      inputSchema: {
        type: 'object',
        properties: {
          kind: { type: 'string', enum: ['bar', 'line', 'pie'], description: 'chart type' },
          title: { type: 'string', description: 'chart title' },
          categories: {
            type: 'array',
            items: { type: 'string' },
            description: 'category (x axis / sector) labels',
          },
          series: {
            type: 'array',
            description:
              'data series; values has the same length as categories, use null for missing data. Pie charts use only the first series',
            items: {
              type: 'object',
              properties: {
                name: { type: 'string' },
                values: { type: 'array', items: { type: ['number', 'null'] } },
              },
              required: ['values'],
            },
          },
          afterBlockIndex: {
            type: 'integer',
            description:
              'insert after this block index; -1 = start of document; omitted = after the block containing the cursor',
          },
        },
        required: ['kind', 'categories', 'series'],
      },
    },
    {
      name: 'edit_chart',
      description:
        'Edit the data of an existing chart in the document (title/category labels/series names/values). Chart blocks in the block list can be edited; ' +
        'the category count and the number of values per series must match the original chart (data points cannot be added or removed).',
      inputSchema: {
        type: 'object',
        properties: {
          blockIndex: { type: 'integer', description: 'block index of the chart' },
          title: { type: 'string', description: 'new title (omit to keep)' },
          categories: {
            type: 'array',
            items: { type: ['string', 'null'] },
            description:
              'new category labels, same length as the original; pass null for positions to keep',
          },
          series: {
            type: 'array',
            description: 'series to change',
            items: {
              type: 'object',
              properties: {
                index: { type: 'integer', description: 'series index (0-based)' },
                name: { type: 'string', description: 'new series name (omit to keep)' },
                values: {
                  type: 'array',
                  items: { type: ['number', 'null'] },
                  description:
                    'new values, same length as the original series; pass null for positions to keep',
                },
              },
              required: ['index'],
            },
          },
        },
        required: ['blockIndex'],
      },
    },
    {
      name: 'set_header_footer',
      description:
        'Set the page header or footer text (the current contents are listed in the message context). Plain text; \\n separates lines; the tokens {PAGE} and {NUMPAGES} become live page-number fields; an empty string clears the text. ' +
        'Per-line alignment/styling and images are preserved. Supported textboxes are patched in their original containers, including compatibility twins; unsupported textbox structures are refused before editing. textboxIndex selects one zero-based textbox from the current strip context and expires when the part changes; omitted replaces all editable text lines. section is zero-based, default 0 for BOTH kinds. scope=section (default) isolates the target and preserves other sections, including inherited/shared parts; scope=shared edits the owning part and affects the sharedWith sections listed in context. view first enables titlePg for that section; even enables document-wide even/odd display.',
      inputSchema: {
        type: 'object',
        properties: {
          kind: { type: 'string', enum: ['header', 'footer'] },
          engineTarget: {
            type: 'object',
            description: 'Revision-bound header/footer/textbox target from query_document.',
          },
          textboxIndex: {
            type: 'integer',
            minimum: 0,
            description:
              'One textbox from fresh section/variant context; omitted edits all supported text lines.',
          },
          section: {
            type: 'integer',
            minimum: 0,
            description: 'Zero-based section index from context; default 0.',
          },
          scope: {
            type: 'string',
            enum: ['section', 'shared'],
            description:
              'Default section preserves non-target sections; shared edits the shared/inherited owning part.',
          },
          text: {
            type: 'string',
            description: 'new text; \\n between lines; may contain {PAGE} / {NUMPAGES}; "" clears',
          },
          view: {
            type: 'string',
            enum: ['default', 'first', 'even'],
            description: 'which variant to write (default when omitted)',
          },
        },
        required: ['text'],
        anyOf: [{ required: ['kind'] }, { required: ['engineTarget'] }],
      },
    },
    PAGE_SETUP_TOOL,
    SECTION_BREAK_TOOL,
    DEFINE_STYLE_TOOL,
    LIST_STYLES_TOOL,
    SET_WATERMARK_TOOL,
    INSERT_TEXT_BOX_TOOL,
    INSERT_PICTURE_TOOL,
    {
      name: 'create_document',
      description:
        'Create a NEW standalone file in the default save folder; the current document is not modified. Word files open in a new editor view and complete their first save before success is returned. Other formats open only when the host provides a viewer; the result reports whether they opened. Use when the user asks to put content into a new/separate document instead of this one. ' +
        "type 'docx' (default) and 'pdf' take the same restricted HTML as insert_content in content; type 'md' takes Markdown source; type 'html' takes a complete standalone HTML page. Images and charts are not supported in the new file's initial content.",
      inputSchema: {
        type: 'object',
        properties: {
          type: {
            type: 'string',
            enum: ['docx', 'pdf', 'md', 'html'],
            description: "target file type (default 'docx')",
          },
          title: { type: 'string', description: 'document title, used as the file name' },
          content: {
            type: 'string',
            description:
              'full document content: restricted HTML for docx/pdf, Markdown for md, a complete HTML page for html',
          },
        },
        required: ['title', 'content'],
      },
    },
    {
      ...MIND_MAP_TOOL,
      description:
        'Validate a complete mind-map tree using the native document limits and block citations. ' +
        'The Harness receives display metadata and is responsible for displaying it. The editor does not render a native chat card in external mode. ' +
        'Supply root with label, optional block index and children. This tool does not modify the document.',
    },
  ]
}
