import { resolve } from 'node:path'
import { z } from 'zod'
import type { OpenDocumentTab, TabKind } from '../../../shared/tabs-api'
import type { McpToolDefinition } from '../mcp-server'
import { familyLabel, generateExtension, type EditorFamily, type SessionFamily } from './formats'
import { sanitizeFileBase, uniquePathIn } from './document-tools'

/**
 * MCP access to the documents the *user* has open — the tab-manager's list of
 * editor tabs, not the single blank session `create_session` owns. The three
 * actions share one tool because they are steps of one workflow: discover,
 * then act on what was discovered.
 *
 * The control object is injected so this module never imports Electron: the
 * shell implements it over TabManager and the per-family bridges, tests pass
 * a fake.
 */

/** tab kind -> format family (the registry's vocabulary, for the type label) */
const FAMILY_BY_KIND: Record<Exclude<TabKind, 'home'>, EditorFamily> = {
  docs: 'docx',
  sheets: 'xlsx',
  slides: 'pptx',
  markdown: 'md',
  html: 'html',
  pdf: 'pdf',
}

/** families whose live content can be read back over MCP */
const READABLE_KINDS: ReadonlySet<TabKind> = new Set<TabKind>([
  'docs',
  'sheets',
  'slides',
  'markdown',
  'html',
])

export interface OpenDocumentsControl {
  /** every editor tab the user can see (Home and chrome-free Present tabs excluded) */
  list: () => Promise<OpenDocumentTab[]>
  /** live content of one open document, unsaved edits included */
  read: (tab: OpenDocumentTab) => Promise<unknown>
  /**
   * Settle the document's unsaved changes and close its tab. `unsaved: 'save'`
   * writes first (to the document's own path, or into the default save folder
   * when it has never been saved); `'discard'` drops the edits.
   */
  close: (
    tab: OpenDocumentTab,
    options: { unsaved: 'save' | 'discard' },
  ) => Promise<{ savedPath?: string }>
}

export interface OpenDocumentsDeps {
  control?: OpenDocumentsControl
  /** folder an untitled dirty document is saved into before closing */
  defaultSaveDir: () => string
}

/** the extension documents of one family carry on disk */
function extensionFor(family: EditorFamily): string {
  if (family === 'md') return '.md'
  if (family === 'pdf') return '.pdf'
  return `.${generateExtension(family as SessionFamily)}`
}

/**
 * Match a caller's target against the open documents: an exact tab id wins,
 * then a path (compared after resolving; case-insensitively on Windows).
 */
export function resolveOpenDocument(
  documents: readonly OpenDocumentTab[],
  target: string,
): OpenDocumentTab | undefined {
  const byId = documents.find((doc) => doc.id === target)
  if (byId) return byId
  const wanted = normalizePath(target)
  return documents.find(
    (doc) => doc.filePath !== undefined && normalizePath(doc.filePath) === wanted,
  )
}

function normalizePath(value: string): string {
  const resolved = resolve(value)
  return process.platform === 'win32' ? resolved.toLowerCase() : resolved
}

/** a guided "no such document" error listing what is actually open */
function noMatch(target: string, documents: readonly OpenDocumentTab[]): Error {
  if (documents.length === 0) {
    return new Error(`no document is open in GenOffice, so "${target}" cannot be resolved`)
  }
  const listing = documents
    .map((doc) => `  ${doc.id}  ${doc.filePath ?? '(never saved)'}  — ${typeLabel(doc)}`)
    .join('\n')
  return new Error(`no open document matches "${target}". Currently open:\n${listing}`)
}

function typeLabel(doc: OpenDocumentTab): string {
  return familyLabel(FAMILY_BY_KIND[doc.kind])
}

/** one row of the `list` action */
function describeDocument(doc: OpenDocumentTab): Record<string, unknown> {
  return {
    id: doc.id,
    type: typeLabel(doc),
    kind: doc.kind,
    title: doc.title,
    path: doc.filePath ?? null,
    saved: doc.filePath !== undefined,
    dirty: doc.dirty,
    active: doc.active,
  }
}

/**
 * Where a dialog-free save should land: the document's own file when it has
 * one, else a unique name derived from its title inside the default folder.
 */
export function closeSavePath(doc: OpenDocumentTab, defaultSaveDir: () => string): string {
  if (doc.filePath) return doc.filePath
  const name = `${sanitizeFileBase(doc.title)}${extensionFor(FAMILY_BY_KIND[doc.kind])}`
  return uniquePathIn(defaultSaveDir(), name)
}

export function createOpenDocumentTools(deps: OpenDocumentsDeps): McpToolDefinition[] {
  const control = deps.control

  return [
    {
      name: 'open_documents',
      description:
        'Work with the documents the user currently has open in GenOffice tabs — not just the ' +
        'one a session opened. Actions: "list" (no target) returns every open document with its ' +
        'id, type, path, title, whether it has ever been saved, and whether it has unsaved ' +
        'changes; "read" returns one document\'s live content including unsaved edits; "close" ' +
        'closes one, saving it first by default (pass unsaved:"discard" to drop the changes). ' +
        'Identify a document by its path, or by the id from "list" when it has never been saved. ' +
        'PDF tabs are listed but cannot be read or closed here: the pdf app is a viewer, so an ' +
        'open PDF has no readable or savable state through this tool.',
      inputSchema: {
        action: z.enum(['list', 'read', 'close']).describe('what to do'),
        target: z
          .string()
          .optional()
          .describe('path or tab id of the document; required for "read" and "close"'),
        unsaved: z
          .enum(['save', 'discard'])
          .optional()
          .describe('"close" only: save before closing (default), or discard the changes'),
      },
      handler: async (args) => {
        if (!control) {
          throw new Error('GenOffice is not running, so its open documents are unreachable')
        }
        const action = String(args.action ?? '')
        if (action === 'list') {
          const documents = await control.list()
          return { documents: documents.map(describeDocument) }
        }

        const target = typeof args.target === 'string' ? args.target.trim() : ''
        if (!target) throw new Error(`"${action}" needs a target (a path, or an id from "list")`)

        const documents = await control.list()
        const doc = resolveOpenDocument(documents, target)
        if (!doc) throw noMatch(target, documents)

        if (action === 'read') {
          if (!READABLE_KINDS.has(doc.kind)) {
            throw new Error(
              `"${doc.title}" is a PDF: the pdf app is a viewer, so an open PDF has no readable ` +
                'content through this tool',
            )
          }
          return {
            id: doc.id,
            type: typeLabel(doc),
            kind: doc.kind,
            title: doc.title,
            path: doc.filePath ?? null,
            dirty: doc.dirty,
            content: await control.read(doc),
          }
        }

        if (action === 'close') {
          const unsaved = args.unsaved === 'discard' ? 'discard' : 'save'
          const { savedPath } = await control.close(doc, { unsaved })
          return {
            closed: true,
            id: doc.id,
            type: typeLabel(doc),
            title: doc.title,
            path: doc.filePath ?? null,
            ...(savedPath ? { savedTo: savedPath } : {}),
            ...(unsaved === 'discard' && doc.dirty ? { discardedUnsavedChanges: true } : {}),
          }
        }

        throw new Error(`unknown action "${action}" — use "list", "read" or "close"`)
      },
    },
  ]
}
