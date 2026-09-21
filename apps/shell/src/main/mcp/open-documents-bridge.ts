import type { WebContents } from 'electron'
import type { OpenDocumentTab } from '../../shared/tabs-api'
import type { OpenDocumentsControl } from './tools/open-documents-tools'
import { closeSavePath, resolveOpenDocumentOfFamily } from './tools/open-documents-tools'
import type { DocsControl } from './tools/document-tools'
import type { SlidesControl } from './tools/slides-tools'
import type { SheetsControl } from './tools/sheets-tools'
import type { EditorFamily } from './tools/formats'

/**
 * The shell's implementation of the `open_documents` control: the three actions
 * over the tab-manager plus each family's own bridge, so an MCP agent reaches
 * the documents the *user* has open (the bridges are addressed by webContents
 * id, so they do not care who opened the tab).
 */

export interface OpenDocumentsBridgeDeps {
  /** the tab-manager's open-document list, dirtiness already resolved */
  list: () => Promise<OpenDocumentTab[]>
  /** the tab's webContents, for the families that must be asked in the renderer */
  webContentsFor: (tabId: string) => WebContents | undefined
  /** remove one tab with no save prompt; false when it is already gone */
  closeTab: (tabId: string) => boolean
  /** folder an untitled dirty document is saved into before closing */
  defaultSaveDir: () => string
  docs?: DocsControl
  sheets?: SheetsControl
  slides?: SlidesControl
  /**
   * slides: drop the session's crash-recovery copies (autosave + untitled draft)
   * on a discard close, so the next open does not offer to restore edits the
   * caller explicitly dropped. Mirrors the interactive "Don't Save" branch.
   */
  slidesDiscard?: (contents: WebContents) => void
  /** markdown: live text + dialog-free save (the app owns the renderer protocol) */
  markdown?: {
    read: (contents: WebContents) => Promise<string>
    save: (contents: WebContents, filePath: string) => Promise<void>
    /** release assets staged next to the document but never written into it */
    discard: (contents: WebContents) => Promise<void>
  }
  /** html: same three, its own renderer protocol */
  html?: {
    read: (contents: WebContents) => Promise<string>
    save: (contents: WebContents, filePath: string) => Promise<void>
    discard: (contents: WebContents) => Promise<void>
  }
}

function requireContents(tab: OpenDocumentTab, deps: OpenDocumentsBridgeDeps): WebContents {
  const contents = deps.webContentsFor(tab.id)
  if (!contents || contents.isDestroyed()) {
    throw new Error(`"${tab.title}" is no longer open`)
  }
  return contents
}

/** deps for the content tools' "target an open document" lookup */
export interface OpenTargetDeps {
  list: () => Promise<OpenDocumentTab[]>
  webContentsFor: (tabId: string) => WebContents | undefined
  /**
   * Make a tab the visible one. Called before a mutating command runs, so the
   * user sees the document the agent is about to change rather than discovering
   * it later. Absent in headless/unit runs.
   */
  activate?: (tabId: string) => void
  /**
   * Bring the window hosting this document forward (un-minimize, show, focus):
   * the shell, or a detached editor window. Separate from `activate` because
   * raising the window is a stronger action than switching a tab inside it.
   */
  revealWindow?: (tabId: string) => void
}

/**
 * Resolver behind the content tools' optional `document` argument: an agent can
 * point an edit at a tab the *user* has open (by tab id or path) instead of the
 * blank one `create_session` opened. Every family's bridge is addressed by
 * webContents id, so a resolved tab needs no session of its own — the caller
 * edits what the user is looking at, and the user decides when to save it.
 *
 * Like `createOpenDocumentsControl`, this lives behind an injected dep so the
 * tool layer never imports Electron.
 */
export function createOpenTargetResolver(
  deps: OpenTargetDeps,
): (target: string, family: EditorFamily, options?: { focus?: boolean }) => Promise<number> {
  return async (target, family, options) => {
    const documents = await deps.list()
    const doc = resolveOpenDocumentOfFamily(documents, target, family)
    const contents = deps.webContentsFor(doc.id)
    if (!contents || contents.isDestroyed()) {
      throw new Error(`"${doc.title}" is no longer open`)
    }
    if (options?.focus && !doc.active) {
      // Showing the tab is a courtesy, not part of the edit: a UI failure here
      // must not fail the command the caller actually asked for.
      try {
        deps.activate?.(doc.id)
        deps.revealWindow?.(doc.id)
      } catch (error) {
        console.warn('[mcp] could not bring the target document into view:', error)
      }
    }
    return contents.id
  }
}

export function createOpenDocumentsControl(deps: OpenDocumentsBridgeDeps): OpenDocumentsControl {
  /** save one document to `filePath` through its family's own writer */
  const saveTo = async (tab: OpenDocumentTab, filePath: string): Promise<void> => {
    switch (tab.kind) {
      case 'docs': {
        if (!deps.docs) throw new Error('the Word editor is unavailable in this build')
        const wcId = requireContents(tab, deps).id
        await deps.docs.runCommand(wcId, 'save_document', { path: filePath, overwrite: true })
        return
      }
      case 'sheets': {
        if (!deps.sheets) throw new Error('the spreadsheet editor is unavailable in this build')
        const wcId = requireContents(tab, deps).id
        await deps.sheets.runCommand(wcId, 'save_sheet', { path: filePath, overwrite: true })
        return
      }
      case 'slides': {
        if (!deps.slides) throw new Error('the presentation editor is unavailable in this build')
        const wcId = requireContents(tab, deps).id
        await deps.slides.saveDeck(wcId, filePath, true)
        return
      }
      case 'markdown': {
        if (!deps.markdown) throw new Error('the Markdown editor is unavailable in this build')
        await deps.markdown.save(requireContents(tab, deps), filePath)
        return
      }
      case 'html': {
        if (!deps.html) throw new Error('the HTML editor is unavailable in this build')
        await deps.html.save(requireContents(tab, deps), filePath)
        return
      }
      case 'pdf':
        throw new Error('a PDF cannot be saved through this tool: the pdf app is a viewer')
      default:
        throw new Error(`cannot save a ${tab.kind} tab`)
    }
  }

  const readDocument = async (tab: OpenDocumentTab): Promise<unknown> => {
    switch (tab.kind) {
      case 'docs': {
        if (!deps.docs) throw new Error('the Word editor is unavailable in this build')
        return deps.docs.runCommand(requireContents(tab, deps).id, 'read_document', {})
      }
      case 'sheets': {
        if (!deps.sheets) throw new Error('the spreadsheet editor is unavailable in this build')
        return deps.sheets.runCommand(requireContents(tab, deps).id, 'read_sheet', {})
      }
      case 'slides': {
        if (!deps.slides) throw new Error('the presentation editor is unavailable in this build')
        return deps.slides.readDeck(requireContents(tab, deps).id)
      }
      case 'markdown':
        if (!deps.markdown) throw new Error('the Markdown editor is unavailable in this build')
        return deps.markdown.read(requireContents(tab, deps))
      case 'html':
        if (!deps.html) throw new Error('the HTML editor is unavailable in this build')
        return deps.html.read(requireContents(tab, deps))
      default:
        throw new Error(`cannot read a ${tab.kind} tab`)
    }
  }

  /** release staged assets a discard leaves behind (markdown and html own files) */
  const discardStagedAssets = async (tab: OpenDocumentTab): Promise<void> => {
    if (tab.kind === 'markdown' && deps.markdown && tab.filePath) {
      await deps.markdown.discard(requireContents(tab, deps)).catch((error) => {
        console.warn('[mcp] markdown discard cleanup incomplete:', error)
      })
    } else if (tab.kind === 'html' && deps.html && tab.filePath) {
      await deps.html.discard(requireContents(tab, deps)).catch((error) => {
        console.warn('[mcp] html discard cleanup incomplete:', error)
      })
    } else if (tab.kind === 'slides' && deps.slidesDiscard) {
      // slides keeps its own crash-recovery copies; dropping the edits must drop
      // them too, or the next open offers them back
      deps.slidesDiscard(requireContents(tab, deps))
    }
  }

  return {
    list: deps.list,

    read: readDocument,

    close: async (tab, { unsaved }) => {
      if (tab.kind === 'pdf') {
        throw new Error(
          'a PDF tab cannot be closed through this tool: the pdf app is a viewer, so it has no ' +
            'unsaved state to settle',
        )
      }
      // the caller's snapshot may predate an edit: settle against the live state
      const live = (await deps.list()).find((doc) => doc.id === tab.id)
      if (!live) throw new Error(`"${tab.title}" is already closed`)
      let savedPath: string | undefined
      if (unsaved === 'save' && live.dirty) {
        savedPath = closeSavePath(live, deps.defaultSaveDir)
        await saveTo(live, savedPath)
      } else if (unsaved === 'discard') {
        await discardStagedAssets(live)
      }
      if (!deps.closeTab(live.id)) {
        throw new Error(
          `"${live.title}" could not be closed: it is already closed or waiting on a prompt`,
        )
      }
      return savedPath ? { savedPath } : {}
    },
  }
}
