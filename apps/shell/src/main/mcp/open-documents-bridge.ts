import type { WebContents } from 'electron'
import type { OpenDocumentTab } from '../../shared/tabs-api'
import type { OpenDocumentsControl } from './tools/open-documents-tools'
import { closeSavePath } from './tools/open-documents-tools'
import type { DocsControl } from './tools/document-tools'
import type { SlidesControl } from './tools/slides-tools'
import type { SheetsControl } from './tools/sheets-tools'

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
      let savedPath: string | undefined
      if (unsaved === 'save' && tab.dirty) {
        savedPath = closeSavePath(tab, deps.defaultSaveDir)
        await saveTo(tab, savedPath)
      } else if (unsaved === 'discard') {
        await discardStagedAssets(tab)
      }
      if (!deps.closeTab(tab.id)) throw new Error(`"${tab.title}" is already closed`)
      return savedPath ? { savedPath } : {}
    },
  }
}
