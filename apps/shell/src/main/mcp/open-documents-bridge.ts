import type { WebContents } from 'electron'
import type { DocumentTabKind, OpenDocumentTab } from '../../shared/tabs-api'
import type { AgentDocumentIo } from '../app-registry'
import type { OpenDocumentsControl } from './tools/open-documents-tools'
import { closeSavePath, resolveOpenDocumentOfFamily } from './tools/open-documents-tools'
import type { EditorFamily } from './tools/formats'

/**
 * The shell's implementation of the `open_documents` control: the four actions
 * over the tab-manager plus each app's own agent surface (usually built from
 * the AppDescriptor registry, `registryAgentIo`), so an MCP agent reaches the
 * documents the *user* has open (the surfaces are addressed by webContents
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
  /**
   * Per-app agent surfaces keyed by tab kind — read/save/discard per family.
   * The shell passes the registry-built map (`registryAgentIo`); tests inject
   * fakes. Absent entries leave the family unreachable through this tool.
   */
  agents?: Partial<Record<DocumentTabKind, AgentDocumentIo>>
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
  /** the agent surface for one open document's app */
  const agentIo = (kind: OpenDocumentTab['kind']): AgentDocumentIo | undefined =>
    deps.agents?.[kind]

  /** save one document to `filePath` through its family's own writer */
  const saveTo = async (tab: OpenDocumentTab, filePath: string): Promise<void> => {
    const io = agentIo(tab.kind)
    if (io?.saveDenied) throw new Error(io.saveDenied)
    if (!io?.save) throw new Error(`cannot save a ${tab.kind} tab`)
    await io.save(requireContents(tab, deps), filePath)
  }

  const readDocument = async (tab: OpenDocumentTab): Promise<unknown> => {
    const io = agentIo(tab.kind)
    if (!io?.read) throw new Error(`cannot read a ${tab.kind} tab`)
    return io.read(requireContents(tab, deps))
  }

  /** release staged assets a discard leaves behind (markdown/html asset files,
   *  slides recovery copies); a failure here must not veto the close */
  const discardStagedAssets = async (tab: OpenDocumentTab): Promise<void> => {
    const io = agentIo(tab.kind)
    if (!io?.discard) return
    if (io.discardNeedsFilePath && !tab.filePath) return
    try {
      await io.discard(requireContents(tab, deps))
    } catch (error) {
      console.warn(`[mcp] ${tab.kind} discard cleanup incomplete:`, error)
    }
  }

  return {
    list: deps.list,

    read: readDocument,

    save: async (tab) => {
      // the caller's snapshot may predate an edit: settle against the live state
      const live = (await deps.list()).find((doc) => doc.id === tab.id)
      if (!live) throw new Error(`"${tab.title}" is already closed`)
      const savedPath = closeSavePath(live, deps.defaultSaveDir)
      await saveTo(live, savedPath)
      return { savedPath }
    },

    close: async (tab, { unsaved }) => {
      // the caller's snapshot may predate an edit: settle against the live state
      const live = (await deps.list()).find((doc) => doc.id === tab.id)
      if (!live) throw new Error(`"${tab.title}" is already closed`)
      let savedPath: string | undefined
      if (unsaved === 'save' && live.dirty) {
        // A tab of a family this tool cannot save (a PDF's annotations) must
        // not lose its unsaved state silently: refuse with the instruction
        // instead; an explicit discard (or a clean document) just closes.
        const refusal = agentIo(live.kind)?.saveCloseDenied
        if (refusal) throw new Error(refusal(live.title))
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
