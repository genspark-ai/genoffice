import './editor-styles'
import { createRoot, type Root } from 'react-dom/client'
import { htmlDir, htmlLang, type Lang } from '@genoffice/i18n'
import { applyAiPanelPrefs, installScreenTips } from '@genoffice/ui/editor'
import { setAltChunkHtmlConverter } from '@genoffice/docx-engine'
import { LocaleProvider, setModuleLang } from './i18n/locale'
import { setAuthorName } from './author-name'
import { getDocsHost, installDocsHost, type DocsHostAdapter } from './host/docs-host'
import type { DocsRuntime } from './runtime/docs-runtime'
import type { DocsCommands } from './runtime/command-router'
import type { DocsInspection } from './sdk/inspection'
import type { DocsAnnotations } from './sdk/annotations'
import type { DocsDocumentCommands } from '../shared/document-host'
import type { DocsEditorUi } from './sdk/file-menu'
import type { DocTheme, UiTheme } from '../shared/ipc'

export { createExternalDocsHost } from './host/docs-host'
export { createQtDocsHost } from './host/qt'
export type { DocsQtChannel, QtSignal } from './host/qt'
export type { DocsHostAdapter, DocsEditorServices } from './host/docs-host'
export type {
  DocsEditorAction,
  DocsHarnessAdapter,
  DocsAiOperationId,
  DocsAiOperation,
  DocsAiInstructionRequest,
} from '../shared/editor-action'
export type { DocsResourceTools } from '../shared/resource-tools'
export { createDocsResourceTools } from './runtime/resource-tools'
export type { DocsResourceOptions } from './runtime/resource-tools'
export type { DocsRuntime, DocsRuntimeClient, DocsRuntimeScope } from './runtime/docs-runtime'
export type { DocsEngineTarget, DocsQueryItem, DocsQueryResult } from '../shared/document-query'
export type {
  DocsDraftInspection,
  DocsDraftSettlement,
  DocsDraftState,
} from '../shared/document-draft'
export type { DocsCommands } from './runtime/command-router'
export type {
  DocsDocumentHost,
  DocsHostedIdentity,
  DocsHostedDocument,
  DocsHostedSaveRequest,
  DocsHostedSaveOutcome,
  DocsDocumentOperation,
} from '../shared/document-host'
export type { DocsDocumentCommands, DocsDocumentState } from '../shared/document-host'
export type {
  DocsEditorUi,
  DocsFileMenu,
  DocsFileMenuItem,
  DocsFileMenuCommand,
  DocsFileMenuConfiguration,
  DocsFileMenuAction,
} from './sdk/file-menu'
export { DOCS_RIBBON_BUILTIN_TABS } from './sdk/ribbon'
export type {
  DocsRibbon,
  DocsRibbonAssistantButton,
  DocsRibbonAction,
  DocsRibbonBuiltinTab,
  DocsRibbonConfiguration,
  DocsRibbonControl,
  DocsRibbonGroup,
  DocsRibbonGroupExtension,
  DocsRibbonMenuItem,
  DocsRibbonTab,
} from './sdk/ribbon'
export type { DocsImageSource } from './ai/image-source'
export type {
  DocsInspection,
  DocsFindResult,
  DocsTextReference,
  DocsTextSnapshot,
} from './sdk/inspection'
export type {
  DocsAnnotations,
  DocsAnnotation,
  DocsDiagnostic,
  DocsRangeStyle,
  DocsMatchRule,
  DocsChecker,
} from './sdk/annotations'
export type {
  DocsOperationResult,
  DocsSaveResult,
  DocsSaveOutcome,
  DocsErrorCode,
} from '../shared/operation-result'
export { docsToolCatalog } from '../shared/tool-catalog'
export { DOCS_RUNTIME_TOOLS } from '../shared/runtime-tools'

export interface DocsEditorHandle {
  readonly ready: Promise<void>
  readonly runtime: DocsRuntime
  readonly commands: DocsCommands
  readonly inspection: DocsInspection
  readonly annotations: DocsAnnotations
  readonly documents: DocsDocumentCommands
  readonly ui: DocsEditorUi
  dispose(): void
}

export interface DocsEditorMountOptions {
  host: DocsHostAdapter
  language?: Lang
  signal?: AbortSignal
}

/** Mount the complete native Word editor using injected services, without an Electron global. */
export async function mountDocsEditor(
  container: HTMLElement,
  options: DocsEditorMountOptions,
): Promise<DocsEditorHandle> {
  options.signal?.throwIfAborted()
  const releaseHost = installDocsHost(options.host)
  const disposers: Array<() => void> = []
  let root: Root | undefined
  let live: Pick<
    DocsEditorHandle,
    'runtime' | 'commands' | 'inspection' | 'annotations' | 'documents' | 'ui'
  > | null = null
  let disposed = false
  let loaded = false
  let delivered = false
  const activeDocuments = () => {
    if (!live || disposed) throw new Error('The document lifecycle API is disposed or not ready')
    return live.documents
  }
  const documents: DocsDocumentCommands = {
    new: async () => activeDocuments().new(),
    open: async (id) => activeDocuments().open(id),
    save: async (saveAs) => activeDocuments().save(saveAs),
    state: () => activeDocuments().state(),
    subscribe: (listener) => activeDocuments().subscribe(listener),
  }
  const lifetime = new AbortController()
  let resolveReady!: () => void
  let rejectReady!: (reason: unknown) => void
  const ready = new Promise<void>((resolve, reject) => {
    resolveReady = resolve
    rejectReady = reject
  })
  const dispose = () => {
    if (disposed) return
    disposed = true
    options.signal?.removeEventListener('abort', dispose)
    try {
      root?.unmount()
      for (const off of disposers.reverse()) off()
      setAltChunkHtmlConverter(null)
    } finally {
      live = null
      releaseHost()
      lifetime.abort()
      if (delivered && !loaded)
        rejectReady(new Error('The editor was disposed before its document opened'))
    }
  }
  options.signal?.addEventListener('abort', dispose, { once: true })
  const startup = <T,>(operation: Promise<T>): Promise<T> =>
    new Promise((resolve, reject) => {
      const canceled = () => reject(new Error('The editor mount was disposed'))
      if (lifetime.signal.aborted) {
        canceled()
        return
      }
      lifetime.signal.addEventListener('abort', canceled, { once: true })
      void operation
        .then(resolve, reject)
        .finally(() => lifetime.signal.removeEventListener('abort', canceled))
    })
  try {
    const services = getDocsHost()
    const [lang, theme, docTheme] = await startup(
      Promise.all([
        options.language
          ? Promise.resolve(options.language)
          : services.getLanguage().catch(() => 'zh' as const),
        services.getTheme().catch(() => 'system' as const),
        services.getDocumentTheme().catch(() => 'follow' as const),
      ]),
    )
    options.signal?.throwIfAborted()
    if (disposed) throw new Error('The editor mount was disposed')
    setModuleLang(lang)
    document.documentElement.lang = htmlLang(lang)
    document.documentElement.dir = htmlDir(lang)
    const applyTheme = (next: UiTheme) => {
      if (next === 'system') document.documentElement.removeAttribute('data-theme')
      else document.documentElement.setAttribute('data-theme', next)
    }
    const applyDocumentTheme = (next: DocTheme) => {
      if (next === 'follow') document.documentElement.removeAttribute('data-doc-theme')
      else document.documentElement.setAttribute('data-doc-theme', next)
    }
    applyTheme(theme)
    applyDocumentTheme(docTheme)
    disposers.push(
      installScreenTips(),
      services.onThemeChanged(applyTheme),
      services.onDocumentThemeChanged(applyDocumentTheme),
    )
    if (services.convertAltChunkHtml)
      setAltChunkHtmlConverter((html) => services.convertAltChunkHtml(html))
    const panelPrefs = await startup(services.getAiPanelPrefs().catch(() => undefined))
    options.signal?.throwIfAborted()
    if (disposed) throw new Error('The editor mount was disposed')
    if (panelPrefs) applyAiPanelPrefs(panelPrefs)
    disposers.push(services.onAiPanelPrefsChanged(applyAiPanelPrefs))
    const author = await startup(services.getAuthorName().catch(() => undefined))
    options.signal?.throwIfAborted()
    if (disposed) throw new Error('The editor mount was disposed')
    if (author !== undefined) setAuthorName(author)
    disposers.push(services.onAuthorNameChanged(setAuthorName))
    const { App } = await startup(import('./App'))
    options.signal?.throwIfAborted()
    root = createRoot(container, { onUncaughtError: (error) => rejectReady(error) })
    root.render(
      <LocaleProvider initial={lang}>
        <App
          embedded
          onRuntime={(value) => {
            live = value
            if (value && !disposed && !loaded) {
              loaded = true
              resolveReady()
            }
          }}
        />
      </LocaleProvider>,
    )
    delivered = true
    return {
      ready,
      get runtime() {
        if (!live || disposed) throw new Error('The document runtime is not ready')
        return live.runtime
      },
      get commands() {
        if (!live || disposed) throw new Error('The document command endpoint is not ready')
        return live.commands
      },
      get inspection() {
        if (!live || disposed) throw new Error('The document inspection API is not ready')
        return live.inspection
      },
      get annotations() {
        if (!live || disposed) throw new Error('The document annotations API is not ready')
        return live.annotations
      },
      get documents() {
        activeDocuments()
        return documents
      },
      get ui() {
        if (!live || disposed) throw new Error('The editor UI is disposed or not ready')
        return live.ui
      },
      dispose,
    }
  } catch (error) {
    // The mount Promise carries startup errors; no consumer has a ready handle yet.
    loaded = true
    dispose()
    throw error
  }
}
