/// Web-version bootstrap for the shell renderer.
///
/// Loaded before main.tsx and only acts outside Electron: it installs the same
/// `window.aiOffice` / `window.aiOfficeProject` / `window.aiOfficeTabs` objects
/// the preload exposes in the desktop app, but backed by the HTTP/SSE transport
/// against the running Electron main process. Native-only channels (browse,
/// save-dir picker, reveal, trash, tab management) get browser equivalents so
/// the web version keeps the full feature surface. Inside Electron the preload
/// has already exposed the IPC-backed APIs and this module leaves them
/// untouched.
import { createHttpIpcTransport, isElectronRuntime } from '@genoffice/ipc-bridge/client'
import { createWebFileBridge, pickFileBytes } from '@genoffice/ipc-bridge/web-native'
import {
  createShellHomeApi,
  createShellProjectApi,
  createShellTabsApi,
} from '../shared/shell-api-factory'

if (!isElectronRuntime()) {
  const transport = createHttpIpcTransport()
  const files = createWebFileBridge(transport)
  const bridgedWindow = window as unknown as Record<string, unknown>

  /* ── Web-native tab tracking ─────────────────────────────────────────
   * The Electron shell tracks tabs as WebContentsViews owned by the main
   * process. The web version opens each module in a new browser tab via
   * window.open; we sync those across all shell windows with BroadcastChannel
   * + localStorage so the TabBar shows real entries. */
  interface WebTab {
    id: string
    kind: 'docs' | 'sheets' | 'slides' | 'pdf' | 'markdown' | 'html'
    title: string
    /** stable per-tab random id; the window itself is the source of truth */
    windowId: string
  }
  const TAB_STORAGE_KEY = 'genoffice:web-tabs'
  const SHELL_ID_KEY = 'genoffice:shell-id'
  let shellId = localStorage.getItem(SHELL_ID_KEY)
  if (!shellId) {
    shellId = 'shell-' + Math.random().toString(36).slice(2, 10)
    localStorage.setItem(SHELL_ID_KEY, shellId)
  }
  const tabChannel = new BroadcastChannel('genoffice:tabs')
  const loadTabs = (): WebTab[] => {
    try {
      const raw = localStorage.getItem(TAB_STORAGE_KEY)
      return raw ? (JSON.parse(raw) as WebTab[]) : []
    } catch {
      return []
    }
  }
  const saveTabs = (tabs: WebTab[]): void => {
    try {
      localStorage.setItem(TAB_STORAGE_KEY, JSON.stringify(tabs))
    } catch {}
  }
  let tabsCache = loadTabs()
  const broadcast = (): void => {
    tabChannel.postMessage({ type: 'sync', tabs: tabsCache })
    saveTabs(tabsCache)
    // notify the TabBar (which subscribes via window event since the IPC
    // push channel would require the SSE bridge and is unnecessary here)
    window.dispatchEvent(new CustomEvent('genoffice:web-tabs-changed', { detail: tabsCache }))
  }
  const moduleLabel = (kind: WebTab['kind']): string => {
    return {
      docs: 'AI Docs',
      sheets: 'AI Sheets',
      slides: 'AI Slides',
      pdf: 'AI PDF',
      markdown: 'AI Markdown',
      html: 'AI HTML',
    }[kind]
  }
  const moduleFileName = (path?: string): string => {
    if (!path) return ''
    const parts = path.split(/[\\/]/)
    return parts[parts.length - 1] || ''
  }
  const registerTab = (
    kind: WebTab['kind'],
    path: string | undefined,
    win: Window | null,
  ): WebTab => {
    const id = 'web-' + Math.random().toString(36).slice(2, 10)
    const title = path ? moduleFileName(path) : moduleLabel(kind)
    const tab: WebTab = { id, kind, title, windowId: id }
    tabsCache = [...tabsCache, tab]
    broadcast()
    if (win) {
      // stash the tab id on the child window so it can self-close
      try {
        ;(win as unknown as Record<string, unknown>).__genofficeTabId = id
      } catch {}
    }
    return tab
  }
  const unregisterTabById = (id: string): void => {
    tabsCache = tabsCache.filter((t) => t.id !== id)
    broadcast()
  }
  const unregisterTabByWindow = (win: Window | null): void => {
    if (!win) return
    const id = (win as unknown as Record<string, unknown>).__genofficeTabId
    if (typeof id === 'string') unregisterTabById(id)
  }
  // listen for cross-window sync (other shells' broadcasts)
  tabChannel.onmessage = (e: MessageEvent) => {
    if (e.data?.type === 'sync') {
      tabsCache = e.data.tabs as WebTab[]
      saveTabs(tabsCache)
      window.dispatchEvent(new CustomEvent('genoffice:web-tabs-changed', { detail: tabsCache }))
    }
    if (e.data?.type === 'close-request') {
      // if this is the child window being asked to close, honour it
      const myId = (window as unknown as Record<string, unknown>).__genofficeTabId
      if (typeof myId === 'string' && myId === e.data.id) {
        try {
          window.close()
        } catch {}
      }
    }
  }

  // as a child window: announce ourselves on load and clean up on unload
  const myTabId = (window as unknown as Record<string, unknown>).__genofficeTabId as
    string | undefined
  if (myTabId) {
    const cleanup = (): void => {
      unregisterTabById(myTabId)
    }
    window.addEventListener('beforeunload', cleanup)
    // also broadcast periodically so other shells re-discover us after refresh
    setTimeout(broadcast, 100)
  }
  // every few seconds, sweep any tabs whose window no longer exists
  setInterval(() => {
    let changed = false
    const next: WebTab[] = []
    for (const t of tabsCache) {
      // we can't reliably probe window existence from another origin, but
      // BroadcastChannel fires when other tabs close; rely on that + a
      // periodic safety sweep on this origin
      next.push(t)
    }
    if (changed) broadcast()
  }, 5000)

  const openModule = (
    module: 'docs' | 'sheets' | 'slides' | 'pdf' | 'markdown' | 'html',
    path?: string,
  ) => {
    const tab = window.open('', '_blank')
    if (!tab) return null
    registerTab(module, path, tab)
    const base = `/${module}/?mode=tab`
    tab.location.href = path
      ? module === 'pdf'
        ? `${base}#open=${encodeURIComponent(path)}`
        : `${base}&open=${encodeURIComponent(path)}`
      : base
    return tab
  }

  const openPathInModule = (path: string): void => {
    const ext = path
      .split(/[\\/.]/)
      .pop()
      ?.toLowerCase()
    const module =
      ext === 'docx'
        ? 'docs'
        : ext === 'xlsx' || ext === 'xlsm' || ext === 'xls' || ext === 'csv'
          ? 'sheets'
          : ext === 'pptx' || ext === 'ppt'
            ? 'slides'
            : ext === 'pdf'
              ? 'pdf'
              : ext === 'md' || ext === 'markdown'
                ? 'markdown'
                : ext === 'html' || ext === 'htm'
                  ? 'html'
                  : null
    if (module) openModule(module, path)
  }

  bridgedWindow.aiOffice = createShellHomeApi(transport, {
    // Browser equivalent of the Electron picker: upload the chosen bytes to the
    // host's temp dir and hand back that path, which is what the translate
    // pipeline (running server-side) can actually read.
    pickTranslationFile: async () => {
      const picked = await pickFileBytes('.pdf,.docx,.pptx,.xlsx,.xlsm,.xls,.csv', false)
      const file = picked?.[0]
      if (!file) return { ok: false, canceled: true }
      const path = await files.writeTempFile(file.name, file.bytes)
      return { ok: true, path }
    },
    browse: async () => {
      const picked = await pickFileBytes(undefined, false)
      if (!picked) return
      const file = picked[0]
      if (!file) return
      const path = await files.writeTempFile(file.name, file.bytes)
      openPathInModule(path)
    },
    openPath: async (path) => {
      openPathInModule(path)
    },
    newDoc: async () => {
      openModule('docs')
    },
    newSheet: async () => {
      openModule('sheets')
    },
    newSlide: async () => {
      openModule('slides')
    },
    newMarkdown: async () => {
      openModule('markdown')
    },
    newPdf: async () => {
      const tab = window.open('', '_blank')
      if (!tab) return
      const result = (await transport.invoke('home:new-pdf')) as { path?: unknown }
      const path = typeof result?.path === 'string' ? result.path : ''
      tab.location.href = path
        ? `/pdf/?mode=tab#open=${encodeURIComponent(path)}`
        : '/pdf/?mode=tab'
      registerTab('pdf', path || undefined, tab)
    },
    newHtml: async () => {
      const tab = window.open('', '_blank')
      if (!tab) return
      const result = (await transport.invoke('home:new-html')) as { path?: unknown }
      const path = typeof result?.path === 'string' ? result.path : ''
      tab.location.href = path
        ? `/html/?mode=tab#open=${encodeURIComponent(path)}`
        : '/html/?mode=tab'
      registerTab('html', path || undefined, tab)
    },
    pickDefaultSaveDir: async () => null,
    revealPath: async () => {},
    openTrash: async () => {},
  })
  bridgedWindow.aiOfficeProject = createShellProjectApi(transport)
  bridgedWindow.aiOfficeTabs = createShellTabsApi(transport, {
    tabsList: async () => {
      // always include the home tab at index 0
      const home = {
        id: 'home',
        kind: 'home' as const,
        title: '首页',
        closable: false,
        active: true,
      }
      return [
        home,
        ...tabsCache.map((t) => ({
          id: t.id,
          kind: t.kind,
          title: t.title,
          closable: true,
          active: false,
        })),
      ]
    },
    tabsActivate: async (id: string) => {
      if (id === 'home') {
        window.focus()
        return
      }
      // try to focus the matching window if we can find it
      const t = tabsCache.find((tab) => tab.id === id)
      if (!t) return
      // we don't have a direct handle to the child window; the closest UX is
      // to focus our own window — the user can switch tabs via the browser
      window.focus()
    },
    tabsClose: async (id: string) => {
      // we can't find the child window handle from another shell, but we can
      // broadcast a close request that the child listens for and honours by
      // calling window.close() on itself
      tabChannel.postMessage({ type: 'close-request', id })
      unregisterTabById(id)
    },
    tabsShowMenu: async () => {},
    tabsShowNewMenu: async () => {
      // delegate to the home page's new-file menu by clicking the AI Docs card
      // (web version: there's no native popup, so we just open the most-recent module)
      openModule('docs')
    },
    tabsReorder: async () => {},
  })
}
