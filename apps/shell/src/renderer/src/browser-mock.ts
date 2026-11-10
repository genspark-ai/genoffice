import type {
  HomeApi,
  ProjectHomeApi,
  RecentEntry,
  RecentQuery,
  TimelineEntryItem,
  UiLanguage,
  UiTheme,
} from '../../shared/home-api'
import type { TabKind, TabsApi, TabSummary } from '../../shared/tabs-api'

if (typeof window !== 'undefined') {
  let mockRecents: RecentEntry[] = [
    {
      path: '/workspace/Executive_Summary_2026.docx',
      name: 'Executive_Summary_2026.docx',
      ext: 'docx',
      mtimeMs: Date.now() - 1000 * 60 * 15,
      sizeBytes: 145200,
      starred: true,
    },
    {
      path: '/workspace/Financial_Model_Q4.xlsx',
      name: 'Financial_Model_Q4.xlsx',
      ext: 'xlsx',
      mtimeMs: Date.now() - 1000 * 60 * 60 * 2,
      sizeBytes: 842000,
      starred: false,
    },
    {
      path: '/workspace/Product_Keynote_v3.pptx',
      name: 'Product_Keynote_v3.pptx',
      ext: 'pptx',
      mtimeMs: Date.now() - 1000 * 60 * 60 * 24,
      sizeBytes: 4210000,
      starred: true,
    },
    {
      path: '/workspace/Design_System_Spec.md',
      name: 'Design_System_Spec.md',
      ext: 'md',
      mtimeMs: Date.now() - 1000 * 60 * 60 * 48,
      sizeBytes: 32400,
      starred: false,
    },
    {
      path: '/workspace/Q4_Security_Audit.pdf',
      name: 'Q4_Security_Audit.pdf',
      ext: 'pdf',
      mtimeMs: Date.now() - 1000 * 60 * 60 * 72,
      sizeBytes: 1205000,
      starred: false,
    },
  ]

  let currentLang: UiLanguage = 'en'
  let currentTheme: UiTheme = 'system'
  try {
    currentTheme = (localStorage.getItem('genoffice-theme') as UiTheme) || 'system'
  } catch {
    currentTheme = 'system'
  }
  const themeListeners: Array<(t: UiTheme) => void> = []

  let mockTabs: (TabSummary & { filePath?: string })[] = [
    { id: 'home', kind: 'home', title: 'Home', closable: false, active: true },
  ]
  const tabListeners: Array<(t: TabSummary[]) => void> = []

  function openTab(kind: TabKind, title: string, filePath?: string) {
    const existing = mockTabs.find(
      (t) => (filePath && t.filePath === filePath) || t.title === title,
    )
    if (existing) {
      mockTabs = mockTabs.map((t) => ({ ...t, active: t.id === existing.id }))
    } else {
      const id = `tab-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`
      mockTabs = mockTabs
        .map((t) => ({ ...t, active: false }))
        .concat({
          id,
          kind,
          title,
          closable: true,
          active: true,
          filePath,
        })
    }
    for (const fn of tabListeners) fn(mockTabs)
  }

  const mockHomeApi = {
    async recents(query?: RecentQuery) {
      const ext = query?.ext
      const text = query?.query?.trim().toLowerCase()
      const byExt = ext ? mockRecents.filter((r) => r.ext === ext) : mockRecents
      const filtered = text ? byExt.filter((r) => r.path.toLowerCase().includes(text)) : byExt
      return {
        entries: filtered,
        total: filtered.length,
        totalAll: mockRecents.length,
      }
    },
    async starred(query?: RecentQuery) {
      const ext = query?.ext
      const stars = mockRecents.filter((r) => r.starred)
      const text = query?.query?.trim().toLowerCase()
      const byExt = ext ? stars.filter((r) => r.ext === ext) : stars
      const filtered = text ? byExt.filter((r) => r.path.toLowerCase().includes(text)) : byExt
      return {
        entries: filtered,
        total: filtered.length,
        totalAll: stars.length,
      }
    },
    async statPaths(paths: string[]) {
      return mockRecents.filter((r) => paths.includes(r.path))
    },
    async toggleStar(path: string) {
      mockRecents = mockRecents.map((r) => (r.path === path ? { ...r, starred: !r.starred } : r))
    },
    async openPath(path: string) {
      const entry = mockRecents.find((r) => r.path === path)
      const kind: TabKind =
        entry?.ext === 'xlsx'
          ? 'sheets'
          : entry?.ext === 'pptx'
            ? 'slides'
            : entry?.ext === 'pdf'
              ? 'pdf'
              : entry?.ext === 'md'
                ? 'markdown'
                : 'docs'
      openTab(kind, entry?.name ?? path, path)
    },
    async browse() {
      openTab('docs', 'Imported Document')
    },
    async newDoc() {
      openTab('docs', 'Untitled Document.docx')
    },
    async newSheet() {
      openTab('sheets', 'Untitled Spreadsheet.xlsx')
    },
    async newSlide() {
      openTab('slides', 'Untitled Presentation.pptx')
    },
    async newMarkdown() {
      openTab('markdown', 'Untitled Note.md')
    },
    async newFile(ext: string) {
      openTab('markdown', `Untitled.${ext}`)
    },
    async removeRecent(paths: string[]) {
      mockRecents = mockRecents.filter((r) => !paths.includes(r.path))
    },
    async revealPath() {},
    async renameFile(path: string, newName: string) {
      mockRecents = mockRecents.map((r) => (r.path === path ? { ...r, name: newName } : r))
      return { ok: true }
    },
    async duplicateFile(path: string) {
      const found = mockRecents.find((r) => r.path === path)
      if (found) {
        mockRecents.unshift({
          ...found,
          path: `${found.path}-copy`,
          name: `Copy of ${found.name}`,
          mtimeMs: Date.now(),
        })
      }
    },
    async deleteFiles(paths: string[]) {
      mockRecents = mockRecents.filter((r) => !paths.includes(r.path))
    },
    async openTrash() {
      return true
    },
    async getLanguage() {
      return currentLang
    },
    async setLanguage(lang: any) {
      currentLang = lang
    },
    async getUpdateChannel() {
      return 'stable'
    },
    async setUpdateChannel() {},
    async accountStatus() {
      return { loggedIn: false }
    },
    async accountLogin() {
      return false
    },
    onAccountLogin() {
      return () => {}
    },
    async openLoginUrl() {},
    async accountLogout() {},
    async getAppVersion() {
      return '0.6.0-web'
    },
    async onboardingSeen() {
      return true
    },
    async setOnboardingSeen() {
      return true
    },
    pathsForFiles(files: File[]) {
      return files.map((f) => f.name)
    },
    async windowMaterialActive() {
      return false
    },
    async getTheme() {
      return currentTheme
    },
    async setTheme(theme: UiTheme) {
      currentTheme = theme
      try {
        localStorage.setItem('genoffice-theme', theme)
      } catch {}
      for (const fn of themeListeners) fn(theme)
    },
    async getDefaultSaveDir() {
      return '/workspace'
    },
    async pickDefaultSaveDir() {
      return '/workspace'
    },
    onThemeChanged(handler: (theme: UiTheme) => void) {
      themeListeners.push(handler)
      return () => {
        const i = themeListeners.indexOf(handler)
        if (i >= 0) themeListeners.splice(i, 1)
      }
    },
    async openGenTeam() {},
    async openCreditUsage() {},
    async cloudProjectsCached() {
      return { available: true, projects: [], syncedAt: Date.now() }
    },
    async cloudProjectsSync() {
      return { available: true, projects: [], syncedAt: Date.now() }
    },
    async openCloudProject() {},
  } as unknown as HomeApi

  const mockProjectApi: ProjectHomeApi = {
    async listProjects() {
      const nowIso = new Date().toISOString()
      return [
        {
          id: 'proj-default',
          name: 'Default project',
          createdAt: nowIso,
          updatedAt: nowIso,
          lastActiveAt: nowIso,
          fileCount: 3,
          isDefault: true,
        },
        {
          id: 'proj-q4',
          name: 'Q4 Product Roadmap',
          createdAt: nowIso,
          updatedAt: nowIso,
          lastActiveAt: nowIso,
          fileCount: 2,
          isDefault: false,
        },
      ]
    },
    async listFiles() {
      return mockRecents.slice(0, 2).map((r) => r.path)
    },
    async createProject(name: string) {
      const nowIso = new Date().toISOString()
      return {
        id: `proj-${Date.now()}`,
        name,
        createdAt: nowIso,
        updatedAt: nowIso,
        lastActiveAt: nowIso,
        fileCount: 0,
        isDefault: false,
      }
    },
    async renameProject() {},
    async deleteProject() {},
    async moveFile() {},
    async getTimeline(projectId: string) {
      // a few recorded chats so the conversation picker has something to show in the browser preview
      const day = 86_400_000
      const rows = (
        path: string,
        chatId: string,
        daysAgo: number,
        ask: string,
      ): TimelineEntryItem[] => {
        const name = path.split('/').pop() ?? path
        const ts = new Date(Date.now() - daysAgo * day).toISOString()
        return [
          { filePath: path, fileName: name, chatId, ts, role: 'user', preview: ask, seq: 1 },
          { filePath: path, fileName: name, chatId, ts, role: 'assistant', preview: 'Done.', seq: 2 },
        ]
      }
      return projectId === 'proj-q4'
        ? [
            ...rows(mockRecents[2].path, 'c-keynote', 4, 'Tighten the Q4 keynote storyline into 8 slides'),
            ...rows(mockRecents[3].path, 'c-spec', 16, 'Summarise the design system spec for engineers'),
          ]
        : [
            ...rows(mockRecents[0].path, 'c-exec', 2, 'Draft an executive summary of 2026 results'),
            ...rows(mockRecents[1].path, 'c-model', 7, 'Add a sensitivity table to the financial model'),
          ]
    },
  }

  const mockTabsApi: TabsApi = {
    async list() {
      return mockTabs
    },
    async activate(id) {
      mockTabs = mockTabs.map((t) => ({ ...t, active: t.id === id }))
      for (const fn of tabListeners) fn(mockTabs)
    },
    async close(id) {
      const idx = mockTabs.findIndex((t) => t.id === id)
      if (idx >= 0) {
        const wasActive = mockTabs[idx].active
        mockTabs = mockTabs.filter((t) => t.id !== id)
        if (wasActive && mockTabs.length > 0) {
          const nextIdx = Math.max(0, idx - 1)
          mockTabs[nextIdx].active = true
        }
        for (const fn of tabListeners) fn(mockTabs)
      }
    },
    async showMenu() {},
    async showTabMenu() {},
    async showAppMenu() {},
    async detach() {},
    async tearOff() {
      return false
    },
    async dragTornWindow() {},
    async dockTornWindow(_index: number) {},
    async endTornDrag() {},
    onDockPreview(_handler: (preview: { x: number } | null) => void) {
      return () => {}
    },
    reportDockIndex(_index: number) {},
    onChromePressed(_handler: () => void) {
      return () => {}
    },
    onStartRename(_handler: (id: string) => void) {
      return () => {}
    },
    async showNewMenu() {
      await mockHomeApi.newDoc()
    },
    async reorder(id, toIndex) {
      const fromIndex = mockTabs.findIndex((t) => t.id === id)
      if (fromIndex >= 0) {
        const [moved] = mockTabs.splice(fromIndex, 1)
        mockTabs.splice(toIndex, 0, moved)
        for (const fn of tabListeners) fn(mockTabs)
      }
    },
    onChanged(handler) {
      tabListeners.push(handler)
      return () => {
        const i = tabListeners.indexOf(handler)
        if (i >= 0) tabListeners.splice(i, 1)
      }
    },
    notifyChromePressed() {},
  }

  if (!window.aiOffice || typeof window.aiOffice.getLanguage !== 'function') {
    window.aiOffice = mockHomeApi
  }
  const win = window as unknown as { aiOfficeProject?: ProjectHomeApi }
  if (!win.aiOfficeProject || typeof win.aiOfficeProject.listProjects !== 'function') {
    win.aiOfficeProject = mockProjectApi
  }
  if (!window.aiOfficeTabs || typeof window.aiOfficeTabs.list !== 'function') {
    window.aiOfficeTabs = mockTabsApi
  }
}
