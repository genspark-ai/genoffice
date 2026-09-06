/// Transport-agnostic construction of the shell renderer bridge APIs.
///
/// One source of truth for the `window.aiOffice` / `window.aiOfficeProject` /
/// `window.aiOfficeTabs` surface: the sandboxed preload builds it on an
/// ipcRenderer transport (Electron), the browser web-bridge builds the exact
/// same object on the HTTP/SSE transport (web version). Channel names, argument
/// shapes, listener wrappers and return coercion are identical — only the
/// transport differs.
import { AI_PROVIDERS, getProviderAdapter } from '@genoffice/ai-provider'
import type { AiSettings } from '@genoffice/ai-provider'
import type { IpcTransport } from '@genoffice/ipc-bridge/client'
import type {
  AccountLoginEvent,
  AccountStatus,
  CloudProjectsSnapshot,
  HomeApi,
  RecentEntry,
  RecentPage,
  RenameResult,
  ProjectHomeApi,
  ProjectSummaryEntry,
  TimelineEntryItem,
  UiLanguage,
} from './home-api'
import { HOME_CHANNELS, PROJECT_CHANNELS } from './home-api'
import type { TabsApi, TabSummary } from './tabs-api'
import { TABS_CHANNELS } from './tabs-api'

const UI_LANGUAGES: readonly UiLanguage[] = [
  'zh',
  'en',
  'ja',
  'ko',
  'fr',
  'de',
  'es',
  'th',
  'id',
  'ru',
  'ar',
  'pt',
  'it',
  'pl',
  'nl',
  'ms',
  'he',
  'hi',
  'zh-TW',
]

function isUiLanguage(value: unknown): value is UiLanguage {
  return UI_LANGUAGES.includes(value as UiLanguage)
}

const EMPTY_PAGE: RecentPage = { entries: [], total: 0, totalAll: 0 }

function asRecentPage(result: unknown): RecentPage {
  if (result && typeof result === 'object' && Array.isArray((result as RecentPage).entries)) {
    return result as RecentPage
  }
  return EMPTY_PAGE
}

export interface ShellApiOverrides {
  /** Web-native browse (browser file picker → open-path). */
  browse?: () => Promise<void>
  /** Web-native default save dir picker (browser download dir is not selectable; returns null). */
  pickDefaultSaveDir?: () => Promise<string | null>
  /** Web-native reveal (no-op in the browser). */
  revealPath?: (path: string) => Promise<void>
  /** Web-native trash (no-op in the browser). */
  openTrash?: () => Promise<void>
  /** Web-native tab management (browser tabs). */
  tabsList?: () => Promise<TabSummary[]>
  tabsActivate?: (id: string) => Promise<void>
  tabsClose?: (id: string) => Promise<void>
  tabsShowMenu?: (x: number, y: number) => Promise<void>
  tabsShowNewMenu?: (x: number, y: number) => Promise<void>
  tabsReorder?: (id: string, toIndex: number) => Promise<void>
}

export function createShellHomeApi(t: IpcTransport, overrides: ShellApiOverrides = {}): HomeApi {
  return {
  async recents(query) {
    return asRecentPage(await t.invoke(HOME_CHANNELS.recents, query))
  },
  async starred(query) {
    return asRecentPage(await t.invoke(HOME_CHANNELS.starred, query))
  },
  async statPaths(paths) {
    const result: unknown = await t.invoke(HOME_CHANNELS.statPaths, paths)
    return Array.isArray(result) ? (result as RecentEntry[]) : []
  },
  async toggleStar(path) {
    if (typeof path !== 'string' || !path) throw new Error('Invalid path.')
    await t.invoke(HOME_CHANNELS.toggleStar, path)
  },
  async openPath(path) {
    if (typeof path !== 'string' || !path) throw new Error('Invalid path.')
    await t.invoke(HOME_CHANNELS.openPath, path)
  },
  async browse() {
    if (overrides.browse) return await overrides.browse()
    await t.invoke(HOME_CHANNELS.browse)
  },
  async newDoc(opts) {
    await t.invoke(HOME_CHANNELS.newDoc, opts)
  },
  async newSheet(opts) {
    await t.invoke(HOME_CHANNELS.newSheet, opts)
  },
  async newSlide(opts) {
    await t.invoke(HOME_CHANNELS.newSlide, opts)
  },
  async newMarkdown(opts) {
    await t.invoke(HOME_CHANNELS.newMarkdown, opts)
  },
  async newPdf(opts) {
    await t.invoke(HOME_CHANNELS.newPdf, opts)
  },
  async removeRecent(paths) {
    await t.invoke(HOME_CHANNELS.removeRecent, paths)
  },
  async revealPath(path) {
    if (overrides.revealPath) return await overrides.revealPath(path)
    if (typeof path !== 'string' || !path) throw new Error('Invalid path.')
    await t.invoke(HOME_CHANNELS.revealPath, path)
  },
  async renameFile(path, newName) {
    if (typeof path !== 'string' || !path) throw new Error('Invalid path.')
    const result: unknown = await t.invoke(HOME_CHANNELS.renameFile, path, newName)
    return (result ?? { ok: false, error: 'Rename failed' }) as RenameResult
  },
  async duplicateFile(path) {
    if (typeof path !== 'string' || !path) throw new Error('Invalid path.')
    await t.invoke(HOME_CHANNELS.duplicateFile, path)
  },
  async deleteFiles(paths) {
    await t.invoke(HOME_CHANNELS.deleteFiles, paths)
  },
  async openTrash() {
    if (overrides.openTrash) return await overrides.openTrash()
    await t.invoke(HOME_CHANNELS.openTrash)
  },
  async getLanguage() {
    const result: unknown = await t.invoke(HOME_CHANNELS.getLanguage)
    return isUiLanguage(result) ? result : 'zh'
  },
  async setLanguage(lang) {
    if (!isUiLanguage(lang)) throw new Error('Invalid language.')
    await t.invoke(HOME_CHANNELS.setLanguage, lang)
  },
  async getUpdateChannel() {
    const result: unknown = await t.invoke(HOME_CHANNELS.getUpdateChannel)
    return result === 'beta' ? 'beta' : 'stable'
  },
  async setUpdateChannel(channel) {
    // validated inline: a runtime import from ../shared/update-api would be
    // shared with the update.ts preload entry and get split into a chunk,
    // which sandboxed preload scripts cannot load (window.aiOffice would
    // silently disappear). Preload entries must stay single-file bundles.
    if (channel !== 'stable' && channel !== 'beta') throw new Error('Invalid update channel.')
    await t.invoke(HOME_CHANNELS.setUpdateChannel, channel)
  },
  async accountStatus() {
    const result: unknown = await t.invoke(HOME_CHANNELS.accountStatus)
    return (result ?? { loggedIn: false }) as AccountStatus
  },
  async accountLogin() {
    const result: unknown = await t.invoke(HOME_CHANNELS.accountLogin)
    return result === true
  },
  onAccountLogin(handler) {
    return t.on(HOME_CHANNELS.accountLoginEvent, (ev) => handler(ev as AccountLoginEvent))
  },
  async openLoginUrl() {
    await t.invoke(HOME_CHANNELS.accountLoginOpenUrl)
  },
  async accountLogout() {
    await t.invoke(HOME_CHANNELS.accountLogout)
  },
  async getAppVersion() {
    const result: unknown = await t.invoke(HOME_CHANNELS.getAppVersion)
    return typeof result === 'string' ? result : ''
  },
  async onboardingSeen() {
    const result: unknown = await t.invoke(HOME_CHANNELS.onboardingSeen)
    return result === true
  },
  async setOnboardingSeen() {
    const result: unknown = await t.invoke(HOME_CHANNELS.setOnboardingSeen)
    return result === true
  },
  async getTheme() {
    const result: unknown = await t.invoke(HOME_CHANNELS.getTheme)
    return result === 'dark' || result === 'light' ? result : 'system'
  },
  async setTheme(theme) {
    if (theme !== 'light' && theme !== 'dark' && theme !== 'system')
      throw new Error('Invalid theme.')
    await t.invoke(HOME_CHANNELS.setTheme, theme)
  },
  async getAnalyticsEnabled() {
    const result: unknown = await t.invoke(HOME_CHANNELS.getAnalyticsEnabled)
    return result !== false
  },
  async setAnalyticsEnabled(enabled) {
    if (typeof enabled !== 'boolean') throw new Error('Invalid analytics consent.')
    const result: unknown = await t.invoke(HOME_CHANNELS.setAnalyticsEnabled, enabled)
    return result === true
  },
  async getDefaultSaveDir() {
    const result: unknown = await t.invoke(HOME_CHANNELS.getDefaultSaveDir)
    return typeof result === 'string' ? result : ''
  },
  async pickDefaultSaveDir() {
    if (overrides.pickDefaultSaveDir) return await overrides.pickDefaultSaveDir()
    const result: unknown = await t.invoke(HOME_CHANNELS.pickDefaultSaveDir)
    return typeof result === 'string' && result ? result : null
  },
  onThemeChanged(handler) {
    return t.on('app:theme-changed', (theme) => {
      if (theme === 'light' || theme === 'dark' || theme === 'system') handler(theme)
    })
  },
  async openGenTeam() {
    await t.invoke(HOME_CHANNELS.openGenTeam)
  },
  async openCreditUsage() {
    await t.invoke(HOME_CHANNELS.openCreditUsage)
  },
  async openGitHubRepo() {
    await t.invoke(HOME_CHANNELS.openGitHubRepo)
  },
  async githubStars() {
    const result: unknown = await t.invoke(HOME_CHANNELS.githubStars)
    return typeof result === 'number' && Number.isFinite(result) ? result : null
  },
  async starPromptShouldShow() {
    const result: unknown = await t.invoke(HOME_CHANNELS.starPromptShouldShow)
    const raw = (result ?? {}) as { show?: unknown; docOpens?: unknown }
    return {
      show: raw.show === true,
      docOpens:
        typeof raw.docOpens === 'number' && Number.isFinite(raw.docOpens) ? raw.docOpens : 0,
    }
  },
  async starPromptAction(action) {
    if (action !== 'starred' && action !== 'later') throw new Error('Invalid star prompt action.')
    await t.invoke(HOME_CHANNELS.starPromptAction, action)
  },
  async cloudProjectsCached() {
    const result: unknown = await t.invoke(HOME_CHANNELS.cloudProjectsCached)
    return asCloudProjectsSnapshot(result)
  },
  async cloudProjectsSync() {
    // failures (network / CLI) resolve to null so the renderer keeps whatever it has
    try {
      const result: unknown = await t.invoke(HOME_CHANNELS.cloudProjects)
      return asCloudProjectsSnapshot(result)
    } catch {
      return null
    }
  },
  async openCloudProject(projectUrl) {
    if (typeof projectUrl !== 'string' || !projectUrl) throw new Error('Invalid project URL.')
    await t.invoke(HOME_CHANNELS.openCloudProject, projectUrl)
  },
  // AI settings channels are registered once by the shell's aggregated docs handlers
  async getAiSettings() {
    return (await t.invoke('ai:get-settings')) as AiSettings
  },
  async setAiSettings(settings) {
    await t.invoke('ai:set-settings', settings)
  },
  getAiProviders() {
    return AI_PROVIDERS.map((meta) => {
      let defaultBaseUrl = ''
      // genspark routes by model and custom has no default — both stay ''
      if (meta.id !== 'genspark' && !meta.needsBaseUrl) {
        defaultBaseUrl = getProviderAdapter(meta.id).resolveEndpoint({
          apiKey: '',
          model: meta.defaultModel,
        }).baseUrl
      }
      return { ...meta, defaultBaseUrl }
    })
  },
  async testAiSettings(settings) {
    const result: unknown = await t.invoke('ai:chat', {
      settings,
      system: 'You are a connectivity test. Reply with the single word OK.',
      user: 'ping',
    })
    const raw = (result ?? {}) as { ok?: unknown; error?: unknown }
    return raw.ok === true
      ? { ok: true }
      : { ok: false, error: typeof raw.error === 'string' ? raw.error : 'Connection failed' }
  },
  }
}
function asCloudProjectsSnapshot(result: unknown): CloudProjectsSnapshot | null {
  if (
    result &&
    typeof result === 'object' &&
    Array.isArray((result as CloudProjectsSnapshot).projects)
  ) {
    return result as CloudProjectsSnapshot
  }
  return null
}

export function createShellProjectApi(t: IpcTransport): ProjectHomeApi {
  return {
  async listProjects() {
    const result: unknown = await t.invoke(PROJECT_CHANNELS.list)
    return Array.isArray(result) ? (result as ProjectSummaryEntry[]) : []
  },
  async listFiles(projectId) {
    const result: unknown = await t.invoke(PROJECT_CHANNELS.files, { projectId })
    return Array.isArray(result)
      ? result.filter((path): path is string => typeof path === 'string')
      : []
  },
  async createProject(name) {
    const result: unknown = await t.invoke(PROJECT_CHANNELS.create, { name })
    return result as ProjectSummaryEntry
  },
  async renameProject(id, name) {
    await t.invoke(PROJECT_CHANNELS.rename, { id, name })
  },
  async deleteProject(id) {
    await t.invoke(PROJECT_CHANNELS.delete, { id })
  },
  async moveFile(filePath, projectId) {
    await t.invoke(PROJECT_CHANNELS.moveFile, { filePath, projectId })
  },
  async getTimeline(projectId, limit) {
    const result: unknown = await t.invoke(PROJECT_CHANNELS.timeline, {
      projectId,
      limit,
    })
    return Array.isArray(result) ? (result as TimelineEntryItem[]) : []
  },
  }
}

export function createShellTabsApi(t: IpcTransport, overrides: ShellApiOverrides = {}): TabsApi {
  return {
  async list() {
    if (overrides.tabsList) return await overrides.tabsList()
    const result: unknown = await t.invoke(TABS_CHANNELS.list)
    return Array.isArray(result) ? (result as TabSummary[]) : []
  },
  async activate(id) {
    if (overrides.tabsActivate) return await overrides.tabsActivate(id)
    await t.invoke(TABS_CHANNELS.activate, id)
  },
  async close(id) {
    if (overrides.tabsClose) return await overrides.tabsClose(id)
    await t.invoke(TABS_CHANNELS.close, id)
  },
  async showMenu(x, y) {
    if (overrides.tabsShowMenu) return await overrides.tabsShowMenu(x, y)
    await t.invoke(TABS_CHANNELS.showMenu, x, y)
  },
  async showNewMenu(x, y) {
    if (overrides.tabsShowNewMenu) return await overrides.tabsShowNewMenu(x, y)
    await t.invoke(TABS_CHANNELS.showNewMenu, x, y)
  },
  async reorder(id, toIndex) {
    if (overrides.tabsReorder) return await overrides.tabsReorder(id, toIndex)
    await t.invoke(TABS_CHANNELS.reorder, id, toIndex)
  },
  onChanged(handler) {
    return t.on(TABS_CHANNELS.changed, (tabs) => handler(tabs as TabSummary[]))
  },
  notifyChromePressed() {
    t.send(TABS_CHANNELS.chromePressed)
  },
  onChromePressed(handler) {
    return t.on('app:chrome-pressed', () => handler())
  },
  }
}
