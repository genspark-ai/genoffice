/// Transport-agnostic construction of the shell renderer bridge APIs.
///
/// One source of truth for the `window.aiOffice` / `window.aiOfficeProject` /
/// `window.aiOfficeTabs` surface: the sandboxed preload builds it on an
/// ipcRenderer transport (Electron), the browser web-bridge builds the exact
/// same object on the HTTP/SSE transport (web version). Channel names, argument
/// shapes, listener wrappers and return coercion are identical — only the
/// transport differs.
import { AI_MEDIA_PROVIDERS, AI_PROVIDERS, AI_SEARCH_PROVIDERS, getProviderAdapter } from '@genoffice/ai-provider/browser'
import type {
  AiMediaProviderConfig,
  AiMediaProviderId,
  AiMediaProviderMeta,
  AiSearchProviderId,
  AiSearchProviderMeta,
  AiSettings,
  CodexModelCatalog,
} from '@genoffice/ai-provider'
import type { IpcTransport } from '@genoffice/ipc-bridge/client'
import type { AiPanelPrefs } from '@genoffice/ui'
import type {
  AccountLoginEvent,
  AccountStatus,
  AiCapabilitiesReport,
  TranslateFilePickResult,
  TranslateFileResult,
  TranslateFileStatus,
  TranslationKbEntry,
  TranslationKbStats,
  AutoSaveDefault,
  CloudProjectsSnapshot,
  HomeApi,
  MarketplaceCategory,
  MarketplaceCategoryInfo,
  MarketplaceArtifactRef,
  MarketplaceEntry,
  MarketplaceUploadEntry,
  MarketplaceUploadPayload,
  ModuleEntry,
  PiResourceReport,
  ModuleKind,
  PluginEntry,
  PluginKind,
  RecentEntry,
  RecentPage,
  RenameResult,
  ProjectHomeApi,
  ProjectSummaryEntry,
  SkillEntry,
  SkillKind,
  TimelineEntryItem,
  UiLanguage,
} from './home-api'
import {
  HOME_CHANNELS,
  PROJECT_CHANNELS,
  type TranslationCoverage,
} from './home-api'
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
  /** Web-native file picker that returns a server-side path (uploads to a temp dir). */
  pickTranslationFile?: (title?: string) => Promise<TranslateFilePickResult>
  /** Web-native browse (browser file picker → open-path). */
  browse?: () => Promise<void>
  /** Web-native default save dir picker (browser download dir is not selectable; returns null). */
  pickDefaultSaveDir?: () => Promise<string | null>
  /** Web-native reveal (no-op in the browser). */
  revealPath?: (path: string) => Promise<void>
  /** Web-native trash (no-op in the browser). */
  openTrash?: () => Promise<void>
  /** Web-native open/new module actions (browser tabs). */
  openPath?: (path: string) => Promise<void>
  newDoc?: (opts?: { projectId?: string }) => Promise<void>
  newSheet?: (opts?: { projectId?: string }) => Promise<void>
  newSlide?: (opts?: { projectId?: string }) => Promise<void>
  newMarkdown?: (opts?: { projectId?: string }) => Promise<void>
  newPdf?: (opts?: { projectId?: string }) => Promise<void>
  newHtml?: (opts?: { projectId?: string }) => Promise<void>
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
      if (overrides.openPath) return await overrides.openPath(path)
      await t.invoke(HOME_CHANNELS.openPath, path)
    },
    async browse() {
      if (overrides.browse) return await overrides.browse()
      await t.invoke(HOME_CHANNELS.browse)
    },
    async newDoc(opts) {
      if (overrides.newDoc) return await overrides.newDoc(opts)
      await t.invoke(HOME_CHANNELS.newDoc, opts)
    },
    async newSheet(opts) {
      if (overrides.newSheet) return await overrides.newSheet(opts)
      await t.invoke(HOME_CHANNELS.newSheet, opts)
    },
    async newSlide(opts) {
      if (overrides.newSlide) return await overrides.newSlide(opts)
      await t.invoke(HOME_CHANNELS.newSlide, opts)
    },
    async newMarkdown(opts) {
      if (overrides.newMarkdown) return await overrides.newMarkdown(opts)
      await t.invoke(HOME_CHANNELS.newMarkdown, opts)
    },
    async newPdf(opts) {
      if (overrides.newPdf) return await overrides.newPdf(opts)
      await t.invoke(HOME_CHANNELS.newPdf, opts)
    },
    async newHtml(opts) {
      if (overrides.newHtml) return await overrides.newHtml(opts)
      await t.invoke(HOME_CHANNELS.newHtml, opts)
    },
    async getAutoSaveDefault() {
      return (await t.invoke(HOME_CHANNELS.getAutoSaveDefault)) as AutoSaveDefault
    },
    async setAutoSaveDefault(on: boolean) {
      await t.invoke(HOME_CHANNELS.setAutoSaveDefault, on)
    },
    async getAiPanelPrefs() {
      return (await t.invoke(HOME_CHANNELS.getAiPanelPrefs)) as AiPanelPrefs
    },
    async setAiPanelPrefs(patch: Partial<AiPanelPrefs>) {
      return (await t.invoke(HOME_CHANNELS.setAiPanelPrefs, patch)) as AiPanelPrefs
    },
    getAiMediaProviders(): AiMediaProviderMeta[] {
      // Synchronous catalog from the bundled registry. Returning the real
      // list (not []) matters: the AI-media pane derives each block's
      // provider options from it, and an empty catalog used to crash the
      // whole settings modal — see the `options[0]` guard in SettingsModal.
      return AI_MEDIA_PROVIDERS
    },
    getAiSearchProviders(): AiSearchProviderMeta[] {
      return AI_SEARCH_PROVIDERS
    },
    async getCodexModels(_cliPath?: string): Promise<CodexModelCatalog> {
      // Codex model catalog is gathered by the main process; the renderer
      // does not yet wire an IPC channel for it. Return an empty catalog.
      return {} as CodexModelCatalog
    },
    async testAiMediaSettings(_input: {
      provider: AiMediaProviderId
      config: AiMediaProviderConfig
    }) {
      // No IPC channel yet; treat as failure until the main process registers
      // a handler for this probe.
      return { ok: false, error: 'testAiMediaSettings not wired' }
    },
    async testAiSearchSettings(_input: { provider: AiSearchProviderId; apiKey: string }) {
      return { ok: false, error: 'testAiSearchSettings not wired' }
    },
    async listModules(): Promise<ModuleEntry[]> {
      const result = (await t.invoke(HOME_CHANNELS.listModules)) as { modules?: ModuleEntry[] }
      return Array.isArray(result?.modules) ? result!.modules : []
    },
    async setModuleEnabled(id: ModuleKind, enabled: boolean): Promise<ModuleEntry[]> {
      const result = (await t.invoke(HOME_CHANNELS.setModuleEnabled, { id, enabled })) as {
        modules?: ModuleEntry[]
      }
      return Array.isArray(result?.modules) ? result!.modules : []
    },
    async reorderModules(order: ModuleKind[]): Promise<ModuleEntry[]> {
      const result = (await t.invoke(HOME_CHANNELS.reorderModules, { order })) as {
        modules?: ModuleEntry[]
      }
      return Array.isArray(result?.modules) ? result!.modules : []
    },
    async resetModules(): Promise<ModuleEntry[]> {
      const result = (await t.invoke(HOME_CHANNELS.resetModules)) as { modules?: ModuleEntry[] }
      return Array.isArray(result?.modules) ? result!.modules : []
    },
    async listSkills(): Promise<SkillEntry[]> {
      const result = (await t.invoke(HOME_CHANNELS.listSkills)) as { skills?: SkillEntry[] }
      return Array.isArray(result?.skills) ? result!.skills : []
    },
    async toggleSkill(id: SkillKind, enabled: boolean): Promise<SkillEntry[]> {
      const result = (await t.invoke(HOME_CHANNELS.toggleSkill, { id, enabled })) as { skills?: SkillEntry[] }
      return Array.isArray(result?.skills) ? result!.skills : []
    },
    async reloadSkill(id: SkillKind): Promise<SkillEntry[]> {
      const result = (await t.invoke(HOME_CHANNELS.reloadSkill, { id })) as { skills?: SkillEntry[] }
      return Array.isArray(result?.skills) ? result!.skills : []
    },
    async installSkill(name: string) {
      return (await t.invoke(HOME_CHANNELS.installSkill, { name })) as {
        ok: boolean
        error?: string
        installed?: SkillEntry
        skills?: SkillEntry[]
        alreadyInstalled?: boolean
      }
    },
    async uninstallSkill(id: SkillKind) {
      return (await t.invoke(HOME_CHANNELS.uninstallSkill, { id })) as {
        ok: boolean
        error?: string
        skills?: SkillEntry[]
      }
    },
    async installPlugin(id: string) {
      return (await t.invoke(HOME_CHANNELS.installPlugin, { id })) as {
        ok: boolean
        error?: string
        installed?: PluginEntry
        plugins?: PluginEntry[]
        alreadyInstalled?: boolean
      }
    },
    async uninstallPlugin(id: PluginKind) {
      return (await t.invoke(HOME_CHANNELS.uninstallPlugin, { id })) as {
        ok: boolean
        error?: string
        plugins?: PluginEntry[]
      }
    },
    async listMarketplaceSkills(): Promise<MarketplaceEntry[]> {
      const result = (await t.invoke(HOME_CHANNELS.listMarketplaceSkills)) as { skills?: MarketplaceEntry[] }
      return Array.isArray(result?.skills) ? result!.skills : []
    },
    async listMarketplacePlugins(): Promise<MarketplaceEntry[]> {
      const result = (await t.invoke(HOME_CHANNELS.listMarketplacePlugins)) as { plugins?: MarketplaceEntry[] }
      return Array.isArray(result?.plugins) ? result!.plugins : []
    },
    async getMarketplaceAndInstalled() {
      return (await t.invoke(HOME_CHANNELS.getMarketplaceAndInstalled)) as {
        marketplaceSkills: MarketplaceEntry[]
        marketplacePlugins: MarketplaceEntry[]
        installedSkills: SkillEntry[]
        installedPlugins: PluginEntry[]
      }
    },
    async marketplaceCategories() {
      const r = (await t.invoke(HOME_CHANNELS.marketplaceCategories)) as { categories?: MarketplaceCategoryInfo[] }
      return { categories: Array.isArray(r?.categories) ? r!.categories : [] }
    },
    async marketplaceSearch(filters: {
      q?: string
      category?: MarketplaceCategory
      type?: 'skill' | 'plugin'
      minRating?: number
      installed?: boolean | 'all'
      sort?: 'popular' | 'rating' | 'newest' | 'name'
    }) {
      const r = (await t.invoke(HOME_CHANNELS.marketplaceSearch, filters)) as {
        skills?: MarketplaceEntry[]
        plugins?: MarketplaceEntry[]
        total?: number
        filters?: typeof filters
      }
      return {
        skills: Array.isArray(r?.skills) ? r!.skills : [],
        plugins: Array.isArray(r?.plugins) ? r!.plugins : [],
        total: typeof r?.total === 'number' ? r!.total : 0,
        filters: r?.filters,
      }
    },
    async marketplaceDetail(id: string, type: 'skill' | 'plugin') {
      return (await t.invoke(HOME_CHANNELS.marketplaceDetail, { id, type })) as {
        ok: boolean
        error?: string
        type?: 'skill' | 'plugin'
        entry?: MarketplaceEntry
        installed?: SkillEntry | PluginEntry | null
      }
    },
    async marketplaceUpload(kind: 'skill' | 'plugin', payload: MarketplaceUploadPayload) {
      return (await t.invoke(HOME_CHANNELS.marketplaceUpload, { kind, payload })) as {
        ok: boolean
        error?: string
        kind?: 'skill' | 'plugin'
        entry?: MarketplaceEntry
        message?: string
        artifact?: MarketplaceArtifactRef | null
        overwritten?: boolean
      }
    },
    async marketplaceListUploads() {
      const r = (await t.invoke(HOME_CHANNELS.marketplaceListUploads)) as {
        uploads?: MarketplaceUploadEntry[]
        error?: string
      }
      return { uploads: Array.isArray(r?.uploads) ? r!.uploads : [], error: r?.error }
    },
    async marketplaceDeleteUpload(kind: 'skill' | 'plugin', id: string) {
      return (await t.invoke(HOME_CHANNELS.marketplaceDeleteUpload, { kind, id })) as {
        ok: boolean
        error?: string
        kind?: 'skill' | 'plugin'
        id?: string
        uninstalled?: boolean
      }
    },
    async listPiResources() {
      return (await t.invoke(HOME_CHANNELS.listPiResources)) as PiResourceReport
    },
    async getAiCapabilities() {
      return (await t.invoke(HOME_CHANNELS.getAiCapabilities)) as AiCapabilitiesReport
    },
    async listTranslationKb(input) {
      return (await t.invoke(HOME_CHANNELS.translationKbList, input ?? {})) as {
        ok: boolean
        entries: TranslationKbEntry[]
      }
    },
    async upsertTranslationKb(entry) {
      return (await t.invoke(HOME_CHANNELS.translationKbUpsert, entry)) as {
        ok: boolean
        error?: string
      }
    },
    async removeTranslationKb(id) {
      return (await t.invoke(HOME_CHANNELS.translationKbRemove, id)) as {
        ok: boolean
        removed: boolean
      }
    },
    async getTranslationKbStats() {
      return (await t.invoke(HOME_CHANNELS.translationKbStats)) as TranslationKbStats
    },
    async getTranslateFileStatus() {
      return (await t.invoke(HOME_CHANNELS.translateFileStatus)) as TranslateFileStatus
    },
    async getTranslationDictionary() {
      return (await t.invoke(HOME_CHANNELS.translateDictionaryStatus)) as {
        ok: boolean
        dictionary: { path: string; terms: number } | null
      }
    },
    async pickTranslationFile(title) {
      if (overrides.pickTranslationFile) return await overrides.pickTranslationFile()
      return (await t.invoke(
        HOME_CHANNELS.pickTranslationFile,
        title ?? '',
      )) as TranslateFilePickResult
    },
    async buildTranslationDictionary(input) {
      return (await t.invoke(HOME_CHANNELS.translateBuildDictionary, input)) as {
        ok: boolean
        dictionaryPath?: string
        kbEntries?: number
        llmEntries?: number
        missed?: string[]
        totalSegments?: number
        error?: string
      }
    },
    async fillTranslationGaps(input) {
      return (await t.invoke(HOME_CHANNELS.translateFillGaps, input)) as {
        ok: boolean
        dictionaryPath?: string
        added?: number
        stillUncovered?: string[]
        coverageBefore?: TranslationCoverage
        coverageAfter?: TranslationCoverage
        elapsedMs?: number
        error?: string
      }
    },
    async translateFileAuto(input) {
      return (await t.invoke(HOME_CHANNELS.translateFileAuto, input)) as TranslateFileResult
    },
    async translateSnippet(input) {
      return (await t.invoke(HOME_CHANNELS.translateSnippet, input)) as {
        ok: boolean
        translation?: string
        status?: 'translated' | 'memory-hit' | 'failed'
        matchedTerms?: string[]
        dictionaryHits?: string[]
        dictionary?: { path: string; terms: number; hits: number } | null
        sourceLang?: string
        targetLang?: string
        elapsedMs?: number
        error?: string
      }
    },
    async marketplaceRate(id: string, kind: 'skill' | 'plugin', rating: number) {
      return (await t.invoke('home:marketplace-rate', { id, kind, rating })) as {
        ok: boolean
        error?: string
        kind?: 'skill' | 'plugin'
        id?: string
        ratingCount?: number
        averageRating?: number
      }
    },
    async resetSkills(): Promise<SkillEntry[]> {
      const result = (await t.invoke(HOME_CHANNELS.resetSkills)) as { skills?: SkillEntry[] }
      return Array.isArray(result?.skills) ? result!.skills : []
    },
    async listPlugins(): Promise<PluginEntry[]> {
      const result = (await t.invoke(HOME_CHANNELS.listPlugins)) as { plugins?: PluginEntry[] }
      return Array.isArray(result?.plugins) ? result!.plugins : []
    },
    async togglePlugin(id: PluginKind, enabled: boolean): Promise<PluginEntry[]> {
      const result = (await t.invoke(HOME_CHANNELS.togglePlugin, { id, enabled })) as { plugins?: PluginEntry[] }
      return Array.isArray(result?.plugins) ? result!.plugins : []
    },
    async reloadPlugin(id: PluginKind): Promise<PluginEntry[]> {
      const result = (await t.invoke(HOME_CHANNELS.reloadPlugin, { id })) as { plugins?: PluginEntry[] }
      return Array.isArray(result?.plugins) ? result!.plugins : []
    },
    async resetPlugins(): Promise<PluginEntry[]> {
      const result = (await t.invoke(HOME_CHANNELS.resetPlugins)) as { plugins?: PluginEntry[] }
      return Array.isArray(result?.plugins) ? result!.plugins : []
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
    async showAppMenu(x, y) {
      await t.invoke(TABS_CHANNELS.showAppMenu, x, y)
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
