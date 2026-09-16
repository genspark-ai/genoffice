import type {
  AiChatResponse,
  AiMediaProviderConfig,
  AiMediaProviderId,
  AiMediaProviderMeta,
  AiProviderMeta,
  AiSearchProviderId,
  AiSearchProviderMeta,
  AiSettings,
  CodexModelCatalog,
} from '@genoffice/ai-provider'
import type { UpdateChannel } from './update-api'
import type { AiPanelPrefs } from '@genoffice/ui/ai-panel-prefs'

/** UI language; kept self-contained here (mirrors Lang in @genoffice/i18n) */
export type UiLanguage =
  | 'zh'
  | 'en'
  | 'ja'
  | 'ko'
  | 'fr'
  | 'de'
  | 'es'
  | 'th'
  | 'id'
  | 'ru'
  | 'ar'
  | 'pt'
  | 'it'
  | 'pl'
  | 'cs'
  | 'nl'
  | 'ms'
  | 'he'
  | 'hi'
  | 'zh-TW'

/** UI theme preference */
export type UiTheme = 'light' | 'dark' | 'system'

/** shell-wide AutoSave default for every editor; updatedAt is 0 until first set */
export interface AutoSaveDefault {
  on: boolean
  updatedAt: number
}

/** a recent file entry shown on the home screen; type derives from the extension */
export interface RecentEntry {
  path: string
  name: string
  /** lowercased extension without the dot ('docx' | 'xlsx' | 'pptx') */
  ext: string
  /** last-modified time, ms since epoch */
  mtimeMs: number
  /** file size in bytes */
  sizeBytes: number
  /** whether the user starred this file */
  starred: boolean
  /** the path failed to stat (disconnected drive, moved, deleted) — kept
      listed like Word's recents instead of silently dropped (r158) */
  missing?: boolean
}

/** paged query for the home file lists */
export interface RecentQuery {
  /** number of entries to skip (default 0) */
  offset?: number
  /** page size; 0 returns no entries but still reports totals (default 50) */
  limit?: number
  /** restrict to one extension ('docx' | 'xlsx' | 'pptx'); omit for all */
  ext?: string
}

export interface RecentPage {
  entries: RecentEntry[]
  /** total matching the query's ext filter */
  total: number
  /** total ignoring the ext filter (for the sidebar counters) */
  totalAll: number
}

export interface HomeApi {
  /** unified recents across document types, newest first (paged) */
  recents(query?: RecentQuery): Promise<RecentPage>
  /** starred files (independent of the recent list), newest first (paged) */
  starred(query?: RecentQuery): Promise<RecentPage>
  /** stat a specific set of paths (project view); unstat-able files come back flagged `missing` */
  statPaths(paths: string[]): Promise<RecentEntry[]>
  /** star / unstar a file */
  toggleStar(path: string): Promise<void>
  /** open an existing file, routing to the right module by extension */
  openPath(path: string): Promise<void>
  /** file picker accepting every supported extension, then routes */
  browse(): Promise<void>
  /** open a docs window at its start screen */
  newDoc(opts?: { projectId?: string }): Promise<void>
  /** open a sheets window */
  newSheet(opts?: { projectId?: string }): Promise<void>
  /** open a slides tab at its start screen (open-a-pptx) */
  newSlide(opts?: { projectId?: string }): Promise<void>
  /** open a blank markdown editor tab */
  newMarkdown(opts?: { projectId?: string }): Promise<void>
  /** open a blank html editor tab */
  newHtml(opts?: { projectId?: string }): Promise<void>
  /** create a blank single-page PDF in the default save folder and open it */
  newPdf(opts?: { projectId?: string }): Promise<void>
  /** drop entries from the recent list (does not touch the files) */
  removeRecent(paths: string[]): Promise<void>
  /** reveal the file in Finder / Explorer */
  revealPath(path: string): Promise<void>
  /** rename the file on disk (same directory) and update the recent list */
  renameFile(path: string, newName: string): Promise<RenameResult>
  /** copy the file next to itself (localized "copy" suffix before .ext) and record it as recent */
  duplicateFile(path: string): Promise<void>
  /** move files to the trash and drop them from the recent list */
  deleteFiles(paths: string[]): Promise<void>
  /** open the OS trash, where deleted files can be restored */
  openTrash(): Promise<void>
  /** current UI language (persisted in userData/app-settings.json) */
  getLanguage(): Promise<UiLanguage>
  /** switch + persist the UI language; main rebuilds its menus to match */
  setLanguage(lang: UiLanguage): Promise<void>
  /** current update channel (persisted in userData/app-settings.json; default 'stable') */
  getUpdateChannel(): Promise<UpdateChannel>
  /** switch + persist the update channel; triggers an immediate update check */
  setUpdateChannel(channel: UpdateChannel): Promise<void>
  /** Genspark account status (gsk login state; to be upgraded to a signup/account system later) */
  accountStatus(): Promise<AccountStatus>
  /** start Genspark login (opens the browser; accountStatus flips to logged-in on completion); returns whether the launch succeeded */
  accountLogin(): Promise<boolean>
  /** progress events for the login started via accountLogin; returns an unsubscribe */
  onAccountLogin(handler: (ev: AccountLoginEvent) => void): () => void
  /** re-open the pending login auth URL in the default browser (rescue when auto-open failed) */
  openLoginUrl(): Promise<void>
  /** log out (clears the saved API key; the login state is shared globally with the gsk CLI) */
  accountLogout(): Promise<void>
  /** app version (from package.json / electron app.getVersion) */
  getAppVersion(): Promise<string>
  /** whether the first-run onboarding has been completed or skipped (persisted in userData/app-settings.json) */
  onboardingSeen(): Promise<boolean>
  /** mark onboarding done; analytics remains enabled unless separately opted out */
  setOnboardingSeen(): Promise<boolean>
  /** current UI theme preference (persisted in userData/app-settings.json) */
  getTheme(): Promise<UiTheme>
  /** switch + persist the UI theme; broadcasts 'app:theme-changed' to all web contents */
  setTheme(theme: UiTheme): Promise<void>
  /** AutoSave default applied by every editor window (persisted in userData/app-settings.json) */
  getAutoSaveDefault(): Promise<AutoSaveDefault>
  /** persist the AutoSave default; broadcasts 'app:auto-save-default-changed' to all web contents */
  setAutoSaveDefault(on: boolean): Promise<void>
  /** whether anonymous usage statistics are enabled (default true in official builds) */
  getAnalyticsEnabled(): Promise<boolean>
  /** persist an explicit analytics opt-in or opt-out */
  setAnalyticsEnabled(enabled: boolean): Promise<boolean>
  /** AI panel text size + chat-input spellcheck (persisted in userData/app-settings.json) */
  getAiPanelPrefs(): Promise<AiPanelPrefs>
  /** merge + persist; broadcasts 'app:ai-panel-prefs-changed' to all web contents */
  setAiPanelPrefs(patch: Partial<AiPanelPrefs>): Promise<AiPanelPrefs>
  /** effective default save folder for new/untitled files (configured in userData/app-settings.json, falls back to <Documents>/GenOffice) */
  getDefaultSaveDir(): Promise<string>
  /** directory picker to change the default save folder; resolves to the new folder, or null when canceled or the pick was unusable */
  pickDefaultSaveDir(): Promise<string | null>
  /** theme switched anywhere (broadcast from the main process) */
  onThemeChanged(handler: (theme: UiTheme) => void): () => void
  /** open the GenTeam community page in the default browser */
  openGenTeam(): Promise<void>
  /** open the Genspark credit-usage page in the default browser */
  openCreditUsage(): Promise<void>
  /** open the public GitHub repository in the default browser */
  openGitHubRepo(): Promise<void>
  /** current stargazer count of the public repo (null while offline / rate-limited) */
  githubStars(): Promise<number | null>
  /** ordered list of quick-create modules (with enabled flag) — drives the home row */
  listModules(): Promise<ModuleEntry[]>
  /** toggle a single module's enabled flag; the order is preserved */
  setModuleEnabled(id: ModuleKind, enabled: boolean): Promise<ModuleEntry[]>
  /** persist a new module order; ids missing from `order` are appended in their existing relative order */
  reorderModules(order: ModuleKind[]): Promise<ModuleEntry[]>
  /** reset all modules to the factory order and enabled=true */
  resetModules(): Promise<ModuleEntry[]>
  /** list all agent skills (built-in + installed from marketplace) */
  listSkills(): Promise<SkillEntry[]>
  /** enable or disable a skill; disabled skills are not loaded by pi's ExtensionRunner */
  toggleSkill(id: SkillKind, enabled: boolean): Promise<SkillEntry[]>
  /** hot-reload a skill (re-run ExtensionRunner registration) */
  reloadSkill(id: SkillKind): Promise<SkillEntry[]>
  /** list marketplace skills (catalog of available 3rd-party skills) */
  listMarketplaceSkills(): Promise<MarketplaceEntry[]>
  /** list marketplace plugins */
  listMarketplacePlugins(): Promise<MarketplaceEntry[]>
  /** composite: marketplace + installed skills + installed plugins in one call */
  getMarketplaceAndInstalled(): Promise<{
    marketplaceSkills: MarketplaceEntry[]
    marketplacePlugins: MarketplaceEntry[]
    installedSkills: SkillEntry[]
    installedPlugins: PluginEntry[]
  }>
  /** W32 — full marketplace v2 surface */
  marketplaceCategories(): Promise<{ categories: MarketplaceCategoryInfo[] }>
  marketplaceSearch(filters: MarketplaceSearchFilters): Promise<MarketplaceSearchResult>
  marketplaceDetail(id: string, type: 'skill' | 'plugin'): Promise<{
    ok: boolean
    error?: string
    type?: 'skill' | 'plugin'
    entry?: MarketplaceEntry
    installed?: SkillEntry | PluginEntry | null
    /** Per-entry pi install state (packageDir, mode, hasCode, skillPath, …).
     *  Shaped by the server's marketplace-detail handler; we only narrow it
     *  at the call site (renderMpCard + detail drawer) so the renderer can
     *  stay loose about optional fields the server may add later. */
    pi?: Record<string, unknown>
  }>
  marketplaceUpload(
    kind: 'skill' | 'plugin',
    payload: MarketplaceUploadPayload,
  ): Promise<{
    ok: boolean
    error?: string
    kind?: 'skill' | 'plugin'
    entry?: MarketplaceEntry
    message?: string
    artifact?: MarketplaceArtifactRef | null
    overwritten?: boolean
  }>
  marketplaceListUploads(): Promise<{ uploads: MarketplaceUploadEntry[]; error?: string }>
  /** Unpublish a marketplace entry by id. Removes its artifact + catalog card
   *  and uninstalls it locally when the entry is currently installed. */
  marketplaceDeleteUpload(
    kind: 'skill' | 'plugin',
    id: string,
  ): Promise<{ ok: boolean; error?: string; kind?: 'skill' | 'plugin'; id?: string; uninstalled?: boolean }>
  /** Real pi resource report: every extension/skill/prompt/theme pi's loader
   *  would pick up for the current agent dir, plus installed packages. */
  listPiResources(): Promise<PiResourceReport>
  marketplaceRate(
    id: string,
    kind: 'skill' | 'plugin',
    rating: number,
  ): Promise<{
    ok: boolean
    error?: string
    kind?: 'skill' | 'plugin'
    id?: string
    ratingCount?: number
    averageRating?: number
  }>
  /** install a skill from the marketplace by id */
  installSkill(name: string): Promise<{ ok: boolean; error?: string; installed?: SkillEntry; skills?: SkillEntry[]; alreadyInstalled?: boolean }>
  /** uninstall a marketplace-installed skill (built-in skills cannot be uninstalled) */
  uninstallSkill(id: SkillKind): Promise<{ ok: boolean; error?: string; skills?: SkillEntry[] }>
  /** install a plugin from the marketplace by id */
  installPlugin(id: string): Promise<{ ok: boolean; error?: string; installed?: PluginEntry; plugins?: PluginEntry[]; alreadyInstalled?: boolean }>
  /** uninstall a marketplace-installed plugin */
  uninstallPlugin(id: PluginKind): Promise<{ ok: boolean; error?: string; plugins?: PluginEntry[] }>
  /** reset all skills to factory defaults */
  resetSkills(): Promise<SkillEntry[]>
  /** list all plugins (advanced extensions) */
  listPlugins(): Promise<PluginEntry[]>
  /** enable or disable a plugin */
  togglePlugin(id: PluginKind, enabled: boolean): Promise<PluginEntry[]>
  /** hot-reload a plugin */
  reloadPlugin(id: PluginKind): Promise<PluginEntry[]>
  /** reset all plugins to factory defaults */
  resetPlugins(): Promise<PluginEntry[]>
  /** whether the one-time "star us" prompt should show now (show:true also counts as shown);
   * docOpens personalizes the card copy ("you've opened N documents") */
  starPromptShouldShow(): Promise<StarPromptShow>
  /** user reacted to the star prompt; 'starred' resolves it permanently */
  starPromptAction(action: StarPromptAction): Promise<void>
  /** locally stored full cloud project list (instant; null when no store or logged out) */
  cloudProjectsCached(): Promise<CloudProjectsSnapshot | null>
  /** sync the full list from Genspark and return it (1 request when nothing changed); null when the sync failed */
  cloudProjectsSync(): Promise<CloudProjectsSnapshot | null>
  /** open a cloud project (relative '/agents?id=...' URL) in the default browser */
  openCloudProject(projectUrl: string): Promise<void>
  /** AI settings (userData/ai-settings.json, shared by every editor); the genspark key never appears here */
  getAiSettings(): Promise<AiSettings>
  /** persist AI settings; open editors pick the change up on their next settings read */
  setAiSettings(settings: AiSettings): Promise<void>
  /** provider catalog with each fixed endpoint's default base URL (empty for genspark/custom) */
  getAiProviders(): AiCatalogEntry[]
  /** live Codex model catalog discovered through the current or overridden app-server */
  getCodexModels(cliPath?: string): Promise<CodexModelCatalog>
  /** one-shot round trip against the given (possibly unsaved) settings — the settings-UI connection test */
  testAiSettings(settings: AiSettings): Promise<AiChatResponse>
  /** image generation / media analysis provider catalog */
  getAiMediaProviders(): AiMediaProviderMeta[]
  /** credential check for a (possibly unsaved) media provider; genspark reports the gsk login state */
  testAiMediaSettings(input: {
    provider: AiMediaProviderId
    config: AiMediaProviderConfig
  }): Promise<{ ok: boolean; error?: string }>
  /** web search provider catalog */
  getAiSearchProviders(): AiSearchProviderMeta[]
  /** one minimal query against the given key (genspark reports the gsk login state) */
  testAiSearchSettings(input: {
    provider: AiSearchProviderId
    apiKey: string
  }): Promise<{ ok: boolean; error?: string }>
  /** Real per-feature availability report from the server. Counts the
   *  zero-config DuckDuckGo fallback as `available` so the UI does not
   *  mislead the user into thinking search is broken without a key. */
  getAiCapabilities(): Promise<AiCapabilitiesReport>
  /** Translation knowledge base — CRUD + resolution. */
  listTranslationKb?(input?: TranslationKbListInput): Promise<{ ok: boolean; entries: TranslationKbEntry[] }>
  upsertTranslationKb?(entry: TranslationKbEntry): Promise<{ ok: boolean; error?: string }>
  removeTranslationKb?(id: string): Promise<{ ok: boolean; removed: boolean }>
  getTranslationKbStats?(): Promise<TranslationKbStats>
  /** Whole-file translation (PDF / XLS(X) / PPTX / DOCX). */
  getTranslateFileStatus?(): Promise<TranslateFileStatus>
  /** Native file picker for the translation pane; returns an absolute path. */
  pickTranslationFile?(title?: string): Promise<TranslateFilePickResult>
  /** Translate a short text snippet using the active provider + KB + memory. */
  translateSnippet?(input: {
    text: string
    sourceLang?: string
    targetLang: string
    customerName?: string
  }): Promise<{
    ok: boolean
    translation?: string
    status?: 'translated' | 'memory-hit' | 'failed'
    matchedTerms?: string[]
    sourceLang?: string
    targetLang?: string
    elapsedMs?: number
    error?: string
  }>
  /** Build a `--dictionary` from a file using the KB + the active provider. */
  buildTranslationDictionary?(input: {
    inputPath: string
    sourceLang?: string
    targetLang: string
    customerName?: string
    glossaryCategory?: string
    useLlm?: boolean
  }): Promise<{
    ok: boolean
    dictionaryPath?: string
    kbEntries?: number
    llmEntries?: number
    missed?: string[]
    totalSegments?: number
    elapsedMs?: number
    segments?: { source: string; target: string; origin: 'kb' | 'llm' }[]
    error?: string
  }>
  /** Build the dictionary, then translate the file in one call. */
  translateFileAuto?(input: {
    inputPath: string
    outputPath?: string
    sourceLang?: string
    targetLang: string
    customerName?: string
    glossaryCategory?: string
    scale?: number
  }): Promise<TranslateFileResult>
}

/** one capability entry — see home:ai-capabilities */
export interface AiCapabilityEntry {
  /** true when at least one backend can serve the request right now */
  available: boolean
  /** primary backend that will be tried first (genspark, duckduckgo, openai, …) */
  via: string
  /** secondary backend used when the primary fails */
  fallback?: string
  /** true when a keyed provider (or genspark credit balance) is set up;
   *  false means the only path is the zero-config fallback */
  configured: boolean
  /** optional human-readable note shown in the UI */
  note?: string
}

/** full capability report from home:ai-capabilities */
export interface AiCapabilitiesReport {
  ok: boolean
  capabilities: {
    search: AiCapabilityEntry
    image_search: AiCapabilityEntry
    image_generation: AiCapabilityEntry
    media_analysis: AiCapabilityEntry
  }
  /** active chat provider id (matches AiSettings.provider) */
  provider: string
  /** false when the user has disabled gsk-backed tools globally */
  gskToolsEnabled: boolean
}

/** Scope ordering for a translation KB entry — most specific first. */
export type TranslationKbScope = 'session' | 'customer' | 'project' | 'company' | 'global'

/** The five schemas the translation knowledge base understands. */
export type TranslationKbSchema =
  | 'term'
  | 'forbidden'
  | 'brand'
  | 'styleRule'
  | 'customerPreference'

/**
 * A single KB row. The backend derives the schema (`trade.translation.*`) from
 * which fields are present, so `kind` is a UI convenience only — read it back
 * with a shape check rather than trusting the wire, and never send it.
 */
export interface TranslationKbEntry {
  id: string
  kind?: TranslationKbSchema
  scope: TranslationKbScope
  priority: number
  /** term */
  sourceTerm?: string
  targetTerm?: string
  category?: string
  /** forbidden */
  forbiddenText?: string
  replacement?: string
  reason?: string
  /** brand */
  word?: string
  policy?: 'neverTranslate' | 'keep' | 'translateAs'
  translateAs?: string
  /** style rule */
  name?: string
  description?: string
  /** customer preference */
  customerName?: string
  preferenceType?: string
  value?: string
  /** optional language filter (term / forbidden / customerPreference) */
  sourceLang?: string
  targetLang?: string
}

export interface TranslationKbListInput {
  schema?: TranslationKbSchema
  scope?: TranslationKbScope
  sourceLang?: string
  targetLang?: string
}

export interface TranslationKbStats {
  ok: boolean
  total: number
  bySchema: Partial<Record<TranslationKbSchema, number>>
  dirty: boolean
}

/** Result of a whole-file translation (PDF / XLS(X) / PPTX / DOCX). */
export interface TranslateFileResult {
  ok: boolean
  stage?: 'dictionary' | 'translate' | 'done'
  outputPath?: string
  bytes?: number
  elapsedMs?: number
  error?: string
  dictionaryPath?: string
  dictionary?: {
    kbEntries?: number
    llmEntries?: number
    missed?: string[]
    totalSegments?: number
    elapsedMs?: number
    segments?: { source: string; target: string; origin: 'kb' | 'llm' }[]
  }
}

export interface TranslateFilePickResult {
  ok: boolean
  /** absolute path on the host filesystem; absent when the user cancelled */
  path?: string
  canceled?: boolean
}

export interface TranslateFileStatus {
  ok: boolean
  available: boolean
  source: 'override' | 'bundled' | 'legacy' | 'missing'
  skillDir: string
  pythonPath: string
  supportedExtensions: string[]
}

export interface AiCatalogEntry extends AiProviderMeta {
  /** default endpoint for fixed-endpoint providers ('' = model-dependent or user-supplied) */
  defaultBaseUrl: string
}

/** 'starred' = went to GitHub or said "already starred" (never prompt again);
 * 'later' = dismissed this time (already counted as shown by the query) */
export type StarPromptAction = 'starred' | 'later'

/** answer to starPromptShouldShow */
export interface StarPromptShow {
  show: boolean
  /** lifetime documents opened — drives the personalized card title */
  docOpens: number
}

export type CloudProjectKind = 'docs' | 'sheets' | 'slides'

/** Identifier for one of the quick-create modules shown on the home page. */
export type ModuleKind = 'docx' | 'xlsx' | 'pptx' | 'md' | 'pdf'

/** One module entry shown in the home-page quick-start row and the module manager. */
export interface ModuleEntry {
  id: ModuleKind
  /** i18n key whose value carries the visible label (matches Home.tsx strings). */
  labelKey: string
  /** File extension used by the file-icon map (no leading dot). */
  ext: string
  /** Subpath under the web origin where the module renders (e.g. '/docs/'). */
  path: string
  /** When false the module is hidden from the quick-start row. */
  enabled: boolean
}

// ── Skills & Plugins ────────────────────────────────────────
// Skills 是 GenOffice 的 agent skill extension(plan §2.1)
// Plugins 是更高级的扩展(multi-agent/audit/local models)
export type SkillKind =
  | 'docs-skill'
  | 'sheets-skill'
  | 'slides-skill'
  | 'office-workflow'
  | 'office-safety'
  | 'frozen-selection'
  | 'verify-response'
  | 'skill-market'
  // marketplace skills (3rd-party)
  | 'notion-sync'
  | 'pdf-ocr-pro'
  | 'github-integration'
  | 'jira-bridge'
  | 'lang-detector'
  // Allow arbitrary future marketplace ids without per-use casts.
  | (string & {})

export type PluginKind =
  | 'agent-team'
  | 'audit-log'
  | 'local-models'
  // marketplace plugins (3rd-party)
  | 'slack-bridge'
  | 'gdrive-export'
  // Allow arbitrary future marketplace ids without per-use casts.
  | (string & {})

export type SkillStatus = 'enabled' | 'disabled' | 'error'

export interface SkillEntry {
  id: SkillKind
  name: string
  description: string
  author: string
  version: string
  package: string
  source: string
  tools: string[]
  scopes: string[]
  status: SkillStatus
  lastLoadedAt: string | null
  error?: string
  builtIn: boolean
}

/** an entry in the marketplace catalog — not yet installed in the local environment */
export type MarketplaceCategory =
  | 'productivity'
  | 'data'
  | 'dev'
  | 'media'
  | 'translation'
  | 'collaboration'
  | 'finance'
  | 'design'

export interface MarketplaceEntry {
  id: string
  name: string
  description: string
  longDescription?: string
  author: string
  version: string
  package: string
  source: string
  tools: string[]
  scopes: string[]
  requirements?: string[]
  category: MarketplaceCategory
  tags: string[]
  /** 0-5 stars, one decimal */
  rating: number
  /** Total download count */
  downloads: number
  /** Featured in the top of the marketplace */
  featured?: boolean
  icon?: string
  homepage?: string
  /** true when this marketplace entry is already installed locally */
  installed: boolean
  /** Artifact the publisher uploaded (real pi extension module or SKILL.md). */
  artifact?: MarketplaceArtifactRef
  /** For a plugin: pi package source (`npm:`/`git:`/`https:`). */
  piPackage?: string
}

/** A publisher-uploaded file that ships with an entry. The catalog card
 *  advertises the file, install writes it to disk. */
export interface MarketplaceArtifactRef {
  filename: string
  bytes: number
  kind: 'extension' | 'skill-md'
  size?: string
}

export interface MarketplaceCategoryInfo {
  id: MarketplaceCategory
  label: string
  count: number
}

export interface MarketplaceSearchFilters {
  q?: string
  category?: MarketplaceCategory
  type?: 'skill' | 'plugin'
  minRating?: number
  installed?: boolean | 'all'
  sort?: 'popular' | 'rating' | 'newest' | 'name'
}

export interface MarketplaceSearchResult {
  skills: MarketplaceEntry[]
  plugins: MarketplaceEntry[]
  total: number
  filters?: MarketplaceSearchFilters
}

export interface MarketplaceUploadPayload {
  id: string
  name: string
  description: string
  /** Optional rich description shown in the detail drawer (plain text up to
   *  ~1024 chars). Falls back to `description` in the UI when absent. */
  longDescription?: string
  version: string
  tools: string[]
  scopes: string[]
  category: MarketplaceCategory
  tags?: string[]
  author?: string
  requirements?: string[]
  icon?: string
  homepage?: string
  /** Real file the publisher attaches. Skills accept a SKILL.md, plugins
   *  accept a pi extension module (.ts/.js). Without this the install only
   *  ships metadata + a synthesized body. */
  artifact?: { filename: string; content: string }
  /** For a plugin: a published pi package source (`npm:` / `git:` /
   *  `https://`). When set the install delegates to pi's package manager
   *  instead of building a local package. */
  piPackage?: string
}

export interface MarketplaceUploadEntry {
  file: string
  kind?: 'skill' | 'plugin'
  id?: string
  name?: string
  uploadedAt?: string
  error?: string
  reviewStatus?: string
  artifact?: { filename: string; kind: string; size?: string } | null
}

export interface PiResourceEntry {
  path: string
  /** file name or skill name, for display */
  name: string
  /** false when the source is registered but disabled/filtered out */
  enabled: boolean
  /** true when the file lives under a GenOffice-managed root */
  managed: boolean
}

export interface PiResourceReport {
  cwd: string
  agentDir: string
  settingsPath: string
  extensions: PiResourceEntry[]
  skills: PiResourceEntry[]
  prompts: PiResourceEntry[]
  themes: PiResourceEntry[]
  /** Counts of resources that come from the user's own pi setup, not GenOffice. */
  external: { extensions: number; skills: number; prompts: number; themes: number }
  packages: Array<{ source: string; scope: string; filtered: boolean; installedPath?: string }>
  diagnostics: Array<{ type: string; message: string; path?: string }>
}

export interface PluginEntry {
  id: PluginKind
  name: string
  description: string
  author: string
  version: string
  package: string
  source: string
  tools: string[]
  scopes: string[]
  requirements: string[]
  status: SkillStatus
  lastLoadedAt: string | null
  error?: string
  builtIn: boolean
}

/** a Genspark web project shown in the home cloud section */
export interface CloudProjectEntry {
  projectId: string
  title: string
  /** module kind derived from the API project type ('docs_agent' → 'docs') */
  kind: CloudProjectKind | 'other'
  /** creation time, ms since epoch (0 when unparsable) */
  ctimeMs: number
  /** relative genspark.ai URL ('/agents?id=...') */
  projectUrl: string
}

/** full local copy of the cloud project list; filtering/paging are client-side */
export interface CloudProjectsSnapshot {
  /** false when gsk is unavailable (CLI missing or not logged in) */
  available: boolean
  /** all projects, newest first */
  projects: CloudProjectEntry[]
  /** ms epoch of the last successful sync (0 when never synced) */
  syncedAt: number
}

export interface AccountStatus {
  /** gsk is installed and logged in */
  loggedIn: boolean
  email?: string
  /** remaining Genspark credits (absent when the balance query failed) */
  creditBalance?: number
}

/** login flow progress pushed from main (gsk login CLI output) */
export interface AccountLoginEvent {
  phase: 'launched' | 'url' | 'success' | 'error'
  url?: string
  expiresInSec?: number
  /** 'network' | 'expired' | raw CLI error text */
  error?: string
}

export interface RenameResult {
  ok: boolean
  /** the new absolute path when ok */
  path?: string
  error?: string
}

// ── Project-related APIs (P1) ────────────────────────────────

export interface ProjectSummaryEntry {
  id: string
  name: string
  createdAt: string
  updatedAt: string
  fileCount: number
  lastActiveAt: string
  isDefault: boolean
}

export interface TimelineEntryItem {
  filePath: string
  fileName: string
  chatId: string
  ts: string
  role: 'user' | 'assistant'
  preview: string
  seq: number
}

export interface ProjectHomeApi {
  /** list all projects (with file count + last-active time) */
  listProjects(): Promise<ProjectSummaryEntry[]>
  /** list existing files currently belonging to a project */
  listFiles(projectId: string): Promise<string[]>
  /** create a project */
  createProject(name: string): Promise<ProjectSummaryEntry>
  /** rename a project */
  renameProject(id: string, name: string): Promise<void>
  /** soft-delete a project */
  deleteProject(id: string): Promise<void>
  /** move a file into the given project */
  moveFile(filePath: string, projectId: string): Promise<void>
  /** fetch the project timeline */
  getTimeline(projectId: string, limit?: number): Promise<TimelineEntryItem[]>
}

export const HOME_CHANNELS = {
  recents: 'home:recents',
  starred: 'home:starred',
  statPaths: 'home:stat-paths',
  toggleStar: 'home:toggle-star',
  openPath: 'home:open-path',
  browse: 'home:browse',
  newDoc: 'home:new-doc',
  newSheet: 'home:new-sheet',
  newSlide: 'home:new-slide',
  newMarkdown: 'home:new-markdown',
  newHtml: 'home:new-html',
  newPdf: 'home:new-pdf',
  removeRecent: 'home:remove-recent',
  revealPath: 'home:reveal-path',
  renameFile: 'home:rename-file',
  duplicateFile: 'home:duplicate-file',
  deleteFiles: 'home:delete-files',
  openTrash: 'home:open-trash',
  getLanguage: 'home:get-language',
  setLanguage: 'home:set-language',
  getUpdateChannel: 'home:get-update-channel',
  setUpdateChannel: 'home:set-update-channel',
  accountStatus: 'home:account-status',
  accountLogin: 'home:account-login',
  accountLoginEvent: 'home:account-login-event',
  accountLoginOpenUrl: 'home:account-login-open-url',
  accountLogout: 'home:account-logout',
  getAppVersion: 'home:get-app-version',
  onboardingSeen: 'home:onboarding-seen',
  setOnboardingSeen: 'home:set-onboarding-seen',
  getTheme: 'home:get-theme',
  setTheme: 'home:set-theme',
  getAutoSaveDefault: 'home:get-auto-save-default',
  setAutoSaveDefault: 'home:set-auto-save-default',
  getAnalyticsEnabled: 'home:get-analytics-enabled',
  setAnalyticsEnabled: 'home:set-analytics-enabled',
  getAiPanelPrefs: 'home:get-ai-panel-prefs',
  setAiPanelPrefs: 'home:set-ai-panel-prefs',
  getDefaultSaveDir: 'home:get-default-save-dir',
  pickDefaultSaveDir: 'home:pick-default-save-dir',
  openGenTeam: 'home:open-genteam',
  openCreditUsage: 'home:open-credit-usage',
  openGitHubRepo: 'home:open-github-repo',
  githubStars: 'home:github-stars',
  starPromptShouldShow: 'home:star-prompt-should-show',
  starPromptAction: 'home:star-prompt-action',
  cloudProjects: 'home:cloud-projects',
  cloudProjectsCached: 'home:cloud-projects-cached',
  openCloudProject: 'home:open-cloud-project',
  listModules: 'home:list-modules',
  setModuleEnabled: 'home:set-module-enabled',
  reorderModules: 'home:reorder-modules',
  resetModules: 'home:reset-modules',
  listSkills: 'home:list-skills',
  toggleSkill: 'home:toggle-skill',
  reloadSkill: 'home:reload-skill',
  installSkill: 'home:install-skill',
  uninstallSkill: 'home:uninstall-skill',
  resetSkills: 'home:reset-skills',
  listPlugins: 'home:list-plugins',
  togglePlugin: 'home:toggle-plugin',
  reloadPlugin: 'home:reload-plugin',
  installPlugin: 'home:install-plugin',
  uninstallPlugin: 'home:uninstall-plugin',
  resetPlugins: 'home:reset-plugins',
  listMarketplaceSkills: 'home:list-marketplace-skills',
  listMarketplacePlugins: 'home:list-marketplace-plugins',
  getMarketplaceAndInstalled: 'home:get-marketplace-and-installed',
  marketplaceCategories: 'home:marketplace-categories',
  marketplaceSearch: 'home:marketplace-search',
  marketplaceDetail: 'home:marketplace-detail',
  marketplaceUpload: 'home:marketplace-upload',
  marketplaceListUploads: 'home:marketplace-list-uploads',
  marketplaceDeleteUpload: 'home:marketplace-delete-upload',
  listPiResources: 'home:list-pi-resources',
  getAiCapabilities: 'home:ai-capabilities',
  translationKbList: 'ai:translation-kb-list',
  translationKbUpsert: 'ai:translation-kb-upsert',
  translationKbRemove: 'ai:translation-kb-remove',
  translationKbStats: 'ai:translation-kb-stats',
  translateFileStatus: 'ai:translate-file-status',
  translateBuildDictionary: 'ai:translate-build-dictionary',
  translateFileAuto: 'ai:translate-file-auto',
  pickTranslationFile: 'home:pick-translation-file',
  translateSnippet: 'home:translate-snippet',
  marketplaceRate: 'home:marketplace-rate',
} as const

export const PROJECT_CHANNELS = {
  list: 'project:list',
  files: 'project:files',
  create: 'project:create',
  rename: 'project:rename',
  delete: 'project:delete',
  moveFile: 'project:moveFile',
  timeline: 'project:timeline',
} as const
