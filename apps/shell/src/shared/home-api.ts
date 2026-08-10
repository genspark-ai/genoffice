import type { UpdateChannel } from './update-api'

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
  | 'nl'
  | 'ms'
  | 'he'
  | 'hi'
  | 'zh-TW'

/** UI theme preference */
export type UiTheme = 'light' | 'dark' | 'system'

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
  /** stat a specific set of paths (project view); missing files are skipped */
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
  /** mark the first-run onboarding as done so it never shows again */
  setOnboardingSeen(): Promise<void>
  /** current UI theme preference (persisted in userData/app-settings.json) */
  getTheme(): Promise<UiTheme>
  /** switch + persist the UI theme; broadcasts 'app:theme-changed' to all web contents */
  setTheme(theme: UiTheme): Promise<void>
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
  /** locally stored full cloud project list (instant; null when no store or logged out) */
  cloudProjectsCached(): Promise<CloudProjectsSnapshot | null>
  /** sync the full list from Genspark and return it (1 request when nothing changed); null when the sync failed */
  cloudProjectsSync(): Promise<CloudProjectsSnapshot | null>
  /** open a cloud project (relative '/agents?id=...' URL) in the default browser */
  openCloudProject(projectUrl: string): Promise<void>
}

export type CloudProjectKind = 'docs' | 'sheets' | 'slides'

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
  getDefaultSaveDir: 'home:get-default-save-dir',
  pickDefaultSaveDir: 'home:pick-default-save-dir',
  openGenTeam: 'home:open-genteam',
  openCreditUsage: 'home:open-credit-usage',
  cloudProjects: 'home:cloud-projects',
  cloudProjectsCached: 'home:cloud-projects-cached',
  openCloudProject: 'home:open-cloud-project',
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
