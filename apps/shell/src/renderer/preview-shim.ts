/**
 * Preview-only shim for the KĀRYA shell renderer when it is served as a
 * plain web page (apps/shell/vite.preview.config.ts) instead of inside
 * Electron, where the preload script would expose `window.aiOffice` through
 * contextBridge. Loaded only by the preview dev server — nothing imports it
 * from the app entry graphs, so it never ships in production bundles.
 *
 * Keeps the Home screen alive with sample data and no-op actions so the
 * branded UI (logo, sidebar, recents list, About dialog) is watchable in a
 * browser preview.
 */
import type { HomeApi, RecentEntry, UiLanguage, UiTheme } from '../shared/home-api'
import type { TabsApi, TabSummary } from '../shared/tabs-api'

const now = Date.now()
const SAMPLE: RecentEntry[] = [
  {
    path: 'C:\\Users\\Demo\\Documents\\KĀRYA\\Q3 roadmap.docx',
    name: 'Q3 roadmap.docx',
    ext: 'docx',
    mtimeMs: now - 45 * 60_000,
    sizeBytes: 24_576,
    starred: true,
  },
  {
    path: 'C:\\Users\\Demo\\Documents\\KĀRYA\\Budget 2026.xlsx',
    name: 'Budget 2026.xlsx',
    ext: 'xlsx',
    mtimeMs: now - 3 * 3_600_000,
    sizeBytes: 98_304,
    starred: false,
  },
  {
    path: 'C:\\Users\\Demo\\Documents\\KĀRYA\\Product launch.pptx',
    name: 'Product launch.pptx',
    ext: 'pptx',
    mtimeMs: now - 26 * 3_600_000,
    sizeBytes: 3_355_443,
    starred: true,
  },
  {
    path: 'C:\\Users\\Demo\\Documents\\KĀRYA\\Annual report.pdf',
    name: 'Annual report.pdf',
    ext: 'pdf',
    mtimeMs: now - 3 * 86_400_000,
    sizeBytes: 1_152_921,
    starred: false,
  },
  {
    path: 'C:\\Users\\Demo\\Documents\\KĀRYA\\Meeting notes.md',
    name: 'Meeting notes.md',
    ext: 'md',
    mtimeMs: now - 10 * 86_400_000,
    sizeBytes: 8_192,
    starred: false,
  },
]

const noop = async (): Promise<void> => {}

const homeApi: HomeApi = {
  recents: async () => ({
    entries: SAMPLE,
    total: SAMPLE.length,
    totalAll: SAMPLE.length,
  }),
  starred: async () => {
    const entries = SAMPLE.filter((e) => e.starred)
    return { entries, total: entries.length, totalAll: SAMPLE.length }
  },
  statPaths: async (paths) => SAMPLE.filter((e) => paths.includes(e.path)),
  toggleStar: noop,
  openPath: noop,
  browse: noop,
  newDoc: noop,
  newSheet: noop,
  newSlide: noop,
  newMarkdown: noop,
  removeRecent: noop,
  revealPath: noop,
  renameFile: async () => ({ ok: true }),
  duplicateFile: noop,
  deleteFiles: noop,
  openTrash: noop,
  getLanguage: async (): Promise<UiLanguage> => 'en',
  setLanguage: noop,
  getUpdateChannel: async () => 'stable' as const,
  setUpdateChannel: noop,
  accountStatus: async () => ({ loggedIn: false }),
  accountLogin: async () => false,
  onAccountLogin: () => () => {},
  openLoginUrl: noop,
  accountLogout: noop,
  getAppVersion: async () => '0.6.0',
  onboardingSeen: async () => true,
  setOnboardingSeen: noop,
  getTheme: async (): Promise<UiTheme> => 'light',
  setTheme: noop,
  getDefaultSaveDir: async () => 'C:\\Users\\Demo\\Documents\\KĀRYA',
  pickDefaultSaveDir: async () => null,
  onThemeChanged: () => () => {},
  openGenTeam: noop,
  openCreditUsage: noop,
  openGitHubRepo: noop,
  githubStars: async () => 1234,
  starPromptShouldShow: async () => ({ show: false, docOpens: 3 }),
  starPromptAction: noop,
  cloudProjectsCached: async () => null,
  cloudProjectsSync: async () => null,
  openCloudProject: noop,
}

// The shell entry always mounts the tab strip (AppFrame → TabBar), so the
// tabs bridge needs a stub too. A single pinned Home tab keeps the strip
// minimal and the content area on the Home screen.
const HOME_TAB: TabSummary = {
  id: 'home',
  kind: 'home',
  title: 'Home',
  closable: false,
  active: true,
}

const tabsApi: TabsApi = {
  list: async () => [HOME_TAB],
  activate: noop,
  close: noop,
  showMenu: noop,
  showNewMenu: noop,
  reorder: noop,
  onChanged: () => () => {},
  notifyChromePressed: () => {},
}

window.aiOffice = homeApi
window.aiOfficeTabs = tabsApi
