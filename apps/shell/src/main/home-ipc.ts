import { existsSync, renameSync, rmSync, statSync } from 'node:fs'
import { spawn } from 'node:child_process'
import { basename, dirname, join } from 'node:path'
import { app, BrowserWindow, dialog, ipcMain, shell } from 'electron'
import {
  DEFAULT_SAVE_DIR_KEY,
  GITHUB_REPO_URL,
  isUsableSaveDir,
  showOpenDialogWithMemory,
} from '@genoffice/electron-utils'
import { atomicWriteFile } from './atomic-write'
import { readAppSettings, writeAppSetting } from './app-settings'
import { TabManager } from './tab-manager'
import { handleDroppedFiles, type DroppedFilesDeps } from './dropped-files'
import type { IndexStore } from './file-index/index-store'
import { FileIndexer } from './file-index/indexer'
import { normalizeFileSearchSettings, probeDecision, SearchReranker } from './file-index/rerank'
import {
  capStatPaths,
  matchesExtFamily,
  normalizeRecentQuery,
  pageRecentPaths,
  statPathEntries,
} from './recent-files'
import { isUserVisibleFile, isMoveSource, type FileTargetSources } from './file-targets'
import { renameStaysInApp } from './app-routing'
import { isSameFile, isValidRawRenameName } from './rename-validation'
import {
  createFolder,
  describeRoot,
  listFolder,
  movePathsInto,
  rebasePath,
  renameFolder,
  type FolderErrors,
} from './folder-tree'
import {
  describeExtraRoot,
  withExtraRoot,
  withoutExtraRoot,
  FOLDER_ROOTS_KEY,
} from './folder-roots'
import { createDefaultAppService, execFileRunner } from './default-app'
import {
  createKnowledgeDocument,
  KB_LIST_CAP,
  kbReadAllows,
  listIndexedKnowledgeFiles,
  listKnowledgeFiles,
  normalizeKbScope,
  notIndexedPaths,
  parseKbOnlyPaths,
  parseKbCreateRequest,
  readKnowledgeFilePage,
  searchKnowledgeBase,
} from './knowledge-base'
import { homeCliContext, runHomeCli, type HomeCliDeps } from './home-cli'
import { createFileViaCli } from './mcp/headless-cli'
import { sanitizeFileBase } from './mcp/tools/document-tools'
import { createCliRunner } from './mcp/cli-runner'
import {
  cloudProjectExternalUrl,
  readCloudProjectsStore,
  syncCloudProjects,
} from './cloud-projects'
import {
  defaultSaveDir,
  projectFileRenamed,
  readRecentFiles,
  readStarredFiles,
  readStarredGroupMap,
  readStarredGroups,
  recordRecentFile,
  removeRecentFiles,
  removeStarredFiles,
  setStarredGroup,
  toggleStarredFile,
  uniquePathIn,
} from './app-registry'
import { showErrorDialog } from './error-dialog'
import type {
  FileSearchPage,
  FileSearchQuery,
  FileSearchRerank,
  FileSearchSettings,
  FolderListing,
  FolderRoot,
  MoveConflictPolicy,
  MoveResult,
  NewFileOpts,
  RecentEntry,
  RecentPage,
  RenameResult,
} from '../shared/home-api'
import { HOME_CHANNELS } from '../shared/home-api'

/**
 * The Home file domain's IPC: the paged recents/starred lists, the local file
 * search and its knowledge-base (Ask AI) readers, new-file creation, the
 * rename/move/delete family, the folder tree over the save folder, and the
 * outgoing links. Extracted from index.ts's registerHomeIpc — the handlers
 * that stayed there are the ones pinned to index's own state (the account
 * login pair, duplicate/delete, the settings-cache family and the star
 * prompt).
 */

/** the locale keys this module resolves through the host (index's tm) */
export type HomeFileString =
  | 'dlgOpenTitle'
  | 'filterSupported'
  | 'filterWord'
  | 'filterExcel'
  | 'filterPpt'
  | 'filterPdf'
  | 'filterMarkdown'
  | 'filterHtml'
  | 'errBadArgs'
  | 'errBadName'
  | 'errBadExtension'
  | 'errMissing'
  | 'errExists'
  | 'errRenameFailed'
  | 'dlgAddFolderRoot'
  | 'errFolderRootUnusable'
  | 'dlgPickSaveDir'
  | 'errSaveDirUnusable'

export interface HomeFileIpcHost {
  translate: (key: HomeFileString, params?: Record<string, string | number>) => string
  appSettingsPath: () => string
  getShellWindow: () => BrowserWindow | null
  getTabManager: () => TabManager | null
  // the search index singletons stay owned by index.ts
  ensureFileIndexer: () => FileIndexer | null
  indexStore: () => IndexStore | null
  refreshIndexer: () => void
  /** bring specific first-sight files into the index now (star/open paths) */
  indexNow: (paths: readonly string[]) => Promise<void>
  readFileSearchSettings: () => FileSearchSettings
  ensureSearchReranker: () => SearchReranker | null
  // routing and new-file creation stay owned by index.ts
  openDocumentPath: (path: string) => boolean
  rememberPendingDir: (kind: string, opts?: NewFileOpts) => void
  newDocTab: () => void
  newSheetTab: () => void | Promise<void>
  newSlideTab: () => void
  newMarkdownTab: () => void
  newHtmlTab: () => void
  newPdfTab: () => void | Promise<void>
  // the file-follow plumbing (recents/stars/tabs re-keying) stays in index.ts
  afterFileMoved: (oldPath: string, newPath: string) => void
  afterFolderMoved: (oldDir: string, newDir: string, filesBefore: readonly string[]) => void
  trackedFilesUnder: (dir: string) => string[]
  fileTargetSources: () => FileTargetSources
  statMaybeFile: (path: string) => { isFile: () => boolean } | null
  // the folder-root set is derived from app-settings plus the registry
  extraFolderRoots: () => string[]
  insideAnyRoot: (path: string) => boolean
  isAnyRoot: (path: string) => boolean
  /** the desktop's stored alias paths (desktop-layout.json), for the kb read gate's location whitelist */
  desktopAliases: () => readonly string[]
  ensureFolderWatchers: () => void
  droppedFilesDeps: () => DroppedFilesDeps
  /** the bundled genoffice CLI runner (index's lazy singleton) */
  cliRunner: () => ReturnType<typeof createCliRunner>
  fetchGithubStars: () => Promise<number | null>
  openDialogExtensions: readonly string[]
}

// GenTeam community page opened from the onboarding's second slide. Stable
// short link served by the genoffice.ai site; it 302s to the tokened invite
// link, which stays out of this repo and rotates server-side.
const GENTEAM_URL = 'https://genoffice.ai/join'

// Genspark credit-usage page opened from the account menu's credits row.
// Kept main-side so the renderer never supplies the URL.
const CREDIT_USAGE_URL = 'https://www.genspark.ai/credit-usage'

/** keys the sidebar filter families stand for, mirroring recent-files.ts */
const SEARCH_EXT_FAMILY: Record<string, readonly string[]> = {
  docx: ['docx', 'doc', 'odt', 'ott'],
  xlsx: ['xlsx', 'xlsm', 'xls', 'csv', 'tsv', 'ods', 'ots'],
  pptx: ['pptx', 'ppt', 'odp', 'otp'],
  md: ['md', 'markdown', 'txt', 'json'],
  html: ['html', 'htm'],
}

export function stringPaths(value: unknown): string[] {
  return Array.isArray(value) ? value.filter((p): p is string => typeof p === 'string') : []
}

function statEntries(paths: string[]): RecentEntry[] {
  return statPathEntries(paths, new Set(readStarredFiles()))
}

export function registerHomeFileIpc(host: HomeFileIpcHost): void {
  ipcMain.handle(HOME_CHANNELS.recents, (_event, query: unknown): RecentPage =>
    pageRecentPaths(readRecentFiles(), query, new Set(readStarredFiles())),
  )

  ipcMain.handle(
    HOME_CHANNELS.searchFiles,
    async (_event, raw: unknown): Promise<FileSearchPage> => {
      const query = (raw && typeof raw === 'object' ? raw : {}) as Partial<FileSearchQuery>
      const indexer = host.ensureFileIndexer()
      const store = host.indexStore()
      if (!indexer || !store) {
        return { hits: [], total: 0, index: { indexed: 0, pending: 0, scanning: false } }
      }
      // an open search box is the moment a stale index shows; rescan at most once a minute
      indexer.refreshIfStale(60_000)
      // the query rides out an in-flight incremental scan (a save the watcher
      // just reported) so a just-saved file is findable on the first try; the
      // cap must outlive the watcher's rescan debounce (1.5s) or the wait
      // resolves before the scan even starts
      await indexer.settle(2_000)
      const q = typeof query.q === 'string' ? query.q.trim().slice(0, 200) : ''
      const filter = typeof query.ext === 'string' ? query.ext : ''
      const exts = filter && filter !== 'all' ? (SEARCH_EXT_FAMILY[filter] ?? [filter]) : undefined
      const offset = Number.isFinite(query.offset) ? Math.max(0, Math.floor(query.offset!)) : 0
      const limit = Number.isFinite(query.limit) ? Math.max(0, Math.floor(query.limit!)) : 50
      const starred = new Set(readStarredFiles())
      // the search runs in the store worker; a failure answers the page the
      // caller can already render (an unanswerable store, an empty result)
      const result = q
        ? await store.search(q, { exts, offset, limit }).catch(() => ({
            hits: [],
            total: 0,
          }))
        : { hits: [], total: 0 }
      return {
        hits: result.hits.map((h) => ({ ...h, starred: starred.has(h.path) })),
        total: result.total,
        index: await indexer.progress(),
      }
    },
  )

  // Knowledge-base channel: the AI-side read access to the user's local corpus
  // (docs wires these into its panel; see knowledge-base.ts). The corpus is the
  // app-settings `kbScope`: 'indexed' (default) is every file the local index
  // covers, 'starred' the original hand-picked set. Path arguments are
  // validated against the scope — a corpus file is the only thing a knowledge
  // tool may read. A request carrying `onlyPaths` (the round scope: an
  // explicit desktop selection) replaces the kbScope for that call: search
  // intersects with the selection and reports the paths the index missed as
  // `notIndexed`; read narrows the same way — a corpus file outside the
  // selection is refused, and a selected one the index missed still reads
  // when it sits at a location a desktop selection can genuinely hold
  // (inside a folder root, on the stored alias list, or among the
  // recents/stars). The answer keeps the bare-array shape for unscoped calls
  // so a renderer that predates onlyPaths (docs' panel) reads it unchanged.
  const kbScope = () => normalizeKbScope(readAppSettings(host.appSettingsPath()).kbScope)

  ipcMain.handle('kb:list', async (): Promise<unknown> => {
    const starred = readStarredFiles()
    const store = host.indexStore()
    if (kbScope() === 'starred' || !store) return listKnowledgeFiles(starred)
    // the recent-files listing is one bounded query in the store worker
    const recent = await store
      .recentFiles(KB_LIST_CAP)
      .catch(() => [] as Awaited<ReturnType<IndexStore['recentFiles']>>)
    return listIndexedKnowledgeFiles(starred, recent)
  })

  ipcMain.handle('kb:search', async (_event, raw: unknown): Promise<unknown> => {
    const query = (raw && typeof raw === 'object' ? raw : {}) as {
      q?: unknown
      limit?: unknown
      onlyPaths?: unknown
    }
    const q = typeof query.q === 'string' ? query.q : ''
    const limitRaw = Number(query.limit)
    const limit = Number.isFinite(limitRaw) ? Math.max(1, Math.floor(limitRaw)) : 8
    // invalid scope payloads are refused outright, before anything runs
    const onlyPaths = parseKbOnlyPaths(query.onlyPaths)
    const store = host.indexStore()
    if (!q.trim() || !store) {
      return onlyPaths ? { hits: [], notIndexed: await notIndexedPaths(store, onlyPaths) } : []
    }
    host.ensureFileIndexer()?.refreshIfStale(60_000)
    // an answer is worth a real wait: let an in-flight scan and its
    // extractions land so a question about a just-saved file cites it
    await host.ensureFileIndexer()?.settle(2_000)
    const starred = readStarredFiles()
    // the round scope replaces the kbScope: the selection is the corpus for
    // this call, whatever the settings say
    const scope = onlyPaths ?? (kbScope() === 'starred' ? starred : null)
    const hits = await searchKnowledgeBase(
      async (qq, pool) =>
        store
          .search(qq, { limit: pool })
          .then((r) =>
            r.hits.map((h) => ({
              path: h.path,
              name: h.name,
              ext: h.ext,
              snippetText: (h.snippet ?? []).map((part) => part.text).join(''),
            })),
          )
          .catch(() => []),
      q,
      scope,
      Math.min(limit, 20),
    )
    return onlyPaths ? { hits, notIndexed: await notIndexedPaths(store, onlyPaths) } : hits
  })

  ipcMain.handle('kb:read', async (_event, raw: unknown): Promise<unknown> => {
    const req = (raw && typeof raw === 'object' ? raw : {}) as {
      path?: unknown
      offset?: unknown
      onlyPaths?: unknown
    }
    const path = typeof req.path === 'string' ? req.path : ''
    const offsetRaw = Number(req.offset)
    const offset = Number.isFinite(offsetRaw) ? Math.floor(offsetRaw) : 0
    const onlyPaths = parseKbOnlyPaths(req.onlyPaths)
    // a starred file stays readable in every scope (even with the index
    // momentarily unavailable); the indexed scope adds "in the index"
    const starred = readStarredFiles().includes(path)
    const inCorpus =
      kbScope() === 'starred'
        ? starred
        : starred ||
          ((await host
            .indexStore()
            ?.has(path)
            .catch(() => false)) ??
            false)
    // the round-scope branch's location whitelist, computed only when it can
    // matter (a scope is attached and the corpus does not already cover the
    // path): a selected file is readable off-corpus only where a desktop
    // selection can genuinely come from — inside a folder root, on the stored
    // alias list, or a tracked recent/starred file. An office-named file
    // anywhere else on disk is refused, whatever the scope says.
    const locationAllowed =
      onlyPaths !== null && !inCorpus
        ? host.insideAnyRoot(path) ||
          host.desktopAliases().includes(path) ||
          readRecentFiles().includes(path) ||
          starred
        : false
    const readable = kbReadAllows(path, {
      inCorpus,
      scopePaths: onlyPaths,
      locationAllowed,
      statFile: host.statMaybeFile,
    })
    if (!path || !readable) {
      throw new Error('path is not in the knowledge base')
    }
    return readKnowledgeFilePage(path, offset)
  })

  ipcMain.handle('kb:create-document', async (_event, raw: unknown): Promise<unknown> => {
    const req = parseKbCreateRequest(raw)
    return createKnowledgeDocument(
      {
        dir: defaultSaveDir,
        uniquePath: uniquePathIn,
        sanitizeBase: sanitizeFileBase,
        writeText: (path, text) => atomicWriteFile(path, Buffer.from(text, 'utf8')),
        markdownToDocx: async (markdown, out) => {
          await createFileViaCli(host.cliRunner(), {
            type: 'docx',
            input: { name: 'content.md', content: markdown },
            out,
          })
        },
        recordRecent: recordRecentFile,
      },
      req,
    )
  })

  // Ask AI building files: the bundled CLI, confined to the conversation's
  // work directory, the save folder and the starred files' folders (home-cli.ts)
  const homeCliDeps: HomeCliDeps = {
    runner: host.cliRunner,
    workRoot: () => join(app.getPath('userData'), 'ask-ai-work'),
    outDir: defaultSaveDir,
    sourcePaths: readStarredFiles,
    recordRecent: recordRecentFile,
  }
  ipcMain.handle('kb:cli-context', (_event, sessionId: unknown): unknown =>
    homeCliContext(homeCliDeps, typeof sessionId === 'string' ? sessionId : ''),
  )
  ipcMain.handle('kb:run-cli', async (_event, raw: unknown): Promise<unknown> => {
    const req = (raw && typeof raw === 'object' ? raw : {}) as {
      sessionId?: unknown
      args?: unknown
      files?: unknown
      sources?: unknown
    }
    return runHomeCli(homeCliDeps, typeof req.sessionId === 'string' ? req.sessionId : '', {
      args: Array.isArray(req.args) ? (req.args as string[]) : [],
      files: Array.isArray(req.files) ? (req.files as { name: string; content: string }[]) : [],
      sources: Array.isArray(req.sources) ? (req.sources as string[]) : [],
    })
  })

  ipcMain.handle(
    HOME_CHANNELS.rerankSearch,
    async (_event, raw: unknown): Promise<FileSearchRerank | null> => {
      const settings = host.readFileSearchSettings()
      if (!settings.rerank) return null
      const query = (raw && typeof raw === 'object' ? raw : {}) as { q?: unknown; paths?: unknown }
      const q = typeof query.q === 'string' ? query.q.trim().slice(0, 200) : ''
      const paths = Array.isArray(query.paths)
        ? query.paths.filter((p): p is string => typeof p === 'string').slice(0, 20)
        : []
      if (!q || paths.length < 2 || !host.ensureFileIndexer() || !host.indexStore()) return null
      const reranker = host.ensureSearchReranker()
      if (!reranker) return null
      return reranker.rerank(q, paths, settings)
    },
  )

  ipcMain.handle(HOME_CHANNELS.getFileSearchSettings, (): FileSearchSettings =>
    host.readFileSearchSettings(),
  )

  ipcMain.handle(
    HOME_CHANNELS.setFileSearchSettings,
    (_event, patch: unknown): FileSearchSettings => {
      const current = host.readFileSearchSettings()
      const p = (patch && typeof patch === 'object' ? patch : {}) as Partial<FileSearchSettings>
      const next = normalizeFileSearchSettings({
        ...current,
        ...p,
        keys: { ...current.keys, ...(p.keys ?? {}) },
      })
      writeAppSetting(host.appSettingsPath(), 'fileSearch', next)
      return next
    },
  )

  ipcMain.handle(HOME_CHANNELS.testFileSearchRerank, (_event, input: unknown) =>
    probeDecision(normalizeFileSearchSettings(input)),
  )

  // Starred files sort by mtime, which requires stat-ing them all first; they are hand-picked and few, so this is fine
  ipcMain.handle(HOME_CHANNELS.starred, (_event, query: unknown): RecentPage => {
    const { offset, limit, ext } = normalizeRecentQuery(query)
    // the Starred view's group pills scope the whole query (totals included)
    const raw = (query && typeof query === 'object' ? query : {}) as { group?: unknown }
    const group = typeof raw.group === 'string' && raw.group ? raw.group : undefined
    const groupOf = readStarredGroupMap()
    const paths = readStarredFiles()
    const scoped = group ? paths.filter((p) => groupOf.get(p) === group) : paths
    const all = statEntries(scoped)
      .sort((a, b) => b.mtimeMs - a.mtimeMs)
      .map((entry) => {
        const g = groupOf.get(entry.path)
        return g ? { ...entry, group: g } : entry
      })
    const filtered = ext ? all.filter((entry) => matchesExtFamily(entry.ext, ext)) : all
    return {
      entries: limit === 0 ? [] : filtered.slice(offset, offset + limit),
      total: filtered.length,
      totalAll: all.length,
    }
  })

  ipcMain.handle(HOME_CHANNELS.starredGroups, (): string[] => readStarredGroups())

  ipcMain.handle(HOME_CHANNELS.setStarredGroup, (_event, paths: unknown, group: unknown) => {
    setStarredGroup(stringPaths(paths), typeof group === 'string' && group ? group : null)
  })

  ipcMain.handle(HOME_CHANNELS.statPaths, (_event, paths: unknown): RecentEntry[] =>
    statEntries(capStatPaths(stringPaths(paths))),
  )

  ipcMain.handle(HOME_CHANNELS.toggleStar, (_event, path: unknown) => {
    if (typeof path !== 'string') return
    toggleStarredFile(path)
    // a just-starred file joins the corpus immediately — extraPaths only
    // catches it at the next scan, and a star is exactly the moment the user
    // expects Ask to see the file
    void host.indexNow([path])
  })

  ipcMain.handle(HOME_CHANNELS.openPath, (_event, path: unknown) => {
    if (typeof path !== 'string' || !path || path.length > 4096) return
    host.openDocumentPath(path)
    // the opened file enters the index now (a first sight outside the roots
    // would wait for the next scan); edits the app makes land through the
    // watcher like any other change
    void host.indexNow([path])
  })

  ipcMain.handle(HOME_CHANNELS.browse, async (event) => {
    const win = BrowserWindow.fromWebContents(event.sender) ?? host.getShellWindow()
    if (!win) return
    const result = await showOpenDialogWithMemory(dialog, win, {
      title: host.translate('dlgOpenTitle'),
      filters: [
        { name: host.translate('filterSupported'), extensions: [...host.openDialogExtensions] },
        { name: host.translate('filterWord'), extensions: ['docx', 'doc', 'odt', 'ott'] },
        {
          name: host.translate('filterExcel'),
          extensions: ['xlsx', 'xlsm', 'xls', 'csv', 'tsv', 'ods', 'ots'],
        },
        { name: host.translate('filterPpt'), extensions: ['pptx', 'ppt', 'odp', 'otp'] },
        { name: host.translate('filterPdf'), extensions: ['pdf'] },
        { name: host.translate('filterMarkdown'), extensions: ['md', 'markdown', 'txt', 'json'] },
        { name: host.translate('filterHtml'), extensions: ['html', 'htm'] },
      ],
      properties: ['openFile', 'multiSelections'],
    })
    if (!result.canceled) for (const path of result.filePaths) host.openDocumentPath(path)
  })

  ipcMain.handle(HOME_CHANNELS.newDoc, (_event, opts?: NewFileOpts) => {
    host.rememberPendingDir('doc', opts)
    host.newDocTab()
  })

  ipcMain.handle(HOME_CHANNELS.newSheet, (_event, opts?: NewFileOpts) => {
    host.rememberPendingDir('sheet', opts)
    void host.newSheetTab()
  })

  ipcMain.handle(HOME_CHANNELS.newSlide, (_event, opts?: NewFileOpts) => {
    host.rememberPendingDir('slide', opts)
    host.newSlideTab()
  })

  ipcMain.handle(HOME_CHANNELS.newMarkdown, (_event, opts?: NewFileOpts) => {
    host.rememberPendingDir('markdown', opts)
    host.newMarkdownTab()
  })

  ipcMain.handle(HOME_CHANNELS.newHtml, (_event, opts?: NewFileOpts) => {
    host.rememberPendingDir('html', opts)
    host.newHtmlTab()
  })

  ipcMain.handle(HOME_CHANNELS.newPdf, (_event, opts?: NewFileOpts) => {
    host.rememberPendingDir('pdf', opts)
    void host.newPdfTab()
  })

  ipcMain.handle(HOME_CHANNELS.removeRecent, (_event, paths: unknown) => {
    const list = stringPaths(paths)
    removeRecentFiles(list)
    // an unavailable entry's star must go with it, or the Starred view keeps
    // a dead dimmed row the recents list no longer shows
    removeStarredFiles(list.filter((p) => !existsSync(p)))
  })

  // Bulk unstar from the Starred view: every row there is a favorite, so the
  // selection action is "unstar", not "remove from recents" — removeRecent
  // keeps existing files' stars on purpose, which made the Starred view's
  // bulk removal a no-op for anything still on disk
  ipcMain.handle(HOME_CHANNELS.unstarPaths, (_event, paths: unknown) => {
    removeStarredFiles(stringPaths(paths))
  })

  ipcMain.handle(HOME_CHANNELS.revealPath, (_event, path: unknown) => {
    if (typeof path === 'string' && existsSync(path)) shell.showItemInFolder(path)
  })

  ipcMain.handle(
    HOME_CHANNELS.renameFile,
    (_event, path: unknown, newName: unknown): RenameResult => {
      if (typeof path !== 'string' || typeof newName !== 'string')
        return { ok: false, error: host.translate('errBadArgs') }
      // Validate the raw name before trimming: trimming first would
      // silently turn "report " into "report" and make the
      // trailing-space gate in isValidRenameName unreachable. Reject
      // with the localized gate instead of renaming to a different
      // name than requested.
      if (!isValidRawRenameName(newName)) return { ok: false, error: host.translate('errBadName') }
      if (!renameStaysInApp(path, newName.trim()))
        return { ok: false, error: host.translate('errBadExtension') }
      if (!isUserVisibleFile(path, host.fileTargetSources()))
        return { ok: false, error: host.translate('errBadArgs') }
      const name = newName.trim()
      if (!existsSync(path)) return { ok: false, error: host.translate('errMissing') }
      const target = join(dirname(path), name)
      if (target === path) return { ok: true, path }
      // A case-only rename (Report.pdf -> report.pdf) hits the source itself on
      // case-insensitive filesystems; only a genuinely different file blocks.
      if (existsSync(target) && !isSameFile(path, target)) {
        return { ok: false, error: host.translate('errExists') }
      }
      try {
        renameSync(path, target)
      } catch (err) {
        return {
          ok: false,
          error: err instanceof Error ? err.message : host.translate('errRenameFailed'),
        }
      }
      host.afterFileMoved(path, target)
      return { ok: true, path: target }
    },
  )

  ipcMain.handle(HOME_CHANNELS.openTrash, () => {
    if (process.platform === 'darwin') {
      void shell.openPath(join(app.getPath('home'), '.Trash'))
    } else if (process.platform === 'win32') {
      spawn('explorer.exe', ['shell:RecycleBin'], { detached: true }).unref()
    } else {
      void shell.openPath(join(app.getPath('home'), '.local', 'share', 'Trash', 'files'))
    }
  })

  ipcMain.handle(HOME_CHANNELS.openHelp, () => {
    host.getTabManager()?.openHelpTab()
  })

  // ── the folder tree over the default save folder ──

  const folderErrors = (): FolderErrors => ({
    badArgs: host.translate('errBadArgs'),
    badName: host.translate('errBadName'),
    missing: host.translate('errMissing'),
    exists: host.translate('errExists'),
    failed: host.translate('errRenameFailed'),
  })
  const insideRoot = (path: unknown): path is string =>
    typeof path === 'string' && host.insideAnyRoot(path)
  const isRoot = (path: string) => host.isAnyRoot(path)

  ipcMain.handle(HOME_CHANNELS.folderRoots, (): FolderRoot[] => {
    // describeRoot creates a missing save folder, so its watcher has something to attach to
    const roots = [
      describeRoot(defaultSaveDir()),
      ...host.extraFolderRoots().map(describeExtraRoot),
    ]
    host.ensureFolderWatchers()
    return roots
  })

  // an added folder joins the tree where it is: nothing on disk is created, copied or moved
  const addFolderRoot = (path: string): FolderRoot | null => {
    const extras = withExtraRoot(host.extraFolderRoots(), defaultSaveDir(), path)
    if (!extras) return null
    writeAppSetting(host.appSettingsPath(), FOLDER_ROOTS_KEY, extras)
    host.ensureFolderWatchers()
    host.refreshIndexer()
    return describeExtraRoot(path)
  }

  ipcMain.handle(HOME_CHANNELS.addFolderRoot, async (): Promise<FolderRoot | null> => {
    const result = await showOpenDialogWithMemory(dialog, host.getShellWindow(), {
      title: host.translate('dlgAddFolderRoot'),
      properties: ['openDirectory'],
    })
    const picked = result.filePaths[0]
    if (result.canceled || !picked) return null
    if (!describeExtraRoot(picked).readable) {
      showErrorDialog(host.getShellWindow(), host.translate('errFolderRootUnusable'), picked)
      return null
    }
    return addFolderRoot(picked)
  })

  ipcMain.handle(HOME_CHANNELS.dropFolderRoots, (_event, paths: unknown): FolderRoot[] => {
    const added: FolderRoot[] = []
    const files: string[] = []
    for (const path of stringPaths(paths)) {
      if (!describeExtraRoot(path).readable) {
        files.push(path)
        continue
      }
      const root = addFolderRoot(path)
      if (root) added.push(root)
    }
    if (files.length > 0) handleDroppedFiles(files, host.droppedFilesDeps())
    return added
  })

  ipcMain.handle(HOME_CHANNELS.removeFolderRoot, (_event, path: unknown) => {
    if (typeof path !== 'string') return
    const extras = withoutExtraRoot(host.extraFolderRoots(), path)
    writeAppSetting(host.appSettingsPath(), FOLDER_ROOTS_KEY, extras)
    host.ensureFolderWatchers()
    host.refreshIndexer()
  })

  ipcMain.handle(HOME_CHANNELS.listFolder, (_event, dir: unknown): FolderListing => {
    if (!insideRoot(dir)) return { dir: String(dir), folders: [], files: [] }
    return listFolder(dir, new Set(readStarredFiles()))
  })

  ipcMain.handle(
    HOME_CHANNELS.createFolder,
    (_event, parent: unknown, name: unknown): RenameResult => {
      if (!insideRoot(parent) || typeof name !== 'string')
        return { ok: false, error: host.translate('errBadArgs') }
      return createFolder(parent, name, folderErrors())
    },
  )

  ipcMain.handle(
    HOME_CHANNELS.renameFolder,
    (_event, dir: unknown, newName: unknown): RenameResult => {
      if (!insideRoot(dir) || isRoot(dir) || typeof newName !== 'string')
        return { ok: false, error: host.translate('errBadArgs') }
      const filesBefore = host.trackedFilesUnder(dir)
      const result = renameFolder(dir, newName, folderErrors())
      if (result.ok && result.path && result.path !== dir) {
        host.afterFolderMoved(dir, result.path, filesBefore)
      }
      return result
    },
  )

  ipcMain.handle(
    HOME_CHANNELS.movePaths,
    async (_event, paths: unknown, targetDir: unknown, policy: unknown): Promise<MoveResult> => {
      const list = stringPaths(paths)
      if (!insideRoot(targetDir)) {
        const error = host.translate('errBadArgs')
        return { moved: [], conflicts: [], failed: list.map((path) => ({ path, error })) }
      }
      const conflictPolicy: MoveConflictPolicy =
        policy === 'replace' || policy === 'keepBoth' || policy === 'skip' ? policy : 'ask'
      const isDir = (p: string) => {
        try {
          return statSync(p).isDirectory()
        } catch {
          return false
        }
      }
      // files may come from anywhere the UI can show (the Recent list); folders only from inside the tree, never a root itself
      const moveSources = {
        ...host.fileTargetSources(),
        isDirectory: isDir,
        isAnyRoot: host.isAnyRoot,
      }
      const sources = list.filter((p) => isMoveSource(p, moveSources))
      const dirFiles = new Map(sources.filter(isDir).map((p) => [p, host.trackedFilesUnder(p)]))
      // 'replace' must not destroy data: the displaced target goes to the trash,
      // and everything keyed on its path (recents, stars, chat history) leaves
      // with it so the incoming file does not inherit another document's record
      const displaced: string[] = []
      const result = movePathsInto(sources, targetDir, conflictPolicy, folderErrors(), {
        replaceExisting: (path) => {
          const parked = join(dirname(path), `.genoffice-replaced-${Date.now()}-${basename(path)}`)
          const files = isDir(path) ? host.trackedFilesUnder(path) : [path]
          renameSync(path, parked)
          return {
            commit: () => {
              displaced.push(parked)
              removeRecentFiles(files)
              removeStarredFiles(files)
              for (const file of files) projectFileRenamed(file, rebasePath(file, path, parked))
            },
            rollback: () => renameSync(parked, path),
          }
        },
      })
      for (const parked of displaced) {
        try {
          await shell.trashItem(parked)
        } catch {
          rmSync(parked, { recursive: true, force: true })
        }
      }
      for (const { from, to } of result.moved) {
        const files = dirFiles.get(from)
        if (files) host.afterFolderMoved(from, to, files)
        else host.afterFileMoved(from, to)
      }
      return result
    },
  )

  ipcMain.handle(HOME_CHANNELS.deleteFolder, async (_event, dir: unknown) => {
    if (!insideRoot(dir) || isRoot(dir)) return
    const files = host.trackedFilesUnder(dir)
    try {
      await shell.trashItem(dir)
    } catch {
      return
    }
    removeRecentFiles(files)
    removeStarredFiles(files)
  })

  ipcMain.handle(HOME_CHANNELS.getDefaultSaveDir, (): string => defaultSaveDir())

  const defaultApp = createDefaultAppService({
    platform: process.platform,
    packaged: app.isPackaged,
    exePath: app.getPath('exe'),
    run: execFileRunner,
    openExternal: (url) => shell.openExternal(url),
  })
  ipcMain.handle(HOME_CHANNELS.getDefaultAppStatus, () => defaultApp.status())
  ipcMain.handle(HOME_CHANNELS.setDefaultApp, () => defaultApp.set())

  ipcMain.handle(HOME_CHANNELS.pickDefaultSaveDir, async (): Promise<string | null> => {
    const result = await showOpenDialogWithMemory(dialog, host.getShellWindow(), {
      title: host.translate('dlgPickSaveDir'),
      defaultPath: defaultSaveDir(),
      properties: ['openDirectory', 'createDirectory'],
    })
    const picked = result.filePaths[0]
    if (result.canceled || !picked) return null
    if (!isUsableSaveDir(picked)) {
      showErrorDialog(host.getShellWindow(), host.translate('errSaveDirUnusable'), picked)
      return null
    }
    writeAppSetting(host.appSettingsPath(), DEFAULT_SAVE_DIR_KEY, picked)
    host.ensureFolderWatchers()
    return picked
  })

  ipcMain.handle(HOME_CHANNELS.openGenTeam, () => {
    shell.openExternal(GENTEAM_URL).catch(() => {
      // no browser handler available; nothing actionable for the user here
    })
  })

  ipcMain.handle(HOME_CHANNELS.openCreditUsage, () => {
    shell.openExternal(CREDIT_USAGE_URL).catch(() => {
      // no browser handler available; nothing actionable for the user here
    })
  })

  ipcMain.handle(HOME_CHANNELS.openGitHubRepo, () => {
    shell.openExternal(GITHUB_REPO_URL).catch(() => {
      // no browser handler available; nothing actionable for the user here
    })
  })

  ipcMain.handle(HOME_CHANNELS.githubStars, () => host.fetchGithubStars())

  const cloudProjectsStorePath = () => join(app.getPath('userData'), 'cloud-projects.json')

  ipcMain.handle(HOME_CHANNELS.cloudProjectsCached, () =>
    readCloudProjectsStore(cloudProjectsStorePath()),
  )

  ipcMain.handle(HOME_CHANNELS.cloudProjects, () => syncCloudProjects(cloudProjectsStorePath()))

  ipcMain.handle(HOME_CHANNELS.openCloudProject, (_event, projectUrl: unknown) => {
    const url = cloudProjectExternalUrl(projectUrl)
    if (url) void shell.openExternal(url)
  })
}
