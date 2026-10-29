import { existsSync, readFileSync } from 'node:fs'
import { basename, extname, join } from 'node:path'
import { app, ipcMain } from 'electron'
import { atomicCopyFile } from './atomic-write'
import { defaultSaveDir, recordRecentFile, uniquePathIn } from '../../../docs/src/main/docs-main'
import { shellTabManager } from './tab-manager'
import { HOME_CHANNELS, type TemplateInfo } from '../shared/home-api'

/** one shipped template as templates.json describes it: TemplateInfo + its file */
interface TemplateEntry extends TemplateInfo {
  /** plain file name inside the templates dir — the manifest never carries a path */
  file: string
}

export interface TemplatesIpcDeps {
  /** where the shipped templates and their manifest live; tests inject a fixture dir */
  templatesDir?: () => string
  /** the user's default save folder the copy lands in */
  defaultSaveDir?: () => string
  /** opens the copied document in the right app tab; the default routes by kind */
  openTemplate?: (filePath: string, kind: TemplateInfo['kind']) => void
}

/** dev ships the templates from the app's own build tree; the packaged app reads
 *  them from resources/templates (electron-builder maps build/templates there) */
function defaultTemplatesDir(): string {
  return app.isPackaged
    ? join(process.resourcesPath, 'templates')
    : join(app.getAppPath(), 'build', 'templates')
}

const TEMPLATE_KINDS: ReadonlySet<string> = new Set(['docx', 'xlsx', 'pptx', 'md'])

function isTemplateEntry(value: unknown): value is TemplateEntry {
  if (value === null || typeof value !== 'object' || Array.isArray(value)) return false
  const entry = value as Record<string, unknown>
  return (
    typeof entry.id === 'string' &&
    entry.id !== '' &&
    typeof entry.title === 'string' &&
    typeof entry.description === 'string' &&
    typeof entry.kind === 'string' &&
    TEMPLATE_KINDS.has(entry.kind) &&
    typeof entry.category === 'string' &&
    typeof entry.file === 'string' &&
    // a plain file name in the templates dir, never a path
    entry.file !== '' &&
    !/[\\/]/.test(entry.file) &&
    basename(entry.file) === entry.file
  )
}

/**
 * The shipped manifest, shape-checked. A missing or corrupt file is the empty
 * list — the template center is an offer, never a reason Home fails to open —
 * and entries that lost their shape are dropped individually, like the todos
 * store's invariant.
 */
function readTemplateManifest(dir: string): TemplateEntry[] {
  let text: string
  try {
    text = readFileSync(join(dir, 'templates.json'), 'utf8')
  } catch {
    return []
  }
  let parsed: unknown
  try {
    parsed = JSON.parse(text)
  } catch {
    return []
  }
  if (!Array.isArray(parsed)) return []
  return parsed.filter(isTemplateEntry)
}

/** the file the copy becomes: the template's title plus the kind's extension,
 *  sanitized the way the document tools sanitize a generated name */
function templateOutputName(entry: TemplateEntry): string {
  const ext = extname(entry.file)
  const base = entry.title
    // eslint-disable-next-line no-control-regex -- control characters are rejected on purpose
    .replace(/[/\\:*?"<>|\u0000-\u001f]/g, '_')
    .trim()
    .slice(0, 80)
    .trim()
  return `${base || basename(entry.file, ext)}${ext}`
}

/**
 * Opens the copied document in the right app tab, the same routes
 * routeDocumentPath takes in main/index.ts, and records it as a recent file.
 * With no shell window left (only detached editors) there is no strip to open
 * a tab in: the file is already saved, so this only warns.
 */
function openTemplateDocument(filePath: string, kind: TemplateInfo['kind']): void {
  recordRecentFile(filePath)
  const tabs = shellTabManager()
  if (!tabs) {
    console.warn(`[shell] template document ${filePath} saved, but the shell window is gone`)
    return
  }
  switch (kind) {
    case 'docx':
      tabs.openDocsTab(filePath)
      break
    case 'xlsx':
      tabs.openSheetsTab(filePath)
      break
    case 'pptx':
      tabs.openSlidesTab(filePath)
      break
    case 'md':
      tabs.openMarkdownTab(filePath)
      break
  }
}

/**
 * The template center's IPC. The renderer sends only a template id, never a
 * path: the id is looked up in the shipped manifest, the file name comes from
 * that entry, and the copy lands under a unique non-clobbering name in the
 * user's default save folder before the matching app tab opens it. Registered
 * from index.ts like the other feature IPC modules.
 */
export function registerTemplatesIpc(deps: TemplatesIpcDeps = {}): void {
  const templatesDir = deps.templatesDir ?? defaultTemplatesDir
  const saveDir = deps.defaultSaveDir ?? defaultSaveDir
  const openTemplate = deps.openTemplate ?? openTemplateDocument

  ipcMain.handle(HOME_CHANNELS.templatesList, (): TemplateInfo[] =>
    readTemplateManifest(templatesDir())
      // only offer what is actually shipped: a pruned install must not surface
      // cards whose file would fail to copy
      .filter((entry) => existsSync(join(templatesDir(), entry.file)))
      .map(({ file: _file, ...info }) => info),
  )

  ipcMain.handle(HOME_CHANNELS.newFromTemplate, async (_event, id: unknown): Promise<void> => {
    if (typeof id !== 'string' || id === '') {
      throw new Error('newFromTemplate needs a template id from the listing')
    }
    const dir = templatesDir()
    const entry = readTemplateManifest(dir).find((candidate) => candidate.id === id)
    if (!entry) throw new Error(`no template with id "${id}"`)
    const source = join(dir, entry.file)
    if (!existsSync(source)) throw new Error(`template ${entry.id} is missing its file`)
    const target = uniquePathIn(saveDir(), templateOutputName(entry))
    await atomicCopyFile(source, target)
    openTemplate(target, entry.kind)
  })
}
