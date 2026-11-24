/// Drag & drop a local document file anywhere in the app window to open it.
///
/// Renderers share one main process (the shell hosts every editor view), so the
/// bridge only has to resolve dropped files to paths and send them over ONE
/// channel; the shell routes each path through its normal open pipeline. The
/// installer runs inside preloads, where `webUtils.getPathForFile` is callable
/// and DOM listeners see real OS drops — paths never cross contextBridge.
///
/// Ownership contract with page-level drop zones: a zone that handles its own
/// file drops (AI attachment panels, slides image insert) cancels the event
/// before it bubbles here (`defaultPrevented`), and this bridge stays out of
/// the way. Text/media drags without OS files are untouched.
import { ipcRenderer, webUtils } from 'electron'

export const DROP_OPEN_CHANNEL = 'app:open-dropped-files'

/** Fire-and-forget record of every OS file drop / clipboard file paste the
 *  preload resolves — the grant-ledger feed for the attachment read channels
 *  (see path-grants). Paths only, no open semantics. */
export const USER_FILE_PATHS_SEEN_CHANNEL = 'app:user-file-paths-seen'

/** Sanitize a raw seen-paths payload: strings only, trimmed, deduped, capped.
 *  No extension filtering — the ledger records what the user physically
 *  handed over (attachments include images), the read handlers keep their
 *  own extension rules. */
export function seenFilePaths(raw: unknown): string[] {
  if (!Array.isArray(raw)) return []
  const seen = new Set<string>()
  const out: string[] = []
  for (const entry of raw as unknown[]) {
    if (typeof entry !== 'string') continue
    const path = entry.trim()
    if (!path || path.length > MAX_DROP_PATH_CHARS || seen.has(path)) continue
    seen.add(path)
    out.push(path)
    if (out.length >= MAX_RESOLVED_DROP_PATHS) break
  }
  return out
}

/** Extensions routed by apps/shell routeDocumentPath — keep in sync there and
 *  with OPEN_DIALOG_EXTENSIONS / OPEN_LOCAL_EXTENSIONS on the home screen. */
export const OPENABLE_DOC_RE =
  /\.(docx|odt|ott|xlsx|xlsm|xls|csv|tsv|ods|ots|pptx|ppt|odp|otp|pdf|md|markdown|html|htm)$/i

/** Recognized-but-unsupported formats: kept in the sent payload so the shell
 *  can show its "not supported" dialog instead of dropping them silently.
 *  Mirrors UNSUPPORTED_DOC_RE in apps/shell/src/main/index.ts. */
export const KNOWN_UNSUPPORTED_DOC_RE = /\.(doc|rtf|pps|xlsb|pages|key|numbers)$/i

/** upper bound on how many files one drop may ask to open */
const MAX_DROPPED_FILES = 20

/** the resolver signature webUtils.getPathForFile satisfies; injectable for tests */
type PathResolver = (file: File) => string

/** Resolve one dropped file, tolerating resolver failures (see droppableFilePaths). */
function tryResolvePath(file: File, getPathForFile: PathResolver): string {
  try {
    return getPathForFile(file).trim()
  } catch {
    return ''
  }
}

/** Early bound for path resolution: downstream caps opens at 20, but resolving
 *  10k dropped files first still costs. Overlong paths are skipped outright. */
export const MAX_RESOLVED_DROP_PATHS = 100
export const MAX_DROP_PATH_CHARS = 4096

/**
 * Resolve an event's dropped files to local paths. Returns null when the drag
 * carries no OS files at all (internal text/element drags), or [] when it does
 * but none resolve (directories, virtual entries) — both mean "not ours".
 */
export function droppableFilePaths(
  ev: Pick<DragEvent, 'dataTransfer'>,
  getPathForFile: PathResolver,
): string[] | null {
  const transfer = ev.dataTransfer
  if (!transfer || !transfer.types.includes('Files')) return null
  const paths: string[] = []
  for (const file of Array.from(transfer.files)) {
    if (paths.length >= MAX_RESOLVED_DROP_PATHS) break
    // A throwing resolver (e.g. a sandboxed entry Electron cannot map) must
    // not abort the whole drop: skip that file like a virtual entry.
    // Non-empty guard covers virtual entries (e.g. page-referenced blobs).
    const path = tryResolvePath(file, getPathForFile)
    if (!path || path.length > MAX_DROP_PATH_CHARS) continue
    paths.push(path)
  }
  return paths
}

/**
 * Sanitize + classify a raw IPC payload: strings only, trimmed, deduped,
 * capped, then split into directly-openable paths and known-unsupported
 * extensions (unique, first-seen order). Anything unrecognized (.png, .zip,
 * nonexistent junk from a hostile sender) falls out silently.
 */
export function partitionDropPayload(raw: unknown): {
  supported: string[]
  unsupportedExts: string[]
} {
  const supported: string[] = []
  const unsupportedExts: string[] = []
  if (!Array.isArray(raw)) return { supported, unsupportedExts }
  const seen = new Set<string>()
  for (const entry of raw as unknown[]) {
    if (typeof entry !== 'string') continue
    const path = entry.trim()
    if (!path || seen.has(path)) continue
    seen.add(path)
    if (OPENABLE_DOC_RE.test(path)) {
      // keep scanning for unsupported entries even after hitting the open cap
      if (supported.length < MAX_DROPPED_FILES) supported.push(path)
    } else if (KNOWN_UNSUPPORTED_DOC_RE.test(path)) {
      const ext = path.slice(path.lastIndexOf('.') + 1).toLowerCase()
      if (!unsupportedExts.includes(ext)) unsupportedExts.push(ext)
    }
  }
  return { supported, unsupportedExts }
}

/** Symbol.for keeps repeated installs idempotent when several bundled copies of
 *  this module end up in one process (mirrors navigation-guard). */
const INSTALLED = Symbol.for('genoffice.drop-open-installed')

/**
 * Preload-side hook: makes `drop` fire for document drags anywhere in the
 * page and forwards recognized files to the shell's router. Intentionally not
 * exposed through contextBridge — page code cannot spoof these events with
 * arbitrary paths because payload construction happens here in the preload
 * world only.
 */
export function installDropOpenBridge(): void {
  const holder = globalThis as Record<symbol, boolean | undefined>
  if (typeof window === 'undefined' || holder[INSTALLED]) return
  holder[INSTALLED] = true

  // Without a canceled dragover Chromium never fires `drop`; canceling here is
  // what lets a document-drag land anywhere that isn't already a drop zone.
  window.addEventListener('dragover', (ev: DragEvent) => {
    if (ev.defaultPrevented) return
    if (!ev.dataTransfer?.types.includes('Files')) return
    ev.preventDefault()
  })

  // Every OS file drop lands in the main process's grant ledger, even when a
  // page zone claims the drop (AI attachments, image insert) and the open
  // bridge stays out of the way: the attachment panels resolve dropped files
  // to paths themselves and hand them to the files-add channels, and those
  // reads verify against the ledger. Spoof-proof for the same reason the
  // open bridge is: payload construction happens here in the preload world,
  // and webUtils.getPathForFile returns '' for page-constructed File objects.
  window.addEventListener(
    'drop',
    (ev: DragEvent) => {
      const paths = droppableFilePaths(ev, (file) => webUtils.getPathForFile(file))
      if (paths && paths.length > 0) {
        ipcRenderer.send(USER_FILE_PATHS_SEEN_CHANNEL, paths.slice(0, MAX_RESOLVED_DROP_PATHS))
      }
    },
    // capture: page zones stopPropagation on their own handlers, and the
    // record must not depend on where the drop landed
    true,
  )

  // File pastes from the OS clipboard (copy a file in Finder, paste into an
  // attachment panel) carry real local paths the same way.
  window.addEventListener(
    'paste',
    (ev: ClipboardEvent) => {
      const files = ev.clipboardData?.files
      if (!files || files.length === 0) return
      const paths: string[] = []
      for (const file of Array.from(files)) {
        if (paths.length >= MAX_RESOLVED_DROP_PATHS) break
        const path = tryResolvePath(file, (f) => webUtils.getPathForFile(f))
        if (path && path.length <= MAX_DROP_PATH_CHARS) paths.push(path)
      }
      if (paths.length > 0) {
        ipcRenderer.send(USER_FILE_PATHS_SEEN_CHANNEL, paths.slice(0, MAX_RESOLVED_DROP_PATHS))
      }
    },
    true,
  )

  window.addEventListener('drop', (ev: DragEvent) => {
    // first: something in the page already claimed this drop (image insert,
    // AI attachments...) — never second-guess it
    if (ev.defaultPrevented) return
    const paths = droppableFilePaths(ev, (file) => webUtils.getPathForFile(file))
    if (!paths) return
    // only recognized documents ride along; stray images/folders are swallowed
    // here so they neither navigate the page nor produce open attempts
    const payload = paths.filter((p) => OPENABLE_DOC_RE.test(p) || KNOWN_UNSUPPORTED_DOC_RE.test(p))
    ev.preventDefault()
    if (payload.length === 0) return
    ipcRenderer.send(DROP_OPEN_CHANNEL, payload.slice(0, MAX_DROPPED_FILES))
  })
}
