import { existsSync, mkdirSync, readdirSync, statSync } from 'node:fs'
import { basename, extname, isAbsolute, join } from 'node:path'
import type { ControlReply, DesktopControlRequest } from '@genoffice/cli/control-protocol'
import type { DesktopPrefs } from '../shared/home-api'
import { atomicCopyFile } from './atomic-write'
import {
  customWallpaperName,
  IMAGE_EXTENSIONS,
  normalizePrefs,
  readDesktopLayout,
  readDesktopPins,
  readDesktopPrefs,
  readWallpaperManifest,
  writeDesktopPrefs,
} from './desktop-ipc'
import { addTodo, readTodos } from './todos-ipc'

/**
 * The `desktop` control domain: the CLI/MCP face of the Home desktop's own
 * state. Everything reads the same files the renderer IPC serves
 * (desktop-layout.json, desktop-prefs.json, todos.json), so a read is always
 * live, and every write goes through those stores' own normalization and
 * atomic write — a CLI edit is indistinguishable from a UI edit on disk.
 *
 * Deps are injected so the module never imports Electron; index.ts wires the
 * real directories and the live window enumeration, tests pass fakes (the
 * desktop-ipc.test.ts style).
 */

export interface DesktopWindowInfo {
  id: string
  kind: string
  title: string
  /** absolute path behind the tab; absent while the document is untitled */
  filePath?: string
  active: boolean
  dirty?: boolean
  /** lives in its own OS window rather than the shell window */
  detached?: boolean
  /** renderer released to free memory; reopens its file when shown */
  dormant?: boolean
  /** excluded from sleeping (tab context menu) */
  pinned?: boolean
}

export interface DesktopControlDeps {
  userDataDir: () => string
  /** dir holding the shipped wallpapers and wallpapers.json */
  wallpapersDir: () => string
  /** dir the picked/copied custom wallpapers live in */
  customWallpapersDir: () => string
  /** the shell window's tabs plus any detached editor windows, as of now */
  listWindows: () => DesktopWindowInfo[] | Promise<DesktopWindowInfo[]>
  /**
   * Called after a write changed the desktop prefs, so a running Home view can
   * re-pull instead of showing yesterday's wallpaper until its next mount.
   */
  notifyDesktopPrefsChanged?: (prefs: DesktopPrefs) => void
}

/** an action failure with the control protocol's stable reason vocabulary */
export class DesktopControlError extends Error {
  constructor(
    readonly reason: 'invalid_argument' | 'file_not_found' | 'unsupported',
    message: string,
    readonly detail?: Record<string, unknown>,
  ) {
    super(message)
    this.name = 'DesktopControlError'
  }
}

/** one desktop control request against live state; throws DesktopControlError */
export async function runDesktopAction(
  deps: DesktopControlDeps,
  request: DesktopControlRequest,
): Promise<Record<string, unknown>> {
  switch (request.action) {
    case 'get-state': {
      const layout = readDesktopLayout(deps.userDataDir())
      return {
        prefs: readDesktopPrefs(deps.userDataDir()) as unknown as Record<string, unknown>,
        layout: {
          positioned: Object.keys(layout.positions).length,
          positions: layout.positions,
          // the alias paths the person dragged onto the desktop (agent-visible
          // so the agent can reason about the surface it is asked to work on)
          aliases: layout.aliases,
        },
        // the pinned document widgets, as desktop-pins.json stores them
        pins: readDesktopPins(deps.userDataDir()),
        todos: readTodos(deps.userDataDir()),
        windows: await deps.listWindows(),
      }
    }
    case 'todo-list':
      return { todos: readTodos(deps.userDataDir()) }
    case 'todo-add': {
      const text = request.text.trim()
      if (!text) throw new DesktopControlError('invalid_argument', 'todo text must not be empty')
      const todos = await addTodo(deps.userDataDir(), text)
      return { added: text, todos }
    }
    case 'windows-list':
      return { windows: await deps.listWindows() }
    case 'set-wallpaper':
      return setWallpaper(deps, request.wallpaper)
  }
}

async function setWallpaper(
  deps: DesktopControlDeps,
  input: string,
): Promise<Record<string, unknown>> {
  const id = await resolveWallpaperId(deps, input)
  // the renderer save path's own invariant, so the file the CLI writes is
  // exactly the shape a desktop-prefs-save would have produced
  const prefs = normalizePrefs({ ...readDesktopPrefs(deps.userDataDir()), wallpaper: id })
  // the same atomic write the renderer's save uses, then the nudge that lets a
  // live Home view pick the change up
  await writeDesktopPrefs(deps.userDataDir(), prefs)
  deps.notifyDesktopPrefsChanged?.(prefs)
  return { wallpaper: prefs.wallpaper }
}

/**
 * What a caller may name: 'default' (the live aurora), a shipped wallpaper
 * ('builtin:teal-deep' or a bare 'teal-deep'), an already-imported custom
 * image ('custom:<name>'), or an absolute path to an image file, which is
 * copied into the custom dir first — the prefs never point outside the
 * stores the renderer reads.
 */
async function resolveWallpaperId(deps: DesktopControlDeps, input: string): Promise<string> {
  const value = input.trim()
  if (value === 'default') return 'default'
  const builtinId = value.startsWith('builtin:') ? value.slice('builtin:'.length) : value
  if (!builtinId.includes('/') && !builtinId.includes('\\') && !isAbsolute(value)) {
    const entry = readWallpaperManifest(deps.wallpapersDir()).find((e) => e.id === builtinId)
    if (entry) return `builtin:${entry.id}`
  }
  if (value.startsWith('custom:')) {
    const name = value.slice('custom:'.length)
    // the read side's own rule (desktop-ipc's wallpaperImage handler): a plain
    // file name that already sits in the custom dir. Anything path-like — a
    // separator, a traversal, an absolute fragment — is refused before any
    // disk touch, so the write can never poison the prefs with an id the read
    // side would reject (and existsSync can never be used as an existence
    // oracle for paths outside the custom dir).
    if (name === '' || name !== basename(name) || /[\\/]/.test(name)) {
      throw new DesktopControlError(
        'invalid_argument',
        `invalid custom wallpaper name: ${name || '(empty)'}`,
      )
    }
    let present: string[]
    try {
      present = readdirSync(deps.customWallpapersDir())
    } catch {
      present = []
    }
    if (!present.includes(name)) {
      throw new DesktopControlError('invalid_argument', `unknown custom wallpaper: ${name}`, {
        suggestion: 'pass an absolute image path to import one',
      })
    }
    return value
  }
  if (!isAbsolute(value)) {
    throw new DesktopControlError('invalid_argument', `unknown wallpaper: ${value}`, {
      available: readWallpaperManifest(deps.wallpapersDir()).map((e) => e.id),
    })
  }
  // an absolute path: a real image file that exists right now
  let stat
  try {
    stat = statSync(value)
  } catch {
    throw new DesktopControlError('file_not_found', `wallpaper image not found: ${value}`)
  }
  if (!stat.isFile()) {
    throw new DesktopControlError('invalid_argument', `not a file: ${value}`)
  }
  const ext = extname(value).slice(1).toLowerCase()
  if (!IMAGE_EXTENSIONS.includes(ext)) {
    throw new DesktopControlError(
      'invalid_argument',
      `not an image (${IMAGE_EXTENSIONS.join('/')}): ${value}`,
    )
  }
  const dir = deps.customWallpapersDir()
  mkdirSync(dir, { recursive: true })
  // the timestamped name with a bumping suffix, exactly like the picker's import
  let attempt = 0
  let name = customWallpaperName(value, attempt)
  while (existsSync(join(dir, name))) {
    attempt += 1
    name = customWallpaperName(value, attempt)
  }
  await atomicCopyFile(value, join(dir, name))
  return `custom:${name}`
}

/**
 * Adapter for the control socket: a DesktopControlError becomes the
 * protocol's structured error; anything unexpected is app_unavailable.
 */
export function desktopControlHandler(
  deps: DesktopControlDeps,
): (request: DesktopControlRequest) => Promise<ControlReply> {
  return async (request) => {
    try {
      return { ok: true, result: await runDesktopAction(deps, request) }
    } catch (error) {
      if (error instanceof DesktopControlError) {
        return {
          ok: false,
          error: {
            reason: error.reason,
            message: error.message,
            ...(error.detail ? { detail: error.detail } : {}),
          },
        }
      }
      throw error
    }
  }
}
