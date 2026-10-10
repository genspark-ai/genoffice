/**
 * Interface scale (issue #1913) — the user's "make the UI bigger" preference.
 *
 * ## Why a CSS variable and not `webContents.setZoomFactor()`
 *
 * The issue suggests scaling the window and dividing each editor's own zoom by
 * the factor to compensate. That compensation is not optional, and it is the
 * whole difficulty: a window zoom changes the CSS-pixel-to-device-pixel mapping
 * for the *entire* document, so a canvas the app already scaled to 150% lands at
 * 150% x 150% the moment the chrome is scaled too. The user zooms the document
 * and the interface follows, which is the complaint restated.
 *
 * So the scale lands on one custom property instead. The main process writes it
 * into every live renderer; each chrome container opts in with
 * `zoom: var(--ui-scale)`, and document canvases — which are siblings of those
 * containers, never children — never see it. The two mechanisms cannot compose
 * because they never overlap.
 *
 * ## What this costs, honestly
 *
 * `zoom` on a container is not the same as re-sizing the chrome in `rem`. It
 * scales the pixels, so a scaled-up ribbon needs proportionally more room and
 * the window's minimum size should scale with it. That is a known rough edge,
 * not something this hides.
 */
import { webContents } from 'electron'

import { readAppSettings, writeAppSetting } from './app-settings'

/** `app-settings.json` key; absent or 1 means "no scaling" */
export const UI_SCALE_KEY = 'uiScale'

/**
 * Presets offered by the View menu.
 *
 * Stops at 150%, short of the 200% #1913 asked for. `zoom` scales pixels and
 * does not reflow, so past 150% the ribbon and the status bar start clipping
 * their right-hand edge — see the screenshots in the issue. 175/200% want the
 * chrome sized in rem off the same variable rather than zoomed; that is a
 * larger change and is not what this lands.
 */
export const UI_SCALE_STEPS = [1, 1.1, 1.25, 1.5] as const

export const UI_SCALE_MIN = UI_SCALE_STEPS[0]
export const UI_SCALE_MAX = UI_SCALE_STEPS[UI_SCALE_STEPS.length - 1]

/**
 * The stored value, clamped into range.
 *
 * `app-settings.json` is a plain JSON file a user can edit, so the number is
 * treated as untrusted: anything unreadable, non-finite, or out of range falls
 * back to 1 rather than propagating a `zoom: 0` or `zoom: 1e9` into the layout.
 */
export function readUiScale(settingsPath: string): number {
  const stored: unknown = readAppSettings(settingsPath)[UI_SCALE_KEY]
  return typeof stored === 'number' && Number.isFinite(stored) ? clampUiScale(stored) : UI_SCALE_MIN
}

export function clampUiScale(scale: number): number {
  return Math.min(UI_SCALE_MAX, Math.max(UI_SCALE_MIN, scale))
}

/** The next preset above / below `scale`, clamped at the ends. */
export function stepUiScale(scale: number, direction: 1 | -1): number {
  const index = UI_SCALE_STEPS.findIndex((step) => step >= scale - 1e-9)
  const from = index === -1 ? UI_SCALE_STEPS.length - 1 : index
  const next = UI_SCALE_STEPS[Math.min(UI_SCALE_STEPS.length - 1, Math.max(0, from + direction))]
  return next ?? UI_SCALE_MIN
}

/**
 * Write the scale into a renderer's `:root`.
 *
 * An inline style, not `insertCSS`: the shared token sheet already declares
 * `--ui-scale: 1` on `:root`, and a second author-level `:root` rule loses to
 * it on sheet order — which reads as "the setting is ignored" while every call
 * reports success. An inline declaration on `documentElement` wins by
 * construction, and `executeJavaScript` needs nothing from the page, so all six
 * editors get it without knowing this module exists.
 */
export function applyUiScaleTo(contents: Electron.WebContents, scale: number): void {
  if (contents.isDestroyed()) return
  void contents
    .executeJavaScript(
      `document.documentElement.style.setProperty('--ui-scale', '${clampUiScale(scale)}')`,
    )
    .catch(() => {
      // A renderer torn down mid-call has nothing left to style.
    })
}

/** Push the scale to every live webContents — shell window, tab views, detached windows. */
export function applyUiScaleToAll(scale: number): void {
  for (const contents of webContents.getAllWebContents()) {
    if (contents.isDestroyed()) continue
    applyUiScaleTo(contents, scale)
  }
}

/**
 * Persist a new scale and apply it everywhere.
 *
 * Persist first, then apply: `writeAppSetting` rethrows when the file is
 * unwritable, and applying first would show a preference that silently reverts
 * on the next launch.
 *
 * Listeners run last because the scale changes more than pixels: the shell
 * positions every editor view below the tab strip, and that strip is `zoom`ed,
 * so the main process has to re-lay-out or the document stays underneath a
 * strip that has quietly grown over it.
 */
export function setUiScale(settingsPath: string, scale: number): number {
  const next = clampUiScale(scale)
  writeAppSetting(settingsPath, UI_SCALE_KEY, next)
  applyUiScaleToAll(next)
  for (const listener of changeListeners) listener(next)
  return next
}

const changeListeners = new Set<(scale: number) => void>()

/**
 * Register what has to happen when the scale changes — today that is the
 * shell's re-layout and its title-bar overlay.
 *
 * A listener registry rather than an import of the shell's window code: the
 * per-editor menu modules live in the editor apps and call `changeUiScale`, so
 * the scale cannot depend on anything in the shell's main entry point.
 */
export function onUiScaleChanged(listener: (scale: number) => void): void {
  changeListeners.add(listener)
}

/** `0` resets; `1` / `-1` step through the presets. */
export function changeUiScale(settingsPath: string, direction: 1 | -1 | 0): number {
  if (direction === 0) return setUiScale(settingsPath, UI_SCALE_MIN)
  return setUiScale(settingsPath, stepUiScale(readUiScale(settingsPath), direction))
}
