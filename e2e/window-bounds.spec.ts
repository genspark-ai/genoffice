import { test, expect } from '@playwright/test'
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { closeAndSaveVideo, launchShell } from './helpers'

/**
 * The shell window remembers its position and size across sessions
 * (genoffice#2049): a moved and resized window comes back at the same
 * geometry after a restart, and a saved position no display can show — an
 * unplugged monitor — is discarded instead of opening off-screen.
 *
 * The geometry is written to userData/app-settings.json, the file the issue
 * names, so the saved state is asserted there; the restored state is read off
 * the live window in the main process.
 */

const BOUNDS = { x: 100, y: 80, width: 900, height: 600 }
const SAVED_KEY = 'windowBounds'

/** The saved window state on disk; {} while absent so polls keep retrying. */
function savedBounds(userDataDir: string): Record<string, unknown> {
  try {
    const settings = JSON.parse(readFileSync(join(userDataDir, 'app-settings.json'), 'utf8')) as {
      [SAVED_KEY]?: unknown
    }
    const bounds = settings[SAVED_KEY]
    return typeof bounds === 'object' && bounds !== null ? (bounds as Record<string, unknown>) : {}
  } catch {
    return {}
  }
}

test.describe('shell: window bounds across sessions', () => {
  test('a moved and resized window reopens at the same geometry', async () => {
    // how far the platform moves a newly created window sideways off the
    // requested x. macOS' window server does this even when the constructor
    // asks for an exact position — reproduced with a hand-made Electron app,
    // so it is a platform quirk, not GenOffice code; on Linux CI the exact
    // rectangle comes back (dx 0). The same nudge applies to the restored
    // window, so measure it here on a throwaway window at the same rectangle
    // and require the restored window to repeat it.
    const first = await launchShell({ onboardingSeen: true, videoDir: 'window-bounds' })
    let nudge: number
    try {
      nudge = await first.app.evaluate(({ BrowserWindow }, bounds) => {
        BrowserWindow.getAllWindows()[0]?.setBounds(bounds)
        const probe = new BrowserWindow({ ...bounds, show: false })
        probe.showInactive()
        const placed = probe.getBounds()
        probe.destroy()
        return placed.x - bounds.x
      }, BOUNDS)
      await expect
        .poll(() => savedBounds(first.userDataDir))
        .toMatchObject({ ...BOUNDS, maximized: false })
      expect(Math.abs(nudge)).toBeLessThan(200)
    } finally {
      await closeAndSaveVideo(first, 'window-bounds-close')
    }

    // ── session 2: the same profile reopens at the saved geometry ──
    const second = await launchShell({ userDataDir: first.userDataDir, videoDir: 'window-bounds' })
    try {
      const restored = await second.app.evaluate(({ BrowserWindow }) => {
        const win = BrowserWindow.getAllWindows()[0]
        return { ...win.getNormalBounds(), maximized: win.isMaximized() }
      })
      expect(restored.width).toBe(BOUNDS.width)
      expect(restored.height).toBe(BOUNDS.height)
      expect(restored.y).toBe(BOUNDS.y)
      expect(restored.maximized).toBe(false)
      expect(restored.x).toBeCloseTo(BOUNDS.x + nudge, 1)
    } finally {
      await closeAndSaveVideo(second, 'window-bounds-restore')
    }
  })

  test('a saved position on a monitor that is gone falls back to the default placement', async () => {
    // a position far off any real display, as if saved on an unplugged monitor
    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'window-bounds',
      settings: { [SAVED_KEY]: { x: 9000, y: 9000, width: 900, height: 600, maximized: false } },
    })
    try {
      const opened = await launched.app.evaluate(({ BrowserWindow, screen }) => {
        const win = BrowserWindow.getAllWindows()[0]
        return {
          bounds: { ...win.getNormalBounds() },
          areas: screen.getAllDisplays().map((display) => display.workArea),
        }
      })
      // the lost position was discarded — the window opens on a display
      expect(opened.bounds.x).not.toBe(9000)
      const onSomeDisplay = opened.areas.some(
        (area) =>
          opened.bounds.x < area.x + area.width &&
          opened.bounds.x + opened.bounds.width > area.x &&
          opened.bounds.y < area.y + area.height &&
          opened.bounds.y + opened.bounds.height > area.y,
      )
      expect(onSomeDisplay).toBe(true)
    } finally {
      await closeAndSaveVideo(launched, 'window-bounds-offscreen')
    }
  })
})
