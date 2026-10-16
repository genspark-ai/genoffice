import { type BrowserWindow, WebContentsView, ipcMain, shell } from 'electron'
import { isAllowedViewUrl } from '@genoffice/agent-core'
import { FEATURE_VIEW_CHANNELS, type SlotViewMountRequest } from '../shared/feature-channels'
import { clampViewRect } from '../shared/feature-view-geometry'

/**
 * The native view slot (K5): a real WebContentsView mounted over the rectangle a
 * slot reserves in the renderer — today a browser panel a feature declared.
 *
 * The renderer never holds a webview tag or a webContents id: it reports a
 * rectangle and an address, the shell decides whether that address is allowed and
 * owns the view. The view is sandboxed with no preload and no node, it can only
 * navigate to allowed addresses, and a window.open goes to the user's browser
 * instead of opening a second, unreported window.
 *
 * One view at a time per slot instance; the map is small and emptied on unmount.
 */
export interface FeaturesViewDeps {
  /** the shell window, or null before it exists / after it closed */
  window(): BrowserWindow | null
}

export interface FeaturesView {
  dispose(): void
}

export function registerFeaturesView(deps: FeaturesViewDeps): FeaturesView {
  const mounted = new Map<string, WebContentsView>()

  const hostWindow = (): BrowserWindow | null => {
    const found = deps.window()
    return found && !found.isDestroyed() ? found : null
  }

  const unmount = (id: string): void => {
    const view = mounted.get(id)
    if (!view) return
    mounted.delete(id)
    hostWindow()?.contentView.removeChildView(view)
    view.webContents.close()
  }

  ipcMain.handle(FEATURE_VIEW_CHANNELS.mount, (event, raw: unknown): boolean => {
    const request = (raw ?? {}) as Partial<SlotViewMountRequest>
    const id = typeof request.id === 'string' ? request.id : ''
    if (!id) return false
    const host = hostWindow()
    // a view belongs to the window whose renderer asked for it
    if (!host || event.sender !== host.webContents) return false

    const url = typeof request.url === 'string' ? request.url : ''
    if (!url) {
      unmount(id)
      return true
    }
    if (!isAllowedViewUrl(url)) {
      unmount(id)
      throw new Error('the native view opens https pages, or http on localhost')
    }
    const rect = clampViewRect(request.rect, host.getContentBounds())
    if (!rect) {
      // the slot is collapsed or off-screen: no view is better than a stray one
      unmount(id)
      return false
    }

    let view = mounted.get(id)
    if (!view) {
      view = new WebContentsView({
        webPreferences: { sandbox: true, contextIsolation: true, nodeIntegration: false },
      })
      view.webContents.setWindowOpenHandler(({ url: target }) => {
        if (isAllowedViewUrl(target)) void shell.openExternal(target)
        return { action: 'deny' }
      })
      view.webContents.on('will-navigate', (event_, target) => {
        if (!isAllowedViewUrl(target)) event_.preventDefault()
      })
      host.contentView.addChildView(view)
      mounted.set(id, view)
      void view.webContents.loadURL(url)
    }
    view.setBounds(rect)
    return true
  })

  ipcMain.handle(FEATURE_VIEW_CHANNELS.unmount, (_event, raw: unknown): boolean => {
    const id = (raw ?? {}) as { id?: unknown }
    if (typeof id.id === 'string') unmount(id.id)
    return true
  })

  return {
    dispose(): void {
      for (const id of [...mounted.keys()]) unmount(id)
      ipcMain.removeHandler(FEATURE_VIEW_CHANNELS.mount)
      ipcMain.removeHandler(FEATURE_VIEW_CHANNELS.unmount)
    },
  }
}
