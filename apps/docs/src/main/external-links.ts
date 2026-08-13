/// Window-open routing for Docs windows: every window.open from a renderer
/// (including document hyperlink clicks) funnels through here. The URL is
/// validated against the Docs protocol allowlist and handed to
/// shell.openExternal; the window itself never navigates or spawns children.
import { safeExternalUrl } from '@genoffice/electron-utils'
import { DOCS_LINK_PROTOCOLS } from '../shared/link-protocols'

type OpenExternal = (url: string) => Promise<void> | void

export interface WindowOpenDetails {
  url: string
}

export interface WindowOpenAction {
  action: 'deny'
}

/**
 * Returns the handler to install via `webContents.setWindowOpenHandler`.
 * `openExternal` is injected so tests can assert routing without Electron.
 */
export function createDocsWindowOpenHandler(
  openExternal: OpenExternal,
): (details: WindowOpenDetails) => WindowOpenAction {
  return ({ url }) => {
    const target = safeExternalUrl(url, { allowedProtocols: DOCS_LINK_PROTOCOLS })
    if (target) void openExternal(target)
    return { action: 'deny' }
  }
}
