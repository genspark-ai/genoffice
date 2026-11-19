/// App-wide navigation lockdown for every webContents in the suite.
///
/// Every renderer is a local single-page app loaded exactly once via
/// loadFile/loadURL (which do not emit will-navigate), so a full-page
/// navigation is never part of normal operation — it can only come from
/// hostile document content or AI-generated markup. The guard blocks all of
/// them except same-URL reloads (used by dev tooling), and installs a
/// default-deny window.open handler.
///
/// Apps that intentionally route window.open links to the system browser
/// (docs/slides/pdf via safeExternalUrl) install their own handler on
/// specific webContents after creation; that replaces the default-deny
/// handler for those contents only.
import type { App } from 'electron'

// Symbol.for: the flag survives multiple bundled copies of this module
// (each app bundle ships its own electron-utils, and the shell hosts all of
// them in one process).
const INSTALLED = Symbol.for('genoffice.navigation-guard-installed')

export function installNavigationGuard(app: App): void {
  const holder = app as unknown as Record<symbol, boolean | undefined>
  if (holder[INSTALLED]) return
  holder[INSTALLED] = true
  app.on('web-contents-created', (_event, contents) => {
    contents.on('will-navigate', (event) => {
      if (event.url !== contents.getURL()) event.preventDefault()
    })
    contents.setWindowOpenHandler(() => ({ action: 'deny' }))
  })
}
