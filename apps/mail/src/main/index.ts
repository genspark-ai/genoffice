import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { ipcMain, shell } from 'electron'
import { stageAttachment } from '@genoffice/mail-protocols'
import { FakeComposeProvider, registerComposeHandlers } from './compose-handlers'
import type { ComposeIpcBus } from './compose-handlers'
// H5 quarantine (additive): staged-temp --> scan --> store IPC handlers.
import { registerQuarantineHandlers } from './quarantine-handlers'
import { startMailStandalone } from './mail-main'

// Standalone mail entry (Lane S owns the original lines; Lane C adds the
// compose wiring below — `mail-main.ts` itself is untouched). The Electron
// ipcMain adapter lives here so `compose-handlers.ts` stays Electron-free
// and vitest-safe. Shell-hosted mode does not load this entry: the merge
// lane calls `registerComposeHandlers` once from the shell main process
// with the same deps shape (provider + stage + stageDir + openExternal).
const electronBus: ComposeIpcBus = {
  handle: (channel, listener) => {
    ipcMain.handle(channel, (_event, payload) => listener(payload))
  },
}

registerComposeHandlers(electronBus, {
  provider: new FakeComposeProvider(),
  stage: stageAttachment,
  stageDir: join(tmpdir(), 'genoffice-mail-compose'),
  openExternal: (url) => shell.openExternal(url),
})
registerQuarantineHandlers(electronBus, {
  openPath: (path) => shell.openPath(path),
})

startMailStandalone()
