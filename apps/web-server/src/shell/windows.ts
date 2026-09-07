/**
 * Web-window channels — open/list/focus. State is in-memory; web windows
 * are URL pointers, not BrowserWindow handles.
 */
import { WEB_WINDOWS, registerHandle } from '../common/index.js'

export function registerWindowHandlers(): void {
  registerHandle('win:new', (_event: unknown, options: unknown) => {
    const opts = options as { url?: string; name?: string } | undefined
    const id = `win-${Date.now()}`
    WEB_WINDOWS.set(id, { url: opts?.url || '/', name: opts?.name || '新窗口' })
    return { id, url: opts?.url || '/' }
  })

  registerHandle('win:list', () => {
    return [...WEB_WINDOWS.entries()].map(([id, win]) => ({ id, ...win }))
  })

  registerHandle('win:focus', (_event: unknown, id: unknown) => {
    if (WEB_WINDOWS.has(id as string)) {
      return { ok: true, id }
    }
    return { ok: false, error: 'Window not found' }
  })
}
