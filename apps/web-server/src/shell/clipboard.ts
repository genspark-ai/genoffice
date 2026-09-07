/**
 * Clipboard channels — `copy/cut/paste` (Electron top-level) and the
 * `clipboard:*` namespaced set. All return the "use the browser's native
 * clipboard" placeholder.
 */
import { registerHandle } from '../common/index.js'

export function registerClipboardHandlers(): void {
  registerHandle('copy', (_event: unknown, _text: unknown) => ({
    ok: true,
    message: '请使用浏览器原生 Ctrl+C / Cmd+C',
  }))
  registerHandle('cut', (_event: unknown, _text: unknown) => ({
    ok: true,
    message: '请使用浏览器原生 Ctrl+X / Cmd+X',
  }))
  registerHandle('paste', () => ({
    text: '',
    message: '请使用浏览器原生 Ctrl+V / Cmd+V',
  }))

  registerHandle('clipboard:copy', (_event: unknown, _text: unknown) => ({
    ok: true,
    message: '请使用浏览器原生 Ctrl+C / Cmd+C',
  }))
  registerHandle('clipboard:cut', (_event: unknown, _text: unknown) => ({
    ok: true,
    message: '请使用浏览器原生 Ctrl+X / Cmd+X',
  }))
  registerHandle('clipboard:paste', () => ({
    text: '',
    message: '请使用浏览器原生 Ctrl+V / Cmd+V',
  }))
}
