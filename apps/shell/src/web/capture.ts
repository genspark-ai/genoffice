import { ipcMain } from 'electron'

/**
 * Mirrors every ipcMain registration into a lookup table so the web bridge can
 * dispatch browser requests into the very same handlers the desktop windows
 * use.
 *
 * This module INSTALLS ITSELF when imported (the call at the bottom runs during
 * module evaluation). ES imports hoist, so any `installWebIpcCapture()` call
 * from shell main would run after docs-main/slides-main — which register a few
 * ipcMain channels at module top level — and those would be silently dropped
 * from the web surface. Importing this module FIRST in shell main puts the
 * wrapper in place before every app module's body evaluates; the exported
 * function stays for clarity and is idempotent.
 */

export type IpcHandler = (...args: unknown[]) => unknown

const invokeHandlers = new Map<string, IpcHandler>()
const sendListeners = new Map<string, IpcHandler[]>()

let installed = false

export function installWebIpcCapture(): void {
  if (installed) return
  installed = true

  const invoke = ipcMain.handle.bind(ipcMain) as (ch: string, fn: unknown) => void
  ;(ipcMain as unknown as { handle: (ch: string, fn: unknown) => void }).handle = (
    channel: string,
    fn: unknown,
  ) => {
    invokeHandlers.set(channel, fn as IpcHandler)
    return invoke(channel, fn)
  }

  const removeHandler = ipcMain.removeHandler.bind(ipcMain)
  ;(ipcMain as unknown as { removeHandler: (ch: string) => void }).removeHandler = (
    channel: string,
  ) => {
    invokeHandlers.delete(channel)
    removeHandler(channel)
  }

  const on = ipcMain.on.bind(ipcMain)
  ;(ipcMain as unknown as { on: (ch: string, fn: unknown) => void }).on = (
    channel: string,
    fn: unknown,
  ) => {
    const list = sendListeners.get(channel) ?? []
    list.push(fn as IpcHandler)
    sendListeners.set(channel, list)
    return on(channel, fn as Parameters<typeof on>[1])
  }

  const once = ipcMain.once.bind(ipcMain)
  ;(ipcMain as unknown as { once: (ch: string, fn: unknown) => void }).once = (
    channel: string,
    fn: unknown,
  ) => {
    const list = sendListeners.get(channel) ?? []
    list.push(fn as IpcHandler)
    sendListeners.set(channel, list)
    return once(channel, fn as Parameters<typeof once>[1])
  }
}

export function getInvokeHandler(channel: string): IpcHandler | undefined {
  return invokeHandlers.get(channel)
}

export function getSendListeners(channel: string): readonly IpcHandler[] {
  return sendListeners.get(channel) ?? []
}

installWebIpcCapture()
