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

  const addSendListener = (channel: string, fn: unknown): void => {
    const list = sendListeners.get(channel) ?? []
    list.push(fn as IpcHandler)
    sendListeners.set(channel, list)
  }

  const on = ipcMain.on.bind(ipcMain)
  ;(ipcMain as unknown as { on: (ch: string, fn: unknown) => void }).on = (
    channel: string,
    fn: unknown,
  ) => {
    addSendListener(channel, fn)
    return on(channel, fn as Parameters<typeof on>[1])
  }

  // the wrapped listener drops itself from the mirror when it fires: an
  // Electron once() runs a single time, and dispatchSend must not keep
  // re-delivering it on every later frame
  const once = ipcMain.once.bind(ipcMain)
  ;(ipcMain as unknown as { once: (ch: string, fn: unknown) => void }).once = (
    channel: string,
    fn: unknown,
  ) => {
    const onceFn = ((...args: unknown[]) => {
      removeSendListener(channel, fn as IpcHandler)
      ;(fn as IpcHandler)(...args)
    }) as IpcHandler
    addSendListener(channel, onceFn)
    return once(channel, onceFn as Parameters<typeof once>[1])
  }

  // removeListener/off are wrapped so a desktop-side removal also drops the
  // mirror entry; when the host lacks them (older mocks), the mirror stays
  // authoritative on its own
  const removeListener = (ipcMain as unknown as Record<string, unknown>).removeListener as
    | ((ch: string, fn: unknown) => void)
    | undefined
  const wrapRemove = (name: 'removeListener' | 'off'): void => {
    ;(ipcMain as unknown as Record<string, unknown>)[name] = (
      channel: string,
      fn: unknown,
    ) => {
      removeSendListener(channel, fn as IpcHandler)
      removeListener?.(channel, fn)
    }
  }
  wrapRemove('removeListener')
  wrapRemove('off')
}

function removeSendListener(channel: string, fn: IpcHandler): void {
  const list = sendListeners.get(channel)
  if (!list) return
  const i = list.indexOf(fn)
  if (i >= 0) list.splice(i, 1)
}

export function getInvokeHandler(channel: string): IpcHandler | undefined {
  return invokeHandlers.get(channel)
}

export function getSendListeners(channel: string): readonly IpcHandler[] {
  return sendListeners.get(channel) ?? []
}

installWebIpcCapture()
