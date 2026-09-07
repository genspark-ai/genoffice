/**
 * Tabs and update channels — tab list/activate/close/reorder/show-menu plus
 * the `update:*` parity set that mirrors `apps/shell/src/main/index.ts`.
 */
import { registerHandle, TABS } from '../common/index.js'

export function registerTabsHandlers(): void {
  registerHandle('tabs:list', () => [...TABS.values()])

  registerHandle('tabs:activate', (_event: unknown, args: unknown) => {
    const { id } = (args || {}) as { id: string }
    return { ok: true, active: id }
  })

  registerHandle('tabs:close', (_event: unknown, args: unknown) => {
    const { id } = (args || {}) as { id: string }
    TABS.delete(id)
    return { ok: true }
  })

  registerHandle('tabs:reorder', (_event: unknown, _args: unknown) => ({ ok: true }))

  registerHandle('tabs:show-menu', (_event: unknown, _args: unknown) => ({ ok: true }))
  registerHandle('tabs:show-new-menu', (_event: unknown, _args: unknown) => ({ ok: true }))
  registerHandle('tabs:chrome-pressed', () => ({ ok: true }))
}

export function registerUpdateHandlers(): void {
  registerHandle('update:get-state', () => ({ status: 'idle', version: '1.0.0' }))
  registerHandle('update:download', () => ({ ok: true, status: 'downloading' }))
  registerHandle('update:install', () => ({ ok: true, status: 'installing' }))
  registerHandle('update:later', () => ({ ok: true }))
  registerHandle('update:open-download', () => ({ ok: true }))
}
