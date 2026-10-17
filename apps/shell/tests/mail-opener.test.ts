import { beforeEach, describe, expect, it, vi } from 'vitest'
import { isValidElement } from 'react'
import type { HomeApi } from '../src/shared/home-api'
import { HOME_CHANNELS } from '../src/shared/home-api'
import type { TabKind } from '../src/shared/tabs-api'

/**
 * Mail P3 testability lane (t_5dcc0c97).
 * Gap A — the mail tab had no automation-reachable opener: the only path was
 * the native "+"/File menu AI Mail item, and QA proved neither Playwright CDP
 * keyboard nor OS SendKeys can drive native TrackPopupMenu HWNDs. The
 * `home:new-mail` HOME_CHANNEL plus `aiOffice.newMail()` (mirroring
 * newDoc..newPdf) gives Playwright a one-line opener:
 * `page.evaluate(() => window.aiOffice.newMail())`.
 * Gap B (tab-strip side) — the mail tab-strip icon must exist for every kind.
 */

const electronMocks = vi.hoisted(() => ({
  exposed: new Map<string, unknown>(),
  invoke: vi.fn(),
  send: vi.fn(),
  on: vi.fn(),
  removeListener: vi.fn(),
}))

vi.mock('electron', () => ({
  contextBridge: {
    exposeInMainWorld: (name: string, api: unknown) => electronMocks.exposed.set(name, api),
  },
  ipcRenderer: {
    invoke: electronMocks.invoke,
    send: electronMocks.send,
    on: electronMocks.on,
    removeListener: electronMocks.removeListener,
  },
  // the preload imports @genoffice/electron-utils (drop-open bridge), which
  // binds webUtils at module scope even though node env never installs it
  webUtils: { getPathForFile: () => '' },
}))

import '../src/preload/index'
import { KIND_ICON } from '../src/renderer/src/TabBar'

const homeApi = electronMocks.exposed.get('aiOffice') as HomeApi

beforeEach(() => {
  electronMocks.invoke.mockReset()
  electronMocks.send.mockReset()
})

describe('home:new-mail channel contract', () => {
  it('registers a dedicated channel mirroring newDoc..newPdf', () => {
    expect(HOME_CHANNELS.newMail).toBe('home:new-mail')
  })
})

describe('aiOffice.newMail() preload bridge', () => {
  it('is exposed on the Home bridge', () => {
    expect(typeof homeApi.newMail).toBe('function')
  })

  it('invokes the channel with no args (fresh mail tab)', async () => {
    electronMocks.invoke.mockResolvedValue(undefined)
    await homeApi.newMail()
    expect(electronMocks.invoke).toHaveBeenCalledWith(HOME_CHANNELS.newMail, undefined)
  })

  it('forwards new-file opts like the other new-file openers', async () => {
    electronMocks.invoke.mockResolvedValue(undefined)
    await homeApi.newMail({ dir: '/tmp' })
    expect(electronMocks.invoke).toHaveBeenCalledWith(HOME_CHANNELS.newMail, { dir: '/tmp' })
  })
})

describe('mail tab-strip icon', () => {
  it('covers every tab kind including mail', () => {
    const kinds: TabKind[] = [
      'home',
      'docs',
      'sheets',
      'slides',
      'pdf',
      'markdown',
      'html',
      'help',
      'mail',
    ]
    expect(Object.keys(KIND_ICON).sort()).toEqual([...kinds].sort())
  })

  it('renders a real element for mail', () => {
    expect(isValidElement(KIND_ICON.mail)).toBe(true)
  })
})
