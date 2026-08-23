import { beforeEach, describe, expect, it, vi } from 'vitest'
import type { HomeApi } from '../src/shared/home-api'
import { HOME_CHANNELS } from '../src/shared/home-api'

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
}))

import '../src/preload/index'

const homeApi = electronMocks.exposed.get('aiOffice') as HomeApi

beforeEach(() => {
  electronMocks.invoke.mockReset()
  electronMocks.send.mockReset()
})

describe('analytics consent preload API', () => {
  it('defaults malformed main-process state to on', async () => {
    electronMocks.invoke.mockResolvedValue('invalid')
    await expect(homeApi.getAnalyticsEnabled()).resolves.toBe(true)
  })

  it('rejects non-boolean consent before invoking the main process', async () => {
    const setAnalytics = homeApi.setAnalyticsEnabled as (enabled: unknown) => Promise<boolean>

    await expect(setAnalytics('true')).rejects.toThrow('Invalid analytics consent')
    expect(electronMocks.invoke).not.toHaveBeenCalled()
  })

  it('forwards explicit booleans and reports persistence success', async () => {
    electronMocks.invoke.mockResolvedValue(true)

    await expect(homeApi.setAnalyticsEnabled(false)).resolves.toBe(true)
    expect(electronMocks.invoke).toHaveBeenLastCalledWith(HOME_CHANNELS.setAnalyticsEnabled, false)

    await expect(homeApi.setOnboardingSeen()).resolves.toBe(true)
    expect(electronMocks.invoke).toHaveBeenLastCalledWith(HOME_CHANNELS.setOnboardingSeen)
  })
})
