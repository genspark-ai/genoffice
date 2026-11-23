import { beforeEach, describe, expect, it, vi } from 'vitest'

const sendMock = vi.fn()
vi.mock('electron', () => ({
  webContents: {
    getAllWebContents: () => [{ isDestroyed: () => false, send: sendMock }],
  },
}))

const closeMock = vi.fn()
const handles: Array<{
  config: { port: number; host: string; token: string }
  close: typeof closeMock
}> = []
let failNext = false
vi.mock('../src/web/index', () => ({
  webBroadcast: vi.fn(),
  startWebServer: async (overrides: Record<string, unknown>) => {
    if (failNext) {
      failNext = false
      throw new Error('listen EADDRINUSE: address already in use')
    }
    const config = {
      port: (overrides.port as number) ?? 8931,
      host: (overrides.host as string) ?? '127.0.0.1',
      token: (overrides.token as string) ?? 'boot-token',
    }
    const handle = { config, close: closeMock }
    handles.push(handle)
    return handle
  },
}))

import {
  lanAddress,
  setWebShareHandleFactory,
  shareUrlFor,
  startWebShare,
  stopWebShare,
  webShareStatus,
} from '../src/web/share'
import { WEB_SHARE_LABELS } from '../../../packages/ui/src/strings-web-share'
import { LANGS } from '@genoffice/i18n'

describe('web share controller', () => {
  beforeEach(() => {
    stopWebShare() // close any handle a previous test left live, first
    handles.length = 0
    closeMock.mockClear()
    sendMock.mockClear()
  })

  it('reports not-running before anything started', () => {
    expect(webShareStatus()).toEqual({ running: false })
  })

  it('starts LAN-reachable with a fresh token and reports the share URL', async () => {
    setWebShareHandleFactory(async (overrides) => {
      const config = {
        port: (overrides.port as number) ?? 8931,
        host: (overrides.host as string) ?? '127.0.0.1',
        token: (overrides.token as string) ?? 'x',
      }
      const handle = { config, close: closeMock }
      handles.push(handle)
      return handle
    })
    const status = await startWebShare()
    expect(status.running).toBe(true)
    expect(status.url).toMatch(/^http:\/\/[\d.]+:8931\/\?token=[0-9a-f]{48}$/)
    expect(handles[0]!.config.host).toBe('0.0.0.0')
    expect(handles[0]!.config.token).not.toBe('boot-token')
    // every open editor hears about it
    expect(sendMock).toHaveBeenCalled()
  })

  it('keeps the live service on a second start (never disturbs sessions)', async () => {
    await startWebShare()
    const again = await startWebShare()
    expect(again.running).toBe(true)
    expect(handles.length).toBe(1)
  })

  it('stop closes the handle and flips every chip back', async () => {
    await startWebShare()
    const status = stopWebShare()
    expect(closeMock).toHaveBeenCalledOnce()
    expect(status.running).toBe(false)
    expect(
      sendMock.mock.calls.some(([, value]) => (value as { running: boolean }).running === false),
    ).toBe(true)
  })

  it('surfaces a start failure as error state instead of throwing', async () => {
    setWebShareHandleFactory(async () => {
      throw new Error('listen EADDRINUSE: address already in use')
    })
    const status = await startWebShare()
    expect(status.running).toBe(false)
    expect(status.error).toContain('EADDRINUSE')
  })
})

describe('share URL derivation', () => {
  it('prefers the first non-internal IPv4', () => {
    const url = shareUrlFor({ port: 8931, host: '0.0.0.0', token: 'abc' }, '10.1.2.3')
    expect(url).toBe('http://10.1.2.3:8931/?token=abc')
  })

  it('falls back to loopback when no LAN interface exists', () => {
    expect(shareUrlFor({ port: 8931, host: '0.0.0.0', token: 'abc' }, '127.0.0.1')).toBe(
      'http://127.0.0.1:8931/?token=abc',
    )
  })

  it('lanAddress reads node interfaces without crashing', () => {
    expect(typeof lanAddress()).toBe('string')
  })
})

describe('web share labels', () => {
  it('covers every UI language with all three strings', () => {
    for (const lang of LANGS) {
      const labels = WEB_SHARE_LABELS[lang]
      expect(labels.share.length, lang).toBeGreaterThan(0)
      expect(labels.shareTip.length, lang).toBeGreaterThan(0)
      expect(labels.stopTip.length, lang).toBeGreaterThan(0)
    }
  })
})
