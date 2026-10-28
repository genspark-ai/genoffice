import { beforeEach, describe, expect, it, vi } from 'vitest'
import type { WebSocket } from 'ws'

// hoisted so the electron factory can reference it
const { ipcMainMock } = vi.hoisted(() => ({
  ipcMainMock: {
    handle: (_channel: string, _fn: unknown) => {},
    on: (_channel: string, _fn: unknown) => {},
    once: (_channel: string, _fn: unknown) => {},
    removeHandler: (_channel: string) => {},
  },
}))

vi.mock('electron', () => ({
  session: { defaultSession: { spellCheckerLanguages: [] } },
  ipcMain: ipcMainMock,
}))
vi.mock('../../docs/src/main/docs-main', () => ({
  queueDocsPendingOpen: (...args: unknown[]) => docsSeeds.push(['open', args]),
  markDocsNewBlank: (...args: unknown[]) => docsSeeds.push(['blank', args]),
}))
vi.mock('../../sheets/src/main/sheets-main', () => ({
  createSheetsWebTab: (wc: unknown, path: string) => sheetsSeeds.push([wc, path]),
  setSheetsNewBlank: () => sheetsSeeds.push(['blank']),
}))
vi.mock('../../slides/src/main/slides-main', () => ({
  trackWebSlidesTab: (wc: unknown, path?: string | null) => otherSeeds.push(['slides', wc, path]),
}))
vi.mock('../../markdown/src/main/markdown-main', () => ({
  trackWebTab: (wc: unknown, path?: string | null) => otherSeeds.push(['markdown', wc, path]),
}))
vi.mock('../../pdf/src/main/pdf-main', () => ({
  trackWebTab: (wc: unknown, path?: string | null) => otherSeeds.push(['pdf', wc, path]),
}))
vi.mock('../../html/src/main/html-main', () => ({
  trackWebTab: (wc: unknown, path?: string | null) => otherSeeds.push(['html', wc, path]),
}))

const docsSeeds: [string, unknown[]][] = []
const sheetsSeeds: unknown[][] = []
const otherSeeds: [string, unknown, unknown][] = []

const ipcRegistry = new Map<string, (...args: unknown[]) => unknown>()
const sendRegistry = new Map<string, ((...args: unknown[]) => unknown)[]>()

ipcMainMock.handle = (channel: string, fn: unknown) => {
  ipcRegistry.set(channel, fn as never)
}
ipcMainMock.on = (channel: string, fn: unknown) => {
  const list = sendRegistry.get(channel) ?? []
  list.push(fn as never)
  sendRegistry.set(channel, list)
}
ipcMainMock.once = ipcMainMock.on
ipcMainMock.removeHandler = (channel: string) => {
  ipcRegistry.delete(channel)
}

import { WebSocketServer } from 'ws'
import { installWebIpcCapture } from '../src/web/capture'
import { attachWebBridge } from '../src/web/bridge'
import { webTabById } from '../src/web/web-contents'
import type { AddressInfo } from 'node:net'

async function connect(
  path: string,
): Promise<{
  ws: WebSocket
  frames: Record<string, unknown>[]
  next: () => Promise<Record<string, unknown>>
}> {
  const WS = (await import('ws')).WebSocket
  const ws = new WS(`ws://127.0.0.1:${port}${path}`)
  const frames: Record<string, unknown>[] = []
  const waiters: ((f: Record<string, unknown>) => void)[] = []
  ws.on('message', (data: Buffer) => {
    const frame = JSON.parse(data.toString('utf8')) as Record<string, unknown>
    const waiter = waiters.shift()
    if (waiter) waiter(frame)
    else frames.push(frame)
  })
  await new Promise<void>((resolve, reject) => {
    ws.once('open', resolve)
    ws.once('error', reject)
  })
  return {
    ws,
    frames,
    next: () => {
      const buffered = frames.shift()
      if (buffered) return Promise.resolve(buffered)
      return new Promise((resolve, reject) => {
        const timer = setTimeout(() => reject(new Error('frame timeout')), 4000)
        waiters.push((f) => {
          clearTimeout(timer)
          resolve(f)
        })
      })
    },
  }
}

let wss: WebSocketServer
let port = 0

beforeEach(() => {
  docsSeeds.length = 0
  sheetsSeeds.length = 0
  otherSeeds.length = 0
  ipcRegistry.clear()
  sendRegistry.clear()
})

async function startBridge(): Promise<void> {
  installWebIpcCapture()
  wss = new WebSocketServer({ port: 0 })
  attachWebBridge(wss, { tabUrl: (tab) => tab.url })
  port = (wss.address() as AddressInfo).port
}

function stopBridge(): Promise<void> {
  for (const client of wss.clients) client.terminate()
  return new Promise((resolve) => wss.close(() => resolve()))
}

describe('web bridge', () => {
  it('dispatches an invoke to the captured handler and replies', async () => {
    await startBridge()
    ipcMainMock.handle('test:echo', (_event, a: string) => ({ echoed: a }))
    const client = await connect('/docs/')
    client.ws.send(JSON.stringify({ t: 'invoke', id: 1, channel: 'test:echo', args: ['hi'] }))
    const reply = await client.next()
    expect(reply).toEqual({ t: 'reply', id: 1, ok: true, value: { echoed: 'hi' } })
    await stopBridge()
  })

  it('reports handler errors instead of crashing the tab', async () => {
    await startBridge()
    ipcMainMock.handle('test:boom', () => {
      throw new Error('kaput')
    })
    const client = await connect('/docs/')
    client.ws.send(JSON.stringify({ t: 'invoke', id: 2, channel: 'test:boom', args: [] }))
    const reply = await client.next()
    expect(reply.ok).toBe(false)
    expect(String(reply.error)).toContain('kaput')
    await stopBridge()
  })

  it('seeds a docs pending open from the connection URL', async () => {
    await startBridge()
    const client = await connect('/docs/?path=%2Ftmp%2Freport.docx')
    await new Promise((r) => setTimeout(r, 100))
    expect(docsSeeds).toEqual([['open', [expect.any(Number), '/tmp/report.docx']]])
    client.ws.close()
    await stopBridge()
  })

  it('marks docs new-blank tabs and registers sheets web tabs', async () => {
    await startBridge()
    const client = await connect('/docs/?new=1')
    await new Promise((r) => setTimeout(r, 100))
    expect(docsSeeds).toEqual([['blank', [expect.any(Number)]]])
    client.ws.close()

    const client2 = await connect('/sheets/?path=%2Ftmp%2Fbook.xlsx')
    await new Promise((r) => setTimeout(r, 100))
    expect(sheetsSeeds).toHaveLength(1)
    expect(sheetsSeeds[0][1]).toBe('/tmp/book.xlsx')
    client2.ws.close()
    await stopBridge()
  })

  it('overrides win:new into an open-url push for the tab', async () => {
    await startBridge()
    const client = await connect('/docs/?path=%2Ftmp%2Fa.docx')
    client.ws.send(
      JSON.stringify({ t: 'invoke', id: 3, channel: 'win:new', args: ['/tmp/b.docx'] }),
    )
    const openUrl = await client.next()
    expect(openUrl).toEqual({ t: 'open-url', url: '/docs/?path=%2Ftmp%2Fb.docx' })
    const reply = await client.next()
    expect(reply.t).toBe('reply')
    client.ws.close()
    await stopBridge()
  })

  it('delivers main-side pushes to the tab and drops the tab state on close', async () => {
    await startBridge()
    const client = await connect('/docs/?path=%2Ftmp%2Fpush.docx')
    await new Promise((r) => setTimeout(r, 100))
    const tab = webTabById((docsSeeds[0]?.[1]?.[0] as number) ?? -1)
    expect(tab).toBeDefined()
    tab!.deliver('menu:command', ['save'])
    const frame = await client.next()
    expect(frame).toEqual({ t: 'event', channel: 'menu:command', args: ['save'] })

    const destroyed: number[] = []
    tab!.on('destroyed', () => destroyed.push(1))
    client.ws.close()
    await new Promise((r) => setTimeout(r, 100))
    expect(destroyed).toHaveLength(1)
    expect(webTabById(tab!.id)).toBeUndefined()
    await stopBridge()
  })

  it('hands binary invoke args and results through the codec', async () => {
    await startBridge()
    ipcMainMock.handle('test:save', (_event, data: unknown) => {
      const bytes = data instanceof ArrayBuffer ? new Uint8Array(data) : (data as Uint8Array)
      expect([...bytes]).toEqual([1, 2, 3])
      return { saved: true, echo: Buffer.from([4, 5]) }
    })
    const client = await connect('/docs/')
    const payload = new Uint8Array([1, 2, 3]).buffer
    client.ws.send(
      JSON.stringify({
        t: 'invoke',
        id: 9,
        channel: 'test:save',
        args: [{ __genofficeBytes: Buffer.from(new Uint8Array(payload)).toString('base64') }],
      }),
    )
    const reply = await client.next()
    expect(reply.ok).toBe(true)
    const value = reply.value as {
      echo: { __genofficeBytes: string; __genofficeBytesView?: boolean }
    }
    expect(Buffer.from(value.echo.__genofficeBytes, 'base64')).toEqual(Buffer.from([4, 5]))
    client.ws.close()
    await stopBridge()
  })

  it('runs fire-and-forget send listeners', async () => {
    await startBridge()
    const seen: unknown[][] = []
    ipcMainMock.on('test:note', (...args: unknown[]) => seen.push(args))
    const client = await connect('/docs/')
    client.ws.send(JSON.stringify({ t: 'send', channel: 'test:note', args: ['hello'] }))
    await new Promise((r) => setTimeout(r, 100))
    expect(seen).toHaveLength(1)
    expect(seen[0][1]).toBe('hello')
    client.ws.close()
    await stopBridge()
  })
})
