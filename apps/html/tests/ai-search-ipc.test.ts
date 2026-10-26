import { beforeEach, describe, expect, it, vi } from 'vitest'
/**
 * Standalone search IPC (src/main/ai-search-ipc.ts): the ai:web-search and
 * ai:image-search channels must be registered exactly once with the same
 * provider chain as sheets/docs/slides (imageSearchTool/webSearchTool over
 * ai-settings.json), and a backend throw must degrade to the shared error
 * shape instead of rejecting the invoke.
 */

const handle = vi.fn()

vi.mock('electron', () => ({
  app: { getPath: vi.fn(() => '/UserData') },
  ipcMain: { handle: (...args: unknown[]) => handle(...(args as [])) },
}))

const webSearchTool = vi.fn()
const imageSearchTool = vi.fn()

vi.mock('@genoffice/ai-search', () => ({
  webSearchTool: (...args: unknown[]) => webSearchTool(...(args as [])),
  imageSearchTool: (...args: unknown[]) => imageSearchTool(...(args as [])),
}))

let registerHtmlAiSearchIpc: typeof import('../src/main/ai-search-ipc').registerHtmlAiSearchIpc

const handlerFor = (channel: string): ((...args: unknown[]) => unknown) => {
  const call = handle.mock.calls.find((c) => c[0] === channel)
  if (!call) throw new Error(`${channel} was not registered`)
  return call[1] as (...args: unknown[]) => unknown
}

beforeEach(async () => {
  // a fresh module per test: the registration flag is module state
  vi.resetModules()
  handle.mockClear()
  webSearchTool.mockReset()
  imageSearchTool.mockReset()
  ;({ registerHtmlAiSearchIpc } = await import('../src/main/ai-search-ipc'))
})

describe('registerHtmlAiSearchIpc', () => {
  it('registers both search channels once; a second call is a no-op', () => {
    registerHtmlAiSearchIpc()
    expect(handle).toHaveBeenCalledTimes(2)
    expect(handle.mock.calls.map((c) => c[0])).toEqual(['ai:web-search', 'ai:image-search'])
    registerHtmlAiSearchIpc()
    expect(handle).toHaveBeenCalledTimes(2)
  })

  it('ai:image-search routes through imageSearchTool with the settings path and default 8', async () => {
    registerHtmlAiSearchIpc()
    imageSearchTool.mockResolvedValue({ images: [], method: 'duckduckgo' })
    const r = (await handlerFor('ai:image-search')({}, 'black cat', undefined)) as {
      method: string
    }
    expect(r.method).toBe('duckduckgo')
    expect(imageSearchTool).toHaveBeenCalledTimes(1)
    const [settingsPath, query, maxResults] = imageSearchTool.mock.calls[0] as unknown[]
    expect(String(settingsPath)).toContain('ai-settings.json')
    expect(query).toBe('black cat')
    expect(maxResults).toBe(8)
  })

  it('ai:image-search passes an explicit maxResults through', async () => {
    registerHtmlAiSearchIpc()
    imageSearchTool.mockResolvedValue({ images: [], method: 'serper' })
    await handlerFor('ai:image-search')({}, 'poster', 3)
    expect(imageSearchTool.mock.calls[0]?.[2]).toBe(3)
  })

  it('ai:image-search degrades a backend throw to the shared error shape', async () => {
    registerHtmlAiSearchIpc()
    imageSearchTool.mockRejectedValue(new Error('boom'))
    const r = (await handlerFor('ai:image-search')({}, 'q', undefined)) as {
      images: unknown[]
      method: string
      error?: string
    }
    expect(r).toMatchObject({ images: [], method: 'error' })
    expect(r.error).toContain('boom')
  })

  it('ai:web-search routes through webSearchTool with the default 6 and degrades throws', async () => {
    registerHtmlAiSearchIpc()
    webSearchTool.mockResolvedValue({ results: [], method: 'serper' })
    const r = (await handlerFor('ai:web-search')({}, 'facts', undefined)) as { method: string }
    expect(r.method).toBe('serper')
    expect(webSearchTool.mock.calls[0]?.[2]).toBe(6)
    webSearchTool.mockRejectedValue(new Error('down'))
    const failed = (await handlerFor('ai:web-search')({}, 'facts', 2)) as {
      results: unknown[]
      method: string
      error?: string
    }
    expect(failed).toMatchObject({ results: [], method: 'error' })
    expect(webSearchTool.mock.calls[1]?.[2]).toBe(2)
    expect(failed.error).toContain('down')
  })
})
