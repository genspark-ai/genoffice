import { afterEach, describe, expect, it, vi } from 'vitest'
import { aiFetch, setRescueFetch } from '../src/fetch'

afterEach(() => {
  vi.unstubAllGlobals()
  setRescueFetch(null)
})

describe('aiFetch', () => {
  it('returns the primary response without touching the rescue path', async () => {
    const ok = new Response('ok')
    vi.stubGlobal('fetch', vi.fn().mockResolvedValue(ok))
    const rescue = vi.fn()
    setRescueFetch(rescue)
    expect(await aiFetch('https://x/', {})).toBe(ok)
    expect(rescue).not.toHaveBeenCalled()
  })

  it('retries over the rescue fetch when the primary fails', async () => {
    vi.stubGlobal('fetch', vi.fn().mockRejectedValue(new Error('fetch failed')))
    const ok = new Response('rescued')
    const rescue = vi.fn().mockResolvedValue(ok)
    setRescueFetch(rescue)
    expect(await aiFetch('https://x/', {})).toBe(ok)
    expect(rescue).toHaveBeenCalledOnce()
  })

  it('throws the primary error when no rescue fetch is set', async () => {
    vi.stubGlobal('fetch', vi.fn().mockRejectedValue(new Error('ECONNRESET')))
    await expect(aiFetch('https://x/', {})).rejects.toThrow('ECONNRESET')
  })

  it('throws the primary error when the rescue fetch also fails', async () => {
    vi.stubGlobal('fetch', vi.fn().mockRejectedValue(new Error('primary down')))
    setRescueFetch(vi.fn().mockRejectedValue(new Error('rescue down')))
    await expect(aiFetch('https://x/', {})).rejects.toThrow('primary down')
  })

  it('does not retry an aborted request', async () => {
    vi.stubGlobal('fetch', vi.fn().mockRejectedValue(new Error('aborted')))
    const rescue = vi.fn()
    setRescueFetch(rescue)
    const controller = new AbortController()
    controller.abort()
    await expect(aiFetch('https://x/', { signal: controller.signal })).rejects.toThrow('aborted')
    expect(rescue).not.toHaveBeenCalled()
  })
})
