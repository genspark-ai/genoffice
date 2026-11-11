/** @vitest-environment jsdom */
import { afterEach, describe, expect, it, vi } from 'vitest'
import { suggestFileName } from '../src/ai-naming-client'

/**
 * The empty-string contract is the whole point of this helper.
 *
 * A file name is the one thing a reader cannot get back easily, so every way
 * the naming attempt can go wrong has to end with no name at all rather than a
 * wrong one. These cases are the ways it can go wrong.
 */

function stubApi(overrides: Record<string, unknown> = {}): void {
  ;(window as unknown as { aiNamingApi: unknown }).aiNamingApi = {
    suggestFileName: vi.fn(async () => ({ ok: true, name: 'Quarterly plan' })),
    ...overrides,
  }
}

afterEach(() => {
  delete (window as unknown as { aiNamingApi?: unknown }).aiNamingApi
  delete (window as unknown as { aiOffice?: unknown }).aiOffice
  vi.restoreAllMocks()
})

describe('suggestFileName', () => {
  it('uses the name the model returned', async () => {
    stubApi()
    expect(await suggestFileName({ content: 'Some document text' })).toBe('Quarterly plan')
  })

  it('offers no name when the model declines', async () => {
    stubApi({ suggestFileName: vi.fn(async () => ({ ok: false, reason: 'no-provider' })) })
    expect(await suggestFileName({ content: 'Some document text' })).toBe('')
  })

  // the bug this whole helper exists to be robust against: an empty name
  // arriving as a *success* would put nothing in the Save As dialog
  it('offers no name when a success carries an empty name', async () => {
    stubApi({ suggestFileName: vi.fn(async () => ({ ok: true, name: '' })) })
    expect(await suggestFileName({ content: 'Some document text' })).toBe('')
  })

  it('offers no name when a success carries only whitespace', async () => {
    stubApi({ suggestFileName: vi.fn(async () => ({ ok: true, name: '   ' })) })
    expect(await suggestFileName({ content: 'Some document text' })).toBe('')
  })

  it('offers no name when the call rejects', async () => {
    stubApi({
      suggestFileName: vi.fn(async () => {
        throw new Error('no handler')
      }),
    })
    expect(await suggestFileName({ content: 'Some document text' })).toBe('')
  })

  it('offers no name when the bridge has no naming channel at all', async () => {
    ;(window as unknown as { aiNamingApi: unknown }).aiNamingApi = {}
    expect(await suggestFileName({ content: 'Some document text' })).toBe('')
  })

  // the namespace the naming bridge must stay out of: `aiOffice` is the shell's
  // home API, and an editor window exposing a partial one would shadow it there
  it('does not look for the naming channel in the shell namespace', async () => {
    ;(window as unknown as { aiOffice: unknown }).aiOffice = {
      suggestFileName: vi.fn(async () => ({ ok: true, name: 'From the shell namespace' })),
    }
    expect(await suggestFileName({ content: 'Some document text' })).toBe('')
  })

  it('never asks about an empty document', async () => {
    const suggest = vi.fn(async () => ({ ok: true, name: 'x' }))
    stubApi({ suggestFileName: suggest })
    expect(await suggestFileName({ content: '   \n  ' })).toBe('')
    expect(suggest).not.toHaveBeenCalled()
  })

  it('hands over exactly the text it was given — the caller decides what the model may see', async () => {
    const suggest = vi.fn(async () => ({ ok: true, name: 'x' }))
    stubApi({ suggestFileName: suggest })
    // a projected view: the withheld words are already gone, replaced by the
    // reader's own label, so the model can still name the file sensibly
    await suggestFileName({ content: 'Customer phone: {{customer phone}}' })
    expect(suggest).toHaveBeenCalledWith({
      content: 'Customer phone: {{customer phone}}',
    })
  })

  it('trims the name it hands back', async () => {
    stubApi({ suggestFileName: vi.fn(async () => ({ ok: true, name: '  Quarterly plan  ' })) })
    expect(await suggestFileName({ content: 'Some document text' })).toBe('Quarterly plan')
  })
})
