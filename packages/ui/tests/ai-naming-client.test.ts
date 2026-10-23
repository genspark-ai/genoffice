/** @vitest-environment jsdom */
import { afterEach, describe, expect, it, vi } from 'vitest'
import { nameForSave } from '../src/ai-naming-client'

/**
 * The fallback contract is the whole point of this helper.
 *
 * A file name is the one thing a reader cannot get back easily, so every way
 * the naming attempt can go wrong has to end with the name the document already
 * had. These cases are the ways it can go wrong.
 */

function stubApi(overrides: Record<string, unknown> = {}): void {
  ;(window as unknown as { aiOffice: unknown }).aiOffice = {
    suggestFileName: vi.fn(async () => ({ ok: true, name: 'Quarterly plan' })),
    getFileNamingEnabled: vi.fn(async () => true),
    setFileNamingEnabled: vi.fn(async (on: boolean) => on),
    ...overrides,
  }
}

afterEach(() => {
  delete (window as unknown as { aiOffice?: unknown }).aiOffice
  vi.restoreAllMocks()
})

const ask = (over: Partial<Parameters<typeof nameForSave>[0]> = {}) =>
  nameForSave({
    content: 'Some document text',
    trigger: 'first-save',
    fallback: 'Untitled.docx',
    ...over,
  })

describe('nameForSave', () => {
  it('uses the name the model returned', async () => {
    stubApi()
    expect(await ask()).toBe('Quarterly plan')
  })

  it('keeps the fallback when the setting is off', async () => {
    stubApi({ getFileNamingEnabled: vi.fn(async () => false) })
    expect(await ask()).toBe('Untitled.docx')
  })

  it('asks anyway on the manual trigger, because the reader asked for it', async () => {
    const suggest = vi.fn(async () => ({ ok: true, name: 'Manual name' }))
    stubApi({ suggestFileName: suggest, getFileNamingEnabled: vi.fn(async () => false) })
    expect(await ask({ trigger: 'manual' })).toBe('Manual name')
    expect(suggest).toHaveBeenCalled()
  })

  it('keeps the fallback when the model declines', async () => {
    stubApi({ suggestFileName: vi.fn(async () => ({ ok: false, reason: 'no-content' })) })
    expect(await ask()).toBe('Untitled.docx')
  })

  // the bug this whole helper exists to be robust against: an empty name
  // arriving as a *success* would rename the reader's file to nothing
  it('keeps the fallback when a success carries an empty name', async () => {
    stubApi({ suggestFileName: vi.fn(async () => ({ ok: true, name: '' })) })
    expect(await ask()).toBe('Untitled.docx')
  })

  it('keeps the fallback when a success carries only whitespace', async () => {
    stubApi({ suggestFileName: vi.fn(async () => ({ ok: true, name: '   ' })) })
    expect(await ask()).toBe('Untitled.docx')
  })

  it('keeps the fallback when the call rejects', async () => {
    stubApi({
      suggestFileName: vi.fn(async () => {
        throw new Error('no handler')
      }),
    })
    expect(await ask()).toBe('Untitled.docx')
  })

  it('keeps the fallback when the bridge has no naming channel at all', async () => {
    ;(window as unknown as { aiOffice: unknown }).aiOffice = {}
    expect(await ask()).toBe('Untitled.docx')
  })

  it('never asks about an empty document', async () => {
    const suggest = vi.fn(async () => ({ ok: true, name: 'x' }))
    stubApi({ suggestFileName: suggest })
    expect(await ask({ content: '   \n  ' })).toBe('Untitled.docx')
    expect(suggest).not.toHaveBeenCalled()
  })

  it('hands over exactly the text it was given — the caller decides what the model may see', async () => {
    const suggest = vi.fn(async () => ({ ok: true, name: 'x' }))
    stubApi({ suggestFileName: suggest })
    // a projected view: the withheld words are already gone, replaced by the
    // reader's own label, so the model can still name the file sensibly
    await ask({ content: 'Customer phone: {{customer phone}}', filePath: '/tmp/a.docx' })
    expect(suggest).toHaveBeenCalledWith({
      content: 'Customer phone: {{customer phone}}',
      trigger: 'first-save',
      filePath: '/tmp/a.docx',
    })
  })

  it('sends a null filePath when the document has never been saved', async () => {
    const suggest = vi.fn(async () => ({ ok: true, name: 'x' }))
    stubApi({ suggestFileName: suggest })
    await ask({ filePath: undefined })
    expect(suggest).toHaveBeenCalledWith(expect.objectContaining({ filePath: null }))
  })
})
