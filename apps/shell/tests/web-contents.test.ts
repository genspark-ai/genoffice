import { describe, expect, it, vi } from 'vitest'

vi.mock('electron', () => ({ session: { defaultSession: {} } }))

import {
  registerWebTab,
  setWebPrintToPDFDelegate,
  unregisterWebTab,
  WebTabContents,
} from '../src/web/web-contents'

describe('web tab contents', () => {
  it('emits destroyed exactly once and runs cleanups', () => {
    const tab = new WebTabContents('/docs/?path=/x.docx')
    const calls: string[] = []
    tab.on('destroyed', () => calls.push('a'))
    tab.once('destroyed', () => calls.push('b'))
    tab.destroy()
    tab.destroy()
    expect(calls).toEqual(['a', 'b'])
  })

  it('pushes through deliver and reports its URL', () => {
    const tab = new WebTabContents('/docs/?path=/x.docx')
    const pushed: unknown[] = []
    tab.deliver = (channel, args) => pushed.push([channel, args])
    tab.send('docs:opened', { ok: true })
    expect(pushed).toEqual([['docs:opened', [{ ok: true }]]])
    expect(tab.getURL()).toBe('/docs/?path=/x.docx')
    expect(tab.isDestroyed()).toBe(false)
  })

  it('delegates printToPDF to the installed pipeline', async () => {
    const tab = new WebTabContents('/docs/?path=/saved.docx')
    setWebPrintToPDFDelegate(async (sourceUrl) => {
      expect(sourceUrl).toContain('/saved.docx')
      return Buffer.from('pdf-bytes')
    })
    await expect(tab.printToPDF({ printBackground: true })).resolves.toEqual(
      Buffer.from('pdf-bytes'),
    )
    setWebPrintToPDFDelegate(null)
    await expect(tab.printToPDF()).rejects.toThrow(/not available in the browser/)
  })

  it('refuses native printing loudly', () => {
    const tab = new WebTabContents('/docs/')
    expect(() => tab.print()).toThrow(/browser print dialog/)
  })

  it('registers and unregisters cleanly', () => {
    const tab = new WebTabContents('/docs/')
    registerWebTab(tab)
    unregisterWebTab(tab)
  })
})
