import { existsSync } from 'node:fs'
import { readFile } from 'node:fs/promises'
import { dirname } from 'node:path'
import { describe, expect, it } from 'vitest'
import { printSlidesHtml, type PrintWindow } from '../src/main/print-window'

class TestPrintWindow implements PrintWindow {
  loadedPath: string | null = null
  loadedHtml = ''
  destroyed = false
  directoryExistedAtDestroy = false
  shown = false
  focused = false
  failLoad = false
  printResult = { success: true, reason: '' }

  async loadFile(path: string): Promise<void> {
    this.loadedPath = path
    this.loadedHtml = await readFile(path, 'utf8')
    if (this.failLoad) throw new Error('load failed')
  }

  webContents = {
    executeJavaScript: async (): Promise<void> => {},
    print: (
      _options: { silent: boolean; printBackground: boolean },
      callback: (success: boolean, reason: string) => void,
    ): void => {
      callback(this.printResult.success, this.printResult.reason)
    },
  }

  show(): void {
    this.shown = true
  }

  focus(): void {
    this.focused = true
  }

  isDestroyed(): boolean {
    return this.destroyed
  }

  destroy(): void {
    this.directoryExistedAtDestroy = !!this.loadedPath && existsSync(dirname(this.loadedPath))
    this.destroyed = true
  }
}

describe('slide print window', () => {
  it('loads a large document from a temporary HTML file and removes it after printing', async () => {
    const win = new TestPrintWindow()
    const html = '<html>' + 'x'.repeat(3 * 1024 * 1024) + '</html>'
    expect(await printSlidesHtml(html, win, 'win32')).toEqual({ ok: true })
    expect(win.loadedPath).toMatch(/genoffice-slides-print-.*slides\.html$/)
    expect(win.loadedHtml).toBe(html)
    expect(win.shown).toBe(true)
    expect(win.focused).toBe(true)
    expect(win.destroyed).toBe(true)
    expect(win.directoryExistedAtDestroy).toBe(true)
    expect(existsSync(dirname(win.loadedPath!))).toBe(false)
  })

  it('cleans up on cancellation and load failure', async () => {
    const canceled = new TestPrintWindow()
    canceled.printResult = { success: false, reason: 'Print job canceled' }
    expect(await printSlidesHtml('<html/>', canceled)).toEqual({ ok: false })
    expect(canceled.destroyed).toBe(true)
    expect(existsSync(dirname(canceled.loadedPath!))).toBe(false)

    const failed = new TestPrintWindow()
    failed.failLoad = true
    expect(await printSlidesHtml('<html/>', failed)).toEqual({
      ok: false,
      error: 'load failed',
    })
    expect(failed.destroyed).toBe(true)
    expect(existsSync(dirname(failed.loadedPath!))).toBe(false)
  })
})
