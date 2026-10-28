/**
 * @vitest-environment jsdom
 */
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { describe, expect, it, vi } from 'vitest'

import { HOME_CHANNELS } from '../src/shared/home-api'
import {
  bindRecentsChanged,
  notifyFilesChanged,
  onFilesChanged,
} from '../src/renderer/src/file-events'

describe('file-events', () => {
  it('delivers a notice to every live subscriber and stops after unsubscribe', () => {
    const a = vi.fn()
    const b = vi.fn()
    const offA = onFilesChanged(a)
    const offB = onFilesChanged(b)
    notifyFilesChanged()
    expect(a).toHaveBeenCalledTimes(1)
    expect(b).toHaveBeenCalledTimes(1)
    offA()
    notifyFilesChanged()
    expect(a).toHaveBeenCalledTimes(1)
    expect(b).toHaveBeenCalledTimes(2)
    offB()
    notifyFilesChanged()
    expect(b).toHaveBeenCalledTimes(2)
  })
})

describe('bindRecentsChanged', () => {
  it('re-fires the main-process recents notice through the in-renderer event', () => {
    let pushed: (() => void) | null = null
    const off = vi.fn()
    ;(window as { aiOffice: unknown }).aiOffice = {
      onRecentsChanged: (handler: () => void) => {
        pushed = handler
        return off
      },
    }
    const notified = vi.fn()
    const detach = onFilesChanged(notified)
    const unbind = bindRecentsChanged()
    pushed!()
    expect(notified).toHaveBeenCalledTimes(1)
    unbind()
    expect(off).toHaveBeenCalledTimes(1)
    detach()
    pushed!()
    expect(notified).toHaveBeenCalledTimes(1)
  })
})

/**
 * The channel is broadcast by the docs main process (its pushRecent funnel) as
 * a string literal — docs cannot import the shell's shared module — so both
 * sides pin the same name and this test is the drift guard for the shell half.
 */
describe('recents-changed channel wiring', () => {
  it('uses the channel name the docs main process broadcasts', () => {
    expect(HOME_CHANNELS.recentsChanged).toBe('home:recents-changed')
  })

  it('the preload exposes the channel to the renderer', () => {
    const preload = readFileSync(join(import.meta.dirname, '../src/preload/index.ts'), 'utf8')
    expect(preload).toContain('onRecentsChanged(handler)')
    expect(preload).toContain('ipcRenderer.on(HOME_CHANNELS.recentsChanged, listener)')
  })

  it('Home mounts the bridge so a save re-pulls the list without waiting for focus', () => {
    const home = readFileSync(join(import.meta.dirname, '../src/renderer/src/Home.tsx'), 'utf8')
    expect(home).toMatch(/useEffect\(\(\) => bindRecentsChanged\(\), \[\]\)/)
  })
})
