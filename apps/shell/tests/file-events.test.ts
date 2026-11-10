import { describe, expect, it, vi } from 'vitest'

import { notifyFilesChanged, onFilesChanged } from '../src/renderer/src/file-events'

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
