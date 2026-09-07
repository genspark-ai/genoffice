import { describe, expect, it, vi } from 'vitest'
import { exportStudioPayload, hmlUnavailableMessage } from '../src/renderer/export-payload'

function studio(partial: {
  hmlSavable?: boolean
  blockers?: { code: string; xmlPath: string; message: string }[]
  exportHml?: () => Promise<Uint8Array>
  getHmlSaveState?: () => Promise<{ hmlSavable: boolean; blockers: { code: string; xmlPath: string; message: string }[] }>
}) {
  const exportHml = partial.exportHml ?? vi.fn(async () => new Uint8Array([3]))
  return {
    exportHwp: vi.fn(async () => new Uint8Array([1])),
    exportHwpx: vi.fn(async () => new Uint8Array([2])),
    exportHml,
    getHmlSaveState:
      partial.getHmlSaveState ??
      vi.fn(async () => ({
        hmlSavable: partial.hmlSavable ?? true,
        blockers: partial.blockers ?? [],
      })),
  }
}

describe('exportStudioPayload', () => {
  it('exports HML only when getHmlSaveState says it is savable', async () => {
    const ready = studio({ hmlSavable: true })
    const ok = await exportStudioPayload(ready)
    expect(ok.hml).toEqual(new Uint8Array([3]))
    expect(ready.exportHml).toHaveBeenCalledOnce()

    const blocked = studio({
      hmlSavable: false,
      blockers: [{ code: 'HML_SOURCE_REQUIRED', xmlPath: '/HWPML', message: 'needs HML source' }],
    })
    const skipped = await exportStudioPayload(blocked)
    expect(skipped.hml).toBeUndefined()
    expect(skipped.hmlBlockers?.[0]?.code).toBe('HML_SOURCE_REQUIRED')
    expect(blocked.exportHml).not.toHaveBeenCalled()
  })

  it('skips HML when getHmlSaveState itself fails', async () => {
    const broken = studio({
      getHmlSaveState: vi.fn(async () => {
        throw new Error('no state')
      }),
    })
    const payload = await exportStudioPayload(broken)
    expect(payload.hml).toBeUndefined()
    expect(broken.exportHml).not.toHaveBeenCalled()
  })
})

describe('hmlUnavailableMessage', () => {
  it('includes blocker text when present', () => {
    expect(hmlUnavailableMessage()).toBe('This document cannot be saved as HML.')
    expect(
      hmlUnavailableMessage([{ code: 'X', xmlPath: '/', message: 'needs HML source' }]),
    ).toBe('This document cannot be saved as HML. needs HML source')
  })
})
