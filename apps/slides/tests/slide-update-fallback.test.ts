import { describe, expect, it, vi } from 'vitest'
import type { RenderSlide } from '@genoffice/pptx-render'
import { mergeBroadcastUpdate } from '../src/shared/slide-update'

const slide = (partPath: string) => ({ partPath }) as RenderSlide

describe('partial broadcast fallback', () => {
  it.each([null, [], [slide('other')]])(
    'refreshes when the base is missing: %s',
    async (current) => {
      const refresh = vi.fn()
      expect(
        mergeBroadcastUpdate(
          current,
          { slides: [slide('changed')], order: ['changed', 'missing'] },
          refresh,
        ),
      ).toBe(current)
      await Promise.resolve()
      expect(refresh).toHaveBeenCalledOnce()
    },
  )

  it('merges a complete base without fetching', async () => {
    const refresh = vi.fn()
    const unchanged = slide('unchanged')
    const changed = slide('changed')
    expect(
      mergeBroadcastUpdate(
        [unchanged],
        { slides: [changed], order: ['unchanged', 'changed'] },
        refresh,
      ),
    ).toEqual([unchanged, changed])
    await Promise.resolve()
    expect(refresh).not.toHaveBeenCalled()
  })
})
