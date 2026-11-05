import { describe, expect, it, vi } from 'vitest'
import type { RenderSlide } from '@genoffice/pptx-render'
import {
  mergeBroadcastUpdate,
  mergeSlideUpdate,
  setDeckRefresh,
  tryMergeSlideUpdate,
} from '../src/shared/slide-update'

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

describe('reply naming a slide this window has not seen', () => {
  it('throws without a registered refresh', () => {
    setDeckRefresh(null)
    expect(() => mergeSlideUpdate([slide('a')], { slides: [], order: ['a', 'missing'] })).toThrow(
      /missing/,
    )
  })

  it('keeps the deck and refetches when a refresh is registered', async () => {
    const refresh = vi.fn()
    setDeckRefresh(refresh)
    try {
      const current = [slide('a')]
      expect(mergeSlideUpdate(current, { slides: [], order: ['a', 'missing'] })).toBe(current)
      await Promise.resolve()
      expect(refresh).toHaveBeenCalledOnce()
    } finally {
      setDeckRefresh(null)
    }
  })
})

describe('tryMergeSlideUpdate', () => {
  it('returns null instead of a stale deck when a slide is unseen', () => {
    expect(tryMergeSlideUpdate([slide('a')], { slides: [], order: ['a', 'missing'] })).toBeNull()
  })

  it('merges when every slide is known', () => {
    const a = slide('a')
    const b = slide('b')
    expect(tryMergeSlideUpdate([a], { slides: [b], order: ['b', 'a'] })).toEqual([b, a])
  })
})
