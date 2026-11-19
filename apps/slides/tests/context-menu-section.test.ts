import { beforeEach, describe, expect, it, vi } from 'vitest'
import { buildCtxItems } from '../src/renderer/context-menu-items'
import * as slideActions from '../src/renderer/slide-actions'
import { t } from '../src/renderer/i18n/locale'
import type { ActionCtx } from '../src/renderer/action-context'
import type { SectionInfo } from '../src/shared/ipc'

vi.mock('../src/renderer/konva-adapter', () => ({ isEditableText: () => false }))
vi.mock('../src/renderer/clipboard-actions', () => ({}))
vi.mock('../src/renderer/slide-actions', () => ({
  removeSectionAt: vi.fn(),
  removeAllSections: vi.fn(),
  removeSectionWithSlides: vi.fn(),
  moveSectionDir: vi.fn(),
}))
vi.mock('../src/renderer/show-actions', () => ({}))
vi.mock('../src/renderer/arrange-actions', () => ({}))
vi.mock('../src/renderer/insert-actions', () => ({}))
vi.mock('../src/renderer/picture-edit-actions', () => ({}))
vi.mock('../src/renderer/table-actions', () => ({}))
vi.mock('../src/renderer/style-actions', () => ({}))

// slide 0 is unsectioned; A = slides 1-2, B = slides 3-4
const twoSections: SectionInfo[] = [
  { id: 'A', name: 'Intro', slideIndices: [1, 2] },
  { id: 'B', name: 'Body', slideIndices: [3, 4] },
]

function makeCtx(sectionId: string | null, sections = twoSections, slideCount = 5) {
  return {
    ctxMenu: { kind: 'section', x: 0, y: 0, sectionId },
    slides: Array.from({ length: slideCount }, () => ({})),
    sections,
    selectedIds: [],
    slide: null,
    current: 0,
    setRenamingSec: vi.fn(),
    setCollapsedSecs: vi.fn(),
  } as unknown as ActionCtx & { setRenamingSec: ReturnType<typeof vi.fn> } & {
    setCollapsedSecs: ReturnType<typeof vi.fn>
  }
}

type Key = Parameters<typeof t>[0]
const find = (ctx: ActionCtx, key: Key) => buildCtxItems(ctx).find((i) => i?.label === t(key))

describe('section header context menu', () => {
  beforeEach(() => vi.clearAllMocks())

  it('lists the PowerPoint item set in order', () => {
    const labels = buildCtxItems(makeCtx('A')).map((i) => i && i.label)
    const keys: Array<Key | null> = [
      'appCtxRenameSection',
      'appCtxRemoveSection',
      'appCtxRemoveAllSections',
      'appCtxRemoveSectionSlides',
      null,
      'appCtxMoveSectionUp',
      'appCtxMoveSectionDown',
      null,
      'appCtxCollapseAll',
      'appCtxExpandAll',
    ]
    expect(labels).toEqual(keys.map((k) => (k ? t(k) : null)))
  })

  it('disables Move Up on the first section and Move Down on the last', () => {
    expect(find(makeCtx('A'), 'appCtxMoveSectionUp')?.disabled).toBe(true)
    expect(find(makeCtx('A'), 'appCtxMoveSectionDown')?.disabled).toBe(false)
    expect(find(makeCtx('B'), 'appCtxMoveSectionUp')?.disabled).toBe(false)
    expect(find(makeCtx('B'), 'appCtxMoveSectionDown')?.disabled).toBe(true)
    find(makeCtx('B'), 'appCtxMoveSectionUp')?.onClick?.()
    expect(slideActions.moveSectionDir).toHaveBeenCalledWith(expect.anything(), 'B', 'up')
  })

  it('gives the default lead section the menu with Rename and moves disabled', () => {
    const ctx = makeCtx(null)
    expect(find(ctx, 'appCtxRenameSection')?.disabled).toBe(true)
    expect(find(ctx, 'appCtxMoveSectionUp')?.disabled).toBe(true)
    expect(find(ctx, 'appCtxMoveSectionDown')?.disabled).toBe(true)
    find(ctx, 'appCtxRenameSection')?.onClick?.()
    expect(ctx.setRenamingSec).not.toHaveBeenCalled()
    find(ctx, 'appCtxRemoveSection')?.onClick?.()
    expect(slideActions.removeSectionAt).toHaveBeenCalledWith(expect.anything(), null)
    expect(find(ctx, 'appCtxRemoveSectionSlides')?.disabled).toBe(false)
    find(ctx, 'appCtxRemoveSectionSlides')?.onClick?.()
    expect(slideActions.removeSectionWithSlides).toHaveBeenCalledWith(expect.anything(), null)
  })

  it('renames with the section name prefilled and dispatches the remove commands', () => {
    const ctx = makeCtx('A')
    find(ctx, 'appCtxRenameSection')?.onClick?.()
    expect(ctx.setRenamingSec).toHaveBeenCalledWith({ id: 'A', value: 'Intro' })
    find(ctx, 'appCtxRemoveSection')?.onClick?.()
    expect(slideActions.removeSectionAt).toHaveBeenCalledWith(expect.anything(), 'A')
    find(ctx, 'appCtxRemoveAllSections')?.onClick?.()
    expect(slideActions.removeAllSections).toHaveBeenCalledTimes(1)
    find(ctx, 'appCtxRemoveSectionSlides')?.onClick?.()
    expect(slideActions.removeSectionWithSlides).toHaveBeenCalledWith(expect.anything(), 'A')
  })

  it('disables Remove Section & Slides when it would delete every slide', () => {
    const whole: SectionInfo[] = [{ id: 'A', name: 'All', slideIndices: [0, 1, 2] }]
    expect(find(makeCtx('A', whole, 3), 'appCtxRemoveSectionSlides')?.disabled).toBe(true)
    // An empty trailing section leaves every slide in the lead group
    const empty: SectionInfo[] = [{ id: 'A', name: 'Tail', slideIndices: [] }]
    expect(find(makeCtx(null, empty, 3), 'appCtxRemoveSectionSlides')?.disabled).toBe(true)
    expect(find(makeCtx('A', empty, 3), 'appCtxRemoveSectionSlides')?.disabled).toBe(false)
  })

  it('Collapse All / Expand All set the whole rail at once', () => {
    const ctx = makeCtx('A')
    find(ctx, 'appCtxCollapseAll')?.onClick?.()
    expect([...ctx.setCollapsedSecs.mock.calls[0]![0]]).toEqual(['A', 'B'])
    find(ctx, 'appCtxExpandAll')?.onClick?.()
    expect(ctx.setCollapsedSecs.mock.calls[1]![0].size).toBe(0)
  })
})
