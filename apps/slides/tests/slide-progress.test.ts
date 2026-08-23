/**
 * Generation progress checklist (todolist-style reminder) verification:
 * after plan_deck plans N pages, buildContext() should report them all unfinished; when a
 * generate_deck run leaves pages failed (cloud generation skips a page after its one retry),
 * each buildContext() round should inject a progress hint ("X still missing, page Y unfinished"),
 * mechanically reminding the AI to top the deck up with another generate_deck append (not
 * relying on a one-shot prompt constraint). Once complete, progress shows done and stops nagging.
 */
import { describe, it, expect } from 'vitest'
import { createSlidesSkill, type DeckAccess } from '../src/renderer/ai/slides-skill'
import type { RenderSlide } from '@genoffice/pptx-render'
import type { AgentToolCall } from '../src/shared/ipc'

/** Minimal mock: cloud page generation + landing that tracks the page count. */
function makeAccess(opts?: { failPages?: number[] }) {
  const failPages = new Set(opts?.failPages ?? [])
  let pages = 0
  const access: DeckAccess = {
    getSlides: () => [] as RenderSlide[],
    getCurrent: () => 0,
    getSelectedIds: () => [],
    applySlide: () => {},
    applyDeck: () => {},
    fitWidthPx: 1280,
    retryBackoffMs: 0,
    isCloudPageGenEnabled: async () => true,
    generatePageCloud: async (args) => {
      if (failPages.has(args.pageIndex)) return { ok: false, error: 'mock fail' }
      return { ok: true, marker: `PAGE${args.pageIndex}:${args.title}` }
    },
    landGeneratedPages: async (
      pageMarkers,
      mode = 'replace',
      _deckName?: string,
      insertAt?: number,
    ) => {
      if (mode === 'insert_at') {
        pages += 1
        return { ok: true, pages, insertedIndex: insertAt }
      }
      if (mode === 'append') {
        const from = pages
        pages += pageMarkers.length
        return { ok: true, pages, appendedFrom: from }
      }
      pages = pageMarkers.length
      return { ok: true, pages }
    },
  }
  return { access, getPages: () => pages }
}

const planCall = (n: number): AgentToolCall => ({
  id: 'c-plan',
  name: 'plan_deck',
  input: {
    core_hook: 'Economic data doubled in a decade',
    style: 'Deep navy + gold data visualization',
    pages: Array.from({ length: n }, (_, i) => ({
      title: `Page ${i + 1} Title`,
      brief: 'brief',
      layout: 'data',
      image_queries: [],
    })),
  },
})

const deckCall = (n: number, insertMode?: 'replace' | 'append'): AgentToolCall => ({
  id: 'c-deck',
  name: 'generate_deck',
  input: {
    core_hook: 'Economic data doubled in a decade',
    style: 'Deep navy + gold data visualization',
    ...(insertMode ? { insert_mode: insertMode } : {}),
    pages: Array.from({ length: n }, (_, i) => ({
      title: `Page ${i + 1} Title`,
      brief: 'brief',
      layout: 'data',
      image_queries: [],
    })),
  },
})

describe('generation progress checklist reminds the AI to fill gaps', () => {
  it('plan_deck plans 5 pages, none generated → buildContext reports all 5 incomplete', async () => {
    const { access } = makeAccess()
    const skill = createSlidesSkill(access)

    // No progress hint before planning
    expect(skill.buildContext?.()).not.toContain('generation-progress')

    await skill.executeTool(planCall(5))
    const afterPlan = skill.buildContext?.() ?? ''
    expect(afterPlan).toContain('generation-progress')
    expect(afterPlan).toContain('5 pages planned')
    expect(afterPlan).toContain('0 generated')
  })

  it('generate_deck skips 4 of 5 failed pages → tool result and later context name the incomplete pages', async () => {
    const { access } = makeAccess({ failPages: [2, 3, 4, 5] })
    const skill = createSlidesSkill(access)

    const res = (await skill.executeTool(deckCall(5))) as { output: string }
    expect(res.output).toContain('still failed after retry')
    expect(res.output).toContain('insert_mode:"append"')

    const ctx = skill.buildContext?.() ?? ''
    expect(ctx).toContain('1 generated')
    expect(ctx).toContain('4 still missing')
    expect(ctx).toContain('page 2')
    expect(ctx).toContain('page 5')
    expect(ctx).toContain('Page 2 Title')
  })

  it('after filling gaps with generate_deck(append) → progress shows complete, no more fill-gap reminder', async () => {
    const { access, getPages } = makeAccess()
    const skill = createSlidesSkill(access)

    await skill.executeTool(planCall(5))
    await skill.executeTool(deckCall(5))
    expect(getPages()).toBe(5)

    const ctx = skill.buildContext?.() ?? ''
    expect(ctx).toContain('all generated')
    expect(ctx).not.toContain('still missing')
  })
})
