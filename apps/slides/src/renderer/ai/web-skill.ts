import { createWebSkill, type AgentSkill } from '@genoffice/agent-core'
import type { RenderSlide } from '@genoffice/pptx-render'
import { renderSlidesToPngBase64 } from '../export-render'

/**
 * The shared browse / extract / load_skill skill, bound to this app's preload
 * bridge and surface.
 *
 * The instructions prompt (user rules + skill catalogue) is assembled in the
 * main process. It is cached here and refreshed in the background whenever it
 * goes stale, so an edit in the settings window reaches the next turn without
 * reopening the document — and without every run site having to remember to
 * await a refresh first.
 */

/** how long a cached instructions prompt is trusted before a background refetch */
const STALE_AFTER_MS = 3_000

/**
 * Live accessors rather than values: the skill is built once per panel, while
 * the deck it renders changes under it on every edit.
 */
export interface SlidesViewSource {
  slides: () => RenderSlide[]
  /** decoded picture cache the off-screen renderer needs to draw images */
  images: () => Map<string, HTMLImageElement>
  /** 0-based index of the slide on screen, for a call that names no page */
  current: () => number
}

export function createAppWebSkill(view: SlidesViewSource): {
  skill: AgentSkill
  refresh: () => Promise<void>
} {
  let instructions = ''
  let hasSkills = false
  let fetchedAt = 0
  let inFlight: Promise<void> | null = null

  const refresh = (): Promise<void> => {
    inFlight ??= (async () => {
      try {
        instructions = await window.slidesApi.aiInstructionsPrompt('pptx')
        hasSkills = instructions.includes('## User skills')
        fetchedAt = Date.now()
      } catch {
        // instructions are an enhancement; a failure must never block a turn
        instructions = ''
        hasSkills = false
      } finally {
        inFlight = null
      }
    })()
    return inFlight
  }

  /** read the cache, kicking off a refetch when it has aged out */
  const current = (): string => {
    if (Date.now() - fetchedAt > STALE_AFTER_MS) void refresh()
    return instructions
  }

  void refresh()

  const skill = createWebSkill({
    surface: 'pptx',
    hasUserSkills: () => hasSkills,
    instructionsPrompt: current,
    bridge: {
      browsePage: (url, opts) => window.slidesApi.aiBrowsePage(url, opts),
      extractPages: (urls, advanced) => window.slidesApi.aiExtractPages(urls, advanced),
      loadSkill: (id) => window.slidesApi.aiSkillBody('pptx', id),
      remember: (text) => window.slidesApi.aiRemember(text),
      forget: (text) => window.slidesApi.aiForget(text),
      viewPageCount: () => view.slides().length,
      // Rendered off-screen by the same path as PPTX export, not captured from
      // the canvas: it reaches any slide without navigating, leaves the user's
      // view alone, and shows the slide as it will actually print rather than
      // at whatever zoom and scroll offset the window happens to be at.
      viewPage: async (page) => {
        const deck = view.slides()
        const index = page === undefined ? view.current() : page - 1
        const slide = deck[index]
        if (!slide) return { ok: false, error: 'That slide does not exist.' }
        try {
          const [png] = await renderSlidesToPngBase64([slide], view.images(), 1)
          return png
            ? { ok: true, images: [png], label: `slide ${index + 1} of ${deck.length}` }
            : { ok: false, error: 'The slide could not be rendered.' }
        } catch (err) {
          return { ok: false, error: err instanceof Error ? err.message : String(err) }
        }
      },
    },
  })

  return { skill, refresh }
}
