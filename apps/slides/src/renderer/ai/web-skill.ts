import { createWebSkill, type AgentSkill } from '@genoffice/agent-core'

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

export function createAppWebSkill(): { skill: AgentSkill; refresh: () => Promise<void> } {
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
    },
  })

  return { skill, refresh }
}
