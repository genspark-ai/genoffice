import { createWebSkill, type AgentSkill } from '@genoffice/agent-core'
import { captureRect } from '@genoffice/ui'

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
        instructions = await window.desktop.aiInstructionsPrompt('docx')
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
    surface: 'docx',
    hasUserSkills: () => hasSkills,
    instructionsPrompt: current,
    bridge: {
      browsePage: (url, opts) => window.desktop.aiBrowsePage(url, opts),
      extractPages: (urls, advanced) => window.desktop.aiExtractPages(urls, advanced),
      loadSkill: (id) => window.desktop.aiSkillBody('docx', id),
      remember: (text) => window.desktop.aiRemember(text),
      forget: (text) => window.desktop.aiForget(text),
      viewPage: async () => {
        const base64 = await window.desktop.aiCapturePage(captureRect('.doc-page'))
        return base64
          ? { ok: true, images: [base64], label: 'the page' }
          : { ok: false, error: 'Nothing could be rendered right now.' }
      },
    },
  })

  return { skill, refresh }
}
