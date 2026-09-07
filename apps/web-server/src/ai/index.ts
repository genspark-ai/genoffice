/**
 * AI capability surface.
 *
 * Sub-files own specific AI sub-domains; this entry just re-exports the
 * registration helpers and the agent-loop SSE helper used by the HTTP
 * layer in `index.ts`.
 */
export { callMiniMax, generateAIResponse, generateAgentResponse } from './minimax.js'
export type { MiniMaxMessage } from './minimax.js'

import { registerAiCoreHandlers } from './chat.js'
import { registerDocAiSkillHandlers } from './doc-skill.js'
import { registerSheetAiSkillHandlers } from './sheet-skill.js'
import { registerSlideAiSkillHandlers } from './slide-skill.js'

/**
 * Register every AI handler in the shared registry. Call once at boot.
 */
export function registerAiHandlers(): void {
  registerAiCoreHandlers()
  registerDocAiSkillHandlers()
  registerSheetAiSkillHandlers()
  registerSlideAiSkillHandlers()
}
