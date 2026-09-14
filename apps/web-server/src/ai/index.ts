/**
 * AI capability surface.
 *
 * Sub-files own specific AI sub-domains; this entry re-exports the
 * registration helpers and the streaming core (settings + provider
 * streaming) used by both the IPC layer and the SSE endpoint.
 */
export { callMiniMax, generateAIResponse } from './minimax.js'
export type { MiniMaxMessage } from './minimax.js'

import {
  registerAiCoreHandlers,
  runProviderStream,
  AI_STREAM_SESSIONS,
} from './chat.js'
import type { AiSettings } from '@genoffice/ai-provider'
import { registerDocAiSkillHandlers } from './doc-skill.js'
import { registerSheetAiSkillHandlers } from './sheet-skill.js'
import { registerSlideAiSkillHandlers } from './slide-skill.js'
import { registerAiMediaSkillHandlers } from './media-skill.js'

// Live settings reference for the SSE handler — it pulls the latest
// persisted AiSettings from the chat module without going through IPC.
import * as chatModule from './chat.js'
export const aiSettings: AiSettings = (chatModule as { aiSettings: AiSettings }).aiSettings

export { runProviderStream, AI_STREAM_SESSIONS }

/**
 * Register every AI handler in the shared registry. Call once at boot.
 *
 * Renderer-side editing skills (docs/sheets/slides) intentionally return
 * `WEB_UNSUPPORTED` here; the matching renderer-side skill is the source of
 * truth. `registerAiMediaSkillHandlers` exposes the real, persisted slide
 * style template storage used by the slide style library.
 */
export function registerAiHandlers(): void {
  registerAiCoreHandlers()
  registerDocAiSkillHandlers()
  registerSheetAiSkillHandlers()
  registerSlideAiSkillHandlers()
  registerAiMediaSkillHandlers()
}
