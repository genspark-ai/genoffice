/**
 * Browser stub for codex-app-server.
 *
 * Used by stream.ts / chat.ts via the alias below so that renderer bundles
 * never pull in the Node-only `codex-app-server.ts` (which imports
 * node:fs/promises stat etc.). The renderer should never actually reach
 * these functions; if it does, throw a clear error so the caller can route
 * through the IPC bridge instead.
 */
import type { AgentMessage, AgentToolDef } from './agent-protocol'
import type { AiChatResponse, AiProviderConfig } from './types'
import type { StreamCallbacks } from './protocols/shared'

export async function streamCodexAppServer(
  _config: AiProviderConfig,
  _system: string,
  _messages: AgentMessage[],
  _tools: AgentToolDef[],
  _maxTokens: number,
  _cb: StreamCallbacks,
): Promise<void> {
  throw new Error(
    'codex-app-server streaming is not available in the renderer; ' +
      'route requests through the main process IPC bridge.',
  )
}

export async function chatCodexAppServer(
  _config: AiProviderConfig,
  _system: string,
  _user: string,
  _signal: AbortSignal,
): Promise<AiChatResponse> {
  throw new Error(
    'codex-app-server chat is not available in the renderer; ' +
      'route requests through the main process IPC bridge.',
  )
}
