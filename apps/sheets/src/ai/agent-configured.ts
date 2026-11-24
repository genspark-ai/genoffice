import {
  AI_PROVIDERS,
  normalizeCustomProfile,
  type AiSettings,
} from '@genoffice/ai-provider/browser'
import { isOfflinePrompt } from './deterministic-planner'

/**
 * Whether the selected provider can serve a real agent run — the gate between
 * runAgent and the offline deterministic micro-DSL planner. Mirrors the
 * main-process admission checks (sheets-main aiChat/agent handlers): CLI-backed
 * providers (Codex) authenticate through the local CLI and may leave the model
 * empty (the CLI picks the account default), so neither an API key nor a model
 * id is required for them.
 */
export function agentConfigured(settings: AiSettings | null | undefined): boolean {
  if (!settings) return false
  const config = settings.providers[settings.provider]
  if (!config) return false
  if (AI_PROVIDERS.find((meta) => meta.id === settings.provider)?.needsCliPath) return true
  if (!config.model) return false
  // Genspark's key never lands in the settings file; the main process injects
  // it from the gsk login state. When logged out, requests return an error
  // guiding sign-in — not intercepted here.
  if (settings.provider === 'genspark' || config.apiKey) return true
  // a key named by environment variable is resolved by the main process per request
  return !!normalizeCustomProfile(config.profile)?.apiKeyEnv
}

/**
 * Where a sent message goes. An incomplete setup (no key, no model) keeps the
 * offline planner for the three commands it understands; anything else runs
 * the agent path so the main process answers with the actual reason instead
 * of the planner's syntax hint (genoffice#2967).
 */
export function sendRoute(
  settings: AiSettings | null | undefined,
  instruction: string,
): 'agent' | 'offline' {
  if (agentConfigured(settings)) return 'agent'
  return settings && !isOfflinePrompt(instruction) ? 'agent' : 'offline'
}
