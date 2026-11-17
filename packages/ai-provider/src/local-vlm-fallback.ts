import { resolveAgyCliPath, type AgyRunDeps } from './agy-cli'
import { agyUsableForDefaults } from './agy-detect'
import { getLocalVlmRuntime } from './local-vlm-manager'
import { LOCAL_VLM_DEFAULT_MODEL } from './local-vlm-meta'
import type { AiProviderConfig } from './types'

export interface LocalVlmFallbackDeps {
  resolveCli?: (cliPath?: string) => Promise<string>
  agyUsable?: (cliPath?: string) => Promise<boolean>
}

/**
 * Returns true when Antigravity CLI (`agy`) is unavailable (not installed or not signed in)
 * and the local LFM3-VL-2B runtime is installed, enabled, and configured to take over.
 */
export async function shouldFallbackFromAgyToLocalVlm(
  cliPath?: string,
  deps: LocalVlmFallbackDeps = {},
): Promise<boolean> {
  const runtime = getLocalVlmRuntime()
  if (!runtime || !runtime.shouldFallbackFromAgy()) return false
  const resolveCli =
    deps.resolveCli ?? ((p?: string) => resolveAgyCliPath(p, undefined as unknown as AgyRunDeps))
  try {
    await resolveCli(cliPath)
  } catch {
    // Fast path: `agy` executable is not present on this machine at all.
    return true
  }
  const checkUsable = deps.agyUsable ?? ((p?: string) => agyUsableForDefaults(p))
  try {
    return !(await checkUsable(cliPath))
  } catch {
    return true
  }
}

export async function ensureLocalVlmEndpoint(
  signal?: AbortSignal,
): Promise<{ baseUrl: string; model: string }> {
  const runtime = getLocalVlmRuntime()
  if (!runtime) {
    throw new Error(
      'Local AI (LFM3-VL-2B) runtime is not initialized. Open Settings → AI Model → Local AI (LFM3-VL-2B).',
    )
  }
  return runtime.ensureRunning(signal)
}

export function localVlmRequestConfig(config?: Partial<AiProviderConfig>): AiProviderConfig {
  return {
    apiKey: '',
    model: config?.model?.trim() || LOCAL_VLM_DEFAULT_MODEL,
  }
}
