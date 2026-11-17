import type { AiProviderId, AiProviderMeta } from './types'

/**
 * Browser-safe metadata, settings schema, and defaults for the local
 * Liquid AI LFM3-VL-2B (Q4_K_M) vision-language model served via llama.cpp.
 * Kept free of Node imports so renderer bundles can safely import it.
 */

export const LOCAL_VLM_PROVIDER_ID = 'local-vlm' satisfies AiProviderId
export const LOCAL_VLM_DEFAULT_MODEL = 'LFM3-VL-2B-Q4_K_M'
export const LOCAL_VLM_HF_REPO = 'LiquidAI/LFM3-VL-2B-GGUF'
export const LOCAL_VLM_MODEL_FILENAME = 'LFM3-VL-2B-Q4_K_M.gguf'
export const LOCAL_VLM_MMPROJ_FILENAME = 'mmproj-LFM3-VL-2B-Q8_0.gguf'

export const LOCAL_VLM_MODEL_URL = `https://huggingface.co/${LOCAL_VLM_HF_REPO}/resolve/main/${LOCAL_VLM_MODEL_FILENAME}`
export const LOCAL_VLM_MMPROJ_URL = `https://huggingface.co/${LOCAL_VLM_HF_REPO}/resolve/main/${LOCAL_VLM_MMPROJ_FILENAME}`
export const LOCAL_VLM_LLAMA_RELEASES_API =
  'https://api.github.com/repos/ggml-org/llama.cpp/releases/latest'

export const LOCAL_VLM_DEFAULT_CTX_SIZE = 32768
export const LOCAL_VLM_MIN_CTX_SIZE = 2048
export const LOCAL_VLM_MAX_CTX_SIZE = 32768
export const LOCAL_VLM_MIN_RAM_MB = 2200
export const LOCAL_VLM_IDLE_TIMEOUT_MS = 180_000

/** Recommended LiquidAI sampling defaults for LFM3-VL-2B GGUF. */
export const LOCAL_VLM_SAMPLING_DEFAULTS: Readonly<Record<string, number>> = {
  temperature: 0.1,
  top_p: 1.0,
  top_k: 50,
  min_p: 0.0,
  repeat_penalty: 1.1,
}

export const LOCAL_VLM_PROVIDER_META: AiProviderMeta = {
  id: LOCAL_VLM_PROVIDER_ID,
  label: 'Local AI (LFM3-VL-2B)',
  models: [LOCAL_VLM_DEFAULT_MODEL],
  defaultModel: LOCAL_VLM_DEFAULT_MODEL,
  keyPlaceholder: '',
}

export const LOCAL_VLM_CAPABILITIES = {
  auth: 'local-vlm',
  vision: true,
  tools: true,
} as const

export function isLocalVlmProvider(provider: string): boolean {
  return provider === LOCAL_VLM_PROVIDER_ID
}

export const LOCAL_VLM_SETTINGS_KEY = 'localVlm'

export interface LocalVlmSettings {
  /** Whether local VLM runtime is enabled for direct use and fallback. */
  enabled: boolean
  /** Automatically fall back to LFM3-VL-2B when Antigravity CLI (agy) is unavailable. */
  autoFallbackFromAgy: boolean
  /** Custom path to LFM3-VL-2B-Q4_K_M.gguf (empty uses userData/models/lfm3-vl-2b/...). */
  modelPath: string
  /** Custom path to mmproj-*.gguf (empty uses userData/models/lfm3-vl-2b/...). */
  mmprojPath: string
  /** Custom path to llama-server binary (empty checks managed runtime then PATH). */
  runtimePath: string
  /** Context size passed to llama-server --ctx-size. */
  ctxSize: number
  /** Optional CPU thread count (0 = llama-server automatic). */
  threads: number
  /** Optional GPU offload layers (-ngl, 0 = CPU only, -1 = all). */
  gpuLayers: number
}

export const DEFAULT_LOCAL_VLM_SETTINGS: LocalVlmSettings = {
  enabled: true,
  autoFallbackFromAgy: true,
  modelPath: '',
  mmprojPath: '',
  runtimePath: '',
  ctxSize: LOCAL_VLM_DEFAULT_CTX_SIZE,
  threads: 0,
  gpuLayers: 0,
}

export interface LocalVlmDownloadProgress {
  active: boolean
  stage: 'runtime' | 'model' | 'mmproj'
  file: string
  receivedBytes: number
  totalBytes?: number
  percent?: number
}

export interface LocalVlmStatus {
  installed: boolean
  modelInstalled: boolean
  mmprojInstalled: boolean
  runtimeInstalled: boolean
  running: boolean
  starting: boolean
  port?: number
  modelPath: string
  mmprojPath: string
  runtimePath?: string
  modelBytes?: number
  mmprojBytes?: number
  availableRamMB: number
  ramSufficient: boolean
  download?: LocalVlmDownloadProgress
  lastError?: string
  settings: LocalVlmSettings
}

function clampInt(value: unknown, min: number, max: number, fallback: number): number {
  if (typeof value !== 'number' || !Number.isFinite(value)) return fallback
  return Math.min(max, Math.max(min, Math.round(value)))
}

function cleanPath(value: unknown, fallback: string): string {
  if (typeof value !== 'string') return fallback
  return value.trim().slice(0, 1024)
}

export function mergeLocalVlmSettings(
  base: LocalVlmSettings | undefined,
  patch: unknown,
): LocalVlmSettings {
  const from = base ?? DEFAULT_LOCAL_VLM_SETTINGS
  const p = (patch && typeof patch === 'object' ? patch : {}) as Record<string, unknown>
  return {
    enabled: typeof p.enabled === 'boolean' ? p.enabled : from.enabled,
    autoFallbackFromAgy:
      typeof p.autoFallbackFromAgy === 'boolean' ? p.autoFallbackFromAgy : from.autoFallbackFromAgy,
    modelPath: cleanPath(p.modelPath, from.modelPath),
    mmprojPath: cleanPath(p.mmprojPath, from.mmprojPath),
    runtimePath: cleanPath(p.runtimePath, from.runtimePath),
    ctxSize: clampInt(p.ctxSize, LOCAL_VLM_MIN_CTX_SIZE, LOCAL_VLM_MAX_CTX_SIZE, from.ctxSize),
    threads: clampInt(p.threads, 0, 256, from.threads),
    gpuLayers: clampInt(p.gpuLayers, -1, 999, from.gpuLayers),
  }
}

export function localVlmSettingsFrom(stored: Record<string, unknown>): LocalVlmSettings {
  return mergeLocalVlmSettings(DEFAULT_LOCAL_VLM_SETTINGS, stored[LOCAL_VLM_SETTINGS_KEY])
}

// ---------------------------------------------------------------------------
// Browser-safe usability state (mirrors agy-default.ts)
// ---------------------------------------------------------------------------

let localVlmUsable: boolean | null = null
const localVlmListeners = new Set<(usable: boolean) => void>()

export function localVlmDefaultsUsable(): boolean {
  return localVlmUsable === true
}

export function localVlmUsabilityKnown(): boolean {
  return localVlmUsable !== null
}

export function setLocalVlmUsable(next: boolean | null): void {
  const before = localVlmUsable === true
  localVlmUsable = next
  if (next !== null && before !== next) {
    for (const listener of [...localVlmListeners]) listener(next)
  }
}

export function onLocalVlmUsableChange(listener: (usable: boolean) => void): () => void {
  localVlmListeners.add(listener)
  return () => localVlmListeners.delete(listener)
}
