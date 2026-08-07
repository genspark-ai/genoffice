import { readFileSync, writeFileSync } from 'node:fs'

function patch(path, from, to) {
  const src = readFileSync(path, 'utf8')
  if (src.includes(to)) return
  if (!src.includes(from)) throw new Error(`Patch anchor not found in ${path}`)
  writeFileSync(path, src.replace(from, to))
}

patch('packages/ai-provider/src/types.ts',
  `export type AiProviderId = 'openrouter' | 'genspark' | 'anthropic' | 'gemini' | 'deepseek' | 'openai' | 'custom'`,
  `export type AiProviderId = 'openrouter' | 'nvidia' | 'genspark' | 'anthropic' | 'gemini' | 'deepseek' | 'openai' | 'custom'`)

patch('packages/ai-provider/src/providers.ts',
  `export const OPENROUTER_BASE_URL = 'https://openrouter.ai/api/v1'`,
  `export const OPENROUTER_BASE_URL = 'https://openrouter.ai/api/v1'\nexport const NVIDIA_NIM_BASE_URL = 'https://integrate.api.nvidia.com/v1'`)

patch('packages/ai-provider/src/providers.ts',
  `  {\n    id: 'openrouter', label: 'OpenRouter (NVIDIA Nemotron)',\n    models: ['nvidia/nemotron-3-ultra-550b-a55b:free', 'nvidia/nemotron-3-super-120b-a12b:free', 'nvidia/nemotron-3-nano-30b-a3b:free', 'openrouter/free'],\n    defaultModel: 'nvidia/nemotron-3-ultra-550b-a55b:free', keyPlaceholder: 'sk-or-v1-...',\n  },`,
  `  {\n    id: 'openrouter', label: 'OpenRouter',\n    models: ['nvidia/nemotron-3-ultra-550b-a55b:free', 'nvidia/nemotron-3-super-120b-a12b:free', 'nvidia/nemotron-3-nano-30b-a3b:free', 'openrouter/free'],\n    defaultModel: 'openrouter/free', keyPlaceholder: 'sk-or-v1-...',\n  },\n  {\n    id: 'nvidia', label: 'NVIDIA NIM (Direct)',\n    models: ['nvidia/nemotron-3-ultra-550b-a55b', 'nvidia/nemotron-3-super-120b-a12b', 'nvidia/nemotron-3-nano-30b-a3b'],\n    defaultModel: 'nvidia/nemotron-3-ultra-550b-a55b', keyPlaceholder: 'nvapi-...',\n  },`)

patch('packages/ai-provider/src/stream.ts',
  `import { OPENROUTER_BASE_URL } from './providers'`,
  `import { NVIDIA_NIM_BASE_URL, OPENROUTER_BASE_URL } from './providers'`)
patch('packages/ai-provider/src/stream.ts',
  `  openrouter: OPENROUTER_BASE_URL,\n  deepseek:`,
  `  openrouter: OPENROUTER_BASE_URL,\n  nvidia: NVIDIA_NIM_BASE_URL,\n  deepseek:`)

// Upgrade shell settings contract from OpenRouter-only to selected provider + both credentials.
patch('apps/shell/src/shared/home-api.ts',
  `export interface OpenRouterSettings {\n  apiKey: string\n  model: string\n}`,
  `export interface OpenRouterSettings {\n  provider: 'openrouter' | 'nvidia'\n  apiKey: string\n  model: string\n  openRouterApiKey?: string\n  nvidiaApiKey?: string\n}`)

// Main process: read/write selected provider while preserving both provider keys.
patch('apps/shell/src/main/index.ts',
  `      const raw = JSON.parse(readFileSync(path, 'utf8')) as { providers?: Record<string, { apiKey?: string; model?: string }> }\n      const cfg = raw.providers?.openrouter\n      return { apiKey: cfg?.apiKey ?? '', model: cfg?.model ?? 'nvidia/nemotron-3-ultra-550b-a55b:free' }`,
  `      const raw = JSON.parse(readFileSync(path, 'utf8')) as { provider?: string; providers?: Record<string, { apiKey?: string; model?: string }> }\n      const provider = raw.provider === 'nvidia' ? 'nvidia' : 'openrouter'\n      const cfg = raw.providers?.[provider]\n      return { provider, apiKey: cfg?.apiKey ?? '', model: cfg?.model ?? (provider === 'nvidia' ? 'nvidia/nemotron-3-ultra-550b-a55b' : 'openrouter/free'), openRouterApiKey: raw.providers?.openrouter?.apiKey ?? '', nvidiaApiKey: raw.providers?.nvidia?.apiKey ?? '' }`)
patch('apps/shell/src/main/index.ts',
  `    const apiKey = typeof input?.apiKey === 'string' ? input.apiKey.trim() : ''\n    const model = typeof input?.model === 'string' && input.model.trim() ? input.model.trim() : 'nvidia/nemotron-3-ultra-550b-a55b:free'`,
  `    const provider = (input as { provider?: unknown } | null)?.provider === 'nvidia' ? 'nvidia' : 'openrouter'\n    const apiKey = typeof input?.apiKey === 'string' ? input.apiKey.trim() : ''\n    const model = typeof input?.model === 'string' && input.model.trim() ? input.model.trim() : (provider === 'nvidia' ? 'nvidia/nemotron-3-ultra-550b-a55b' : 'openrouter/free')`)
patch('apps/shell/src/main/index.ts',
  `    providers.openrouter = { apiKey, model }\n    saved.provider = 'openrouter'`,
  `    providers[provider] = { apiKey, model }\n    saved.provider = provider`)

// Preload returns selected provider too.
patch('apps/shell/src/preload/index.ts',
  `    const value = result as { apiKey?: unknown; model?: unknown } | null\n    return { apiKey: typeof value?.apiKey === 'string' ? value.apiKey : '', model: typeof value?.model === 'string' ? value.model : 'nvidia/nemotron-3-ultra-550b-a55b:free' }`,
  `    const value = result as { provider?: unknown; apiKey?: unknown; model?: unknown; openRouterApiKey?: unknown; nvidiaApiKey?: unknown } | null\n    const provider = value?.provider === 'nvidia' ? 'nvidia' : 'openrouter'\n    return { provider, apiKey: typeof value?.apiKey === 'string' ? value.apiKey : '', model: typeof value?.model === 'string' ? value.model : (provider === 'nvidia' ? 'nvidia/nemotron-3-ultra-550b-a55b' : 'openrouter/free'), openRouterApiKey: typeof value?.openRouterApiKey === 'string' ? value.openRouterApiKey : '', nvidiaApiKey: typeof value?.nvidiaApiKey === 'string' ? value.nvidiaApiKey : '' }`)

// Home UI: add provider selector and swap model/key values safely.
patch('apps/shell/src/renderer/src/Home.tsx',
  `  const [openRouterKey, setOpenRouterKey] = useState('')`,
  `  const [aiProvider, setAiProvider] = useState<'openrouter' | 'nvidia'>('openrouter')\n  const [openRouterKey, setOpenRouterKey] = useState('')\n  const [nvidiaKey, setNvidiaKey] = useState('')`)
patch('apps/shell/src/renderer/src/Home.tsx',
  `void window.aiOffice.getOpenRouterSettings?.().then((cfg) => { if (alive) { setOpenRouterKey(cfg.apiKey); setOpenRouterModel(cfg.model) } })`,
  `void window.aiOffice.getOpenRouterSettings?.().then((cfg) => { if (alive) { setAiProvider(cfg.provider); setOpenRouterKey(cfg.openRouterApiKey ?? (cfg.provider === 'openrouter' ? cfg.apiKey : '')); setNvidiaKey(cfg.nvidiaApiKey ?? (cfg.provider === 'nvidia' ? cfg.apiKey : '')); setOpenRouterModel(cfg.model) } })`)
patch('apps/shell/src/renderer/src/Home.tsx',
  `<label>Provider<input value="OpenRouter" disabled /></label>\n            <label>OpenRouter API Key<input type="password" value={openRouterKey} placeholder="sk-or-v1-..." onChange={(e) => { setOpenRouterKey(e.target.value); setAiSaved(false) }} /></label>\n            <label>Model<select value={openRouterModel} onChange={(e) => { setOpenRouterModel(e.target.value); setAiSaved(false) }}><option value="nvidia/nemotron-3-ultra-550b-a55b:free">NVIDIA Nemotron 3 Ultra (Free)</option><option value="openrouter/free">OpenRouter Free Router</option></select></label>`,
  `<label>Provider<select value={aiProvider} onChange={(e) => { const p = e.target.value as 'openrouter' | 'nvidia'; setAiProvider(p); setOpenRouterModel(p === 'nvidia' ? 'nvidia/nemotron-3-ultra-550b-a55b' : 'openrouter/free'); setAiSaved(false) }}><option value="openrouter">OpenRouter</option><option value="nvidia">NVIDIA NIM (Direct)</option></select></label>\n            <label>{aiProvider === 'nvidia' ? 'NVIDIA API Key' : 'OpenRouter API Key'}<input type="password" value={aiProvider === 'nvidia' ? nvidiaKey : openRouterKey} placeholder={aiProvider === 'nvidia' ? 'nvapi-...' : 'sk-or-v1-...'} onChange={(e) => { aiProvider === 'nvidia' ? setNvidiaKey(e.target.value) : setOpenRouterKey(e.target.value); setAiSaved(false) }} /></label>\n            <label>Model<select value={openRouterModel} onChange={(e) => { setOpenRouterModel(e.target.value); setAiSaved(false) }}>{aiProvider === 'nvidia' ? <><option value="nvidia/nemotron-3-ultra-550b-a55b">Nemotron 3 Ultra</option><option value="nvidia/nemotron-3-super-120b-a12b">Nemotron 3 Super</option><option value="nvidia/nemotron-3-nano-30b-a3b">Nemotron 3 Nano</option></> : <><option value="openrouter/free">OpenRouter Free Router</option><option value="nvidia/nemotron-3-ultra-550b-a55b:free">Nemotron 3 Ultra (Free)</option><option value="nvidia/nemotron-3-super-120b-a12b:free">Nemotron 3 Super (Free)</option></>}</select></label>`)
patch('apps/shell/src/renderer/src/Home.tsx',
  `window.aiOffice.setOpenRouterSettings({ apiKey: openRouterKey, model: openRouterModel })`,
  `window.aiOffice.setOpenRouterSettings({ provider: aiProvider, apiKey: aiProvider === 'nvidia' ? nvidiaKey : openRouterKey, model: openRouterModel })`)

console.log('Added balanced OpenRouter + direct NVIDIA NIM provider support.')
