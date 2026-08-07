import { readFileSync, writeFileSync } from 'node:fs'

function patch(path, from, to) {
  const src = readFileSync(path, 'utf8')
  if (src.includes(to)) return
  if (!src.includes(from)) throw new Error(`Patch anchor not found in ${path}`)
  writeFileSync(path, src.replace(from, to))
}

// Main-process AI route is authoritative. Do not trust renderer settings that may
// have been loaded before the OpenRouter key was saved from the Home window.
patch(
  'apps/docs/src/main/docs-main.ts',
  `    const { requestId, settings, system, messages } = request\n    const tools = request.tools ?? []\n    const maxTokens = request.maxTokens ?? 8192\n    const provider = settings.provider\n    let config = settings.providers?.[provider]`,
  `    const { requestId, system, messages } = request\n    const tools = request.tools ?? []\n    const maxTokens = request.maxTokens ?? 8192\n    const stored = readJson<Partial<AiSettings> & LegacyAiSettings>(SETTINGS_PATH(), {})\n    const settings = resolveAiSettings(stored, defaultAiSettings())\n    const provider = settings.provider\n    let config = settings.providers?.[provider]`,
)

patch(
  'apps/docs/src/main/docs-main.ts',
  `  ipcMain.handle('ai:chat', async (_event, request: AiChatRequest) => {\n    const { settings, system, user } = request\n    const provider = settings.provider\n    let config = settings.providers?.[provider]`,
  `  ipcMain.handle('ai:chat', async (_event, request: AiChatRequest) => {\n    const { system, user } = request\n    const stored = readJson<Partial<AiSettings> & LegacyAiSettings>(SETTINGS_PATH(), {})\n    const settings = resolveAiSettings(stored, defaultAiSettings())\n    const provider = settings.provider\n    let config = settings.providers?.[provider]`,
)

// The old chat UI queried Genspark status after EVERY AI error. Since a direct
// OpenRouter user is naturally not logged into Genspark, any OpenRouter error was
// incorrectly turned into a Genspark sign-in prompt.
patch(
  'apps/docs/src/renderer/ai/AiPanel.tsx',
  `          // Signed-out failures get an inline sign-in button; detected via\n          // gsk status rather than matching the localized error text\n          void window.desktop\n            .aiGskStatus()\n            .then((status) => {\n              if (status.loggedIn) return\n              setChat((prev) => {\n                const next = [...prev]\n                const last = next.at(-1)\n                if (last?.role === 'assistant' && last.error) {\n                  next[next.length - 1] = { ...last, loginRequired: true }\n                }\n                return next\n              })\n            })\n            .catch(() => {})\n          setBusy(false)`,
  `          // Direct OpenRouter errors remain provider errors; never fall back to Genspark auth.\n          setBusy(false)`,
)

patch(
  'apps/docs/src/renderer/ai/AiPanel.tsx',
  `              {entry.loginRequired && (\n                <button className="ai-login-btn" onClick={() => void window.desktop.aiGskLogin()}>\n                  {t('aiGskLoginBtn')}\n                </button>\n              )}\n`,
  ``,
)

console.log('OpenRouter chat route fixed: saved settings are authoritative and Genspark login fallback is disabled.')
