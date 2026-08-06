import { readFileSync, writeFileSync } from 'node:fs'

function patch(path, from, to) {
  const src = readFileSync(path, 'utf8')
  if (src.includes(to)) return
  if (!src.includes(from)) throw new Error(`Patch anchor not found in ${path}`)
  writeFileSync(path, src.replace(from, to))
}

// Stop the shared AI IPC from forcing every saved configuration back to Genspark.
patch(
  'apps/docs/src/main/docs-main.ts',
  `    // AI features all go through Genspark (gsk login); legacy settings with another provider are reset\n    settings.provider = 'genspark'\n    return settings`,
  `    return settings`,
)

// Home API: expose the shared ai-settings.json used by Docs/Sheets/Slides.
patch(
  'apps/shell/src/shared/home-api.ts',
  `  /** app version (from package.json / electron app.getVersion) */\n  getAppVersion(): Promise<string>`,
  `  /** OpenRouter/Nemotron settings shared by all editor AI panels */\n  getOpenRouterSettings(): Promise<OpenRouterSettings>\n  setOpenRouterSettings(settings: OpenRouterSettings): Promise<void>\n  /** app version (from package.json / electron app.getVersion) */\n  getAppVersion(): Promise<string>`,
)
patch(
  'apps/shell/src/shared/home-api.ts',
  `export interface AccountStatus {`,
  `export interface OpenRouterSettings {\n  apiKey: string\n  model: string\n}\n\nexport interface AccountStatus {`,
)
patch(
  'apps/shell/src/shared/home-api.ts',
  `  accountLogout: 'home:account-logout',\n  getAppVersion: 'home:get-app-version',`,
  `  accountLogout: 'home:account-logout',\n  getOpenRouterSettings: 'home:get-openrouter-settings',\n  setOpenRouterSettings: 'home:set-openrouter-settings',\n  getAppVersion: 'home:get-app-version',`,
)

// Preload bridge.
patch(
  'apps/shell/src/preload/index.ts',
  `  async getAppVersion() {`,
  `  async getOpenRouterSettings() {\n    const result: unknown = await ipcRenderer.invoke(HOME_CHANNELS.getOpenRouterSettings)\n    const value = result as { apiKey?: unknown; model?: unknown } | null\n    return {\n      apiKey: typeof value?.apiKey === 'string' ? value.apiKey : '',\n      model: typeof value?.model === 'string' ? value.model : 'nvidia/nemotron-3-ultra-550b-a55b:free',\n    }\n  },\n  async setOpenRouterSettings(settings) {\n    await ipcRenderer.invoke(HOME_CHANNELS.setOpenRouterSettings, settings)\n  },\n  async getAppVersion() {`,
)

// Main-process persistence. This writes the exact ai-settings.json consumed by registerAiIpc().
patch(
  'apps/shell/src/main/index.ts',
  `  ipcMain.handle(HOME_CHANNELS.getAppVersion, (): string => app.getVersion())`,
  `  ipcMain.handle(HOME_CHANNELS.getOpenRouterSettings, () => {\n    const path = join(app.getPath('userData'), 'ai-settings.json')\n    try {\n      const raw = JSON.parse(readFileSync(path, 'utf8')) as { providers?: Record<string, { apiKey?: string; model?: string }> }\n      const cfg = raw.providers?.openrouter\n      return { apiKey: cfg?.apiKey ?? '', model: cfg?.model ?? 'nvidia/nemotron-3-ultra-550b-a55b:free' }\n    } catch {\n      return { apiKey: '', model: 'nvidia/nemotron-3-ultra-550b-a55b:free' }\n    }\n  })\n\n  ipcMain.handle(HOME_CHANNELS.setOpenRouterSettings, (_event, value: unknown) => {\n    const input = value as { apiKey?: unknown; model?: unknown } | null\n    const apiKey = typeof input?.apiKey === 'string' ? input.apiKey.trim() : ''\n    const model = typeof input?.model === 'string' && input.model.trim()\n      ? input.model.trim()\n      : 'nvidia/nemotron-3-ultra-550b-a55b:free'\n    const path = join(app.getPath('userData'), 'ai-settings.json')\n    let saved: Record<string, unknown> = {}\n    try { saved = JSON.parse(readFileSync(path, 'utf8')) as Record<string, unknown> } catch { /* first run */ }\n    const providers = (saved.providers && typeof saved.providers === 'object' ? saved.providers : {}) as Record<string, unknown>\n    providers.openrouter = { apiKey, model }\n    saved.provider = 'openrouter'\n    saved.providers = providers\n    writeFileSync(path, JSON.stringify(saved, null, 2))\n  })\n\n  ipcMain.handle(HOME_CHANNELS.getAppVersion, (): string => app.getVersion())`,
)

// Account menu UI: OpenRouter settings are available from the Home screen.
patch(
  'apps/shell/src/renderer/src/Home.tsx',
  `  const [appVersion, setAppVersion] = useState('')`,
  `  const [appVersion, setAppVersion] = useState('')\n  const [showAiSettings, setShowAiSettings] = useState(false)\n  const [openRouterKey, setOpenRouterKey] = useState('')\n  const [openRouterModel, setOpenRouterModel] = useState('nvidia/nemotron-3-ultra-550b-a55b:free')\n  const [aiSaved, setAiSaved] = useState(false)`,
)
patch(
  'apps/shell/src/renderer/src/Home.tsx',
  `    void window.aiOffice.getAppVersion?.().then((v) => {\n      if (alive && v) setAppVersion(v)\n    })`,
  `    void window.aiOffice.getAppVersion?.().then((v) => {\n      if (alive && v) setAppVersion(v)\n    })\n    void window.aiOffice.getOpenRouterSettings?.().then((cfg) => {\n      if (!alive) return\n      setOpenRouterKey(cfg.apiKey)\n      setOpenRouterModel(cfg.model)\n    })`,
)
patch(
  'apps/shell/src/renderer/src/Home.tsx',
  `          <div className="account-menu-divider" />\n          <div\n            className="lang-row-wrap"`,
  `          <div className="account-menu-divider" />\n          <button\n            className="account-menu-item"\n            role="menuitem"\n            onClick={() => { setShowAiSettings(true); setMenuOpen(false) }}\n          >\n            <svg width="16" height="16" viewBox="0 0 16 16" fill="none" aria-hidden="true"><path d="M8 2.2a5.8 5.8 0 1 0 0 11.6A5.8 5.8 0 0 0 8 2.2Zm0 3.1v5.4M5.3 8h5.4" stroke="currentColor" strokeWidth="1.2" strokeLinecap="round"/></svg>\n            <span>AI Settings · OpenRouter</span>\n          </button>\n          <div\n            className="lang-row-wrap"`,
)
patch(
  'apps/shell/src/renderer/src/Home.tsx',
  `      {!menuOpen && waiting && authUrl && (`,
  `      {showAiSettings && (\n        <div className="or-settings-overlay" onClick={() => setShowAiSettings(false)}>\n          <div className="or-settings-card" role="dialog" aria-modal="true" aria-label="OpenRouter AI Settings" onClick={(e) => e.stopPropagation()}>\n            <div className="or-settings-head"><div><h3>AI Settings</h3><p>OpenRouter · NVIDIA Nemotron</p></div><button onClick={() => setShowAiSettings(false)} aria-label="Close">×</button></div>\n            <label>Provider<input value="OpenRouter" disabled /></label>\n            <label>OpenRouter API Key<input type="password" value={openRouterKey} placeholder="sk-or-v1-..." onChange={(e) => { setOpenRouterKey(e.target.value); setAiSaved(false) }} /></label>\n            <label>Model<select value={openRouterModel} onChange={(e) => { setOpenRouterModel(e.target.value); setAiSaved(false) }}><option value="nvidia/nemotron-3-ultra-550b-a55b:free">NVIDIA Nemotron 3 Ultra (Free)</option><option value="openrouter/free">OpenRouter Free Router</option></select></label>\n            <p className="or-settings-note">Your key is stored locally in GenOffice user data and is not built into the EXE.</p>\n            <div className="or-settings-actions"><span>{aiSaved ? 'Saved ✓' : ''}</span><button onClick={() => void window.aiOffice.setOpenRouterSettings({ apiKey: openRouterKey, model: openRouterModel }).then(() => setAiSaved(true))}>Save AI Settings</button></div>\n          </div>\n        </div>\n      )}\n      {!menuOpen && waiting && authUrl && (`,
)

// Modal styling.
const cssPath = 'apps/shell/src/renderer/src/home.css'
let css = readFileSync(cssPath, 'utf8')
if (!css.includes('.or-settings-overlay')) {
  css += `\n\n/* OpenRouter / Nemotron settings */\n.or-settings-overlay{position:fixed;inset:0;background:rgba(15,23,42,.28);display:flex;align-items:center;justify-content:center;z-index:10000;backdrop-filter:blur(2px)}\n.or-settings-card{width:min(460px,calc(100vw - 40px));background:#fff;border:1px solid #e5e7eb;border-radius:16px;box-shadow:0 24px 70px rgba(15,23,42,.22);padding:22px;color:#111827}\n.or-settings-head{display:flex;align-items:flex-start;justify-content:space-between;margin-bottom:18px}.or-settings-head h3{font-size:20px;margin:0 0 4px}.or-settings-head p{margin:0;color:#667085;font-size:13px}.or-settings-head button{border:0;background:transparent;font-size:25px;line-height:1;color:#667085;cursor:pointer}\n.or-settings-card label{display:block;font-size:12px;font-weight:600;color:#475467;margin:13px 0}.or-settings-card input,.or-settings-card select{display:block;width:100%;box-sizing:border-box;margin-top:6px;border:1px solid #d0d5dd;border-radius:9px;padding:10px 11px;background:#fff;color:#101828;font:inherit;outline:none}.or-settings-card input:focus,.or-settings-card select:focus{border-color:#667085;box-shadow:0 0 0 3px rgba(102,112,133,.12)}.or-settings-card input:disabled{background:#f9fafb;color:#667085}.or-settings-note{font-size:12px;line-height:1.5;color:#667085;margin:14px 0}.or-settings-actions{display:flex;align-items:center;justify-content:space-between;min-height:38px}.or-settings-actions span{font-size:12px;color:#067647}.or-settings-actions button{border:0;border-radius:9px;background:#101828;color:#fff;padding:10px 15px;font-weight:600;cursor:pointer}\n`
  writeFileSync(cssPath, css)
}

console.log('OpenRouter/Nemotron fork patches applied.')
