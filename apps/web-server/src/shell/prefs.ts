/**
 * Web-server fallback for the shell-side `app:*` preference channels the docs /
 * sheets / slides renderers expect. The Electron shell owns the real
 * persisted values (`apps/shell/src/main/index.ts`); the standalone web build
 * returns the package defaults so the renderer doesn't 404 every time it boots.
 *
 * Channels registered:
 *  - app:get-auto-save-default
 *  - app:set-auto-save-default
 *  - app:auto-save-default-changed (push event)
 *  - app:get-ai-panel-prefs
 *  - app:set-ai-panel-prefs
 *  - app:ai-panel-prefs-changed (push event)
 */
import { registerHandle } from '../common/index.js'

type AutoSaveDefault = { on: boolean; updatedAt: number }
type AiFontSize = 'default' | 'large' | 'xlarge' | 'custom'
type AiPanelPrefs = {
  fontSize: AiFontSize
  customFontSize: number
  spellcheck: boolean
}

const DEFAULT_AUTO_SAVE: AutoSaveDefault = { on: false, updatedAt: 0 }
const DEFAULT_AI_PANEL_PREFS: AiPanelPrefs = {
  fontSize: 'default',
  customFontSize: 14,
  spellcheck: true,
}

function isAutoSaveDefault(value: unknown): value is AutoSaveDefault {
  if (!value || typeof value !== 'object') return false
  const v = value as Record<string, unknown>
  return typeof v.on === 'boolean' && typeof v.updatedAt === 'number'
}

function isAiPanelPrefs(value: unknown): value is AiPanelPrefs {
  if (!value || typeof value !== 'object') return false
  const v = value as Record<string, unknown>
  if (v.fontSize !== 'default' && v.fontSize !== 'large' && v.fontSize !== 'xlarge' && v.fontSize !== 'custom') return false
  if (typeof v.customFontSize !== 'number') return false
  if (typeof v.spellcheck !== 'boolean') return false
  return true
}

export function registerPrefsHandlers(): void {
  let autoSave: AutoSaveDefault = { ...DEFAULT_AUTO_SAVE }
  let aiPanel: AiPanelPrefs = { ...DEFAULT_AI_PANEL_PREFS }

  registerHandle('app:get-auto-save-default', () => ({ ...autoSave }))
  registerHandle('app:set-auto-save-default', (_event: unknown, value: unknown) => {
    if (!isAutoSaveDefault(value)) return { ok: false, error: 'invalid AutoSaveDefault payload' }
    autoSave = { ...value }
    return { ok: true }
  })

  registerHandle('app:get-ai-panel-prefs', () => ({ ...aiPanel }))
  registerHandle('app:set-ai-panel-prefs', (_event: unknown, patch: unknown) => {
    if (!patch || typeof patch !== 'object') {
      return { ok: false, error: 'invalid AiPanelPrefs payload' }
    }
    const next: AiPanelPrefs = { ...aiPanel, ...(patch as Partial<AiPanelPrefs>) }
    if (!isAiPanelPrefs(next)) return { ok: false, error: 'invalid AiPanelPrefs payload' }
    aiPanel = next
    return { ok: true, prefs: { ...aiPanel } }
  })
}
