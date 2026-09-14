/**
 * Module management channels — list, update enabled/disabled state and reorder
 * the quick-create modules shown on the home page. Persisted to a small JSON
 * file under DATA_DIR so the order survives restarts.
 */
import { existsSync, readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { DATA_DIR, registerHandle } from '../common/index.js'

export type ModuleKind = 'docx' | 'xlsx' | 'pptx' | 'md' | 'pdf' | 'html'

export interface ModuleEntry {
  id: ModuleKind
  /** display label, l10n key (e.g. 'newDoc' / 'newSheet') */
  labelKey: string
  /** file extension (without dot) */
  ext: string
  /** module subdir under the web root */
  path: string
  enabled: boolean
}

const MODULES_FILE = join(DATA_DIR, 'modules.json')

const DEFAULT_MODULES: ModuleEntry[] = [
  { id: 'docx', labelKey: 'newDoc', ext: 'docx', path: '/docs/', enabled: true },
  { id: 'xlsx', labelKey: 'newSheet', ext: 'xlsx', path: '/sheets/', enabled: true },
  { id: 'pptx', labelKey: 'newSlide', ext: 'pptx', path: '/slides/', enabled: true },
  { id: 'md', labelKey: 'newMarkdown', ext: 'md', path: '/markdown/', enabled: true },
  { id: 'pdf', labelKey: 'newPdf', ext: 'pdf', path: '/pdf/', enabled: true },
  { id: 'html', labelKey: 'newHtml', ext: 'html', path: '/html/', enabled: true },
]

let cached: ModuleEntry[] | null = null

function loadModules(): ModuleEntry[] {
  if (cached) return cached
  try {
    if (existsSync(MODULES_FILE)) {
      const parsed = JSON.parse(readFileSync(MODULES_FILE, 'utf-8')) as Partial<ModuleEntry>[]
      if (Array.isArray(parsed) && parsed.length > 0) {
        cached = DEFAULT_MODULES.map((def) => {
          const override = parsed.find((p) => p && p.id === def.id)
          return {
            ...def,
            enabled: typeof override?.enabled === 'boolean' ? override.enabled : def.enabled,
          }
        })
        return cached
      }
    }
  } catch {}
  cached = DEFAULT_MODULES.map((m) => ({ ...m }))
  return cached
}

function saveModules(modules: ModuleEntry[]): void {
  cached = modules
  try {
    writeFileSync(MODULES_FILE, JSON.stringify(modules, null, 2))
  } catch {}
}

export function registerModuleHandlers(): void {
  registerHandle('home:list-modules', () => {
    return { modules: loadModules() }
  })

  registerHandle('home:set-module-enabled', (_event: unknown, args: unknown) => {
    const { id, enabled } = (args || {}) as { id: ModuleKind; enabled: boolean }
    const modules = loadModules().map((m) => (m.id === id ? { ...m, enabled: !!enabled } : m))
    saveModules(modules)
    return { ok: true, modules }
  })

  registerHandle('home:reorder-modules', (_event: unknown, args: unknown) => {
    const { order } = (args || {}) as { order: ModuleKind[] }
    const modules = loadModules()
    const map = new Map(modules.map((m) => [m.id, m]))
    const reordered: ModuleEntry[] = []
    for (const id of order || []) {
      const entry = map.get(id)
      if (entry) {
        reordered.push(entry)
        map.delete(id)
      }
    }
    for (const remaining of map.values()) reordered.push(remaining)
    saveModules(reordered)
    return { ok: true, modules: reordered }
  })

  registerHandle('home:reset-modules', () => {
    cached = null
    saveModules(DEFAULT_MODULES.map((m) => ({ ...m })))
    return { ok: true, modules: loadModules() }
  })
}
