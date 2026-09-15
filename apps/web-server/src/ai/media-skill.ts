/**
 * AI media handlers — slide style template persistence.
 *
 * These handlers back the renderer-side "slide style library" feature:
 * the user can save the current slide theme as a named template, list all
 * saved templates, and load one back. The renderer produces the actual
 * style payload (colors / fonts / layouts); the web-server stores it as
 * JSON under DATA_DIR and hands the renderer back its saved shape on load.
 *
 * No LLM is involved. Real edits continue to live in the renderer-side
 * `slides-skill.ts`; this module is storage + IPC only.
 */
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'

import { DATA_DIR, registerHandle } from '../common/index'
import { CorruptError, InvalidArgumentError, NotFoundError } from './errors'

export interface SlideStyleTemplate {
  id: string
  name: string
  /** ISO timestamp of last save */
  updatedAt: string
  /** arbitrary key/value bag (colors, fonts, layouts, …) */
  style: Record<string, unknown>
}

const TEMPLATES_DIR = join(DATA_DIR, 'slide-style-templates')
const INDEX_FILE = join(TEMPLATES_DIR, 'index.json')

function ensureDir(): void {
  if (!existsSync(TEMPLATES_DIR)) {
    mkdirSync(TEMPLATES_DIR, { recursive: true })
  }
}

function readIndex(): SlideStyleTemplate[] {
  try {
    if (!existsSync(INDEX_FILE)) return []
    const raw = readFileSync(INDEX_FILE, 'utf8')
    const parsed = JSON.parse(raw) as unknown
    return Array.isArray(parsed) ? (parsed as SlideStyleTemplate[]) : []
  } catch (err) {
    console.warn('[ai:media] failed to read style template index:', err)
    return []
  }
}

function writeIndex(items: SlideStyleTemplate[]): void {
  ensureDir()
  writeFileSync(INDEX_FILE, JSON.stringify(items, null, 2), 'utf8')
}

function safeId(input: string): string {
  return input.replace(/[^a-zA-Z0-9._-]/g, '_').slice(0, 64) || 'template'
}

export function registerAiMediaSkillHandlers(): void {
  registerHandle('ai:save-style-template', async (_event: unknown, args: unknown) => {
    const { id, name, style } = args as {
      id?: string
      name: string
      style: Record<string, unknown>
    }
    if (!name || typeof name !== 'string') {
      throw new InvalidArgumentError('ai:save-style-template', 'name is required')
    }
    ensureDir()
    const templateId = safeId(id || `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`)
    const record: SlideStyleTemplate = {
      id: templateId,
      name,
      updatedAt: new Date().toISOString(),
      style: style ?? {},
    }
    writeFileSync(
      join(TEMPLATES_DIR, `${templateId}.json`),
      JSON.stringify(record, null, 2),
      'utf8',
    )
    const all = readIndex().filter(t => t.id !== templateId)
    all.push(record)
    writeIndex(all)
    return { ok: true, template: record }
  })

  registerHandle('ai:list-style-templates', async () => {
    const items = readIndex().sort((a, b) => b.updatedAt.localeCompare(a.updatedAt))
    return { templates: items }
  })

  registerHandle('ai:load-style-template', async (_event: unknown, args: unknown) => {
    const { id } = args as { id: string }
    if (!id || typeof id !== 'string') {
      throw new InvalidArgumentError('ai:load-style-template', 'id is required')
    }
    const file = join(TEMPLATES_DIR, `${safeId(id)}.json`)
    if (!existsSync(file)) {
      throw new NotFoundError('ai:load-style-template', `template '${id}' not found`)
    }
    try {
      const raw = readFileSync(file, 'utf8')
      const parsed = JSON.parse(raw) as SlideStyleTemplate
      return { template: parsed }
    } catch (err) {
      throw new CorruptError(
        'ai:load-style-template',
        err instanceof Error ? err.message : String(err),
      )
    }
  })
}
