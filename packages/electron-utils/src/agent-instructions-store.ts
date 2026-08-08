/**
 * On-disk store for user-authored agent instructions.
 *
 * Layout under userData:
 *   agent-rules.json     { global?, docx?, pptx?, sheets?, pdf? }
 *   agent-skills/<id>.md  one markdown file per skill, front matter and all
 *   agent-memory.json    [{ id, text, createdAt }] the agent's own notes
 *
 * Memory is JSON rather than one file per entry: entries are one-sentence
 * facts the agent writes and rewrites, not documents anyone would edit by
 * hand, so the per-file affordance skills need buys nothing here.
 *
 * Skills are plain files on purpose: the user can drop a `skill.md` in with a
 * file manager, edit one in their own editor, or keep them in a git repo, and
 * the app picks the change up on the next read. That is also why the id is the
 * filename stem rather than a generated key stored in an index.
 *
 * The base directory is injected rather than read from `app.getPath` so the
 * store is unit-testable without Electron.
 */
import { existsSync, mkdirSync, readdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import {
  MAX_MEMORIES,
  coerceScope,
  normalizeMemoryText,
  parseSkillMarkdown,
  serializeSkillMarkdown,
  type AgentRules,
  type InstructionScope,
  type UserMemory,
  type UserSkill,
} from '@genoffice/agent-core'

const RULES_FILE = 'agent-rules.json'
const MEMORY_FILE = 'agent-memory.json'
const SKILLS_DIR = 'agent-skills'
/** guards against a pasted novel becoming the system prompt on every turn */
const MAX_RULE_CHARS = 20_000
const MAX_SKILL_CHARS = 200_000

export class AgentInstructionsStore {
  constructor(private readonly baseDir: string) {}

  private get rulesPath(): string {
    return join(this.baseDir, RULES_FILE)
  }

  private get skillsDir(): string {
    return join(this.baseDir, SKILLS_DIR)
  }

  private get memoryPath(): string {
    return join(this.baseDir, MEMORY_FILE)
  }

  // ── rules ─────────────────────────────────────────────────────────

  readRules(): AgentRules {
    try {
      if (!existsSync(this.rulesPath)) return {}
      const raw: unknown = JSON.parse(readFileSync(this.rulesPath, 'utf-8'))
      if (typeof raw !== 'object' || raw === null) return {}
      const out: AgentRules = {}
      for (const [key, value] of Object.entries(raw as Record<string, unknown>)) {
        if (typeof value !== 'string') continue
        const scope = coerceScope(key)
        // coerceScope falls back to 'global'; only keep a key that really is one
        if (key === scope || coerceScope(key) === scope) out[scope] = value.slice(0, MAX_RULE_CHARS)
      }
      return out
    } catch {
      // a corrupted file must not take AI features down with it
      return {}
    }
  }

  writeRules(rules: AgentRules): AgentRules {
    const clean: AgentRules = {}
    for (const [key, value] of Object.entries(rules ?? {})) {
      if (typeof value !== 'string') continue
      const trimmed = value.trim()
      if (trimmed) clean[coerceScope(key)] = trimmed.slice(0, MAX_RULE_CHARS)
    }
    mkdirSync(this.baseDir, { recursive: true })
    writeFileSync(this.rulesPath, JSON.stringify(clean, null, 2))
    return clean
  }

  // ── skills ────────────────────────────────────────────────────────

  listSkills(): UserSkill[] {
    try {
      if (!existsSync(this.skillsDir)) return []
      return readdirSync(this.skillsDir)
        .filter((f) => f.toLowerCase().endsWith('.md'))
        .map((file) => {
          const id = file.replace(/\.md$/i, '')
          try {
            return parseSkillMarkdown(readFileSync(join(this.skillsDir, file), 'utf-8'), id)
          } catch {
            return null
          }
        })
        .filter((s): s is UserSkill => s !== null)
        .sort((a, b) => a.name.localeCompare(b.name))
    } catch {
      return []
    }
  }

  /**
   * Create or overwrite a skill. `id` is slugified so it is always a safe
   * filename; a fresh id gains a numeric suffix rather than clobbering an
   * existing skill.
   */
  saveSkill(input: {
    id?: string
    name: string
    description?: string
    scopes?: unknown[]
    body: string
    enabled?: boolean
  }): UserSkill {
    const scopes = normalizeScopes(input.scopes)
    const skill: UserSkill = {
      id: input.id ? safeId(input.id) : this.freshId(safeId(input.name) || 'skill'),
      name: (input.name || '').trim() || 'Untitled skill',
      description: (input.description ?? '').trim(),
      scopes,
      body: (input.body ?? '').slice(0, MAX_SKILL_CHARS),
      enabled: input.enabled !== false,
    }
    mkdirSync(this.skillsDir, { recursive: true })
    writeFileSync(join(this.skillsDir, `${skill.id}.md`), serializeSkillMarkdown(skill))
    return skill
  }

  /**
   * Import a raw `skill.md`. Metadata comes from its front matter, so a file
   * written elsewhere keeps its own name, description and scopes.
   */
  importSkillMarkdown(source: string, filename: string): UserSkill {
    const stem = filename.replace(/\.md$/i, '')
    const parsed = parseSkillMarkdown(source.slice(0, MAX_SKILL_CHARS), safeId(stem) || 'skill')
    return this.saveSkill({
      id: this.freshId(parsed.id),
      name: parsed.name,
      description: parsed.description,
      scopes: parsed.scopes,
      body: parsed.body,
      enabled: parsed.enabled,
    })
  }

  deleteSkill(id: string): void {
    const path = join(this.skillsDir, `${safeId(id)}.md`)
    if (existsSync(path)) rmSync(path)
  }

  // ── memory ────────────────────────────────────────────────────────

  /** Recorded preferences, newest first. A corrupt file reads as none. */
  readMemories(): UserMemory[] {
    try {
      if (!existsSync(this.memoryPath)) return []
      const raw: unknown = JSON.parse(readFileSync(this.memoryPath, 'utf-8'))
      if (!Array.isArray(raw)) return []
      return raw
        .map((entry) => {
          const e = entry as { id?: unknown; text?: unknown; createdAt?: unknown }
          const text = normalizeMemoryText(e.text)
          const id = String(e.id ?? '')
          if (!text || !id) return null
          const createdAt = Number(e.createdAt)
          return { id, text, createdAt: Number.isFinite(createdAt) ? createdAt : 0 }
        })
        .filter((m): m is UserMemory => m !== null)
        .sort((a, b) => b.createdAt - a.createdAt)
    } catch {
      // a corrupted file must not take AI features down with it
      return []
    }
  }

  /**
   * Record one preference. Returns null when there is nothing worth storing.
   * Restating an existing preference refreshes it instead of adding a copy, so
   * a chatty agent cannot fill the prompt budget with duplicates.
   */
  addMemory(text: unknown): UserMemory | null {
    const clean = normalizeMemoryText(text)
    if (!clean) return null
    const existing = this.readMemories()
    const duplicate = existing.find((m) => m.text.toLowerCase() === clean.toLowerCase())
    const entry: UserMemory = duplicate
      ? { ...duplicate, createdAt: Date.now() }
      : { id: freshMemoryId(), text: clean, createdAt: Date.now() }
    this.writeMemories([entry, ...existing.filter((m) => m.id !== entry.id)].slice(0, MAX_MEMORIES))
    return entry
  }

  /** Returns false when there was nothing with that id to remove. */
  deleteMemory(id: string): boolean {
    const existing = this.readMemories()
    const next = existing.filter((m) => m.id !== id)
    if (next.length === existing.length) return false
    this.writeMemories(next)
    return true
  }

  private writeMemories(memories: UserMemory[]): void {
    mkdirSync(this.baseDir, { recursive: true })
    writeFileSync(this.memoryPath, JSON.stringify(memories, null, 2))
  }

  /** unused stem, suffixing `-2`, `-3`… when the name is already taken */
  private freshId(base: string): string {
    const stem = base || 'skill'
    if (!existsSync(join(this.skillsDir, `${stem}.md`))) return stem
    for (let n = 2; n < 500; n++) {
      const candidate = `${stem}-${n}`
      if (!existsSync(join(this.skillsDir, `${candidate}.md`))) return candidate
    }
    return `${stem}-${Date.now()}`
  }
}

function freshMemoryId(): string {
  return `m${Date.now().toString(36)}${Math.floor(Math.random() * 1e6).toString(36)}`
}

/** filename-safe, lowercase, no traversal */
function safeId(raw: string): string {
  return String(raw ?? '')
    .toLowerCase()
    .replace(/[^a-z0-9._-]+/g, '-')
    .replace(/^[-.]+|[-.]+$/g, '')
    .slice(0, 60)
}

function normalizeScopes(raw: unknown[] | undefined): InstructionScope[] {
  const list = Array.isArray(raw) ? raw : []
  const out: InstructionScope[] = []
  for (const item of list) {
    const scope = coerceScope(item)
    if (!out.includes(scope)) out.push(scope)
  }
  return out.length ? out : ['global']
}
