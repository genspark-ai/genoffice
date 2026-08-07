/**
 * On-disk store for user-authored agent instructions.
 *
 * Layout under userData:
 *   agent-rules.json     { global?, docx?, pptx?, sheets?, pdf? }
 *   agent-skills/<id>.md  one markdown file per skill, front matter and all
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
  coerceScope,
  parseSkillMarkdown,
  serializeSkillMarkdown,
  type AgentRules,
  type InstructionScope,
  type UserSkill,
} from '@genoffice/agent-core'

const RULES_FILE = 'agent-rules.json'
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
