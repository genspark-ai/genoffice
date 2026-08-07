/**
 * User-authored agent instructions: **rules** and **skills**.
 *
 * These are the user's own text, not code — distinct from `AgentSkill` in
 * ./skill.ts, which is a compiled-in capability domain (tools + executor).
 * Naming is deliberately `UserSkill` here to keep the two apart.
 *
 * - **Rules** are always-on style/convention text. One global block plus one
 *   per app surface, concatenated into the system prompt every turn.
 * - **Skills** are markdown playbooks that are *not* pasted in wholesale.
 *   Only each skill's name and description reach the system prompt; the agent
 *   calls `load_skill` to pull a body in when a task actually calls for it.
 *   That keeps a large library from crowding out the context window.
 *
 * Pure functions only: file storage lives in the Electron main process.
 */

/**
 * Where an instruction applies. 'docx' / 'pptx' / 'sheets' / 'pdf' name the
 * editor surfaces; 'global' applies everywhere. ('slides' is accepted as an
 * alias for 'pptx' when parsing user-written front matter, since the UI and
 * the product name disagree about which word to use.)
 */
export type InstructionScope = 'global' | 'docx' | 'pptx' | 'sheets' | 'pdf'

export const INSTRUCTION_SCOPES: readonly InstructionScope[] = [
  'global',
  'docx',
  'pptx',
  'sheets',
  'pdf',
]

/** The surface asking for instructions — every scope except 'global'. */
export type AppSurface = Exclude<InstructionScope, 'global'>

export function isInstructionScope(value: unknown): value is InstructionScope {
  return INSTRUCTION_SCOPES.includes(value as InstructionScope)
}

/** Accept the aliases users actually type before falling back to 'global'. */
export function coerceScope(value: unknown): InstructionScope {
  const raw = String(value ?? '')
    .trim()
    .toLowerCase()
  if (raw === 'slides' || raw === 'slide' || raw === 'ppt') return 'pptx'
  if (raw === 'doc' || raw === 'docs' || raw === 'word') return 'docx'
  if (raw === 'sheet' || raw === 'xlsx' || raw === 'excel') return 'sheets'
  return isInstructionScope(raw) ? raw : 'global'
}

/** Always-on instruction text, one block per scope. */
export type AgentRules = Partial<Record<InstructionScope, string>>

/** A user-authored markdown playbook. */
export interface UserSkill {
  /** stable id; also the storage filename stem */
  id: string
  name: string
  /** one line telling the agent when this skill is worth loading */
  description: string
  scopes: InstructionScope[]
  /** markdown body, everything after the front matter */
  body: string
  /** false hides it from the agent without deleting it */
  enabled: boolean
}

/** A skill as advertised to the model, before its body is loaded. */
export type UserSkillSummary = Pick<UserSkill, 'id' | 'name' | 'description'>

export const LOAD_SKILL_TOOL = 'load_skill'

// ── front matter ────────────────────────────────────────────────────

const FRONT_MATTER = /^---\r?\n([\s\S]*?)\r?\n---\r?\n?/

/**
 * Parse a `skill.md`: optional YAML-ish front matter (`name`, `description`,
 * `scope`/`scopes`, `enabled`) followed by the markdown body. Deliberately not
 * a full YAML parser — these are hand-written files with flat scalar keys, and
 * a missing or malformed header degrades to sensible defaults rather than
 * rejecting the file, so an uploaded skill is never silently lost.
 */
export function parseSkillMarkdown(source: string, fallbackId: string): UserSkill {
  // strip a UTF-8 BOM: editors on Windows add one, and it would otherwise
  // stop the front-matter delimiter from matching at position 0
  const text = source.charCodeAt(0) === 0xfeff ? source.slice(1) : source
  const match = FRONT_MATTER.exec(text)
  const body = (match ? text.slice(match[0].length) : text).trim()
  const meta = new Map<string, string>()
  if (match) {
    for (const line of match[1]!.split(/\r?\n/)) {
      const kv = /^([A-Za-z_][\w-]*)\s*:\s*(.*)$/.exec(line.trim())
      if (kv) meta.set(kv[1]!.toLowerCase(), stripQuotes(kv[2]!.trim()))
    }
  }
  const rawScopes = meta.get('scopes') ?? meta.get('scope') ?? ''
  const scopes = parseScopeList(rawScopes)
  const enabled = meta.get('enabled')
  return {
    id: fallbackId,
    name: meta.get('name') || firstHeading(body) || fallbackId,
    description: meta.get('description') || firstParagraph(body),
    scopes: scopes.length ? scopes : ['global'],
    body,
    enabled: enabled === undefined ? true : enabled !== 'false',
  }
}

/** Serialize back to `skill.md` so a round trip through the editor is lossless. */
export function serializeSkillMarkdown(skill: UserSkill): string {
  const front = [
    '---',
    `name: ${skill.name}`,
    `description: ${skill.description}`,
    `scopes: ${skill.scopes.join(', ')}`,
    `enabled: ${skill.enabled}`,
    '---',
  ].join('\n')
  return `${front}\n\n${skill.body.trim()}\n`
}

function parseScopeList(raw: string): InstructionScope[] {
  const parts = raw
    .replace(/^\[|\]$/g, '')
    .split(/[,\s]+/)
    .map((s) => stripQuotes(s).trim())
    .filter(Boolean)
  const out: InstructionScope[] = []
  for (const part of parts) {
    const scope = coerceScope(part)
    if (!out.includes(scope)) out.push(scope)
  }
  return out
}

function stripQuotes(s: string): string {
  return s.replace(/^['"]|['"]$/g, '')
}

function firstHeading(body: string): string {
  return /^#{1,6}\s+(.+)$/m.exec(body)?.[1]?.trim() ?? ''
}

function firstParagraph(body: string): string {
  for (const line of body.split(/\r?\n/)) {
    const trimmed = line.trim()
    if (trimmed && !trimmed.startsWith('#')) return trimmed.slice(0, 200)
  }
  return ''
}

// ── selection + prompt assembly ─────────────────────────────────────

/** Does this scope list apply to the surface currently asking? */
export function scopeApplies(scopes: InstructionScope[], surface: AppSurface): boolean {
  return scopes.includes('global') || scopes.includes(surface)
}

/** Enabled skills that apply to this surface. */
export function skillsForSurface(skills: UserSkill[], surface: AppSurface): UserSkill[] {
  return skills.filter((s) => s.enabled && scopeApplies(s.scopes, surface))
}

/**
 * The rules section for one surface: the global block, then the surface's own,
 * so a more specific rule reads as a refinement of the general one.
 */
export function buildRulesPrompt(rules: AgentRules, surface: AppSurface): string {
  const blocks = [rules.global, rules[surface]].map((b) => (b ?? '').trim()).filter(Boolean)
  if (!blocks.length) return ''
  return ['## User rules', 'Follow these for every response in this app.', ...blocks].join('\n\n')
}

/**
 * The skills catalogue: names and descriptions only, plus how to fetch a body.
 * Returns '' when nothing applies, so no empty heading reaches the prompt.
 */
export function buildSkillsPrompt(skills: UserSkill[], surface: AppSurface): string {
  const applicable = skillsForSurface(skills, surface)
  if (!applicable.length) return ''
  const lines = applicable.map(
    (s) => `- \`${s.id}\` — **${s.name}**${s.description ? `: ${s.description}` : ''}`,
  )
  return [
    '## User skills',
    `These are step-by-step playbooks the user wrote. Only the titles are listed here. When a task matches one, call \`${LOAD_SKILL_TOOL}\` with its id to read the full instructions **before** acting, then follow them.`,
    ...lines,
  ].join('\n')
}

/** Body text returned by the `load_skill` tool. */
export function skillBodyForTool(skills: UserSkill[], surface: AppSurface, id: string): string {
  const skill = skillsForSurface(skills, surface).find((s) => s.id === id)
  if (!skill) {
    const available = skillsForSurface(skills, surface)
      .map((s) => s.id)
      .join(', ')
    return `No skill "${id}" available here.${available ? ` Available: ${available}` : ''}`
  }
  return `# ${skill.name}\n\n${skill.body}`
}

/** Combined instruction block appended to an app's system prompt. */
export function buildInstructionsPrompt(
  rules: AgentRules,
  skills: UserSkill[],
  surface: AppSurface,
): string {
  return [buildRulesPrompt(rules, surface), buildSkillsPrompt(skills, surface)]
    .filter(Boolean)
    .join('\n\n')
}
