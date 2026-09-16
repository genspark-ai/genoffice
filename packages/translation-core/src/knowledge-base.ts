/**
 * Translation Knowledge Base (KB) — 5-schema model with priority resolution.
 *
 * Mirrors the `trade.translation.*` schema set from LumosAI's translate-config
 * skill, so users who already maintain a translation KB in
 * `~/.lumosai/knowledge/translation_kb.json` can import it directly. GenOffice
 * keeps its own copy under `~/.genoffice/translation-kb.json`; the two files
 * stay independent but share shape.
 *
 * Five schemas (see {@link SchemaId}):
 *
 *   - trade.translation.term            — mandatory source -> target mappings
 *   - trade.translation.forbidden       — banned translations + replacements
 *   - trade.translation.brand           — brand policy (neverTranslate / keep / translateAs)
 *   - trade.translation.styleRule       — tone / formality / voice rules
 *   - trade.translation.customerPreference — per-customer overrides
 *
 * Priority resolution (see {@link KnowledgeBase.resolve}):
 *
 *   session > customer > project > company > global
 *
 * A `customer` rule wins over a `company` rule with the same source term. When
 * multiple rules in the same scope match, the higher `priority` value wins
 * (ties broken by lexicographic `id` so the resolver is deterministic).
 *
 * The KB is JSON-backed so it survives process restarts. `load()` reads from
 * disk; `save()` writes atomically via a `.tmp` + rename. Tests inject an
 * in-memory filesystem via {@link KnowledgeBaseOptions.fileSystem}.
 */

import { mkdir, readFile, writeFile, rename } from 'node:fs/promises'
import path from 'node:path'

export const SCOPES = ['session', 'customer', 'project', 'company', 'global'] as const
export type Scope = (typeof SCOPES)[number]

export const SCHEMA_IDS = [
  'trade.translation.term',
  'trade.translation.forbidden',
  'trade.translation.brand',
  'trade.translation.styleRule',
  'trade.translation.customerPreference',
] as const
export type SchemaId = (typeof SCHEMA_IDS)[number]

export type SchemaKey = 'term' | 'forbidden' | 'brand' | 'styleRule' | 'customerPreference'

const KEY_TO_SCHEMA: Record<SchemaKey, SchemaId> = {
  term: 'trade.translation.term',
  forbidden: 'trade.translation.forbidden',
  brand: 'trade.translation.brand',
  styleRule: 'trade.translation.styleRule',
  customerPreference: 'trade.translation.customerPreference',
}

const SCHEMA_TO_KEY: Record<SchemaId, SchemaKey> = {
  'trade.translation.term': 'term',
  'trade.translation.forbidden': 'forbidden',
  'trade.translation.brand': 'brand',
  'trade.translation.styleRule': 'styleRule',
  'trade.translation.customerPreference': 'customerPreference',
}

interface LangScoped {
  sourceLang?: string | undefined
  targetLang?: string | undefined
}

export interface TermEntry extends LangScoped {
  id: string
  scope: Scope
  priority: number
  sourceTerm: string
  targetTerm: string
  category?: string | undefined
  notes?: string | undefined
}

export interface ForbiddenEntry extends LangScoped {
  id: string
  scope: Scope
  priority: number
  forbiddenText: string
  replacement?: string | undefined
  reason?: string | undefined
  severity?: 'warning' | 'error' | undefined
}

export interface BrandEntry {
  id: string
  scope: Scope
  priority: number
  word: string
  policy: 'neverTranslate' | 'keep' | 'translateAs'
  translateAs?: string | undefined
}

export interface StyleRuleEntry {
  id: string
  scope: Scope
  priority: number
  name: string
  description: string
}

export interface CustomerPreferenceEntry extends LangScoped {
  id: string
  scope: Scope
  priority: number
  customerName: string
  preferenceType: string
  value: string
}

export type KBEntry =
  | TermEntry
  | ForbiddenEntry
  | BrandEntry
  | StyleRuleEntry
  | CustomerPreferenceEntry

export type KBStore = Partial<Record<SchemaId, KBEntry[]>>

export interface KnowledgeBaseFileSystem {
  mkdir(path: string, options?: { recursive?: boolean }): Promise<void>
  readFile(path: string, encoding: 'utf8'): Promise<string>
  writeFile(path: string, contents: string, encoding: 'utf8'): Promise<void>
  rename(from: string, to: string): Promise<void>
}

const defaultFileSystem: KnowledgeBaseFileSystem = {
  mkdir: (p, o) => mkdir(p, o).then(() => undefined),
  readFile: (p, e) => readFile(p, e) as Promise<string>,
  writeFile: (p, c, e) => writeFile(p, c, e),
  rename: (from, to) => rename(from, to).then(() => undefined),
}

export interface KnowledgeBaseOptions {
  /** Override the on-disk path. Defaults to `~/.genoffice/translation-kb.json`. */
  filePath?: string
  /** Inject a custom filesystem implementation (tests). */
  fileSystem?: KnowledgeBaseFileSystem
  /** Pre-seeded entries; useful for tests and one-off constructors. */
  seed?: KBStore
}

export interface ResolvedRules {
  terms: TermEntry[]
  forbidden: ForbiddenEntry[]
  brands: BrandEntry[]
  styleRules: StyleRuleEntry[]
  customerPreferences: CustomerPreferenceEntry[]
  /** Human-readable prompt block; empty string when there are no rules. */
  promptBlock: string
}

export interface KBListFilters {
  schema?: SchemaId | SchemaKey | undefined
  scope?: Scope | undefined
  sourceLang?: string | undefined
  targetLang?: string | undefined
}

/**
 * Default on-disk location. `GENOFFICE_TRANSLATION_KB` overrides it so tests
 * and portable installs can point at a scratch file without touching the
 * user's real KB.
 */
const DEFAULT_FILE_PATH =
  process.env.GENOFFICE_TRANSLATION_KB ??
  path.join(
    process.env.HOME ?? path.join(path.sep, 'tmp'),
    '.genoffice',
    'translation-kb.json',
  )

export class KnowledgeBase {
  private store: KBStore
  private readonly filePath: string
  private readonly fs: KnowledgeBaseFileSystem
  private dirty = false

  constructor(opts: KnowledgeBaseOptions = {}) {
    this.filePath = opts.filePath ?? DEFAULT_FILE_PATH
    this.fs = opts.fileSystem ?? defaultFileSystem
    this.store = cloneStore(opts.seed ?? {})
  }

  /** Read the on-disk store. Returns an empty store if the file is missing. */
  async load(): Promise<void> {
    try {
      const raw = await this.fs.readFile(this.filePath, 'utf8')
      const parsed: unknown = JSON.parse(raw)
      if (!isKBStore(parsed)) throw new Error('translation-kb.json: invalid shape')
      this.store = parsed
      this.dirty = false
    } catch (error) {
      if (isNotFound(error)) {
        this.store = {}
        this.dirty = false
        return
      }
      throw error
    }
  }

  /** Write the current store atomically (`.tmp` + rename). */
  async save(): Promise<void> {
    const dir = path.dirname(this.filePath)
    await this.fs.mkdir(dir, { recursive: true })
    const tmp = this.filePath + '.tmp'
    await this.fs.writeFile(tmp, JSON.stringify(this.store, null, 2), 'utf8')
    await this.fs.rename(tmp, this.filePath)
    this.dirty = false
  }

  /** Whether the in-memory state diverges from disk. */
  isDirty(): boolean {
    return this.dirty
  }

  /** Look up entries in a schema (or across all schemas). */
  list(filters: KBListFilters = {}): KBEntry[] {
    const schemas = filters.schema
      ? [normalizeSchemaId(filters.schema)]
      : (SCHEMA_IDS as readonly SchemaId[])
    const out: KBEntry[] = []
    for (const schema of schemas) {
      const entries = this.store[schema]
      if (!entries) continue
      for (const entry of entries) {
        if (filters.scope && entry.scope !== filters.scope) continue
        if (filters.sourceLang || filters.targetLang) {
          const lang = entry as LangScoped
          if (filters.sourceLang && lang.sourceLang && lang.sourceLang !== filters.sourceLang) continue
          if (filters.targetLang && lang.targetLang && lang.targetLang !== filters.targetLang) continue
        }
        out.push(entry)
      }
    }
    return out
  }

  /** Insert or replace an entry by id. Returns the resolved entry. */
  upsert(entry: KBEntry): KBEntry {
    const schemaId = schemaForEntry(entry)
    const bucket = this.store[schemaId] ?? []
    const next = bucket.filter((e) => e.id !== entry.id)
    next.push(entry)
    this.store[schemaId] = next
    this.dirty = true
    return entry
  }

  /** Remove an entry by id. Returns true when something was removed. */
  remove(id: string): boolean {
    let removed = false
    for (const schema of SCHEMA_IDS) {
      const bucket = this.store[schema]
      if (!bucket) continue
      const next = bucket.filter((e) => {
        if (e.id !== id) return true
        removed = true
        return false
      })
      if (removed) this.store[schema] = next
    }
    return removed
  }

  /**
   * Resolve the active rule set for a language pair (and optional customer /
   * glossary category). Returns the rules in priority order and a
   * human-readable prompt block.
   *
   * `category` is matched against the entry's `category` (term) when present;
   * category-specific rules win over generic ones with the same priority.
   * `customerName` filters {@link CustomerPreferenceEntry} entries.
   */
  resolve(opts: {
    sourceLang: string
    targetLang: string
    category?: string | undefined
    customerName?: string | undefined
  }): ResolvedRules {
    const terms: TermEntry[] = []
    const forbidden: ForbiddenEntry[] = []
    const brands: BrandEntry[] = []
    const styleRules: StyleRuleEntry[] = []
    const customerPreferences: CustomerPreferenceEntry[] = []

    const passesLang = (entry: LangScoped): boolean => {
      if (entry.sourceLang && entry.sourceLang !== opts.sourceLang) return false
      if (entry.targetLang && entry.targetLang !== opts.targetLang) return false
      return true
    }

    for (const schema of SCHEMA_IDS) {
      const entries = this.store[schema] ?? []
      for (const entry of entries) {
        if (schema === 'trade.translation.term') {
          const e = entry as TermEntry
          if (!passesLang(e)) continue
          if (opts.category && e.category && e.category !== opts.category) continue
          terms.push(e)
        } else if (schema === 'trade.translation.forbidden') {
          const e = entry as ForbiddenEntry
          if (!passesLang(e)) continue
          forbidden.push(e)
        } else if (schema === 'trade.translation.brand') {
          brands.push(entry as BrandEntry)
        } else if (schema === 'trade.translation.styleRule') {
          styleRules.push(entry as StyleRuleEntry)
        } else {
          const e = entry as CustomerPreferenceEntry
          if (!passesLang(e)) continue
          if (opts.customerName && e.customerName !== opts.customerName) continue
          customerPreferences.push(e)
        }
      }
    }

    sortByPriority(terms)
    sortByPriority(forbidden)
    sortByPriority(brands)
    sortByPriority(styleRules)
    sortByPriority(customerPreferences)

    const promptBlock = renderPromptBlock({
      sourceLang: opts.sourceLang,
      targetLang: opts.targetLang,
      terms,
      forbidden,
      brands,
      styleRules,
      customerPreferences,
    })

    return { terms, forbidden, brands, styleRules, customerPreferences, promptBlock }
  }
}

/* ------------------------------------------------------------------ */
/* Helpers.                                                            */
/* ------------------------------------------------------------------ */

function cloneStore(store: KBStore): KBStore {
  const out: KBStore = {}
  for (const schema of SCHEMA_IDS) {
    const bucket = store[schema]
    if (bucket && bucket.length > 0) out[schema] = [...bucket]
  }
  return out
}

function normalizeSchemaId(raw: SchemaId | SchemaKey): SchemaId {
  return (KEY_TO_SCHEMA as Record<string, SchemaId>)[raw] ?? (raw as SchemaId)
}

function schemaForEntry(entry: KBEntry): SchemaId {
  if ('sourceTerm' in entry && typeof (entry as TermEntry).sourceTerm === 'string') {
    return 'trade.translation.term'
  }
  if ('forbiddenText' in entry && typeof (entry as ForbiddenEntry).forbiddenText === 'string') {
    return 'trade.translation.forbidden'
  }
  if ('word' in entry && 'policy' in entry) {
    return 'trade.translation.brand'
  }
  if ('name' in entry && 'description' in entry) {
    return 'trade.translation.styleRule'
  }
  return 'trade.translation.customerPreference'
}

function sortByPriority<T extends { scope: Scope; priority: number; id: string }>(entries: T[]): void {
  entries.sort((a, b) => {
    const sa = SCOPES.indexOf(a.scope)
    const sb = SCOPES.indexOf(b.scope)
    if (sa !== sb) return sa - sb
    if (a.priority !== b.priority) return b.priority - a.priority
    return a.id.localeCompare(b.id)
  })
}

interface PromptBlockInput {
  sourceLang: string
  targetLang: string
  terms: TermEntry[]
  forbidden: ForbiddenEntry[]
  brands: BrandEntry[]
  styleRules: StyleRuleEntry[]
  customerPreferences: CustomerPreferenceEntry[]
}

function renderPromptBlock(input: PromptBlockInput): string {
  const sections: string[] = []
  if (input.terms.length > 0) {
    const rows = input.terms
      .slice(0, 30)
      .map((t) => `  - "${t.sourceTerm}" -> "${t.targetTerm}"`)
      .join('\n')
    sections.push(`Mandatory terms:\n${rows}`)
  }
  if (input.forbidden.length > 0) {
    const rows = input.forbidden
      .slice(0, 20)
      .map((f) => {
        const replacement = f.replacement ? ` -> use "${f.replacement}"` : ''
        return `  - never write "${f.forbiddenText}"${replacement}`
      })
      .join('\n')
    sections.push(`Forbidden translations:\n${rows}`)
  }
  if (input.brands.length > 0) {
    const rows = input.brands
      .slice(0, 30)
      .map((b) => {
        if (b.policy === 'neverTranslate') return `  - "${b.word}" must remain untranslated`
        if (b.policy === 'translateAs') return `  - "${b.word}" -> "${b.translateAs ?? ''}"`
        return `  - "${b.word}" - keep as-is`
      })
      .join('\n')
    sections.push(`Brand rules:\n${rows}`)
  }
  if (input.styleRules.length > 0) {
    const rows = input.styleRules
      .slice(0, 10)
      .map((s) => `  - [${s.scope}/${s.priority}] ${s.name}: ${s.description}`)
      .join('\n')
    sections.push(`Style rules:\n${rows}`)
  }
  if (input.customerPreferences.length > 0) {
    const rows = input.customerPreferences
      .slice(0, 10)
      .map((p) => `  - ${p.customerName}: ${p.preferenceType}=${p.value}`)
      .join('\n')
    sections.push(`Customer preferences:\n${rows}`)
  }
  if (sections.length === 0) return ''
  return `Translation rules (${input.sourceLang} -> ${input.targetLang}):\n${sections.join('\n\n')}`
}

function isKBStore(value: unknown): value is KBStore {
  if (value === null || typeof value !== 'object') return false
  const obj = value as Record<string, unknown>
  for (const key of Object.keys(obj)) {
    if (!(SCHEMA_IDS as readonly string[]).includes(key)) return false
    const bucket = obj[key]
    if (!Array.isArray(bucket)) return false
  }
  return true
}

function isNotFound(error: unknown): boolean {
  return (
    typeof error === 'object' &&
    error !== null &&
    'code' in error &&
    (error as { code?: string }).code === 'ENOENT'
  )
}

export { KEY_TO_SCHEMA as SCHEMA_KEY_TO_ID, SCHEMA_TO_KEY }
