/**
 * pi resource bridge.
 *
 * GenOffice does not re-implement resource discovery. Marketplace skills and
 * plugins are handed to pi's own settings + package manager, so one install is
 * visible to (a) the embedded pi session, (b) the `pi` CLI pointed at the same
 * agent dir, and (c) this server's Settings UI — from a single source of truth
 * on disk:
 *
 *   DATA_DIR/pi-agent/          pi's agent dir (settings.json, auth, npm cache)
 *   DATA_DIR/pi-skills/<id>/    installed marketplace skill (SKILL.md per id)
 *   DATA_DIR/pi-plugins/<id>/   installed marketplace plugin (a pi package)
 *
 * A "plugin" is a pi package: a directory with `package.json` (`pi` manifest or
 * conventional `extensions/` + `skills/` dirs). Plugins that ship an uploaded
 * extension source get a real `extensions/index.ts`; plugins without one are
 * installed as guidance-only packages (a generated SKILL.md) and never fake a
 * tool. Every registration goes through pi's SettingsManager, and every read
 * goes through pi's DefaultPackageManager — no bespoke loader, no duplicated
 * frontmatter rules.
 *
 * Disabled resources are moved out of the discovery path (skills) or registered
 * as `{ source, autoload: false }` packages (plugins) so "disabled" means the
 * agent really stops seeing them instead of only flipping a UI flag.
 */
import {
  existsSync,
  mkdirSync,
  readdirSync,
  readFileSync,
  renameSync,
  rmSync,
  statSync,
  writeFileSync,
} from 'node:fs'
import { homedir } from 'node:os'
import { dirname, join, resolve, sep } from 'node:path'
import {
  DefaultPackageManager,
  SettingsManager,
  loadSkillsFromDir,
  type PackageSource,
  type Skill,
} from '@earendil-works/pi-coding-agent'
import { DATA_DIR } from '../common/index'

/** Working directory pi resolves relative package paths against. */
export const PI_CWD = join(DATA_DIR, 'pi-cwd')
/** pi's agent dir — holds settings.json, auth.json and the npm package cache. */
export const PI_AGENT_DIR = join(DATA_DIR, 'pi-agent')
/** Skills installed from the marketplace, one `<id>/SKILL.md` per skill. */
export const PI_SKILLS_DIR = join(DATA_DIR, 'pi-skills')
/** Disabled skills are parked here, outside pi's discovery path. */
export const PI_SKILLS_DISABLED_DIR = join(DATA_DIR, 'pi-skills-disabled')
/** Local plugin packages built from marketplace entries. */
export const PI_PLUGIN_DIR = join(DATA_DIR, 'pi-plugins')
/** Scratch space for validating an uploaded SKILL.md before it is published. */
export const PI_STAGING_DIR = join(DATA_DIR, 'pi-staging')
/** Wrapper SKILL.md files generated from LumosAI's bundled translate suite —
 *  GenOffice owns the wrapper frontmatter, the upstream LumosAI scripts are
 *  untouched. Listed here so the pi-session bridge can hand this directory
 *  to `DefaultResourceLoader.additionalSkillPaths`. */
export const LUMOS_SKILLS_WRAPPER_DIR = join(DATA_DIR, 'lumos-skill-wrappers')

for (const dir of [PI_CWD, PI_AGENT_DIR, PI_SKILLS_DIR, PI_PLUGIN_DIR, PI_STAGING_DIR]) {
  try {
    mkdirSync(dir, { recursive: true })
  } catch {
    /* a read-only data dir must not crash the server at import time */
  }
}

/** Managed roots — resources under these are the ones GenOffice owns. */
const MANAGED_ROOTS = [PI_SKILLS_DIR, PI_PLUGIN_DIR, PI_AGENT_DIR]

function isManagedPath(path: string): boolean {
  const abs = resolve(path)
  return MANAGED_ROOTS.some((root) => abs === root || abs.startsWith(root + sep))
}

/** Open pi's settings for our agent dir. A fresh manager per call keeps
 *  externally edited settings.json visible without a reload dance. */
export function openPiSettings(): SettingsManager {
  return SettingsManager.create(PI_CWD, PI_AGENT_DIR)
}

export function piSettingsPath(): string {
  return join(PI_AGENT_DIR, 'settings.json')
}

function openPackageManager(settings = openPiSettings()): DefaultPackageManager {
  return new DefaultPackageManager({ cwd: PI_CWD, agentDir: PI_AGENT_DIR, settingsManager: settings })
}

/** Plain source string of a settings entry, whether it is a bare string or an
 *  `{ source, autoload }` filter object. */
export function packageSourceOf(entry: PackageSource): string {
  return typeof entry === 'string' ? entry : entry.source
}

// ── skill dir registration ──────────────────────────────────────────


/** LumosAI bundles its translate suite under `~/.lumos/bundled-skills/<hash>/translate*`.
 *  Without the wrapper below pi's loader silently skips those directories (it only
 *  understands the Agent Skills frontmatter shape — `name`, `description`, etc.) and
 *  the GenOffice UI's marketplace entry becomes the only place those skills exist.
 *  We materialise a tiny `SKILL.md` wrapper per sibling LumosAI skill so the pi agent
 *  can `bash` into `scripts/translate.py` exactly like a native skill. The LumosAI
 *  scripts themselves are untouched — GenOffice owns only the wrapper frontmatter.
 *
 *  Returned `registered` lists the wrapper SKILL.md paths pi now sees.
 */
export async function ensureLumosSkillsRegistered(): Promise<{ registered: string[]; alreadyHad: string[] }> {
  const lumosHome = process.env.LUMOS_HOME ?? join(homedir(), '.lumos')
  const bundledRoot = join(lumosHome, 'bundled-skills')
  if (!existsSync(bundledRoot)) return { registered: [], alreadyHad: [] }

  const settings = openPiSettings()
  const existing = new Set(settings.getSkillPaths())
  const registered: string[] = []
  const alreadyHad: string[] = []
  const wrapDirs: string[] = []

  // Pick the newest hash; older hashes are stale snapshots we do not shadow.
  let newestHash: string | null = null
  let newestMtime = 0
  for (const entry of readdirSync(bundledRoot)) {
    const candidate = join(bundledRoot, entry)
    let mtime = 0
    try {
      mtime = statSync(candidate).mtimeMs
    } catch {
      continue
    }
    if (!newestHash || mtime > newestMtime) {
      newestHash = entry
      newestMtime = mtime
    }
  }
  if (!newestHash) return { registered: [], alreadyHad: [] }

  const skillRoot = join(bundledRoot, newestHash)
  const wrapperRoot = join(DATA_DIR, 'lumos-skill-wrappers')
  mkdirSync(wrapperRoot, { recursive: true })

  for (const entry of readdirSync(skillRoot)) {
    const candidate = join(skillRoot, entry, 'SKILL.md')
    if (!existsSync(candidate)) continue
    const wrapperDir = join(wrapperRoot, entry)
    const wrapperSkill = join(wrapperDir, 'SKILL.md')
    if (existsSync(wrapperSkill)) {
      alreadyHad.push(wrapperSkill)
      wrapDirs.push(wrapperDir)
      continue
    }
    mkdirSync(wrapperDir, { recursive: true })
    const body = readFileSync(candidate, 'utf-8')
    // Pi reads only frontmatter; everything below is for the agent. We rewrite the
    // body so the agent sees a clean description + a direct bash pointer at the
    // upstream LumosAI script. The LumosAI frontmatter (command_dispatch, etc.) is
    // dropped — pi would warn on those unknown fields anyway.
    const wrapper = renderLumosSkillWrapper({ name: entry, source: candidate, body })
    writeFileSync(wrapperSkill, wrapper, 'utf-8')
    registered.push(wrapperSkill)
    wrapDirs.push(wrapperDir)
  }

  // Register the wrapper root exactly once. Re-adding the same path is a no-op
  // for pi, but we still dedupe so the settings file does not drift.
  if (!existing.has(wrapperRoot)) {
    settings.setSkillPaths([...settings.getSkillPaths(), wrapperRoot])
  }
  // `flush()` is async on the SettingsManager; fire-and-forget — the path is in
  // settings.json either way and pi reads on every session start.
  try {
    await settings.flush()
  } catch (err: unknown) {
    console.warn('[pi-resources] settings.flush failed:', err)
  }

  return { registered, alreadyHad }
}

/**
 * Unescape a YAML single-line scalar. Handles backslash escapes (\\, \", \n,
 * \t, \r, \0) and unknown escapes are kept as-is. Block scalars also benefit
 * from this since they may contain literal `\` characters.
 */
function unescapeYamlScalar(text: string): string {
  return text.replace(/\\(.)/g, (_match, ch: string) => {
    switch (ch) {
      case 'n': return '\\n'
      case 't': return '\\t'
      case 'r': return '\\r'
      case '"': return '"'
      case "'": return "'"
      case '\\': return '\\'
      case '0': return '\\0'
      default: return ch
    }
  })
}

/**
 * Collapse doubled apostrophes inside a YAML single-quoted scalar
 * (YAML uses `''` to escape a literal `'`).
 */
function unescapeYamlSingleQuoted(text: string): string {
  return text.replace(/''/g, "'")
}

/**
 * Walk a string from index 0, treating every `\X` (YAML escape) as a single
 * token, and return the index of the first unescaped `"` (or `'`). Used by
 * the double-quoted and single-quoted scalar branches so a description like
 * `app_update(action=\"check\")` is not cut short at the first escaped quote.
 */
function indexOfClosingQuote(haystack: string, quote: '"' | "'"): number {
  let k = 0
  while (k < haystack.length) {
    const ch = haystack[k]
    if (ch === '\\' && quote === '"' && k + 1 < haystack.length) {
      k += 2 // skip \X as one token
      continue
    }
    if (ch === "'" && quote === "'" && haystack[k + 1] === "'") {
      k += 2 // skip '' as one token (YAML single-quoted escape)
      continue
    }
    if (ch === quote) return k
    k++
  }
  return -1
}

export function extractLumosDescription(body: string, fallback: string): string {
  // LumosAI SKILL.md frontmatter uses YAML multi-line forms — folded `>`
  // /literal `|` block scalars, single- or double-quoted strings that span
  // many lines. We parse the top-level description by reading the line after
  // `description:` and continuing based on the first character of its content.
  // If we cannot find one we fall back to the file's name.
  const start = body.indexOf('---')
  if (start < 0) return fallback
  const fmEnd = body.indexOf('\n---', start + 3)
  if (fmEnd < 0) return fallback
  const frontmatter = body.slice(0, fmEnd)
  const lines = frontmatter.split('\n')
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i]
    const m = line.match(/^description:\s*(.*)$/)
    if (!m) continue
    const content = m[1]
    const trimmed = content.trim()
    if (!trimmed) continue
    const firstChar = trimmed[0]
    if (firstChar === '"') {
      const afterQuote = trimmed.slice(trimmed.indexOf('"') + 1)
      const closing = indexOfClosingQuote(afterQuote, '"')
      if (closing >= 0) {
        return unescapeYamlScalar(afterQuote.slice(0, closing)).replace(/\s+/g, ' ').trim()
      }
      let buffer = afterQuote
      for (let j = i + 1; j < lines.length; j++) {
        const idx = indexOfClosingQuote(lines[j], '"')
        if (idx >= 0) {
          buffer += ' ' + lines[j].slice(0, idx)
          return unescapeYamlScalar(buffer).replace(/\s+/g, ' ').trim()
        }
        buffer += ' ' + lines[j]
      }
      return unescapeYamlScalar(buffer).replace(/\s+/g, ' ').trim()
    }
    if (firstChar === "'") {
      const afterQuote = trimmed.slice(trimmed.indexOf("'") + 1)
      const closing = indexOfClosingQuote(afterQuote, "'")
      if (closing >= 0) {
        return unescapeYamlSingleQuoted(
          unescapeYamlScalar(afterQuote.slice(0, closing)),
        ).replace(/\s+/g, ' ').trim()
      }
      let buffer = afterQuote
      for (let j = i + 1; j < lines.length; j++) {
        const idx = indexOfClosingQuote(lines[j], "'")
        if (idx >= 0) {
          buffer += ' ' + lines[j].slice(0, idx)
          return unescapeYamlSingleQuoted(
            unescapeYamlScalar(buffer),
          ).replace(/\s+/g, ' ').trim()
        }
        buffer += ' ' + lines[j]
      }
      return unescapeYamlSingleQuoted(
        unescapeYamlScalar(buffer),
      ).replace(/\s+/g, ' ').trim()
    }
    if (firstChar === '>' || firstChar === '|') {
      // Folded/literal block scalar — collect every indented line until a
      // non-indented line (which is the next YAML key).
      let buffer = ''
      for (let j = i + 1; j < lines.length; j++) {
        if (lines[j].startsWith(' ') || lines[j].startsWith('\t')) {
          buffer += (buffer ? ' ' : '') + lines[j].trim()
        } else {
          break
        }
      }
      return unescapeYamlScalar(buffer).trim() || fallback
    }
    // Plain scalar on the same line.
    return trimmed
  }
  return fallback
}




/**
 * Encode a string so it can sit inside a YAML double-quoted scalar: backslash
 * and double-quote are the only two characters that need escaping.
 */
function yamlDoubleQuoted(text: string): string {
  return text.replace(/\\/g, '\\\\').replace(/"/g, '\\"')
}

function renderLumosSkillWrapper(opts: {
  name: string
  source: string
  body: string
}): string {
  const description = extractLumosDescription(opts.body, `LumosAI translation skill — ${opts.name}`).slice(0, 1024)
  // Re-slug the directory name so pi accepts it (lowercase a-z, 0-9, hyphen only,
  // no leading/trailing hyphens, no consecutive hyphens, max 64 chars).
  const slug = opts.name
    .toLowerCase()
    .replace(/[^a-z0-9-]/g, '-')
    .replace(/-+/g, '-')
    .replace(/^-|-$/g, '')
    .slice(0, 64) || 'skill'
  const scriptPath = join(dirname(opts.source), 'scripts', 'translate.py')
  const fallback = `python3 ${scriptPath} <input> <output> --dictionary <dict.json>`
  return [
    '---',
    `name: ${slug}`,
    // Keep the description short and behavioural per pi.dev guidance; the full
    // upstream text lives in the body so the agent has it when it bash-calls
    // the script.
    `description: "${yamlDoubleQuoted(description).slice(0, 1024)}"`,
    // Pi ignores unknown metadata fields, so we drop the LumosAI-specific
    // `metadata.hermes` block (which only ever meant something to LumosAI).
    '---',
    '',
    `# ${opts.name}`,
    '',
    'This is a GenOffice wrapper around a LumosAI skill. The upstream `SKILL.md` lives',
    `at \`${opts.source}\` and is kept verbatim. Pi only reads the frontmatter above; the`,
    'body below is the agent-facing playbook and points at the upstream Python script.',
    '',
    '## Setup',
    '',
    'The upstream script depends on the LumosAI Python stack:',
    '',
    '```bash',
    'pip install pypdfium2 pdfplumber reportlab python-docx python-pptx openpyxl xlrd Pillow',
    '```',
    '',
    '## When to load',
    '',
    description,
    '',
    '## How to run',
    '',
    'Call the upstream translate script via `bash`. The skill directory contains the',
    'Python handler that knows the file format:',
    '',
    '```bash',
    fallback,
    '```',
    '',
    '### Generating a dictionary first',
    '',
    'The script accepts `--dictionary <path-to-json>` with `{ "source": "target" }`',
    'mappings. To get one for an unknown file:',
    '',
    '1. Run the upstream script in `--dry-run` mode (it will list untranslated strings).',
    '2. Or build one with the GenOffice KB + LLM pass:',
    '   `POST /api/ipc/ai:translate-build-dictionary`',
    `3. Or mine the file yourself with \`pdfplumber\` / \`openpyxl\` and translate the`,
    '   unique strings.',
    '',
    '## Supported formats',
    '',
    'PDF (.pdf), Excel (.xls/.xlsx), PowerPoint (.pptx), Word (.docx).',
    '',
    '## Failure modes',
    '',
    '- **Missing dependency**: the Python stack needs `pypdfium2`, `pdfplumber`,',
    '  `reportlab`, `python-docx`, `python-pptx`, `openpyxl`, `xlrd`, `Pillow`.',
    '  `/opt/homebrew/bin/python3` ships them on macOS; otherwise `pip install -r`',
    '  the LumosAI requirements.',
    '- **No --dictionary**: the script reformats only — coverage will be 0%.',
    '- **Source not in supported extensions**: refused up-front with the extension list.',
  ].join('\n')
}


/**
 * Materialise SKILL.md files for every built-in skill. Without this the
 * marketplace catalog entries (`DEFAULT_SKILLS` in `skills.ts`) are pure UI
 * metadata — there is nothing on disk for pi's loader to discover, so the
 * agent has no idea the tools (`read_blocks`, `web_search`, etc.) even exist.
 *
 * Idempotent: re-running only writes a SKILL.md when the file is missing.
 * Returns the paths that were written vs. already on disk so callers can log
 * a one-liner that shows the bootstrap moved.
 */
export function ensureBuiltInSkillsMaterialized(): { written: string[]; skipped: string[] } {
  const written: string[] = []
  const skipped: string[] = []
  mkdirSync(PI_SKILLS_DIR, { recursive: true })
  for (const entry of BUILT_IN_SKILLS) {
    const dir = join(PI_SKILLS_DIR, entry.id)
    const skillPath = join(dir, 'SKILL.md')
    if (existsSync(skillPath)) {
      skipped.push(skillPath)
      continue
    }
    mkdirSync(dir, { recursive: true })
    writeFileSync(skillPath, renderBuiltInSkillMarkdown(entry), 'utf-8')
    written.push(skillPath)
  }
  return { written, skipped }
}

/** All built-in skills with the metadata needed to render a pi-compatible
 *  SKILL.md. Keep this list in sync with `DEFAULT_SKILLS` in `skills.ts` —
 *  the two are the same set of tools, just rendered differently (UI vs. pi). */
const BUILT_IN_SKILLS: Array<{
  id: string
  name: string
  description: string
  version: string
  author: string
  tools: string[]
  scopes: string[]
  category: string
  tags?: string[]
}> = [
  { id: 'docs-skill', name: 'Docs Skill', description: 'Word 文档操作工具集(读块/写块/替换/格式/搜索/...)', version: '0.85.1', author: 'GenOffice', tools: ['read_blocks', 'write_block', 'replace_document', 'insert_blocks', 'delete_blocks', 'format_blocks', 'search_blocks', 'find_replace', 'set_page_margins', 'headings_outline'], scopes: ['files:read', 'files:write', 'docs:edit'], category: 'productivity', tags: ['docs', 'word', 'office'] },
  { id: 'sheets-skill', name: 'Sheets Skill', description: 'Excel 表格操作工具集(读 cell/写 cell/公式/图表/筛选/排序)', version: '0.85.1', author: 'GenOffice', tools: ['read_range', 'write_range', 'apply_formula', 'create_chart', 'sort_range', 'filter_range'], scopes: ['files:read', 'files:write', 'sheets:edit'], category: 'productivity', tags: ['sheets', 'excel', 'office'] },
  { id: 'slides-skill', name: 'Slides Skill', description: 'PPT 幻灯片操作工具集(读 slide/写 slide/插入图/布局/演讲备注)', version: '0.85.1', author: 'GenOffice', tools: ['read_slides', 'write_slide', 'insert_image', 'set_layout', 'set_speaker_notes'], scopes: ['files:read', 'files:write', 'slides:edit'], category: 'productivity', tags: ['slides', 'ppt', 'office'] },
  { id: 'office-workflow', name: 'Office Workflow', description: '跨 Office 工作流编排(docs→sheets→slides 数据流)', version: '0.85.1', author: 'GenOffice', tools: ['cross_office_workflow', 'office_data_pipeline'], scopes: ['files:read', 'files:write'], category: 'productivity', tags: ['office', 'workflow'] },
  { id: 'office-safety', name: 'Office Safety', description: 'Office 操作安全检查(用户确认 + 范围限制 + 撤销支持)', version: '0.85.1', author: 'GenOffice', tools: ['confirm_destructive_op', 'scope_check'], scopes: ['safety'], category: 'productivity', tags: ['safety'] },
  { id: 'frozen-selection', name: 'Frozen Selection', description: '冻结 AI 选区,防止后续修改误伤用户意图', version: '0.85.1', author: 'GenOffice', tools: ['freeze_selection', 'unfreeze_selection', 'list_frozen'], scopes: ['docs:edit'], category: 'productivity', tags: ['safety'] },
  { id: 'verify-response', name: 'Verify Response', description: '验证 AI 响应与原文档的一致性(diff + 摘要回归)', version: '0.85.1', author: 'GenOffice', tools: ['verify_response', 'summarize_diff'], scopes: ['ai:stream'], category: 'productivity', tags: ['safety', 'verify'] },
  { id: 'skill-market', name: 'Skill Marketplace', description: 'Skills 市场(浏览/搜索/安装/卸载第三方 skill)', version: '0.85.1', author: 'GenOffice', tools: ['list_marketplace', 'search_skills', 'install_skill', 'uninstall_skill'], scopes: ['network:out'], category: 'dev', tags: ['marketplace'] },
  { id: 'web-search', name: 'Web Search', description: '通过 DuckDuckGo HTML 提供零配置网页搜索,无 API key', version: '1.0.0', author: 'GenOffice', tools: ['web_search'], scopes: ['network:out'], category: 'media', tags: ['search'] },
  { id: 'image-search', name: 'Image Search', description: 'DuckDuckGo 图片搜索 + 图片下载,无 API key,base64 直接喂给多模态模型', version: '1.0.0', author: 'GenOffice', tools: ['image_search', 'fetch_image'], scopes: ['network:out'], category: 'media', tags: ['images'] },
  { id: 'ocr', name: 'OCR Image', description: '读取本地/网络图片为 base64 data URI,供多模态模型做文字识别', version: '1.0.0', author: 'GenOffice', tools: ['ocr_image'], scopes: ['files:read', 'network:out'], category: 'media', tags: ['ocr'] },
  { id: 'agent-team', name: 'Agent Team', description: '多 Agent 团队协作(5 个内置角色 + request_review 工具)', version: '0.85.1', author: 'GenOffice', tools: ['request_review', 'run_agent', 'coordinate_team'], scopes: ['ai:stream', 'agents:multi'], category: 'dev', tags: ['agents'] },
  { id: 'audit-log', name: 'Audit Log', description: '企业级审计日志(3 个 sink: file/otlp/console + 自动 tool_call/result 配对)', version: '0.85.1', author: 'GenOffice', tools: ['export_audit', 'tail_audit'], scopes: ['audit:write'], category: 'dev', tags: ['audit', 'enterprise'] },
  { id: 'local-models', name: 'Local Models (Ollama)', description: 'Ollama 本地模型 provider(createOllamaProvider + installLocalModels)', version: '0.85.1', author: 'GenOffice', tools: ['list_ollama_models', 'install_ollama_model', 'pull_ollama_model'], scopes: ['network:out', 'providers:add'], category: 'dev', tags: ['local', 'ollama'] },
]


function renderBuiltInSkillMarkdown(entry: typeof BUILT_IN_SKILLS[number]): string {
  const scopes = entry.scopes.join(', ')
  // `tags` and `allowed-tools` must be YAML lists per the Agent Skills spec
  // (https://pi.dev/docs/latest/skills). A bare scalar string would parse as
  // one value and lose list semantics downstream.
  const tagsList = (entry.tags ?? []).map((t) => `  - ${t}`).join('\n')
  const toolsList = entry.tools.map((t) => `  - ${t}`).join('\n')
  return [
    '---',
    `name: ${entry.id}`,
    `description: "${yamlDoubleQuoted(entry.description.replace(/[\r\n]+/g, ' ')).slice(0, 1024)}"`,
    `version: ${entry.version}`,
    `author: ${entry.author}`,
    `category: ${entry.category}`,
    tagsList ? `tags:\n${tagsList}` : '',
    // Pre-approve these tools so the agent does not ask for permission to call
    // them — they are part of the host's own UI surface, not third-party actions.
    toolsList ? `allowed-tools:\n${toolsList}` : '',
    '---',
    '',
    `# ${entry.name}`,
    '',
    entry.description,
    '',
    '## Tools',
    '',
    `This skill registers the following tools: ${entry.tools.join(', ')}.`,
    'Each is implemented as a TypeScript pi extension in the GenOffice agent-skills',
    'package and is wired into the host pi session automatically.',
    '',
    '## Required permissions',
    '',
    scopes + '.',
    '',
    '## When to load',
    '',
    `Load when the user asks about ${entry.tags?.join(', ') ?? entry.category} tasks.`,
  ].join('\n')
}


/**
 * Make sure pi's settings point at the marketplace skills directory. Without
 * this the installed SKILL.md files sit on disk and no pi session ever sees
 * them — the exact gap that made "install" a UI-only illusion.
 */
export async function ensureSkillDirRegistered(): Promise<void> {
  const settings = openPiSettings()
  const paths = settings.getSkillPaths()
  if (paths.includes(PI_SKILLS_DIR)) return
  // Keep the user's own pi skill paths untouched; we only add ours.
  settings.setSkillPaths([...paths, PI_SKILLS_DIR])
  await settings.flush()
}

/** Path of an installed skill's directory (active or parked). */
export function skillDir(id: string, enabled = true): string {
  return join(enabled ? PI_SKILLS_DIR : PI_SKILLS_DISABLED_DIR, id)
}

export interface SkillToggleResult {
  ok: boolean
  /** true when a directory actually moved (i.e. the agent's view changed) */
  moved: boolean
  /** where the skill lives after the call */
  path?: string
  error?: string
}

/**
 * Enable/disable one installed skill for real: the SKILL.md directory moves
 * between the discovery root and the parked root. pi's loader only scans the
 * configured directory, so a parked skill is genuinely invisible to the agent.
 */
export async function setSkillEnabled(id: string, enabled: boolean): Promise<SkillToggleResult> {
  await ensureSkillDirRegistered()
  const active = skillDir(id, true)
  const parked = skillDir(id, false)
  try {
    if (!enabled && existsSync(active)) {
      mkdirSync(PI_SKILLS_DISABLED_DIR, { recursive: true })
      rmSync(parked, { recursive: true, force: true })
      renameSync(active, parked)
      return { ok: true, moved: true, path: parked }
    }
    if (enabled && existsSync(parked)) {
      mkdirSync(PI_SKILLS_DIR, { recursive: true })
      rmSync(active, { recursive: true, force: true })
      renameSync(parked, active)
      return { ok: true, moved: true, path: active }
    }
    return { ok: true, moved: false, path: enabled ? active : parked }
  } catch (err) {
    return { ok: false, moved: false, error: err instanceof Error ? err.message : String(err) }
  }
}

/** Remove an installed skill from both the active and parked roots. */
export function removeSkillFromPi(id: string): void {
  rmSync(skillDir(id, true), { recursive: true, force: true })
  rmSync(skillDir(id, false), { recursive: true, force: true })
}

/**
 * Validate a SKILL.md body with pi's own loader: stage it in a scratch
 * directory and check that pi parses it into a skill without diagnostics.
 *
 * Staging matters: pi skips dot-prefixed directories and refuses a skill with
 * no description, so running the publisher's file through the real parser is
 * the only way to promise "this is what the agent will see". The scratch root
 * lives outside the skills directory so a staging copy can never be picked up
 * by a concurrent discovery pass.
 */
export function validateSkillMarkdown(
  id: string,
  markdown: string,
): { ok: true; name: string; description: string } | { ok: false; error: string } {
  const stagedDir = join(PI_STAGING_DIR, `validate-${id}-${Date.now().toString(36)}`)
  try {
    rmSync(stagedDir, { recursive: true, force: true })
    mkdirSync(stagedDir, { recursive: true })
    writeFileSync(join(stagedDir, 'SKILL.md'), markdown, 'utf-8')
    const parsed = loadSkillsFromDir({ dir: stagedDir, source: 'genoffice-validation' })
    const match = parsed.skills.find((s) => resolve(s.filePath) === resolve(join(stagedDir, 'SKILL.md')))
    if (!match) {
      const detail = parsed.diagnostics.map((d) => d.message).join('; ')
      return {
        ok: false,
        error: detail || 'missing a frontmatter description (pi refuses to load a skill without one)',
      }
    }
    if (parsed.diagnostics.length > 0) {
      return { ok: false, error: parsed.diagnostics.map((d) => d.message).join('; ') }
    }
    return { ok: true, name: match.name, description: match.description }
  } finally {
    rmSync(stagedDir, { recursive: true, force: true })
  }
}

// ── plugin packages ─────────────────────────────────────────────────

export interface PluginArtifact {
  filename: string
  content: string
}

export interface PluginPackageInput {
  id: string
  name: string
  description: string
  version: string
  author?: string
  longDescription?: string
  tools?: string[]
  scopes?: string[]
  requirements?: string[]
  category?: string
  homepage?: string
  /** Uploaded pi extension module — the plugin's executable part. */
  artifact?: PluginArtifact
}

export interface PluginPackageInfo {
  ok: boolean
  packageDir?: string
  /** Absolute paths pi will load extensions from (empty for guidance-only). */
  extensions?: string[]
  /** Absolute paths of SKILL.md files the package contributes. */
  skills?: string[]
  /** true when the package ships executable extension code */
  hasCode?: boolean
  error?: string
}

export function pluginPackageDir(id: string): string {
  return join(PI_PLUGIN_DIR, id)
}

/** Only plain extension module names are allowed — no path traversal, no
 *  nested lookups. pi imports `.ts`/`.js` from the package's extensions dir. */
export function sanitizeExtensionFilename(name: string): string | null {
  const base = name.split(/[\\/]/).pop() ?? ''
  if (!/^[A-Za-z0-9._-]{1,64}$/.test(base)) return null
  if (!/\.(ts|mts|js|mjs)$/.test(base)) return null
  if (base.startsWith('.')) return null
  return base
}

function renderPluginSkill(input: PluginPackageInput): string {
  const tools = input.tools?.length ? input.tools.join(', ') : '(none)'
  const scopes = input.scopes?.length ? input.scopes.join(', ') : '(none)'
  const requirements = input.requirements?.length ? input.requirements.join(', ') : '(none)'
  const instructions = input.artifact
    ? [
        `This plugin ships a pi extension (${input.artifact.filename}) that registers its tools. Use them directly; do not invent tool names.`,
        `Tools: ${tools}.`,
        `Required setup: ${requirements}. Ask the user to configure anything missing instead of guessing credentials.`,
      ]
    : [
        `This plugin is installed as guidance only — it ships no executable extension, so its tools (${tools}) are NOT registered.`,
        'Treat the capability as advisory: explain the workflow, and say plainly that the executable part is not installed when the user asks for automation.',
        `Required setup: ${requirements}.`,
      ]
  return [
    '---',
    `name: ${input.id}`,
    `description: "${yamlDoubleQuoted(input.description.replace(/\s+/g, ' ')).slice(0, 280)}"`,
    `display_name: ${input.name.replace(/\s+/g, ' ')}`,
    '---',
    '',
    `# ${input.name}`,
    '',
    input.longDescription?.trim() || input.description,
    '',
    '## Instructions',
    '',
    ...instructions.map((line) => `- ${line}`),
    '',
    `Permission scopes declared by the plugin: ${scopes}.`,
  ].join('\n')
}

/**
 * Materialize a marketplace plugin as a pi package and register it in pi's
 * settings. Idempotent: re-installing overwrites the package contents and keeps
 * a single settings entry.
 */
export async function installPluginPackage(input: PluginPackageInput): Promise<PluginPackageInfo> {
  const dir = pluginPackageDir(input.id)
  try {
    rmSync(dir, { recursive: true, force: true })
    mkdirSync(dir, { recursive: true })

    const extensionName = input.artifact ? sanitizeExtensionFilename(input.artifact.filename) : null
    if (input.artifact && !extensionName) {
      return { ok: false, error: `Unsupported extension filename "${input.artifact.filename}"` }
    }

    // The `pi` manifest only lists directories that exist: pi reads globs
    // literally, and a manifest entry pointing at a missing dir is a load error.
    const manifest = {
      name: `@genoffice-plugin/${input.id}`,
      version: input.version,
      description: input.description,
      author: input.author || 'GenOffice Marketplace',
      keywords: ['pi-package', 'genoffice-plugin'],
      ...(input.homepage ? { homepage: input.homepage } : {}),
      pi: {
        ...(extensionName ? { extensions: ['./extensions'] } : {}),
        skills: ['./skills'],
      },
    }
    writeFileSync(join(dir, 'package.json'), JSON.stringify(manifest, null, 2), 'utf-8')

    const extensions: string[] = []
    if (extensionName && input.artifact) {
      const extDir = join(dir, 'extensions')
      mkdirSync(extDir, { recursive: true })
      const extPath = join(extDir, extensionName)
      writeFileSync(extPath, input.artifact.content, 'utf-8')
      extensions.push(extPath)
    }

    const skills: string[] = []
    const skillRoot = join(dir, 'skills', input.id)
    mkdirSync(skillRoot, { recursive: true })
    const skillPath = join(skillRoot, 'SKILL.md')
    writeFileSync(skillPath, renderPluginSkill(input), 'utf-8')
    skills.push(skillPath)

    writeFileSync(
      join(dir, 'README.md'),
      `# ${input.name}\n\n${input.longDescription?.trim() || input.description}\n\n` +
        `Installed by GenOffice from the plugin marketplace. Package layout:\n\n` +
        `- \`package.json\` — pi package manifest\n` +
        (extensionName ? `- \`extensions/${extensionName}\` — pi extension module\n` : '') +
        `- \`skills/${input.id}/SKILL.md\` — agent guidance\n`,
      'utf-8',
    )
    writeFileSync(
      join(dir, '.genoffice-plugin.json'),
      JSON.stringify(
        {
          id: input.id,
          version: input.version,
          author: input.author ?? null,
          category: input.category ?? null,
          tools: input.tools ?? [],
          scopes: input.scopes ?? [],
          requirements: input.requirements ?? [],
          hasCode: !!extensionName,
          artifact: extensionName ? input.artifact!.filename : null,
          installedAt: new Date().toISOString(),
        },
        null,
        2,
      ),
      'utf-8',
    )

    await registerPluginPackage(input.id, true)
    return { ok: true, packageDir: dir, extensions, skills, hasCode: !!extensionName }
  } catch (err) {
    return { ok: false, error: err instanceof Error ? err.message : String(err) }
  }
}

/** Read the manifest GenOffice wrote next to an installed plugin package. */
export function readPluginManifest(id: string): {
  hasCode: boolean
  artifact: string | null
  installedAt: string | null
} | null {
  const file = join(pluginPackageDir(id), '.genoffice-plugin.json')
  if (!existsSync(file)) return null
  try {
    const raw = JSON.parse(readFileSync(file, 'utf-8')) as {
      hasCode?: boolean
      artifact?: string | null
      installedAt?: string
    }
    return {
      hasCode: raw.hasCode === true,
      artifact: raw.artifact ?? null,
      installedAt: raw.installedAt ?? null,
    }
  } catch {
    return null
  }
}

/**
 * Register (or disable) a plugin package with pi. Disabled packages stay in
 * settings as `{ source, autoload: false }`, which pi resolves to "this package
 * contributes nothing" — the same mechanism `pi config` uses to disable
 * resources from installed packages.
 */
async function registerPluginPackage(id: string, enabled: boolean): Promise<void> {
  const dir = pluginPackageDir(id)
  const settings = openPiSettings()
  const kept = settings.getPackages().filter((entry) => resolve(packageSourceOf(entry)) !== resolve(dir))
  const entry: PackageSource = enabled ? dir : { source: dir, autoload: false }
  settings.setPackages([...kept, entry])
  await settings.flush()
}

export interface PluginToggleResult {
  ok: boolean
  enabled?: boolean
  error?: string
}

export async function setPluginPackageEnabled(id: string, enabled: boolean): Promise<PluginToggleResult> {
  if (!existsSync(pluginPackageDir(id))) {
    return { ok: false, error: `Plugin package "${id}" is not installed` }
  }
  try {
    await registerPluginPackage(id, enabled)
    return { ok: true, enabled }
  } catch (err) {
    return { ok: false, error: err instanceof Error ? err.message : String(err) }
  }
}

/**
 * Enable/disable any registered package source (local path or `npm:`/`git:`)
 * with the same `autoload` mechanism pi uses, so a disabled plugin stops
 * contributing extensions/skills on the next resource reload.
 */
export async function setPackageSourceEnabled(
  source: string,
  enabled: boolean,
): Promise<{ ok: boolean; error?: string }> {
  try {
    const settings = openPiSettings()
    const entries = settings.getPackages()
    const present = entries.some((entry) => packageSourceOf(entry) === source)
    if (!present && !enabled) return { ok: true }
    const kept = entries.filter((entry) => packageSourceOf(entry) !== source)
    const next: PackageSource = enabled ? source : { source, autoload: false }
    settings.setPackages([...kept, next])
    await settings.flush()
    return { ok: true }
  } catch (err) {
    return { ok: false, error: err instanceof Error ? err.message : String(err) }
  }
}

/** Unregister a plugin package and delete its files. */
export async function removePluginPackage(id: string): Promise<{ ok: boolean; error?: string }> {
  const dir = pluginPackageDir(id)
  try {
    const settings = openPiSettings()
    const kept = settings.getPackages().filter((entry) => resolve(packageSourceOf(entry)) !== resolve(dir))
    if (kept.length !== settings.getPackages().length) {
      settings.setPackages(kept)
      await settings.flush()
    }
    rmSync(dir, { recursive: true, force: true })
    return { ok: true }
  } catch (err) {
    return { ok: false, error: err instanceof Error ? err.message : String(err) }
  }
}

// ── remote pi packages (npm / git / absolute path) ───────────────────

export function isRemotePackageSource(source: string): boolean {
  return /^(npm:|git:|https?:)/.test(source)
}

/**
 * Install a real pi package (npm:/git:) through pi's package manager. The
 * package lands in pi's own cache under the agent dir and is persisted to
 * settings, so its extensions/skills/prompts/themes load exactly like any
 * other pi package.
 */
export async function installPiPackage(
  source: string,
): Promise<{ ok: boolean; source?: string; error?: string }> {
  const trimmed = source.trim()
  if (!/^(npm:|git:|https?:)/.test(trimmed)) {
    return { ok: false, error: 'Package source must start with npm:, git: or https://' }
  }
  try {
    const settings = openPiSettings()
    const manager = openPackageManager(settings)
    await manager.installAndPersist(trimmed, { local: false })
    return { ok: true, source: trimmed }
  } catch (err) {
    return { ok: false, error: err instanceof Error ? err.message : String(err) }
  }
}

export async function removePiPackage(source: string): Promise<{ ok: boolean; error?: string }> {
  try {
    const settings = openPiSettings()
    const manager = openPackageManager(settings)
    await manager.removeAndPersist(source, { local: false })
    return { ok: true }
  } catch (err) {
    return { ok: false, error: err instanceof Error ? err.message : String(err) }
  }
}

// ── resolution: what pi would actually load right now ───────────────

export interface ResolvedResource {
  path: string
  /** file name or skill name, for display */
  name: string
  /** false when the source is registered but disabled/filtered out */
  enabled: boolean
  /** true when the file lives under a GenOffice-managed root */
  managed: boolean
}

export interface PiResourceReport {
  cwd: string
  agentDir: string
  settingsPath: string
  /** pi's own resolution of registered packages + settings paths */
  extensions: ResolvedResource[]
  skills: ResolvedResource[]
  prompts: ResolvedResource[]
  themes: ResolvedResource[]
  /** resources that come from the user's global pi setup, not GenOffice */
  external: { extensions: number; skills: number; prompts: number; themes: number }
  packages: Array<{ source: string; scope: string; filtered: boolean; installedPath?: string }>
  diagnostics: Array<{ type: string; message: string; path?: string }>
}

function resourceName(path: string): string {
  const parts = path.split(sep)
  const base = parts[parts.length - 1] ?? path
  if (base === 'SKILL.md') return parts[parts.length - 2] ?? base
  return base
}

/**
 * Ask pi what it would load for our agent dir. This calls pi's
 * DefaultPackageManager (path-level resolution, no code execution) so the
 * report is the loader's own view rather than a GenOffice re-derivation.
 * Missing remote packages are skipped instead of triggering an install.
 */
export async function resolvePiResources(): Promise<PiResourceReport> {
  await ensureSkillDirRegistered()
  const settings = openPiSettings()
  const manager = openPackageManager(settings)
  const resolved = await manager.resolve(async () => 'skip')

  const map = (list: Array<{ path: string; enabled: boolean }>): ResolvedResource[] =>
    list.map((entry) => ({
      path: entry.path,
      name: resourceName(entry.path),
      enabled: entry.enabled,
      managed: isManagedPath(entry.path),
    }))

  const extensions = map(resolved.extensions)
  const skills = map(resolved.skills)
  const prompts = map(resolved.prompts)
  const themes = map(resolved.themes)

  const external = {
    extensions: extensions.filter((e) => !e.managed).length,
    skills: skills.filter((e) => !e.managed).length,
    prompts: prompts.filter((e) => !e.managed).length,
    themes: themes.filter((e) => !e.managed).length,
  }

  return {
    cwd: PI_CWD,
    agentDir: PI_AGENT_DIR,
    settingsPath: piSettingsPath(),
    extensions,
    skills,
    prompts,
    themes,
    external,
    packages: manager.listConfiguredPackages(),
    diagnostics: settings.drainErrors().map((e) => ({
      type: 'settings',
      message: e.error.message,
      ...(e.path ? { path: e.path } : {}),
    })),
  }
}

/** Installed skill files with their parsed frontmatter, straight from pi's
 *  loader. `enabled` mirrors whether the file sits in the discovery root. */
export function installedSkills(): Array<Skill & { enabled: boolean; managed: boolean }> {
  const parse = (dir: string, enabled: boolean) => {
    if (!existsSync(dir)) return []
    const parsed = loadSkillsFromDir({ dir, source: 'genoffice-marketplace' })
    return parsed.skills.map((skill) => ({ ...skill, enabled, managed: isManagedPath(skill.filePath) }))
  }
  return [...parse(PI_SKILLS_DIR, true), ...parse(PI_SKILLS_DISABLED_DIR, false)]
}

/** Size guard for uploaded artifacts (extension modules and SKILL.md bodies). */
export function artifactBytes(content: string): number {
  return Buffer.byteLength(content, 'utf-8')
}

export function readableBytes(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`
  return `${(bytes / (1024 * 1024)).toFixed(2)} MB`
}

/** True when the path exists and is a file (used by upload validation). */
export function isFile(path: string): boolean {
  try {
    return statSync(path).isFile()
  } catch {
    return false
  }
}
