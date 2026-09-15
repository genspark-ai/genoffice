/**
 * Skills & Plugins management channels
 *
 * GenOffice 的核心 Agent 完全以 pi 为核心(W25 验证)。agent-skills 包内置 11
 * 个 skill extension(7 个 office skill + 4 个 plugin)。settings → "Skills &
 * Plugins" 让用户:
 *
 * - 查看所有内置 skills 和 plugins(来自 @genoffice/agent-skills)
 * - 启用/禁用每个 skill(enabled=false 时 ExtensionRunner 不加载)
 * - 从 marketplace 安装/卸载 skill(模拟或真实 fs 路径)
 * - 查看每个 skill 的元信息(版本/作者/描述/工具数)
 * - 重新加载 skill(从 pi ExtensionRunner 重启)
 *
 * 持久化到 DATA_DIR/skills.json 和 plugins.json。
 *
 * 设计要点:
 * - skills 是 GenOffice 自己实现的 agent skill(对应 plan §2.1 的 extension)
 * - plugins 是更高级的扩展(多 Agent 团队/审计/本地模型,plan §2.1 的 agent-team)
 * - 真实生产路径是从 `@genoffice/agent-skills/src/extensions/*.ts` 读取
 * - 这里用元数据 catalog 来驱动 UI,实际启用通过 toggleSkill 触发重新加载
 */
import { existsSync, readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { DATA_DIR, registerHandle } from '../common/index.js'

export type SkillKind =
  | 'docs-skill'
  | 'sheets-skill'
  | 'slides-skill'
  | 'office-workflow'
  | 'office-safety'
  | 'frozen-selection'
  | 'verify-response'
  | 'skill-market'

export type PluginKind = 'agent-team' | 'audit-log' | 'local-models'

export type SkillStatus = 'enabled' | 'disabled' | 'error'

export interface SkillEntry {
  id: SkillKind
  /** Display name shown in the UI */
  name: string
  /** Short description */
  description: string
  /** Author / package maintainer */
  author: string
  /** Semantic version */
  version: string
  /** Source package */
  package: string
  /** Relative source path under the package */
  source: string
  /** Tool names contributed by this skill */
  tools: string[]
  /** Permission scopes the skill declares (e.g. "files:write", "ai:stream") */
  scopes: string[]
  /** Current effective state */
  status: SkillStatus
  /** ISO timestamp when last loaded by ExtensionRunner; null if never */
  lastLoadedAt: string | null
  /** Optional error message when status === 'error' */
  error?: string
  /** Whether this is a GenOffice built-in skill (cannot be uninstalled) */
  builtIn: boolean
}

export interface PluginEntry {
  id: PluginKind
  name: string
  description: string
  author: string
  version: string
  package: string
  source: string
  /** Tools contributed */
  tools: string[]
  /** Permission scopes the plugin declares (e.g. "agents:multi", "audit:write") */
  scopes: string[]
  /** Resource requirements (e.g. "ollama-runtime") */
  requirements: string[]
  status: SkillStatus
  lastLoadedAt: string | null
  error?: string
  builtIn: boolean
}

// ── BUILT-IN CATALOG ────────────────────────────────────────
// 元数据与 packages/agent-skills/src/extensions/*.ts 实际对应。
// 这些 skill 来自 plan §2.1 设计的 7 个 office skill + 4 个 plugin。
const DEFAULT_SKILLS: SkillEntry[] = [
  {
    id: 'docs-skill',
    name: 'Docs Skill',
    description: 'Word 文档操作工具集(读块/写块/替换/格式/搜索/...)',
    author: 'GenOffice',
    version: '0.85.1',
    package: '@genoffice/agent-skills',
    source: 'src/extensions/docs-skill.ts',
    tools: [
      'read_blocks',
      'write_block',
      'replace_document',
      'insert_blocks',
      'delete_blocks',
      'format_blocks',
      'search_blocks',
      'find_replace',
      'set_page_margins',
      'headings_outline',
    ],
    scopes: ['files:read', 'files:write', 'docs:edit'],
    status: 'enabled',
    lastLoadedAt: null,
    builtIn: true,
  },
  {
    id: 'sheets-skill',
    name: 'Sheets Skill',
    description: 'Excel 表格操作工具集(读 cell/写 cell/公式/图表/筛选/排序)',
    author: 'GenOffice',
    version: '0.85.1',
    package: '@genoffice/agent-skills',
    source: 'src/extensions/sheets-skill.ts',
    tools: ['read_range', 'write_range', 'apply_formula', 'create_chart', 'sort_range', 'filter_range'],
    scopes: ['files:read', 'files:write', 'sheets:edit'],
    status: 'enabled',
    lastLoadedAt: null,
    builtIn: true,
  },
  {
    id: 'slides-skill',
    name: 'Slides Skill',
    description: 'PPT 幻灯片操作工具集(读 slide/写 slide/插入图/布局/演讲备注)',
    author: 'GenOffice',
    version: '0.85.1',
    package: '@genoffice/agent-skills',
    source: 'src/extensions/slides-skill.ts',
    tools: ['read_slides', 'write_slide', 'insert_image', 'set_layout', 'set_speaker_notes'],
    scopes: ['files:read', 'files:write', 'slides:edit'],
    status: 'enabled',
    lastLoadedAt: null,
    builtIn: true,
  },
  {
    id: 'office-workflow',
    name: 'Office Workflow',
    description: '跨 Office 工作流编排(docs→sheets→slides 数据流)',
    author: 'GenOffice',
    version: '0.85.1',
    package: '@genoffice/agent-skills',
    source: 'src/extensions/office-workflow.ts',
    tools: ['cross_office_workflow', 'office_data_pipeline'],
    scopes: ['files:read', 'files:write'],
    status: 'enabled',
    lastLoadedAt: null,
    builtIn: true,
  },
  {
    id: 'office-safety',
    name: 'Office Safety',
    description: 'Office 操作安全检查(用户确认 + 范围限制 + 撤销支持)',
    author: 'GenOffice',
    version: '0.85.1',
    package: '@genoffice/agent-skills',
    source: 'src/extensions/office-safety.ts',
    tools: ['confirm_destructive_op', 'scope_check'],
    scopes: ['safety'],
    status: 'enabled',
    lastLoadedAt: null,
    builtIn: true,
  },
  {
    id: 'frozen-selection',
    name: 'Frozen Selection',
    description: '冻结 AI 选区,防止后续修改误伤用户意图',
    author: 'GenOffice',
    version: '0.85.1',
    package: '@genoffice/agent-skills',
    source: 'src/extensions/frozen-selection.ts',
    tools: ['freeze_selection', 'unfreeze_selection', 'list_frozen'],
    scopes: ['docs:edit'],
    status: 'enabled',
    lastLoadedAt: null,
    builtIn: true,
  },
  {
    id: 'verify-response',
    name: 'Verify Response',
    description: '验证 AI 响应与原文档的一致性(diff + 摘要回归)',
    author: 'GenOffice',
    version: '0.85.1',
    package: '@genoffice/agent-skills',
    source: 'src/extensions/verify-response.ts',
    tools: ['verify_response', 'summarize_diff'],
    scopes: ['ai:stream'],
    status: 'enabled',
    lastLoadedAt: null,
    builtIn: true,
  },
  {
    id: 'skill-market',
    name: 'Skill Marketplace',
    description: 'Skills 市场(浏览/搜索/安装/卸载第三方 skill)',
    author: 'GenOffice',
    version: '0.85.1',
    package: '@genoffice/agent-skills',
    source: 'src/extensions/skill-market.ts',
    tools: ['list_marketplace', 'search_skills', 'install_skill', 'uninstall_skill'],
    scopes: ['network:out'],
    status: 'enabled',
    lastLoadedAt: null,
    builtIn: true,
  },
]

const DEFAULT_PLUGINS: PluginEntry[] = [
  {
    id: 'agent-team',
    name: 'Agent Team',
    description: '多 Agent 团队协作(5 个内置角色 + request_review 工具)',
    author: 'GenOffice',
    version: '0.85.1',
    package: '@genoffice/agent-skills',
    source: 'src/extensions/agent-team.ts',
    tools: ['request_review', 'run_agent', 'coordinate_team'],
    scopes: ['ai:stream', 'agents:multi'],
    requirements: [],
    status: 'enabled',
    lastLoadedAt: null,
    builtIn: true,
  },
  {
    id: 'audit-log',
    name: 'Audit Log',
    description: '企业级审计日志(3 个 sink: file/otlp/console + 自动 tool_call/result 配对)',
    author: 'GenOffice',
    version: '0.85.1',
    package: '@genoffice/agent-skills',
    source: 'src/extensions/audit-log.ts',
    tools: ['export_audit', 'tail_audit'],
    scopes: ['audit:write'],
    requirements: [],
    status: 'enabled',
    lastLoadedAt: null,
    builtIn: true,
  },
  {
    id: 'local-models',
    name: 'Local Models (Ollama)',
    description: 'Ollama 本地模型 provider(createOllamaProvider + installLocalModels)',
    author: 'GenOffice',
    version: '0.85.1',
    package: '@genoffice/agent-skills',
    source: 'src/extensions/local-models.ts',
    tools: ['list_ollama_models', 'install_ollama_model', 'pull_ollama_model'],
    scopes: ['network:out', 'providers:add'],
    requirements: ['ollama-runtime'],
    status: 'disabled',
    lastLoadedAt: null,
    builtIn: true,
  },
]

const SKILLS_FILE = join(DATA_DIR, 'skills.json')
const PLUGINS_FILE = join(DATA_DIR, 'plugins.json')

let skillsCache: SkillEntry[] | null = null
let pluginsCache: PluginEntry[] | null = null

function loadSkills(): SkillEntry[] {
  if (skillsCache) return skillsCache
  try {
    if (existsSync(SKILLS_FILE)) {
      const parsed = JSON.parse(readFileSync(SKILLS_FILE, 'utf-8')) as Partial<SkillEntry>[]
      if (Array.isArray(parsed) && parsed.length > 0) {
        skillsCache = DEFAULT_SKILLS.map((def) => {
          const override = parsed.find((p) => p && p.id === def.id)
          return {
            ...def,
            status: override?.status === 'enabled' || override?.status === 'disabled' || override?.status === 'error'
              ? override.status
              : def.status,
            lastLoadedAt: typeof override?.lastLoadedAt === 'string' ? override.lastLoadedAt : def.lastLoadedAt,
          }
        })
        return skillsCache
      }
    }
  } catch {}
  skillsCache = DEFAULT_SKILLS.map((s) => ({ ...s }))
  return skillsCache
}

function saveSkills(skills: SkillEntry[]): void {
  skillsCache = skills
  try {
    writeFileSync(
      SKILLS_FILE,
      JSON.stringify(
        skills.map((s) => ({ id: s.id, status: s.status, lastLoadedAt: s.lastLoadedAt })),
        null,
        2,
      ),
    )
  } catch {}
}

function loadPlugins(): PluginEntry[] {
  if (pluginsCache) return pluginsCache
  try {
    if (existsSync(PLUGINS_FILE)) {
      const parsed = JSON.parse(readFileSync(PLUGINS_FILE, 'utf-8')) as Partial<PluginEntry>[]
      if (Array.isArray(parsed) && parsed.length > 0) {
        pluginsCache = DEFAULT_PLUGINS.map((def) => {
          const override = parsed.find((p) => p && p.id === def.id)
          return {
            ...def,
            status:
              override?.status === 'enabled' || override?.status === 'disabled' || override?.status === 'error'
                ? override.status
                : def.status,
            lastLoadedAt: typeof override?.lastLoadedAt === 'string' ? override.lastLoadedAt : def.lastLoadedAt,
          }
        })
        return pluginsCache
      }
    }
  } catch {}
  pluginsCache = DEFAULT_PLUGINS.map((p) => ({ ...p }))
  return pluginsCache
}

function savePlugins(plugins: PluginEntry[]): void {
  pluginsCache = plugins
  try {
    writeFileSync(
      PLUGINS_FILE,
      JSON.stringify(
        plugins.map((p) => ({ id: p.id, status: p.status, lastLoadedAt: p.lastLoadedAt })),
        null,
        2,
      ),
    )
  } catch {}
}

export function registerSkillHandlers(): void {
  // ── SKILLS ──
  registerHandle('home:list-skills', () => {
    return { skills: loadSkills() }
  })

  registerHandle('home:toggle-skill', (_event: unknown, args: unknown) => {
    const { id, enabled } = (args || {}) as { id: SkillKind; enabled: boolean }
    const skills = loadSkills().map((s) =>
      s.id === id
        ? {
            ...s,
            status: (enabled ? 'enabled' : 'disabled') as SkillStatus,
            lastLoadedAt: enabled ? new Date().toISOString() : s.lastLoadedAt,
            error: undefined,
          }
        : s,
    )
    saveSkills(skills)
    return { ok: true, skills }
  })

  registerHandle('home:reload-skill', (_event: unknown, args: unknown) => {
    const { id } = (args || {}) as { id: SkillKind }
    const skills = loadSkills().map((s) =>
      s.id === id
        ? { ...s, status: 'enabled' as SkillStatus, lastLoadedAt: new Date().toISOString(), error: undefined }
        : s,
    )
    saveSkills(skills)
    return { ok: true, skills }
  })

  registerHandle('home:install-skill', (_event: unknown, args: unknown) => {
    const { name } = (args || {}) as { name: string }
    // 模拟从 marketplace 安装
    const skills = loadSkills()
    const exists = skills.some((s) => s.id === (name as SkillKind))
    if (exists) {
      return { ok: false, error: `Skill "${name}" already installed` }
    }
    // 真实生产路径:从 marketplace fetch + dynamic import + ExtensionRunner.register
    // 这里返回 install success,UI 可继续触发 reload
    return {
      ok: true,
      installed: { id: name, name, marketplace: true },
    }
  })

  registerHandle('home:uninstall-skill', (_event: unknown, args: unknown) => {
    const { id } = (args || {}) as { id: SkillKind }
    const skill = loadSkills().find((s) => s.id === id)
    if (!skill) return { ok: false, error: `Skill "${id}" not found` }
    if (skill.builtIn) return { ok: false, error: `Cannot uninstall built-in skill "${id}"` }
    const skills = loadSkills().filter((s) => s.id !== id)
    saveSkills(skills)
    return { ok: true, skills }
  })

  registerHandle('home:reset-skills', () => {
    skillsCache = null
    saveSkills(DEFAULT_SKILLS.map((s) => ({ ...s })))
    return { ok: true, skills: loadSkills() }
  })

  // ── PLUGINS ──
  registerHandle('home:list-plugins', () => {
    return { plugins: loadPlugins() }
  })

  registerHandle('home:toggle-plugin', (_event: unknown, args: unknown) => {
    const { id, enabled } = (args || {}) as { id: PluginKind; enabled: boolean }
    const plugins = loadPlugins().map((p) =>
      p.id === id
        ? {
            ...p,
            status: (enabled ? 'enabled' : 'disabled') as SkillStatus,
            lastLoadedAt: enabled ? new Date().toISOString() : p.lastLoadedAt,
            error: undefined,
          }
        : p,
    )
    savePlugins(plugins)
    return { ok: true, plugins }
  })

  registerHandle('home:reload-plugin', (_event: unknown, args: unknown) => {
    const { id } = (args || {}) as { id: PluginKind }
    const plugins = loadPlugins().map((p) =>
      p.id === id
        ? { ...p, status: 'enabled' as SkillStatus, lastLoadedAt: new Date().toISOString(), error: undefined }
        : p,
    )
    savePlugins(plugins)
    return { ok: true, plugins }
  })

  registerHandle('home:reset-plugins', () => {
    pluginsCache = null
    savePlugins(DEFAULT_PLUGINS.map((p) => ({ ...p })))
    return { ok: true, plugins: loadPlugins() }
  })

  // ── 复合接口: 一次返回 skills + plugins ──
  registerHandle('home:get-skills-and-plugins', () => {
    return { skills: loadSkills(), plugins: loadPlugins() }
  })
}
