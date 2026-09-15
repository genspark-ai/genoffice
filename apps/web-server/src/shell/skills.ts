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
  // marketplace skills (3rd-party)
  | 'notion-sync'
  | 'pdf-ocr-pro'
  | 'github-integration'
  | 'jira-bridge'
  | 'lang-detector'
  | (string & {}) // allow other marketplace ids without cast at every use site

export type PluginKind =
  | 'agent-team'
  | 'audit-log'
  | 'local-models'
  // marketplace plugins (3rd-party)
  | 'slack-bridge'
  | 'gdrive-export'
  | (string & {}) // allow other marketplace ids without cast at every use site

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
        const builtIns = DEFAULT_SKILLS.map((def) => {
          const override = parsed.find((p) => p && p.id === def.id)
          return {
            ...def,
            status: override?.status === 'enabled' || override?.status === 'disabled' || override?.status === 'error'
              ? override.status
              : def.status,
            lastLoadedAt: typeof override?.lastLoadedAt === 'string' ? override.lastLoadedAt : def.lastLoadedAt,
          }
        })
        // Keep marketplace-installed skills (present in skills.json but not in
        // DEFAULT_SKILLS) — without this they vanish on the first reload.
        const builtInIds = new Set(DEFAULT_SKILLS.map((d) => d.id as string))
        const installed = parsed.filter(
          (p): p is SkillEntry =>
            !!p && typeof p.id === 'string' && !builtInIds.has(p.id) && typeof p.name === 'string',
        )
        skillsCache = [...builtIns, ...installed.map((p) => ({ ...p, builtIn: false }))]
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
        const builtIns = DEFAULT_PLUGINS.map((def) => {
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
        // Keep marketplace-installed plugins (present in plugins.json but not
        // in DEFAULT_PLUGINS) — without this they vanish on the first reload.
        const builtInIds = new Set(DEFAULT_PLUGINS.map((d) => d.id as string))
        const installed = parsed.filter(
          (p): p is PluginEntry =>
            !!p && typeof p.id === 'string' && !builtInIds.has(p.id) && typeof p.name === 'string',
        )
        pluginsCache = [...builtIns, ...installed.map((p) => ({ ...p, builtIn: false }))]
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


// ── MARKETPLACE CATALOG ────────────────────────────────────
// 模拟的市场数据。真实生产会 fetch marketplace API。
// 这些 skill/plugin 是第三方扩展,用户从市场浏览并安装到自己的环境。
const MARKETPLACE_SKILLS: Array<Omit<SkillEntry, 'status' | 'lastLoadedAt' | 'builtIn'>> = [
  {
    id: 'notion-sync',
    name: 'Notion Sync',
    description: '双向同步 Notion workspace 与本地文档,自动转换格式',
    author: 'Community',
    version: '1.4.2',
    package: '@marketplace/notion-sync',
    source: 'src/notion-sync.ts',
    tools: ['sync_workspace', 'import_page', 'export_doc', 'resolve_links'],
    scopes: ['network:out', 'files:read', 'files:write'],
  },
  {
    id: 'pdf-ocr-pro',
    name: 'PDF OCR Pro',
    description: '高级 OCR(中英文混合 + 表格识别 + 手写体)',
    author: 'OCR Labs',
    version: '2.0.1',
    package: '@marketplace/pdf-ocr-pro',
    source: 'src/pdf-ocr-pro.ts',
    tools: ['ocr_page', 'ocr_table', 'ocr_handwriting', 'ocr_batch'],
    scopes: ['files:read', 'pdf:read'],
  },
  {
    id: 'github-integration',
    name: 'GitHub Integration',
    description: '从 GitHub 仓库读取 issue/PR/commit 并生成 Office 文档',
    author: 'DevTools Collective',
    version: '0.9.0',
    package: '@marketplace/github-integration',
    source: 'src/github.ts',
    tools: ['fetch_issue', 'fetch_pr', 'fetch_commit', 'export_to_doc'],
    scopes: ['network:out', 'docs:edit'],
  },
  {
    id: 'jira-bridge',
    name: 'Jira Bridge',
    description: '把 Jira ticket 转换为 Office 任务清单和进度报告',
    author: 'Atlassian Tools',
    version: '1.2.0',
    package: '@marketplace/jira-bridge',
    source: 'src/jira.ts',
    tools: ['fetch_tickets', 'sync_sprint', 'export_progress'],
    scopes: ['network:out', 'sheets:edit'],
  },
  {
    id: 'lang-detector',
    name: 'Language Detector',
    description: '检测 Office 文档语言并自动应用翻译预设',
    author: 'i18n Group',
    version: '3.1.0',
    package: '@marketplace/lang-detector',
    source: 'src/lang-detector.ts',
    tools: ['detect_language', 'apply_translate_preset'],
    scopes: ['ai:stream'],
  },
]

const MARKETPLACE_PLUGINS: Array<Omit<PluginEntry, 'status' | 'lastLoadedAt' | 'builtIn'>> = [
  {
    id: 'slack-bridge',
    name: 'Slack Bridge',
    description: '把 Office AI 操作日志转发到 Slack 频道',
    author: 'Workspace Integrations',
    version: '1.0.0',
    package: '@marketplace/slack-bridge',
    source: 'src/slack.ts',
    tools: ['send_notification', 'request_approval', 'sync_channel'],
    scopes: ['network:out'],
    requirements: ['slack-workspace-token'],
  },
  {
    id: 'gdrive-export',
    name: 'Google Drive Export',
    description: '导出 Office 文档到 Google Drive(双向同步)',
    author: 'Cloud Sync',
    version: '2.3.1',
    package: '@marketplace/gdrive-export',
    source: 'src/gdrive.ts',
    tools: ['upload_doc', 'sync_folder', 'resolve_permissions'],
    scopes: ['network:out', 'files:write'],
    requirements: ['google-oauth-client'],
  },
]

function listMarketplaceSkills() {
  const installed = loadSkills()
  return MARKETPLACE_SKILLS.map((entry) => ({
    ...entry,
    installed: installed.some((s) => s.id === entry.id),
  }))
}

function listMarketplacePlugins() {
  const installed = loadPlugins()
  return MARKETPLACE_PLUGINS.map((entry) => ({
    ...entry,
    installed: installed.some((p) => p.id === entry.id),
  }))
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

  // ── MARKETPLACE ──
  registerHandle('home:list-marketplace-skills', () => {
    return { skills: listMarketplaceSkills() }
  })

  registerHandle('home:list-marketplace-plugins', () => {
    return { plugins: listMarketplacePlugins() }
  })

  registerHandle('home:install-skill', (_event: unknown, args: unknown) => {
    // 真实安装:从 marketplace 取元数据,合并到 installed skills
    // 兼容两种参数风格:前端 IPC 客户端 installSkill(name) 用 {name},marketplace UI 用 {id}
    const { id: idArg, name } = (args || {}) as { id?: string; name?: string }
    const id = idArg || name
    if (!id) return { ok: false, error: 'Missing skill id/name' }
    const entry = MARKETPLACE_SKILLS.find((s) => s.id === id)
    if (!entry) {
      // 也支持已存在的内置/已安装 skill 的 "重新激活" 触发
      const existing = loadSkills().find((s) => s.id === id)
      if (!existing) return { ok: false, error: `Skill "${id}" not found in marketplace` }
      return { ok: true, installed: existing, alreadyInstalled: true }
    }
    const skills = loadSkills()
    if (skills.some((s) => s.id === id)) {
      const existing = skills.find((s) => s.id === id)!
      return { ok: true, installed: existing, alreadyInstalled: true, skills }
    }
    const installedEntry: SkillEntry = {
      ...entry,
      status: 'enabled',
      lastLoadedAt: new Date().toISOString(),
      builtIn: false,
    }
    const next = [...skills, installedEntry]
    saveSkills(next)
    return { ok: true, installed: installedEntry, skills: next }
  })

  registerHandle('home:install-plugin', (_event: unknown, args: unknown) => {
    const { id } = (args || {}) as { id: string }
    const entry = MARKETPLACE_PLUGINS.find((p) => p.id === id)
    if (!entry) {
      const existing = loadPlugins().find((p) => p.id === id)
      if (!existing) return { ok: false, error: `Plugin "${id}" not found in marketplace` }
      return { ok: true, installed: existing, alreadyInstalled: true }
    }
    const plugins = loadPlugins()
    if (plugins.some((p) => p.id === id)) {
      const existing = plugins.find((p) => p.id === id)!
      return { ok: true, installed: existing, alreadyInstalled: true, plugins }
    }
    const installedEntry: PluginEntry = {
      ...entry,
      status: 'enabled',
      lastLoadedAt: new Date().toISOString(),
      builtIn: false,
    }
    const next = [...plugins, installedEntry]
    savePlugins(next)
    return { ok: true, installed: installedEntry, plugins: next }
  })

  registerHandle('home:uninstall-plugin', (_event: unknown, args: unknown) => {
    const { id } = (args || {}) as { id: PluginKind }
    const plugin = loadPlugins().find((p) => p.id === id)
    if (!plugin) return { ok: false, error: `Plugin "${id}" not found` }
    if (plugin.builtIn) return { ok: false, error: `Cannot uninstall built-in plugin "${id}"` }
    const plugins = loadPlugins().filter((p) => p.id !== id)
    savePlugins(plugins)
    return { ok: true, plugins }
  })

  registerHandle('home:get-marketplace-and-installed', () => {
    return {
      marketplaceSkills: listMarketplaceSkills(),
      marketplacePlugins: listMarketplacePlugins(),
      installedSkills: loadSkills(),
      installedPlugins: loadPlugins(),
    }
  })

  // ── 复合接口: 一次返回 skills + plugins ──

  registerHandle('home:get-skills-and-plugins', () => {
    return { skills: loadSkills(), plugins: loadPlugins() }
  })
}
