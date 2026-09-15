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
import { existsSync, mkdirSync, readdirSync, readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { DATA_DIR, registerHandle } from '../common/index'

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
// ── MARKETPLACE CATALOG (扩展 schema) ──────────────────────
// 每个 marketplace 扩展携带 category / tags / rating / downloads / featured 等
// 元数据,供 UI 搜索 / 分类 / 排序使用。运行时,通过 ExtensionAPI 注册到 pi
// AgentRunner(详见 agent-runtime + extensions/)。

export type MarketplaceCategory =
  | 'productivity'
  | 'data'
  | 'dev'
  | 'media'
  | 'translation'
  | 'collaboration'
  | 'finance'
  | 'design'

export interface MarketplaceSkillEntry {
  id: string
  name: string
  description: string
  longDescription?: string
  author: string
  version: string
  package: string
  source: string
  tools: string[]
  scopes: string[]
  category: MarketplaceCategory
  tags: string[]
  /** Stars from 0 to 5 (one decimal) */
  rating: number
  /** Total download count */
  downloads: number
  /** Featured in the top of the marketplace */
  featured?: boolean
  /** Optional upstream icon (single emoji) */
  icon?: string
  homepage?: string
}

export interface MarketplacePluginEntry {
  id: string
  name: string
  description: string
  longDescription?: string
  author: string
  version: string
  package: string
  source: string
  tools: string[]
  scopes: string[]
  requirements: string[]
  category: MarketplaceCategory
  tags: string[]
  rating: number
  downloads: number
  featured?: boolean
  icon?: string
  homepage?: string
}

const MARKETPLACE_SKILLS: MarketplaceSkillEntry[] = [
  {
    id: 'notion-sync',
    name: 'Notion Sync',
    description: '双向同步 Notion workspace 与本地文档,自动转换格式',
    longDescription: '把 Notion 页面、database、block 实时映射到 GenOffice doc 树,保留富文本、嵌套列表与引用。',
    author: 'Community',
    version: '1.4.2',
    package: '@marketplace/notion-sync',
    source: 'src/notion-sync.ts',
    tools: ['sync_workspace', 'import_page', 'export_doc', 'resolve_links'],
    scopes: ['network:out', 'files:read', 'files:write'],
    category: 'productivity',
    tags: ['notion', 'sync', 'workspace', 'wiki'],
    rating: 4.6,
    downloads: 18420,
    featured: true,
    icon: 'N',
  },
  {
    id: 'pdf-ocr-pro',
    name: 'PDF OCR Pro',
    description: '高级 OCR(中英文混合 + 表格识别 + 手写体)',
    longDescription: '基于多模型融合的 OCR pipeline:中英双语、表格结构还原、手写体识别、批量离线处理。',
    author: 'OCR Labs',
    version: '2.0.1',
    package: '@marketplace/pdf-ocr-pro',
    source: 'src/pdf-ocr-pro.ts',
    tools: ['ocr_page', 'ocr_table', 'ocr_handwriting', 'ocr_batch'],
    scopes: ['files:read', 'pdf:read'],
    category: 'media',
    tags: ['ocr', 'pdf', 'scanning', 'handwriting'],
    rating: 4.8,
    downloads: 25104,
    featured: true,
    icon: 'O',
  },
  {
    id: 'github-integration',
    name: 'GitHub Integration',
    description: '从 GitHub 仓库读取 issue/PR/commit 并生成 Office 文档',
    longDescription: '把仓库活动一键导出为周报、变更日志、Roadmap 表格;支持 Markdown / sheet / slide 多种渲染。',
    author: 'DevTools Collective',
    version: '0.9.0',
    package: '@marketplace/github-integration',
    source: 'src/github.ts',
    tools: ['fetch_issue', 'fetch_pr', 'fetch_commit', 'export_to_doc'],
    scopes: ['network:out', 'docs:edit'],
    category: 'dev',
    tags: ['github', 'git', 'issues', 'release-notes'],
    rating: 4.4,
    downloads: 9821,
    icon: 'G',
  },
  {
    id: 'jira-bridge',
    name: 'Jira Bridge',
    description: '把 Jira ticket 转换为 Office 任务清单和进度报告',
    longDescription: 'JQL 查询结果自动生成甘特图、燃尽图、Owner 看板;支持自定义字段映射与导出为 docs/sheets。',
    author: 'Atlassian Tools',
    version: '1.2.0',
    package: '@marketplace/jira-bridge',
    source: 'src/jira.ts',
    tools: ['fetch_tickets', 'sync_sprint', 'export_progress'],
    scopes: ['network:out', 'sheets:edit'],
    category: 'dev',
    tags: ['jira', 'agile', 'project-management'],
    rating: 4.2,
    downloads: 5340,
    icon: 'J',
  },
  {
    id: 'lang-detector',
    name: 'Language Detector',
    description: '检测 Office 文档语言并自动应用翻译预设',
    longDescription: '识别 80+ 语言,按文档使用频率推荐翻译策略、术语表、字体回退方案。',
    author: 'i18n Group',
    version: '3.1.0',
    package: '@marketplace/lang-detector',
    source: 'src/lang-detector.ts',
    tools: ['detect_language', 'apply_translate_preset'],
    scopes: ['ai:stream'],
    category: 'translation',
    tags: ['i18n', 'language', 'detection'],
    rating: 4.5,
    downloads: 11028,
    icon: 'L',
  },
  {
    id: 'figma-export',
    name: 'Figma Export',
    description: '把 Figma frame 一键导出为 slides 与插图',
    longDescription: '通过 Figma API 拉取 frame,转换为 PPT 占位符与嵌入 PNG/SVG;保留图层命名便于回链。',
    author: 'Design Bridge',
    version: '1.0.4',
    package: '@marketplace/figma-export',
    source: 'src/figma.ts',
    tools: ['fetch_frame', 'render_svg', 'insert_to_slides'],
    scopes: ['network:out', 'slides:edit'],
    category: 'design',
    tags: ['figma', 'design', 'slides', 'assets'],
    rating: 4.3,
    downloads: 3812,
    icon: 'F',
  },
  {
    id: 'youtube-transcript',
    name: 'YouTube Transcript',
    description: '把 YouTube 视频转写为可编辑文档与摘要',
    longDescription: '拉取字幕/ASR,按章节切分,生成中英对照稿、关键观点摘要、提问清单。',
    author: 'Media AI',
    version: '0.7.2',
    package: '@marketplace/youtube-transcript',
    source: 'src/youtube.ts',
    tools: ['fetch_transcript', 'summarize_video', 'extract_chapters'],
    scopes: ['network:out', 'ai:stream'],
    category: 'media',
    tags: ['youtube', 'video', 'transcript', 'summary'],
    rating: 4.1,
    downloads: 6720,
    icon: 'Y',
  },
  {
    id: 'csv-data-viz',
    name: 'CSV Data Viz',
    description: 'CSV → 自动选型图表并嵌入 docs/sheets',
    longDescription: '基于列类型分布自动推荐柱/线/散点/饼图,支持自定义配色与导出 PNG。',
    author: 'Data Suite',
    version: '2.1.0',
    package: '@marketplace/csv-data-viz',
    source: 'src/csv-viz.ts',
    tools: ['auto_chart', 'recommend_palette', 'export_png'],
    scopes: ['files:read', 'charts:write'],
    category: 'data',
    tags: ['csv', 'chart', 'visualization', 'analytics'],
    rating: 4.7,
    downloads: 14210,
    featured: true,
    icon: 'C',
  },
  {
    id: 'linear-sync',
    name: 'Linear Sync',
    description: 'Linear 项目同步到 Office 任务清单',
    longDescription: 'Linear issues / projects / cycles 同步为 docs / sheets / slides;支持双向状态映射。',
    author: 'PM Tools',
    version: '0.5.1',
    package: '@marketplace/linear-sync',
    source: 'src/linear.ts',
    tools: ['fetch_issues', 'sync_status', 'export_roadmap'],
    scopes: ['network:out', 'sheets:edit'],
    category: 'dev',
    tags: ['linear', 'project-management', 'sync'],
    rating: 4.0,
    downloads: 2108,
    icon: 'L',
  },
  {
    id: 'finance-spreadsheet',
    name: 'Finance Spreadsheet',
    description: '财务报表模板与公式助手',
    longDescription: '内置 30+ 报表模板(损益/资产负债/现金流),含合规公式与多币种换算。',
    author: 'FinanceLab',
    version: '1.3.0',
    package: '@marketplace/finance-spreadsheet',
    source: 'src/finance.ts',
    tools: ['apply_template', 'compute_metric', 'fx_convert'],
    scopes: ['sheets:edit', 'ai:stream'],
    category: 'finance',
    tags: ['finance', 'accounting', 'spreadsheet', 'templates'],
    rating: 4.6,
    downloads: 7842,
    featured: true,
    icon: '$',
  },
  {
    id: 'web-clipper',
    name: 'Web Clipper',
    description: '把任意网页正文/图片剪藏到 doc',
    longDescription: '基于 Readability 提取正文,保留图片与代码块,自动添加来源元数据。',
    author: 'Productivity Co',
    version: '0.8.0',
    package: '@marketplace/web-clipper',
    source: 'src/web-clipper.ts',
    tools: ['clip_page', 'extract_readability', 'import_to_doc'],
    scopes: ['network:out', 'docs:edit'],
    category: 'productivity',
    tags: ['clip', 'web', 'readability'],
    rating: 4.2,
    downloads: 4501,
    icon: 'W',
  },
  {
    id: 'audio-transcribe',
    name: 'Audio Transcribe',
    description: '音频文件转写 + 说话人分离',
    longDescription: '支持 mp3 / wav / m4a,说话人 diarization,导出为带时间戳的 transcript 文档。',
    author: 'Speech Lab',
    version: '1.1.2',
    package: '@marketplace/audio-transcribe',
    source: 'src/audio.ts',
    tools: ['transcribe', 'diarize', 'export_doc'],
    scopes: ['files:read', 'ai:stream'],
    category: 'media',
    tags: ['audio', 'speech', 'transcription', 'meeting'],
    rating: 4.5,
    downloads: 6230,
    icon: 'A',
  },
  {
    id: 'diagram-mindmap',
    name: 'Mindmap Diagrams',
    description: '思维导图自动生成 + 主题样式',
    longDescription: '从纯文本大纲生成可缩放思维导图,支持 Mermaid / PlantUML / SVG 导出。',
    author: 'Visual Co',
    version: '0.9.3',
    package: '@marketplace/diagram-mindmap',
    source: 'src/mindmap.ts',
    tools: ['build_mindmap', 'export_mermaid', 'export_svg'],
    scopes: ['ai:stream', 'charts:write'],
    category: 'design',
    tags: ['mindmap', 'diagram', 'mermaid', 'planning'],
    rating: 4.4,
    downloads: 3411,
    icon: 'M',
  },
]

const MARKETPLACE_PLUGINS: MarketplacePluginEntry[] = [
  {
    id: 'slack-bridge',
    name: 'Slack Bridge',
    description: '把 Office AI 操作日志转发到 Slack 频道',
    longDescription: '所有 agent 行为可以异步通知到 Slack 频道,支持 thread 关联与回滚审批。',
    author: 'Workspace Integrations',
    version: '1.0.0',
    package: '@marketplace/slack-bridge',
    source: 'src/slack.ts',
    tools: ['send_notification', 'request_approval', 'sync_channel'],
    scopes: ['network:out'],
    requirements: ['slack-workspace-token'],
    category: 'collaboration',
    tags: ['slack', 'notification', 'chatops'],
    rating: 4.4,
    downloads: 8804,
    featured: true,
    icon: 'S',
  },
  {
    id: 'gdrive-export',
    name: 'Google Drive Export',
    description: '导出 Office 文档到 Google Drive(双向同步)',
    longDescription: 'OAuth 授权后,支持单文件 / 文件夹粒度的双向同步;权限映射与版本控制。',
    author: 'Cloud Sync',
    version: '2.3.1',
    package: '@marketplace/gdrive-export',
    source: 'src/gdrive.ts',
    tools: ['upload_doc', 'sync_folder', 'resolve_permissions'],
    scopes: ['network:out', 'files:write'],
    requirements: ['google-oauth-client'],
    category: 'collaboration',
    tags: ['gdrive', 'sync', 'cloud', 'oauth'],
    rating: 4.3,
    downloads: 12480,
    featured: true,
    icon: 'G',
  },
  {
    id: 'teams-bridge',
    name: 'MS Teams Bridge',
    description: 'Microsoft Teams 双向桥接',
    longDescription: '把 Office 操作结果同步到 Teams 频道 / chat,支持 adaptive card 与回链。',
    author: 'Enterprise Suite',
    version: '1.2.0',
    package: '@marketplace/teams-bridge',
    source: 'src/teams.ts',
    tools: ['send_card', 'sync_chat', 'request_approval'],
    scopes: ['network:out'],
    requirements: ['ms-tenant-id', 'ms-app-registration'],
    category: 'collaboration',
    tags: ['teams', 'enterprise', 'adaptive-card'],
    rating: 4.1,
    downloads: 3104,
    icon: 'T',
  },
  {
    id: 'audit-elk',
    name: 'Audit Log → ELK',
    description: '审计日志转发到 Elasticsearch / Kibana',
    longDescription: '所有 agent 行为可推送到 ELK,支持自定义 index template 与字段映射。',
    author: 'Compliance Co',
    version: '0.6.0',
    package: '@marketplace/audit-elk',
    source: 'src/audit-elk.ts',
    tools: ['push_event', 'apply_template', 'rotate_index'],
    scopes: ['audit:write', 'network:out'],
    requirements: ['elk-endpoint'],
    category: 'dev',
    tags: ['audit', 'elk', 'logging', 'compliance'],
    rating: 4.0,
    downloads: 1248,
    icon: 'E',
  },
  {
    id: 'cloud-storage-s3',
    name: 'Cloud Storage (S3)',
    description: '把 Office 文件备份到任意 S3 兼容存储',
    longDescription: '支持 AWS S3 / MinIO / R2,按对象锁 / 版本 / 生命周期策略保留。',
    author: 'Cloud Native',
    version: '1.0.5',
    package: '@marketplace/cloud-storage-s3',
    source: 'src/s3.ts',
    tools: ['upload', 'restore', 'lifecycle_apply'],
    scopes: ['network:out', 'files:read'],
    requirements: ['s3-bucket', 's3-access-key'],
    category: 'data',
    tags: ['s3', 'backup', 'storage', 'minio'],
    rating: 4.5,
    downloads: 5632,
    icon: '☁',
  },
  {
    id: 'local-models-ollama',
    name: 'Local Models (Ollama)',
    description: '本地 Ollama 模型路由',
    longDescription: '通过 Ollama 路由本地 LLM(llama3 / qwen / mistral);完全离线,适合隐私场景。',
    author: 'Privacy First',
    version: '0.4.1',
    package: '@marketplace/local-models-ollama',
    source: 'src/ollama.ts',
    tools: ['route_request', 'list_models', 'warmup'],
    scopes: ['ai:stream', 'network:local'],
    requirements: ['ollama-runtime'],
    category: 'data',
    tags: ['ollama', 'local', 'privacy', 'offline'],
    rating: 4.6,
    downloads: 7104,
    featured: true,
    icon: 'O',
  },
  {
    id: 'analytics-posthog',
    name: 'Analytics (PostHog)',
    description: '把产品事件转发到 PostHog',
    longDescription: '为 GenOffice 操作事件提供产品分析;支持自定义 event / property。',
    author: 'Analytics Suite',
    version: '0.3.2',
    package: '@marketplace/analytics-posthog',
    source: 'src/posthog.ts',
    tools: ['capture', 'identify', 'flush'],
    scopes: ['network:out'],
    requirements: ['posthog-api-key'],
    category: 'data',
    tags: ['analytics', 'posthog', 'product'],
    rating: 3.9,
    downloads: 612,
    icon: 'P',
  },
  {
    id: 'design-tokens',
    name: 'Design Tokens Sync',
    description: '把 Figma tokens 注入到文档与样式',
    longDescription: '从 Figma / Style Dictionary 拉取 design tokens,应用到 docs / sheets / slides 主题。',
    author: 'DesignOps',
    version: '0.7.0',
    package: '@marketplace/design-tokens',
    source: 'src/design-tokens.ts',
    tools: ['fetch_tokens', 'apply_theme', 'export_css'],
    scopes: ['network:out', 'docs:edit'],
    requirements: ['figma-personal-token'],
    category: 'design',
    tags: ['design-system', 'figma', 'tokens'],
    rating: 4.2,
    downloads: 2104,
    icon: 'D',
  },
]

/* ── Published catalog ───────────────────────────────────────────────
 * Extensions submitted through `home:marketplace-upload` are persisted as
 * `<kind>.<id>.json` under DATA_DIR/marketplace-uploads and join the live
 * catalog immediately: publishing makes an extension searchable, installable
 * and durable across restarts. The upload response still reports a pending
 * review status so the UI can label it honestly. */
const UPLOAD_DIR = join(DATA_DIR, 'marketplace-uploads')
try {
  mkdirSync(UPLOAD_DIR, { recursive: true })
} catch {}

interface UploadedFile {
  kind: 'skill' | 'plugin'
  entry: MarketplaceSkillEntry | MarketplacePluginEntry
}

let uploadedCache: UploadedFile[] | null = null

function loadUploaded(): UploadedFile[] {
  if (uploadedCache) return uploadedCache
  const out: UploadedFile[] = []
  try {
    if (existsSync(UPLOAD_DIR)) {
      for (const file of readdirSync(UPLOAD_DIR)) {
        if (!file.endsWith('.json')) continue
        try {
          const raw = JSON.parse(readFileSync(join(UPLOAD_DIR, file), 'utf-8')) as {
            kind?: unknown
            entry?: unknown
          }
          const entry = raw.entry as (MarketplaceSkillEntry | MarketplacePluginEntry) | undefined
          if (!entry || typeof entry.id !== 'string') continue
          out.push({
            kind: raw.kind === 'plugin' ? 'plugin' : 'skill',
            entry: {
              ...entry,
              rating: typeof entry.rating === 'number' ? entry.rating : 0,
              downloads: typeof entry.downloads === 'number' ? entry.downloads : 0,
              tags: Array.isArray(entry.tags) ? entry.tags : [],
              tools: Array.isArray(entry.tools) ? entry.tools : [],
              scopes: Array.isArray(entry.scopes) ? entry.scopes : [],
            },
          })
        } catch {
          // an unparseable upload must not break the whole catalog
        }
      }
    }
  } catch {}
  uploadedCache = out
  return out
}

/** a publish rewrites the directory; drop the memo so it is re-read */
function invalidateUploads(): void {
  uploadedCache = null
}

/** curated skills + everything published at runtime (published ids win) */
function allMarketplaceSkills(): MarketplaceSkillEntry[] {
  const uploaded = loadUploaded().filter((u) => u.kind === 'skill').map((u) => u.entry)
  const publishedIds = new Set(uploaded.map((u) => u.id))
  return [...MARKETPLACE_SKILLS.filter((s) => !publishedIds.has(s.id)), ...uploaded] as MarketplaceSkillEntry[]
}

function allMarketplacePlugins(): MarketplacePluginEntry[] {
  const uploaded = loadUploaded().filter((u) => u.kind === 'plugin').map((u) => u.entry)
  const publishedIds = new Set(uploaded.map((u) => u.id))
  return [...MARKETPLACE_PLUGINS.filter((p) => !publishedIds.has(p.id)), ...uploaded] as MarketplacePluginEntry[]
}

function listMarketplaceSkills(): MarketplaceSkillEntry[] {
  const installed = loadSkills()
  return allMarketplaceSkills().map((entry) => ({
    ...entry,
    installed: installed.some((s) => s.id === entry.id),
  }))
}

function listMarketplacePlugins(): MarketplacePluginEntry[] {
  const installed = loadPlugins()
  return allMarketplacePlugins().map((entry) => ({
    ...entry,
    installed: installed.some((p) => p.id === entry.id),
  }))
}

export function listMarketplaceCategories(): Array<{
  id: MarketplaceCategory
  label: string
  count: number
}> {
  const cats: Record<MarketplaceCategory, string> = {
    productivity: '生产力',
    data: '数据',
    dev: '开发',
    media: '媒体',
    translation: '翻译',
    collaboration: '协作',
    finance: '财务',
    design: '设计',
  }
  const out: Record<string, number> = {}
  ;[...allMarketplaceSkills(), ...allMarketplacePlugins()].forEach((e) => {
    out[e.category] = (out[e.category] ?? 0) + 1
  })
  return (Object.keys(cats) as MarketplaceCategory[]).map((id) => ({
    id,
    label: cats[id],
    count: out[id] ?? 0,
  }))
}

export interface MarketplaceSearchFilters {
  q?: string
  category?: MarketplaceCategory
  type?: 'skill' | 'plugin'
  minRating?: number
  installed?: boolean | 'all'
  sort?: 'popular' | 'rating' | 'newest' | 'name'
}

export function searchMarketplace(filters: MarketplaceSearchFilters): {
  skills: MarketplaceSkillEntry[]
  plugins: MarketplacePluginEntry[]
  total: number
} {
  const q = (filters.q ?? '').trim().toLowerCase()
  const installedSkills = new Set(loadSkills().map((s) => s.id))
  const installedPlugins = new Set(loadPlugins().map((p) => p.id))
  const minRating = filters.minRating ?? 0
  const sort = filters.sort ?? 'popular'

  function matches(entry: { id: string; name: string; description: string; longDescription?: string; tags: string[]; author: string; rating: number }): boolean {
    if (q.length > 0) {
      const haystack = [
        entry.id,
        entry.name,
        entry.description,
        entry.longDescription ?? '',
        entry.tags.join(' '),
        entry.author,
      ]
        .join(' \u0001 ')
        .toLowerCase()
      if (!haystack.includes(q)) return false
    }
    if (entry.rating < minRating) return false
    return true
  }

  function sortBy<T extends { rating: number; downloads: number; name: string }>(arr: T[]): T[] {
    const sorted = [...arr]
    switch (sort) {
      case 'rating':
        sorted.sort((a, b) => b.rating - a.rating || b.downloads - a.downloads)
        break
      case 'newest':
        sorted.sort((a, b) => b.name.localeCompare(a.name))
        break
      case 'name':
        sorted.sort((a, b) => a.name.localeCompare(b.name))
        break
      case 'popular':
      default:
        sorted.sort((a, b) => b.downloads - a.downloads)
    }
    return sorted
  }

  let skills = allMarketplaceSkills().filter((e) => {
    if (!matches(e)) return false
    if (filters.category && e.category !== filters.category) return false
    if (filters.installed === true && !installedSkills.has(e.id)) return false
    if (filters.installed === false && installedSkills.has(e.id)) return false
    return true
  }).map((e) => ({ ...e, installed: installedSkills.has(e.id) }))

  let plugins = allMarketplacePlugins().filter((e) => {
    if (!matches(e)) return false
    if (filters.category && e.category !== filters.category) return false
    if (filters.installed === true && !installedPlugins.has(e.id)) return false
    if (filters.installed === false && installedPlugins.has(e.id)) return false
    return true
  }).map((e) => ({ ...e, installed: installedPlugins.has(e.id) }))

  if (filters.type === 'skill') plugins = []
  if (filters.type === 'plugin') skills = []

  return {
    skills: sortBy(skills),
    plugins: sortBy(plugins),
    total: skills.length + plugins.length,
  }
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
    const entry = allMarketplaceSkills().find((s) => s.id === id)
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
    const entry = allMarketplacePlugins().find((p) => p.id === id)
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

  registerHandle('home:marketplace-categories', () => {
    return { categories: listMarketplaceCategories() }
  })

  registerHandle('home:marketplace-search', (_event: unknown, args: unknown) => {
    const a = (args || {}) as {
      q?: string
      category?: MarketplaceCategory
      type?: 'skill' | 'plugin'
      minRating?: number
      installed?: boolean | 'all'
      sort?: 'popular' | 'rating' | 'newest' | 'name'
    }
    return { ...searchMarketplace(a), filters: a }
  })

  registerHandle('home:marketplace-detail', (_event: unknown, args: unknown) => {
    const { id, type } = (args || {}) as { id?: string; type?: 'skill' | 'plugin' }
    if (!id) return { ok: false, error: 'Missing id' }
    if (type === 'plugin') {
      const entry = allMarketplacePlugins().find((p) => p.id === id)
      if (!entry) return { ok: false, error: `Plugin "${id}" not found in marketplace` }
      const installed = loadPlugins().find((p) => p.id === id)
      return { ok: true, type: 'plugin', entry, installed: installed ?? null }
    }
    const entry = allMarketplaceSkills().find((s) => s.id === id)
    if (!entry) return { ok: false, error: `Skill "${id}" not found in marketplace` }
    const installed = loadSkills().find((s) => s.id === id)
    return { ok: true, type: 'skill', entry, installed: installed ?? null }
  })

  // ── 插件上传:用户提交 marketplace 扩展元数据,服务端校验 + 落盘 + 立即可见 ──

  function validateUpload(
    raw: unknown,
    expectedKind: 'skill' | 'plugin',
  ): { ok: true; entry: MarketplaceSkillEntry | MarketplacePluginEntry } | { ok: false; error: string } {
    if (!raw || typeof raw !== 'object') {
      return { ok: false, error: 'Payload must be a JSON object' }
    }
    const r = raw as Record<string, unknown>
    const id = String(r.id ?? '').trim()
    const name = String(r.name ?? '').trim()
    const description = String(r.description ?? '').trim()
    const author = String(r.author ?? '').trim() || 'Anonymous'
    const version = String(r.version ?? '').trim()
    const pkg = String(r.package ?? '').trim() || `@marketplace/${id}`
    const source = String(r.source ?? '').trim() || `src/${id}.ts`
    const tools = Array.isArray(r.tools) ? r.tools.map((t) => String(t)) : []
    const scopes = Array.isArray(r.scopes) ? r.scopes.map((s) => String(s)) : []
    const category = String(r.category ?? '') as MarketplaceCategory
    const validCategories: MarketplaceCategory[] = [
      'productivity',
      'data',
      'dev',
      'media',
      'translation',
      'collaboration',
      'finance',
      'design',
    ]
    if (!/^[a-z][a-z0-9-]{1,40}$/.test(id)) {
      return { ok: false, error: 'id must match /^[a-z][a-z0-9-]{1,40}$/' }
    }
    if (name.length < 2 || name.length > 60) {
      return { ok: false, error: 'name must be 2-60 chars' }
    }
    if (description.length < 10 || description.length > 280) {
      return { ok: false, error: 'description must be 10-280 chars' }
    }
    if (!/^\d+\.\d+\.\d+/.test(version)) {
      return { ok: false, error: 'version must be semver (e.g. 1.0.0)' }
    }
    if (tools.length === 0) {
      return { ok: false, error: 'tools[] must contain at least one tool name' }
    }
    if (!validCategories.includes(category)) {
      return { ok: false, error: `category must be one of ${validCategories.join(', ')}` }
    }
    if (expectedKind === 'skill') {
      if (MARKETPLACE_SKILLS.some((s) => s.id === id)) {
        return { ok: false, error: `Skill "${id}" already exists in marketplace` }
      }
      return {
        ok: true,
        entry: {
          id,
          name,
          description,
          author,
          version,
          package: pkg,
          source,
          tools,
          scopes,
          category,
          tags: Array.isArray(r.tags) ? r.tags.map((t) => String(t)) : [],
          rating: typeof r.rating === 'number' ? Math.max(0, Math.min(5, r.rating)) : 0,
          downloads: 0,
          icon: typeof r.icon === 'string' ? r.icon.slice(0, 4) : '?',
        },
      }
    }
    if (MARKETPLACE_PLUGINS.some((p) => p.id === id)) {
      return { ok: false, error: `Plugin "${id}" already exists in marketplace` }
    }
    return {
      ok: true,
      entry: {
        id,
        name,
        description,
        author,
        version,
        package: pkg,
        source,
        tools,
        scopes,
        requirements: Array.isArray(r.requirements) ? r.requirements.map((t) => String(t)) : [],
        category,
        tags: Array.isArray(r.tags) ? r.tags.map((t) => String(t)) : [],
        rating: typeof r.rating === 'number' ? Math.max(0, Math.min(5, r.rating)) : 0,
        downloads: 0,
        icon: typeof r.icon === 'string' ? r.icon.slice(0, 4) : '?',
      },
    }
  }

  registerHandle('home:marketplace-upload', (_event: unknown, args: unknown) => {
    const a = (args || {}) as { kind?: 'skill' | 'plugin'; payload?: unknown }
    const kind = a.kind === 'plugin' ? 'plugin' : 'skill'
    const validated = validateUpload(a.payload, kind)
    if (validated.ok === false) return { ok: false, error: validated.error }
    const entry = validated.entry
    try {
      const file = join(UPLOAD_DIR, `${kind}.${entry.id}.json`)
      writeFileSync(
        file,
        JSON.stringify(
          { kind, entry, uploadedAt: new Date().toISOString(), reviewStatus: 'pending' },
          null,
          2,
        ),
      )
      invalidateUploads()
    } catch (err) {
      return { ok: false, error: `Write failed: ${err instanceof Error ? err.message : String(err)}` }
    }
    return {
      ok: true,
      kind,
      entry: { ...validated.entry, installed: false },
      reviewStatus: 'pending',
      message:
        '已发布到本地 marketplace。扩展已立即出现在市场中,可搜索/安装;提交 GenOffice 团队审核后即可进入公共目录。',
    }
  })

  registerHandle('home:marketplace-list-uploads', () => {
    try {
      if (!existsSync(UPLOAD_DIR)) return { uploads: [] }
      const files = readdirSync(UPLOAD_DIR).filter((f) => f.endsWith('.json'))
      const items = files.map((f) => {
        try {
          const raw = JSON.parse(readFileSync(join(UPLOAD_DIR, f), 'utf-8'))
          return { file: f, kind: raw.kind, id: raw.entry?.id, name: raw.entry?.name, uploadedAt: raw.uploadedAt }
        } catch {
          return { file: f, error: 'unparseable' }
        }
      })
      return { uploads: items }
    } catch (err) {
      return { uploads: [], error: err instanceof Error ? err.message : String(err) }
    }
  })

  // ── 复合接口: 一次返回 skills + plugins ──

  registerHandle('home:get-skills-and-plugins', () => {
    return { skills: loadSkills(), plugins: loadPlugins() }
  })
}
