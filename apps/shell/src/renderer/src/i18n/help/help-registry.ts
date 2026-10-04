/**
 * The in-app manual's topic registry (issue #1520).
 *
 * Each topic is one markdown file per language under ./topics (`<id>.zh.md`,
 * `<id>.en.md`), bundled at build time via import.meta.glob(?raw). A topic is
 * discovered by listing it here — the registry is what search indexes and the
 * sidebar renders, so a feature without a manual entry is a visible gap in
 * this list rather than an invisible one.
 */

export type HelpGroupId = 'start' | 'apps' | 'ai' | 'manage'

export interface HelpGroup {
  id: HelpGroupId
  titleZh: string
  titleEn: string
}

export const HELP_GROUPS: HelpGroup[] = [
  { id: 'start', titleZh: '入门', titleEn: 'Getting started' },
  { id: 'apps', titleZh: '应用', titleEn: 'Applications' },
  { id: 'ai', titleZh: 'AI 能力', titleEn: 'AI features' },
  { id: 'manage', titleZh: '文件与设置', titleEn: 'Files & settings' },
]

export interface HelpTopicMeta {
  id: string
  group: HelpGroupId
  /** within-group sort key */
  order: number
  titleZh: string
  titleEn: string
  /** search terms beyond the title (both languages welcome) */
  keywords: string[]
}

/** Authoritative topic list. Body files live in ./topics/<id>.<lang>.md */
export const HELP_TOPICS: HelpTopicMeta[] = [
  {
    id: 'getting-started',
    group: 'start',
    order: 1,
    titleZh: '快速入门：界面与基本操作',
    titleEn: 'Quick start: the interface and the basics',
    keywords: ['界面', 'overview', '标签页', 'tab', '窗口', '保存', 'save', '快捷键', 'shortcuts'],
  },
  {
    id: 'home-screen',
    group: 'start',
    order: 2,
    titleZh: '主屏（Home）：文件都在哪',
    titleEn: 'The Home screen: where your files live',
    keywords: [
      '最近',
      'recents',
      '收藏',
      'starred',
      '文件夹',
      'folders',
      '云项目',
      'cloud',
      '回收站',
      'trash',
      '搜索',
      'search',
    ],
  },
  {
    id: 'tabs-and-windows',
    group: 'start',
    order: 3,
    titleZh: '标签页与窗口管理',
    titleEn: 'Tabs and window management',
    keywords: [
      '重命名',
      'rename',
      '双击',
      'double-click',
      '拖拽',
      'drag',
      '分离',
      'detach',
      '排序',
      'reorder',
    ],
  },
  {
    id: 'docs',
    group: 'apps',
    order: 1,
    titleZh: 'Docs：文字处理',
    titleEn: 'Docs: word processing',
    keywords: [
      'word',
      '文档',
      '样式',
      'styles',
      '表格',
      'table',
      '页眉',
      'header',
      '目录',
      'toc',
      '批注',
      'comments',
      'docx',
    ],
  },
  {
    id: 'sheets',
    group: 'apps',
    order: 2,
    titleZh: 'Sheets：电子表格',
    titleEn: 'Sheets: spreadsheets',
    keywords: ['excel', '表格', '公式', 'formula', 'xlsx', 'csv', 'tsv', 'sheet', '单元格', 'cell'],
  },
  {
    id: 'slides',
    group: 'apps',
    order: 3,
    titleZh: 'Slides：演示文稿',
    titleEn: 'Slides: presentations',
    keywords: [
      'ppt',
      'pptx',
      '幻灯片',
      '演示',
      '版式',
      'layout',
      '母版',
      'master',
      '图表',
      'chart',
    ],
  },
  {
    id: 'pdf',
    group: 'apps',
    order: 4,
    titleZh: 'PDF：阅读、注释与涂黑',
    titleEn: 'PDF: reading, annotating and redacting',
    keywords: [
      '高亮',
      'highlight',
      '便签',
      'note',
      '涂黑',
      'redact',
      '签名',
      'signature',
      '表单',
      'form',
      '合并',
      'merge',
      '拆分',
      'split',
    ],
  },
  {
    id: 'markdown',
    group: 'apps',
    order: 5,
    titleZh: 'Markdown 编辑器',
    titleEn: 'The Markdown editor',
    keywords: ['md', 'markdown', '预览', 'preview', 'gfm'],
  },
  {
    id: 'html',
    group: 'apps',
    order: 6,
    titleZh: 'HTML 编辑器',
    titleEn: 'The HTML editor',
    keywords: ['网页', 'html', '预览', 'preview', '检查器', 'inspector'],
  },
  {
    id: 'ai-panel',
    group: 'ai',
    order: 1,
    titleZh: 'AI 助手面板',
    titleEn: 'The AI assistant panel',
    keywords: [
      'ai',
      '助手',
      'assistant',
      '面板',
      'panel',
      '生成',
      'generate',
      '回滚',
      'rollback',
      '撤销',
    ],
  },
  {
    id: 'ai-models',
    group: 'ai',
    order: 2,
    titleZh: 'AI 模型与设置',
    titleEn: 'AI models and settings',
    keywords: [
      '模型',
      'model',
      'provider',
      'key',
      'byok',
      '端点',
      'endpoint',
      'genspark',
      '登录',
      'login',
    ],
  },
  {
    id: 'fonts',
    group: 'ai',
    order: 3,
    titleZh: '字体：系统字体与可下载字体',
    titleEn: 'Fonts: system faces and downloadable families',
    keywords: [
      '字体',
      'font',
      '下载',
      'download',
      'cjk',
      '中文',
      '日文',
      '韩文',
      'install',
      'catalog',
    ],
  },
  {
    id: 'file-ops',
    group: 'manage',
    order: 1,
    titleZh: '文件操作：重命名、删除、导出',
    titleEn: 'File operations: rename, delete, export',
    keywords: [
      '重命名',
      'rename',
      '删除',
      'delete',
      '导出',
      'export',
      'pdf',
      '另存为',
      'save as',
      '复制',
      'duplicate',
    ],
  },
  {
    id: 'settings-integrations',
    group: 'manage',
    order: 2,
    titleZh: '设置、语言、主题与 MCP 集成',
    titleEn: 'Settings, language, theme and MCP integrations',
    keywords: [
      '设置',
      'settings',
      '语言',
      'language',
      '主题',
      'theme',
      '深色',
      'dark',
      'mcp',
      '默认应用',
      'default app',
    ],
  },
]

/** topic bodies per language, keyed `<id>.<lang>`; missing pairs fall back to en */
const bodies = import.meta.glob<string>('./topics/*.md', {
  query: '?raw',
  import: 'default',
  eager: true,
})

/** screenshots referenced from topic bodies as ![alt](img/<name>.png) */
const images = import.meta.glob<string>('./topics/img/*.png', {
  query: '?url',
  import: 'default',
  eager: true,
})

/** Bundled asset URL for a topic image href, or undefined when absent */
export function helpImage(href: string): string | undefined {
  return images[`./topics/${href.replace(/^\.\//, '')}`]
}

export function helpBody(id: string, lang: string): string | null {
  const short = lang.startsWith('zh') ? 'zh' : 'en'
  const direct = bodies[`./topics/${id}.${short}.md`]
  if (direct !== undefined) return direct
  return bodies[`./topics/${id}.en.md`] ?? null
}

/** topics whose title or body mentions the query (case-insensitive, both langs) */
export function searchTopics(query: string, lang: string): Set<string> {
  const q = query.trim().toLowerCase()
  if (!q) return new Set(HELP_TOPICS.map((t) => t.id))
  const hits = new Set<string>()
  for (const t of HELP_TOPICS) {
    const hay = [
      t.titleZh,
      t.titleEn,
      ...t.keywords,
      helpBody(t.id, lang) ?? '',
      helpBody(t.id, 'en') ?? '',
    ]
      .join('\n')
      .toLowerCase()
    if (hay.includes(q)) hits.add(t.id)
  }
  return hits
}
