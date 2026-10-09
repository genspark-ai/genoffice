/**
 * The in-app manual's topic registry (issue #1520).
 *
 * Each topic is one markdown file per language under ./topics (`<id>.zh.md`,
 * `<id>.en.md`), bundled at build time via import.meta.glob(?raw). A topic is
 * discovered by listing it here — the registry is what search indexes and the
 * sidebar renders, so a feature without a manual entry is a visible gap in
 * this list rather than an invisible one.
 */

import { LANGS } from '@genoffice/i18n'
import { topicTitle } from './help-titles'

export type HelpGroupId = 'start' | 'apps' | 'ai' | 'manage'

export interface HelpGroup {
  id: HelpGroupId
}

export const HELP_GROUPS: HelpGroup[] = [
  { id: 'start' },
  { id: 'apps' },
  { id: 'ai' },
  { id: 'manage' },
]

export interface HelpTopicMeta {
  id: string
  group: HelpGroupId
  /** within-group sort key */
  order: number
  /** search terms beyond the title (both languages welcome) */
  keywords: string[]
}

/** Authoritative topic list. Body files live in ./topics/<id>.<lang>.md */
export const HELP_TOPICS: HelpTopicMeta[] = [
  {
    id: 'install',
    group: 'start',
    order: 1,
    keywords: ['安装', 'install', '下载', 'download', 'dmg', 'deb', 'rpm', 'appimage', '安装包'],
  },
  {
    id: 'getting-started',
    group: 'start',
    order: 2,
    keywords: ['界面', 'overview', '标签页', 'tab', '窗口', '保存', 'save', '快捷键', 'shortcuts'],
  },
  {
    id: 'home-screen',
    group: 'start',
    order: 3,
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
    order: 4,
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
    id: 'shortcuts',
    group: 'start',
    order: 5,
    keywords: [
      '快捷键',
      'shortcuts',
      '键盘',
      'keyboard',
      'hotkeys',
      '快捷方式',
      'command',
      'chord',
      'mac',
      'windows',
      'linux',
      'ctrl',
    ],
  },
  {
    id: 'docs',
    group: 'apps',
    order: 1,
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
    keywords: ['excel', '表格', '公式', 'formula', 'xlsx', 'csv', 'tsv', 'sheet', '单元格', 'cell'],
  },
  {
    id: 'slides',
    group: 'apps',
    order: 3,
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
    keywords: ['md', 'markdown', '预览', 'preview', 'gfm'],
  },
  {
    id: 'html',
    group: 'apps',
    order: 6,
    keywords: ['网页', 'html', '预览', 'preview', '检查器', 'inspector'],
  },
  {
    id: 'ai-panel',
    group: 'ai',
    order: 1,
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
  {
    id: 'cli',
    group: 'manage',
    order: 3,
    keywords: ['命令行', 'cli', 'genoffice', 'convert', 'render', '命令行工具', '终端'],
  },
  {
    id: 'mcp',
    group: 'manage',
    order: 4,
    keywords: [
      'mcp',
      'model context protocol',
      '编程助手',
      'coding agent',
      'claude',
      'codex',
      'cursor',
      '智能体',
    ],
  },
]

/**
 * Topic bodies per language, keyed `<id>.<lang>`; missing pairs fall back to en.
 *
 * Lazy on purpose. Inlined eagerly, 294 markdown files pushed the shell
 * renderer's entry chunk from 1.36 MB to 2.63 MB, and Home parses that on
 * every launch to open a screen most sessions never reach. Each body is its
 * own chunk now, fetched when its topic is opened.
 */
const bodies = import.meta.glob<string>('./topics/*.md', {
  query: '?raw',
  import: 'default',
})

/** screenshots referenced from topic bodies as ![alt](img/<name>.png) */
const images = import.meta.glob<string>('./topics/img/*.png', {
  query: '?url',
  import: 'default',
  eager: true,
})

/**
 * Bundled asset URL for a topic image href, or undefined when absent.
 *
 * A body writes `img/<name>.png` and never names a language: the figures are
 * the running UI, so a body in a language with no figure of its own falls back
 * to the English one rather than to a broken image. Where a figure *was*
 * captured per language (`<name>.<lang>.png`) that one wins, so a reader whose
 * language has real UI in the shot is not shown the English chrome instead.
 */
export function helpImage(href: string, lang?: string): string | undefined {
  const name = href.replace(/^\.\//, '')
  const suffix = '.png'
  if (!name.endsWith(suffix)) return images[`./topics/${name}`]
  const stem = name.slice(0, -suffix.length)
  const suffixLang = lang ? helpLangSuffix(lang) : null
  if (suffixLang) {
    const localized = images[`./topics/${stem}.${suffixLang}${suffix}`]
    if (localized !== undefined) return localized
  }
  return images[`./topics/${stem}.en${suffix}`] ?? images[`./topics/${name}`]
}

/**
 * The bundle suffix a UI language maps to.
 *
 * `zh-TW` is a real file suffix, not a `zh` variant: the two scripts are
 * written differently enough that one shared body would leave half of each
 * reading in the wrong orthography. Everything else is its own suffix, and an
 * unknown one falls through to English rather than to a file that is not there.
 */
export function helpLangSuffix(lang: string): string {
  return (LANGS as readonly string[]).includes(lang) ? lang : 'en'
}

/** A topic's body in the reader's language, falling back to English. */
export async function helpBody(id: string, lang: string): Promise<string | null> {
  const direct = bodies[`./topics/${id}.${helpLangSuffix(lang)}.md`]
  if (direct) return (await direct()) as string
  const english = bodies[`./topics/${id}.en.md`]
  if (english) return (await english()) as string
  return null
}

/** True when the topic has a body in this language rather than the English fallback. */
export function helpHasBody(id: string, lang: string): boolean {
  // the key is in the map without the file having been fetched, so this stays
  // synchronous where a caller only needs to know the pair exists
  return bodies[`./topics/${id}.${helpLangSuffix(lang)}.md`] !== undefined
}

/**
 * Topics whose title or body mentions the query (case-insensitive, both langs).
 *
 * Reads every body, so it is a real cost: it runs when the reader types, not on
 * mount, which is why the screen only calls it once the box has something in
 * it. The results are cached per language by the caller.
 */
export async function searchTopics(query: string, lang: string): Promise<Set<string>> {
  const q = query.trim().toLowerCase()
  if (!q) return new Set(HELP_TOPICS.map((t) => t.id))
  const hits = new Set<string>()
  await Promise.all(
    HELP_TOPICS.map(async (t) => {
      const [own, english] = await Promise.all([helpBody(t.id, lang), helpBody(t.id, 'en')])
      // both the reader's title and the English one, so a query typed in either
      // language finds the topic whichever body is on screen
      const hay = [
        topicTitle(t.id, lang),
        topicTitle(t.id, 'en'),
        ...t.keywords,
        own ?? '',
        english ?? '',
      ]
        .join('\n')
        .toLowerCase()
      if (hay.includes(q)) hits.add(t.id)
    }),
  )
  return hits
}
