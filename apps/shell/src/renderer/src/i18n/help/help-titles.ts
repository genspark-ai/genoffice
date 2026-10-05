/**
 * Per-language titles for the manual's sidebar and article headings.
 *
 * The registry holds the structure — which topics exist, their group and order
 * — and this holds the words, one entry per UI language. They are separate
 * because they change for different reasons and on different schedules: a new
 * topic lands in the registry, a new language lands here, and a language with
 * no body yet still needs a title so the sidebar never shows an empty row.
 */

import type { HelpGroupId } from './help-registry'

/** The languages the manual has titles for; `@genoffice/i18n`'s list, minus none. */
export const HELP_TITLE_LANGS = [
  'zh',
  'zh-TW',
  'en',
  'ja',
  'ko',
  'fr',
  'de',
  'es',
  'th',
  'id',
  'ru',
  'ar',
  'pt',
  'it',
  'pl',
  'cs',
  'nl',
  'ms',
  'he',
  'hi',
  'vi',
] as const

export type HelpTitleLang = (typeof HELP_TITLE_LANGS)[number]

/** Sidebar group headings, per language. */
const GROUP_TITLES: Record<HelpGroupId, Record<HelpTitleLang, string>> = {
  start: {
    zh: '入门',
    'zh-TW': '入門',
    en: 'Getting started',
    ja: 'はじめに',
    ko: '시작하기',
    fr: 'Bien démarrer',
    de: 'Erste Schritte',
    es: 'Primeros pasos',
    th: 'เริ่มต้น',
    id: 'Memulai',
    ru: 'Начало работы',
    ar: 'البداية',
    pt: 'Primeiros passos',
    it: 'Per iniziare',
    pl: 'Pierwsze kroki',
    cs: 'Začínáme',
    nl: 'Aan de slag',
    ms: 'Mula',
    he: 'תחילת העבודה',
    hi: 'शुरुआत',
    vi: 'Bắt đầu',
  },
  apps: {
    zh: '应用',
    'zh-TW': '應用程式',
    en: 'Applications',
    ja: 'アプリ',
    ko: '앱',
    fr: 'Applications',
    de: 'Anwendungen',
    es: 'Aplicaciones',
    th: 'แอปพลิเคชัน',
    id: 'Aplikasi',
    ru: 'Приложения',
    ar: 'التطبيقات',
    pt: 'Aplicações',
    it: 'Applicazioni',
    pl: 'Aplikacje',
    cs: 'Aplikace',
    nl: 'Toepassingen',
    ms: 'Aplikasi',
    he: 'יישומים',
    hi: 'ऐप्स',
    vi: 'Ứng dụng',
  },
  ai: {
    zh: 'AI 能力',
    'zh-TW': 'AI 功能',
    en: 'AI features',
    ja: 'AI 機能',
    ko: 'AI 기능',
    fr: 'Fonctions IA',
    de: 'KI-Funktionen',
    es: 'Funciones de IA',
    th: 'ความสามารถ AI',
    id: 'Fitur AI',
    ru: 'Возможности ИИ',
    ar: 'ميزات الذكاء الاصطناعي',
    pt: 'Recursos de IA',
    it: 'Funzioni IA',
    pl: 'Funkcje AI',
    cs: 'Funkce AI',
    nl: 'AI-functies',
    ms: 'Ciri AI',
    he: 'יכולות AI',
    hi: 'AI सुविधाएँ',
    vi: 'Tính năng AI',
  },
  manage: {
    zh: '文件与设置',
    'zh-TW': '檔案與設定',
    en: 'Files & settings',
    ja: 'ファイルと設定',
    ko: '파일 및 설정',
    fr: 'Fichiers et paramètres',
    de: 'Dateien & Einstellungen',
    es: 'Archivos y ajustes',
    th: 'ไฟล์และการตั้งค่า',
    id: 'Berkas & pengaturan',
    ru: 'Файлы и настройки',
    ar: 'الملفات والإعدادات',
    pt: 'Arquivos e configurações',
    it: 'File e impostazioni',
    pl: 'Pliki i ustawienia',
    cs: 'Soubory a nastavení',
    nl: 'Bestanden en instellingen',
    ms: 'Fail & tetapan',
    he: 'קבצים והגדרות',
    hi: 'फ़ाइलें और सेटिंग्स',
    vi: 'Tệp và cài đặt',
  },
}

/** Article titles, per language, keyed by topic id. A language lands here when
 *  its bodies land in ./topics; until then the title falls back to English. */
const TOPIC_TITLES: Record<string, Record<HelpTitleLang, string>> = {
  'getting-started': {
    zh: '快速入门：界面与基本操作',
    en: 'Quick start: the interface and the basics',
  },
  'home-screen': {
    zh: '主屏（Home）：文件都在哪',
    en: 'The Home screen: where your files live',
  },
  'tabs-and-windows': {
    zh: '标签页与窗口管理',
    en: 'Tabs and window management',
  },
  docs: {
    zh: 'Docs：文字处理',
    en: 'Docs: word processing',
  },
  sheets: {
    zh: 'Sheets：电子表格',
    en: 'Sheets: spreadsheets',
  },
  slides: {
    zh: 'Slides：演示文稿',
    en: 'Slides: presentations',
  },
  pdf: {
    zh: 'PDF：阅读、注释与涂黑',
    en: 'PDF: reading, annotating and redacting',
  },
  markdown: {
    zh: 'Markdown 编辑器',
    en: 'The Markdown editor',
  },
  html: {
    zh: 'HTML 编辑器',
    en: 'The HTML editor',
  },
  'ai-panel': {
    zh: 'AI 助手面板',
    en: 'The AI assistant panel',
  },
  'ai-models': {
    zh: 'AI 模型与设置',
    en: 'AI models and settings',
  },
  fonts: {
    zh: '字体：系统字体与可下载字体',
    en: 'Fonts: system faces and downloadable families',
  },
  'file-ops': {
    zh: '文件操作：重命名、删除、导出',
    en: 'File operations: rename, delete, export',
  },
  'settings-integrations': {
    zh: '设置、语言、主题与 MCP 集成',
    en: 'Settings, language, theme and MCP integrations',
  },
}

/** Sidebar group heading in the reader's language; English when unknown. */
export function groupTitle(id: HelpGroupId, lang: string): string {
  const row = GROUP_TITLES[id]
  return (row?.[lang as HelpTitleLang] ?? row?.en) as string
}

/** Article title in the reader's language; falls back to the topic id, which
 *  is never pretty, so a language added without its titles shows up as a gap
 *  rather than as a blank row. */
export function topicTitle(id: string, lang: string): string {
  const row = TOPIC_TITLES[id]
  return (row?.[lang as HelpTitleLang] ?? row?.en ?? id) as string
}
