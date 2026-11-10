import type { Editor } from '@tiptap/core'
import type { Node as ProseMirrorNode } from '@tiptap/pm/model'
import { DATE_LOCALES, getLang, t, type StringKey } from '../i18n/locale'
import type { PmNode } from './convert'

/** hidden bookmarks that let "Remove Current Cover Page" find a gallery cover after a save/reopen */
export const COVER_START_MARK = '_GenOfficeCoverPage'
export const COVER_END_MARK = '_GenOfficeCoverPageEnd'

/**
 * Built-in cover page library (the preset gallery behind Insert → Cover
 * Page). Everything is built from layout primitives the engine
 * already supports: paragraph shading / single-line borders / alignment /
 * spacing + text color / size / bold / italic. No drawing objects, so the
 * result holds up in Word after saving.
 */

export interface CoverPara {
  text?: string
  bold?: boolean
  italic?: boolean
  /** hex without '#' */
  color?: string
  sizeHalfPoints?: number
  align?: 'left' | 'center' | 'right'
  /** twips */
  spaceBefore?: number
  spaceAfter?: number
  shadingFill?: string
  /** subset of "tblr" */
  borders?: string
  indentLeft?: number
  indentRight?: number
}

export interface CoverPreset {
  id: string
  name: string
  paras: CoverPara[]
}

const TITLE = () => t('editorCoverTitle')
const SUBTITLE = () => t('editorCoverSubtitle')
const AUTHOR = () => t('editorCoverAuthor')

function today(): string {
  return new Date().toLocaleDateString(DATE_LOCALES[getLang()], {
    year: 'numeric',
    month: 'long',
    day: 'numeric',
  })
}

/** name/paras are both lazily evaluated so they follow the current UI language */
function preset(id: string, nameKey: StringKey, paras: () => CoverPara[]): CoverPreset {
  return {
    id,
    get name() {
      return t(nameKey)
    },
    get paras() {
      return paras()
    },
  }
}

/* palette */
const NAVY = '1F3864'
const BLUE = '2B579A'
const TEAL = '1C7C74'
const GRAY = '595959'
const LIGHT = 'A6A6A6'
const DARK = '262626'
const RED = 'C0504D'
const GOLD = 'BF8F00'

export const COVER_PRESETS: CoverPreset[] = [
  preset('classic', 'editorCoverClassic', () => [
    { spaceBefore: 3600 },
    { text: TITLE(), bold: true, sizeHalfPoints: 72, align: 'center', spaceAfter: 240 },
    {
      text: SUBTITLE(),
      color: GRAY,
      sizeHalfPoints: 28,
      align: 'center',
      borders: 't',
      spaceAfter: 4800,
    },
    { text: AUTHOR(), sizeHalfPoints: 24, align: 'center' },
    { text: today(), color: GRAY, sizeHalfPoints: 22, align: 'center' },
  ]),
  preset('banded', 'editorCoverBanded', () => [
    { spaceBefore: 2400 },
    {
      text: TITLE(),
      bold: true,
      color: 'FFFFFF',
      sizeHalfPoints: 64,
      align: 'center',
      shadingFill: NAVY,
      spaceBefore: 600,
      spaceAfter: 600,
    },
    {
      text: SUBTITLE(),
      color: 'FFFFFF',
      sizeHalfPoints: 26,
      align: 'center',
      shadingFill: BLUE,
      spaceAfter: 5200,
    },
    { text: `${AUTHOR()} · ${today()}`, color: GRAY, sizeHalfPoints: 22, align: 'center' },
  ]),
  preset('boxed', 'editorCoverBoxed', () => [
    { spaceBefore: 3000 },
    {
      text: TITLE(),
      bold: true,
      sizeHalfPoints: 60,
      align: 'center',
      borders: 'tblr',
      spaceBefore: 400,
      spaceAfter: 400,
      indentLeft: 720,
      indentRight: 720,
    },
    {
      text: SUBTITLE(),
      color: GRAY,
      sizeHalfPoints: 26,
      align: 'center',
      spaceBefore: 360,
      spaceAfter: 5200,
    },
    { text: AUTHOR(), sizeHalfPoints: 24, align: 'center' },
    { text: today(), color: GRAY, sizeHalfPoints: 22, align: 'center' },
  ]),
  preset('sideline', 'editorCoverSideline', () => [
    { spaceBefore: 3600 },
    {
      text: TITLE(),
      bold: true,
      color: NAVY,
      sizeHalfPoints: 64,
      align: 'left',
      borders: 'l',
      indentLeft: 360,
      spaceAfter: 200,
    },
    {
      text: SUBTITLE(),
      color: GRAY,
      sizeHalfPoints: 26,
      align: 'left',
      borders: 'l',
      indentLeft: 360,
      spaceAfter: 5600,
    },
    { text: AUTHOR(), sizeHalfPoints: 22, align: 'left' },
    { text: today(), color: GRAY, sizeHalfPoints: 22, align: 'left' },
  ]),
  preset('modern', 'editorCoverModern', () => [
    { spaceBefore: 4200 },
    { text: TITLE(), bold: true, color: DARK, sizeHalfPoints: 80, align: 'left', spaceAfter: 120 },
    { text: SUBTITLE(), color: TEAL, sizeHalfPoints: 28, align: 'left', spaceAfter: 5600 },
    { text: `${AUTHOR()} · ${today()}`, color: LIGHT, sizeHalfPoints: 22, align: 'left' },
  ]),
  preset('elegant', 'editorCoverElegant', () => [
    { spaceBefore: 3600 },
    {
      text: TITLE(),
      italic: true,
      sizeHalfPoints: 64,
      align: 'center',
      borders: 'tb',
      spaceBefore: 300,
      spaceAfter: 300,
      indentLeft: 1080,
      indentRight: 1080,
    },
    {
      text: SUBTITLE(),
      italic: true,
      color: GRAY,
      sizeHalfPoints: 24,
      align: 'center',
      spaceBefore: 400,
      spaceAfter: 5200,
    },
    { text: `${AUTHOR()} · ${today()}`, color: GRAY, sizeHalfPoints: 20, align: 'center' },
  ]),
  preset('minimal', 'editorCoverMinimal', () => [
    { spaceBefore: 5400 },
    { text: TITLE(), color: LIGHT, sizeHalfPoints: 88, align: 'left', spaceAfter: 300 },
    { text: SUBTITLE(), color: GRAY, sizeHalfPoints: 24, align: 'left', spaceAfter: 4800 },
    { text: today(), color: LIGHT, sizeHalfPoints: 20, align: 'left' },
  ]),
  preset('dark', 'editorCoverDark', () => [
    { shadingFill: DARK, spaceBefore: 2000, spaceAfter: 0 },
    {
      text: TITLE(),
      bold: true,
      color: 'FFFFFF',
      sizeHalfPoints: 68,
      align: 'center',
      shadingFill: DARK,
      spaceBefore: 800,
      spaceAfter: 200,
    },
    {
      text: SUBTITLE(),
      color: 'D9D9D9',
      sizeHalfPoints: 26,
      align: 'center',
      shadingFill: DARK,
      spaceAfter: 800,
    },
    { shadingFill: DARK, spaceAfter: 4400 },
    { text: `${AUTHOR()} · ${today()}`, color: GRAY, sizeHalfPoints: 22, align: 'center' },
  ]),
  preset('accent', 'editorCoverAccent', () => [
    { spaceBefore: 3600 },
    { text: TITLE(), bold: true, color: RED, sizeHalfPoints: 68, align: 'left', spaceAfter: 160 },
    {
      text: SUBTITLE(),
      color: GRAY,
      sizeHalfPoints: 26,
      align: 'left',
      borders: 'b',
      spaceAfter: 5600,
    },
    { text: AUTHOR(), sizeHalfPoints: 22, align: 'right' },
    { text: today(), color: GRAY, sizeHalfPoints: 20, align: 'right' },
  ]),
  preset('badge', 'editorCoverBadge', () => [
    { spaceBefore: 3200 },
    { text: '· · ·', color: GOLD, sizeHalfPoints: 28, align: 'center', spaceAfter: 200 },
    {
      text: TITLE(),
      bold: true,
      sizeHalfPoints: 60,
      align: 'center',
      borders: 'tb',
      spaceBefore: 240,
      spaceAfter: 240,
      indentLeft: 1440,
      indentRight: 1440,
    },
    {
      text: SUBTITLE(),
      color: GOLD,
      sizeHalfPoints: 24,
      align: 'center',
      spaceBefore: 300,
      spaceAfter: 5200,
    },
    { text: `${AUTHOR()} · ${today()}`, color: GRAY, sizeHalfPoints: 20, align: 'center' },
  ]),
  preset('facet', 'editorCoverFacet', () => [
    { spaceBefore: 4200 },
    { text: TITLE(), bold: true, color: TEAL, sizeHalfPoints: 64, align: 'right', spaceAfter: 160 },
    {
      text: SUBTITLE(),
      color: GRAY,
      sizeHalfPoints: 26,
      align: 'right',
      borders: 'b',
      spaceAfter: 5600,
    },
    { text: `${AUTHOR()} · ${today()}`, color: LIGHT, sizeHalfPoints: 20, align: 'right' },
  ]),
  preset('annual', 'editorCoverAnnual', () => [
    { spaceBefore: 3000 },
    {
      text: String(new Date().getFullYear()),
      bold: true,
      color: BLUE,
      sizeHalfPoints: 144,
      align: 'left',
      spaceAfter: 0,
    },
    {
      text: TITLE(),
      bold: true,
      sizeHalfPoints: 56,
      align: 'left',
      borders: 't',
      spaceBefore: 200,
      spaceAfter: 120,
    },
    { text: SUBTITLE(), color: GRAY, sizeHalfPoints: 24, align: 'left', spaceAfter: 4800 },
    { text: AUTHOR(), color: GRAY, sizeHalfPoints: 22, align: 'left' },
  ]),
]

function paraNode(p: CoverPara): PmNode {
  const marks: Array<{ type: string; attrs?: Record<string, unknown> }> = []
  if (p.bold) marks.push({ type: 'bold' })
  if (p.italic) marks.push({ type: 'italic' })
  if (p.color || p.sizeHalfPoints) {
    marks.push({
      type: 'docTextStyle',
      attrs: {
        color: p.color ?? null,
        sizeHalfPoints: p.sizeHalfPoints ?? null,
        font: null,
        highlight: null,
        vertAlign: null,
      },
    })
  }
  return {
    type: 'docParagraph',
    attrs: {
      docxIndex: null,
      aiChanged: false,
      align: p.align ?? null,
      spaceBefore: p.spaceBefore ?? null,
      spaceAfter: p.spaceAfter ?? null,
      shadingFill: p.shadingFill ?? null,
      borders: p.borders ?? null,
      indentLeft: p.indentLeft ?? null,
      indentRight: p.indentRight ?? null,
    },
    ...(p.text ? { content: [{ type: 'text', text: p.text, marks }] } : {}),
  }
}

/** Cover paragraphs + one trailing page-break empty paragraph (pushes existing content to page 2) */
export function buildCoverNodes(preset: CoverPreset): PmNode[] {
  const paras = preset.paras.map(paraNode)
  paras[0].attrs = { ...paras[0].attrs, hiddenBookmarks: [COVER_START_MARK] }
  return [
    ...paras,
    {
      type: 'docParagraph',
      attrs: {
        docxIndex: null,
        aiChanged: false,
        pageBreakBefore: true,
        hiddenBookmarks: [COVER_END_MARK],
      },
    },
  ]
}

/** document range of the gallery cover: first start-marked block through the first end-marked block */
export function coverPageRange(doc: ProseMirrorNode): { from: number; to: number } | null {
  let from = -1
  let to = -1
  let pos = 0
  for (let i = 0; i < doc.childCount; i++) {
    const child = doc.child(i)
    const marks = child.attrs.hiddenBookmarks as string[] | null | undefined
    if (from < 0 && marks?.includes(COVER_START_MARK)) from = pos
    if (to < 0 && marks?.includes(COVER_END_MARK)) to = pos + child.nodeSize
    pos += child.nodeSize
  }
  return from >= 0 && to > from ? { from, to } : null
}

/** Insert the cover at the very start of the document (Word behavior: cover is always page 1, a previous gallery cover is replaced) */
export function insertCoverPage(editor: Editor, preset: CoverPreset): void {
  const existing = coverPageRange(editor.state.doc)
  const chain = editor.chain().focus()
  if (existing) chain.deleteRange(existing)
  chain.insertContentAt(0, buildCoverNodes(preset)).run()
}

export function removeCoverPage(editor: Editor): boolean {
  const range = coverPageRange(editor.state.doc)
  if (!range) return false
  return editor.chain().focus().deleteRange(range).run()
}
