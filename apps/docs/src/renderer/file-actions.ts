/**
 * Document lifecycle: open/new (full state reset from a parsed docx), the
 * save pipeline (PM doc → save plan → docx bytes → reparse), and PDF export.
 * Extracted from App.tsx; the App component passes a FileActionContext built
 * fresh per call (wrapped in useCallback with the original dependency lists)
 * so state never goes stale.
 */
import type { Editor } from '@tiptap/core'
import { history } from '@tiptap/pm/history'
import {
  applyPageNumType,
  applySectionSettings,
  applySectionStartType,
  buildBlankDocx,
  findChartWorkbookPath,
  parseChartPartXml,
  parseDocx,
  patchChartPartXml,
  patchChartWorkbookXlsxBase64,
  readDocxPartBase64,
  readPageColor,
  readSections,
  readSectionSettings,
  saveDocx,
  type CommentInfo,
  type DocProtection,
  type HeaderFooter,
  type NoteInfo,
  type SectionInfo,
  type SectionSettings,
  type SourceInfo,
  type StyleUpsert,
  type ThemeColors,
  type ThemeFonts,
} from '@genoffice/docx-engine'
import type { Dispatch, SetStateAction } from 'react'
import type { OpenFileResult } from '../shared/ipc'
import {
  hfVariantsFromParsed,
  type DocState,
  type HfVariantKey,
  type HfVariantsState,
  type HfView,
  type PendingNumbering,
} from './doc-state'
import { docStyleCss } from './doc-style-css'
import type { CompareEntry } from './editor/compare'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from './editor/convert'
import {
  annotationsFromParsed,
  buildInkImages,
  type InkAnnotation,
  type InkTool,
} from './editor/ink'
import { t, getLang } from './i18n/locale'
import { isBlankDocument } from './ai/protocol'
import { isDocDirty } from './doc-dirty'
import { createSaveSerializer } from './save-until-persisted'
import { checkMissingFonts, collectDocFonts } from './font-check'
import { defaultEastAsiaFontFor } from './font-list'
import { hasPrintableHeaderFooter } from './pagination'
import { showToast } from './components/toast-bus'

/** The App state the file actions need; built fresh per call. */
export interface FileActionContext {
  editor: Editor | null
  doc: DocState | null
  dirtyRef: { current: boolean }
  saveInFlightRef: { current: boolean }
  saveIncompleteRef: { current: boolean }
  pendingMixedExportRef: { current: boolean | string }
  setStatus: (status: string) => void
  setRecent: (paths: string[]) => void
  setShowAi: (show: boolean) => void
  setDoc: Dispatch<SetStateAction<DocState | null>>
  setAiPanelKey: Dispatch<SetStateAction<number>>
  setDocCss: (css: string) => void
  setShowPagePreview: (show: boolean) => void
  section: SectionSettings | null
  sectionDirty: boolean
  sections: SectionInfo[]
  sectionsDirty: number[]
  trailingStartType: SectionInfo['startType'] | null
  setSection: (value: SectionSettings | null) => void
  setSections: (value: SectionInfo[]) => void
  setSectionDirty: (dirty: boolean) => void
  setSectionsDirty: (value: number[]) => void
  setTrailingStartType: (value: SectionInfo['startType'] | null) => void
  pageColor: string | null
  pageColorDirty: boolean
  setPageColor: (value: string | null) => void
  setPageColorDirty: (dirty: boolean) => void
  header: HeaderFooter | null
  headerDirty: boolean
  footer: HeaderFooter | null
  footerDirty: boolean
  setHeader: (value: HeaderFooter | null) => void
  setHeaderDirty: (dirty: boolean) => void
  setFooter: (value: HeaderFooter | null) => void
  setFooterDirty: (dirty: boolean) => void
  hfVariants: HfVariantsState
  hfVariantsDirty: HfVariantKey[]
  setHfVariants: (value: HfVariantsState) => void
  setHfVariantsDirty: (value: HfVariantKey[]) => void
  sectionHfEdits: Record<string, HeaderFooter>
  setSectionHfEdits: (value: Record<string, HeaderFooter>) => void
  titlePg: boolean
  titlePgDirty: boolean
  evenOddHf: boolean
  evenOddHfDirty: boolean
  setTitlePg: (value: boolean) => void
  setTitlePgDirty: (dirty: boolean) => void
  setEvenOddHf: (value: boolean) => void
  setEvenOddHfDirty: (dirty: boolean) => void
  setHfView: (view: HfView) => void
  pgNumEdit: { fmt?: string; start?: number } | null
  pgNumDirtySections: number[]
  setPgNumEdit: (value: { fmt?: string; start?: number } | null) => void
  setPgNumDirtySections: (value: number[]) => void
  pendingNumbering: PendingNumbering
  numberingDirty: boolean
  setPendingNumbering: (value: PendingNumbering) => void
  styleUpserts: Record<string, StyleUpsert>
  setStyleUpserts: (value: Record<string, StyleUpsert>) => void
  comments: CommentInfo[]
  commentsDirty: boolean
  setComments: (value: CommentInfo[]) => void
  setCommentsDirty: (dirty: boolean) => void
  setShowComments: (show: boolean) => void
  setCommentComposing: (composing: boolean) => void
  watermark: string | null
  watermarkDirty: boolean
  setWatermark: (value: string | null) => void
  setWatermarkDirty: (dirty: boolean) => void
  inkAnnotations: InkAnnotation[]
  inksDirty: boolean
  setInkAnnotations: (value: InkAnnotation[]) => void
  setInksDirty: (dirty: boolean) => void
  setInkTool: (tool: InkTool) => void
  footnotes: NoteInfo[]
  endnotes: NoteInfo[]
  notesDirty: boolean
  setFootnotes: (value: NoteInfo[]) => void
  setEndnotes: (value: NoteInfo[]) => void
  setNotesDirty: (dirty: boolean) => void
  sources: SourceInfo[]
  sourcesDirty: boolean
  setSources: (value: SourceInfo[]) => void
  setSourcesDirty: (dirty: boolean) => void
  themeFonts: ThemeFonts | null
  themeFontsDirty: boolean
  themeColors: ThemeColors | null
  themeColorsDirty: boolean
  setThemeFonts: (value: ThemeFonts | null) => void
  setThemeFontsDirty: (dirty: boolean) => void
  setThemeColors: (value: ThemeColors | null) => void
  setThemeColorsDirty: (dirty: boolean) => void
  setTrackChanges: (on: boolean) => void
  protection: DocProtection | null
  protectionDirty: boolean
  setProtection: (value: DocProtection | null) => void
  setProtectionDirty: (dirty: boolean) => void
  setCompareResult: (value: { otherName: string; entries: CompareEntry[] } | null) => void
}

/** Drop the undo stack: undo across an open/reparse boundary resurrects stale
 *  docxIndex anchors (corrupting the next save) or the previous document. */
function resetEditorHistory(editor: Editor): void {
  const plugin = editor.state.plugins.find((p) =>
    String((p as unknown as { key: string }).key).startsWith('history$'),
  )
  if (!plugin) return
  editor.unregisterPlugin('history')
  editor.registerPlugin(history((plugin.spec as { config?: object }).config))
}

export async function loadFile(
  ctx: FileActionContext,
  result: OpenFileResult | null,
): Promise<void> {
  if (!result || !ctx.editor) return
  try {
    const parsed = await parseDocx(new Uint8Array(result.data))
    ctx.editor.storage.listNumbering.defs = parsed.numbering
    ctx.editor.commands.setContent(blocksToPmDoc(parsed.blocks) as never)
    resetEditorHistory(ctx.editor)
    noteDocumentSwapped()
    ctx.setDoc({ parsed, filePath: result.path, fileName: result.name, hash: result.hash })
    ctx.setAiPanelKey((k) => k + 1)
    ctx.setDocCss(docStyleCss(parsed))
    ctx.setSection(readSectionSettings(parsed))
    ctx.setSections(readSections(parsed))
    ctx.setSectionDirty(false)
    ctx.setPageColor(readPageColor(parsed))
    ctx.setPageColorDirty(false)
    ctx.setHeader(
      parsed.headerText || parsed.headerHasPageNumber || parsed.headerParas?.length
        ? {
            text: parsed.headerText ?? '',
            pageNumber: parsed.headerHasPageNumber,
            paras: parsed.headerParas ?? undefined,
          }
        : null,
    )
    ctx.setHeaderDirty(false)
    ctx.setFooter(
      parsed.footerText || parsed.footerHasPageNumber || parsed.footerParas?.length
        ? {
            text: parsed.footerText ?? '',
            pageNumber: parsed.footerHasPageNumber,
            paras: parsed.footerParas ?? undefined,
          }
        : null,
    )
    ctx.setFooterDirty(false)
    ctx.setHfVariants(hfVariantsFromParsed(parsed))
    ctx.setHfVariantsDirty([])
    ctx.setTitlePg(parsed.titlePg ?? false)
    ctx.setTitlePgDirty(false)
    ctx.setEvenOddHf(parsed.evenAndOddHeaders ?? false)
    ctx.setEvenOddHfDirty(false)
    ctx.setHfView('default')
    ctx.setShowComments(false)
    ctx.setComments(parsed.comments)
    ctx.setCommentsDirty(false)
    ctx.setWatermark(parsed.watermarkText ?? null)
    ctx.setWatermarkDirty(false)
    ctx.setInkAnnotations(annotationsFromParsed(parsed.inks))
    ctx.setInksDirty(false)
    ctx.setInkTool('select')
    ctx.setFootnotes(parsed.footnotes)
    ctx.setEndnotes(parsed.endnotes)
    ctx.setNotesDirty(false)
    ctx.setSources(parsed.sources)
    ctx.setSourcesDirty(false)
    ctx.setThemeFonts(parsed.themeFonts ?? null)
    ctx.setThemeFontsDirty(false)
    ctx.setThemeColors(parsed.themeColors ?? null)
    ctx.setThemeColorsDirty(false)
    ctx.setCommentComposing(false)
    ctx.setTrackChanges(false)
    ctx.setProtection(parsed.protection)
    ctx.setProtectionDirty(false)
    ctx.setCompareResult(null)
    ctx.dirtyRef.current = false
    const missing = checkMissingFonts(collectDocFonts(parsed))
    const verticalText = readSections(parsed).some((s) => s.settings.textDirection)
    if (verticalText) {
      // visible degradation: vertical writing is not rendered yet, never silently
      ctx.setStatus(t('appVerticalTextNotice'))
    } else if (missing.length > 0) {
      const names = missing
        .slice(0, 3)
        .map((m) => (m.substitute ? `${m.name} → ${m.substitute}` : m.name))
        .join(', ')
      ctx.setStatus(t('appFontsMissing', { names: missing.length > 3 ? `${names}…` : names }))
    } else {
      ctx.setStatus(t('appOpenedFile', { name: result.name }))
    }
    void window.desktop.getRecentFiles().then(ctx.setRecent)
  } catch (err) {
    ctx.setStatus(t('appOpenFailed', { error: String(err) }))
  }
}

/** new document from the built-in blank template (AI can then generate into it) */
export async function newFile(ctx: FileActionContext): Promise<boolean | undefined> {
  if (!ctx.editor) return
  try {
    const bytes = await buildBlankDocx({ eastAsiaFont: defaultEastAsiaFontFor(getLang()) })
    const parsed = await parseDocx(bytes)
    ctx.editor.storage.listNumbering.defs = parsed.numbering
    ctx.editor.commands.setContent(blocksToPmDoc(parsed.blocks) as never)
    resetEditorHistory(ctx.editor)
    noteDocumentSwapped()
    ctx.setDoc({ parsed, filePath: null, fileName: t('appUntitledDocx'), hash: '', isBlank: true })
    ctx.setAiPanelKey((k) => k + 1)
    ctx.setDocCss(docStyleCss(parsed))
    ctx.setSection(readSectionSettings(parsed))
    ctx.setSections(readSections(parsed))
    ctx.setSectionDirty(false)
    ctx.setPageColor(readPageColor(parsed))
    ctx.setPageColorDirty(false)
    ctx.setHeader(null)
    ctx.setHeaderDirty(false)
    ctx.setFooter(null)
    ctx.setFooterDirty(false)
    ctx.setShowComments(false)
    ctx.setComments([])
    ctx.setCommentsDirty(false)
    ctx.setWatermark(null)
    ctx.setWatermarkDirty(false)
    ctx.setInkAnnotations([])
    ctx.setInksDirty(false)
    ctx.setInkTool('select')
    ctx.setFootnotes(parsed.footnotes)
    ctx.setEndnotes(parsed.endnotes)
    ctx.setNotesDirty(false)
    ctx.setSources([])
    ctx.setSourcesDirty(false)
    ctx.setThemeFonts(parsed.themeFonts ?? null)
    ctx.setThemeFontsDirty(false)
    ctx.setThemeColors(parsed.themeColors ?? null)
    ctx.setThemeColorsDirty(false)
    ctx.setCommentComposing(false)
    ctx.setTrackChanges(false)
    ctx.setProtection(null)
    ctx.setProtectionDirty(false)
    ctx.setCompareResult(null)
    ctx.dirtyRef.current = false
    ctx.setShowAi(true)
    ctx.setStatus(t('appNewDocCreated'))
    return true
  } catch (err) {
    ctx.setStatus(t('appNewFailed', { error: String(err) }))
    return false
  }
}

/** Plain text of a PM node's inline content. */
function pmNodeText(node: PmNode): string {
  if (node.text) return node.text
  return (node.content ?? []).map(pmNodeText).join('')
}

/** Sanitize a heading into a safe filename base: strip illegal path chars, collapse whitespace, cap length; null if invalid. (Mirrors slides' draft naming.) */
function sanitizeFileBaseName(raw: string): string | null {
  const cleaned = raw
    // eslint-disable-next-line no-control-regex -- stripping control chars is the point here
    .replace(/[\\/:*?"<>|\u0000-\u001f]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    // Strip leading/trailing dots (Windows disallows a trailing dot; a hidden-file prefix is meaningless here)
    .replace(/^\.+|\.+$/g, '')
    .trim()
  if (!cleaned) return null
  return cleaned.length > 40 ? cleaned.slice(0, 40).trim() : cleaned
}

/**
 * Docs' analog of slides' deckName: a still-untitled document takes its file name
 * from the first heading, so the first silent save (and with it the shell tab
 * title) picks up a content-derived name instead of "Untitled.docx".
 */
function deriveAutoFileName(editor: Editor): string | null {
  const body = (editor.getJSON() as PmNode).content ?? []
  for (const node of body) {
    if (node.type !== 'docHeading') continue
    const base = sanitizeFileBaseName(pmNodeText(node))
    if (base) return `${base}.docx`
  }
  return null
}

/**
 * Serialize the current editor/document state to .docx bytes — the shared
 * serialization half of save(); no dialogs, no state changes. Also used for
 * crash-recovery copies.
 */
export async function buildDocBytes(ctx: FileActionContext): Promise<Uint8Array | null> {
  const { doc, editor } = ctx
  if (!doc || !editor) return null
  const plan = pmDocToSavePlan(editor.getJSON() as PmNode, doc.parsed.blocks)
  // chart data edits patch the chart's own zip part, not the body XML
  const partXml: Record<string, string> = {}
  const partBinary: Record<string, string> = {}
  for (const { partPath, patch } of plan.chartPatches) {
    const originalPart = doc.parsed.extras.chartParts[partPath]
    if (originalPart) {
      const patchedXml = patchChartPartXml(originalPart, patch)
      partXml[partPath] = patchedXml
      // Also update the embedded workbook so Word's "Edit Data" shows correct data
      const wbPath = await findChartWorkbookPath(doc.parsed.internal.originalBytes, partPath)
      if (wbPath) {
        const existingBase64 = await readDocxPartBase64(doc.parsed.internal.originalBytes, wbPath)
        if (existingBase64) {
          const display = parseChartPartXml(patchedXml, partPath)
          if (display) {
            const namedSeries = display.series.map((s, i) => ({
              name: s.name ?? `Series${i + 1}`,
              values: s.values as (number | null)[],
            }))
            const updated = await patchChartWorkbookXlsxBase64(
              existingBase64,
              display.categories,
              namedSeries,
            )
            if (updated) partBinary[wbPath] = updated
          }
        }
      }
    }
  }
  // Ink must be passed whenever any annotation exists (not only when
  // dirty): a regenerated anchor paragraph loses its ink run, and the
  // engine re-injects the full layer from this list.
  const inks =
    ctx.inksDirty || ctx.inkAnnotations.length > 0
      ? buildInkImages(ctx.inkAnnotations, plan.saveBlockIndexByDocx)
      : undefined
  // page-setup edits for non-final sections: rewrite the sectPr inside their section-break paragraphs (the final section uses options.section)
  let saveBlocks = plan.saveBlocks
  const dirtySectionIdxs = [...new Set([...ctx.sectionsDirty, ...ctx.pgNumDirtySections])]
  if (dirtySectionIdxs.length > 0) {
    const rewrites = new Map<number, string>()
    for (const si of dirtySectionIdxs) {
      const sec = ctx.sections[si]
      if (!sec || si === ctx.sections.length - 1) continue
      const blk = doc.parsed.blocks.find((b) => b.docxIndex === sec.lastBlockIndex)
      if (!blk?.originalXml || !sec.sectPrXml) continue
      let sectPr = applySectionSettings(sec.sectPrXml, sec.settings)
      sectPr = applySectionStartType(sectPr, sec.startType)
      // touch w:pgNumType only when the page-number format was edited (avoids dropping unmodeled attrs like chapStyle)
      if (ctx.pgNumDirtySections.includes(si)) {
        sectPr = applyPageNumType(sectPr, sec.pageNumberFmt, sec.pageNumberStart)
      }
      rewrites.set(sec.lastBlockIndex, blk.originalXml.replace(sec.sectPrXml, sectPr))
    }
    saveBlocks = saveBlocks.map((fb) =>
      fb.kind === 'original' && rewrites.has(fb.docxIndex)
        ? { kind: 'xml' as const, xml: rewrites.get(fb.docxIndex)!, docxIndex: fb.docxIndex }
        : fb,
    )
  }
  // header/footer edits for non-final sections: the engine writes parts/references per section
  const sectionHf = Object.entries(ctx.sectionHfEdits).map(([key, hf]) => {
    const [lastBlockIndex, kind] = key.split(':')
    return { lastBlockIndex: Number(lastBlockIndex), kind: kind as 'header' | 'footer', hf }
  })
  const bytes = await saveDocx(doc.parsed, saveBlocks, {
    section: ctx.sectionDirty && ctx.section ? ctx.section : undefined,
    sectionStartType: ctx.trailingStartType ?? undefined,
    pgNumType: ctx.pgNumEdit ?? undefined,
    sectionHf: sectionHf.length > 0 ? sectionHf : undefined,
    numbering: ctx.numberingDirty ? ctx.pendingNumbering : undefined,
    styleUpserts:
      Object.keys(ctx.styleUpserts).length > 0 ? Object.values(ctx.styleUpserts) : undefined,
    pageColor: ctx.pageColorDirty ? ctx.pageColor : undefined,
    header: ctx.headerDirty && ctx.header ? ctx.header : undefined,
    footer: ctx.footerDirty && ctx.footer ? ctx.footer : undefined,
    headerFirst:
      ctx.hfVariantsDirty.includes('headerFirst') && ctx.hfVariants.headerFirst
        ? ctx.hfVariants.headerFirst
        : undefined,
    footerFirst:
      ctx.hfVariantsDirty.includes('footerFirst') && ctx.hfVariants.footerFirst
        ? ctx.hfVariants.footerFirst
        : undefined,
    headerEven:
      ctx.hfVariantsDirty.includes('headerEven') && ctx.hfVariants.headerEven
        ? ctx.hfVariants.headerEven
        : undefined,
    footerEven:
      ctx.hfVariantsDirty.includes('footerEven') && ctx.hfVariants.footerEven
        ? ctx.hfVariants.footerEven
        : undefined,
    titlePg: ctx.titlePgDirty ? ctx.titlePg : undefined,
    evenAndOddHeaders: ctx.evenOddHfDirty ? ctx.evenOddHf : undefined,
    partXml: Object.keys(partXml).length > 0 ? partXml : undefined,
    partBinary: Object.keys(partBinary).length > 0 ? partBinary : undefined,
    comments: ctx.commentsDirty ? ctx.comments : undefined,
    protection: ctx.protectionDirty ? ctx.protection : undefined,
    inks,
    watermark: ctx.watermarkDirty ? ctx.watermark : undefined,
    footnotes: ctx.notesDirty ? ctx.footnotes : undefined,
    endnotes: ctx.notesDirty ? ctx.endnotes : undefined,
    sources: ctx.sourcesDirty ? ctx.sources : undefined,
    themeFonts: ctx.themeFontsDirty && ctx.themeFonts ? ctx.themeFonts : undefined,
    themeColors: ctx.themeColorsDirty && ctx.themeColors ? ctx.themeColors : undefined,
  })
  return bytes
}

/**
 * Crash-recovery copy: serialize the dirty document and hand the
 * bytes to the main process, which stores them under userData. Best-effort —
 * a failure only means this tick's copy is skipped.
 *
 * A dirty document that never got a path can't have a recovery copy (the main
 * process only accepts allowlisted paths), so it is silently persisted into the
 * default folder first (same save-new path as the first manual save); every
 * later tick then covers it like any opened file.
 */
export async function writeRecoveryCopy(ctx: FileActionContext): Promise<void> {
  const { doc, editor } = ctx
  if (!doc || !editor || ctx.saveInFlightRef.current || !isDocDirty(ctx)) return
  if (!doc.filePath) {
    if (isBlankDocument(editor)) return
    if (editor.view.composing) return
    const active = document.activeElement as HTMLElement | null
    if (active?.closest('td[contenteditable], .doc-textbox')) return
    await save(ctx, false, true)
    return
  }
  try {
    const bytes = await buildDocBytes(ctx)
    if (!bytes) return
    const buffer = bytes.buffer.slice(
      bytes.byteOffset,
      bytes.byteOffset + bytes.byteLength,
    ) as ArrayBuffer
    await window.desktop.writeRecoveryCopy(doc.filePath, buffer)
  } catch {
    /* best-effort */
  }
}

const runSerializedSave = createSaveSerializer()

/**
 * Path assigned by the first save of a still-pathless document (silent save-new
 * or Save As). A queued save whose ctx snapshot predates that first save still
 * sees `doc.filePath === null`; without this it would re-run the save-new path
 * and create a duplicate file. Reset whenever a different document is loaded.
 */
let pathlessDocSavedPath: string | null = null

export function noteDocumentSwapped(): void {
  pathlessDocSavedPath = null
}

export function save(ctx: FileActionContext, saveAs: boolean, auto = false): Promise<boolean> {
  // A save arriving mid-flight waits for the current one instead of failing.
  // Reuse the finished pass only when it left nothing behind — judged by the
  // composite dirty check (header/section/theme edits do not set dirtyRef), plus
  // the raced-with-typing flag. A pass that left anything runs its own pass;
  // saveOnce resolves a stale pathless snapshot via pathlessDocSavedPath, so
  // the retry can no longer create a duplicate file.
  return runSerializedSave(
    () => saveOnce(ctx, saveAs, auto),
    () => !saveAs && !ctx.saveIncompleteRef.current && !isDocDirty(ctx),
  )
}

async function saveOnce(ctx: FileActionContext, saveAs: boolean, auto: boolean): Promise<boolean> {
  const { doc, editor } = ctx
  if (!doc || !editor) return false
  ctx.saveInFlightRef.current = true
  ctx.saveIncompleteRef.current = false
  try {
    // flush pending in-place table cell / textbox edits into the PM doc first
    window.dispatchEvent(new Event('ai-docs-commit-tables'))
    // identity snapshot: detects edits that arrive while the save is in flight
    const docSnapshot = editor.state.doc
    const selectionPos = editor.state.selection.from
    const bytes = await buildDocBytes(ctx)
    if (!bytes) return false
    const buffer = bytes.buffer.slice(
      bytes.byteOffset,
      bytes.byteOffset + bytes.byteLength,
    ) as ArrayBuffer
    // a pathless snapshot may belong to a document that an earlier queued pass
    // already landed on disk — overwrite that file instead of creating another
    let savedPath = doc.filePath ?? pathlessDocSavedPath
    if (saveAs || !savedPath) {
      // A never-saved document still called "Untitled" gets a name derived from its first heading
      const autoName =
        !doc.filePath && doc.fileName === t('appUntitledDocx') ? deriveAutoFileName(editor) : null
      // Save As keeps the dialog; a new document's first save lands silently in the default folder
      const result = saveAs
        ? await window.desktop.saveDocxAs(autoName ?? doc.fileName, buffer)
        : await window.desktop.saveDocxNew(autoName ?? doc.fileName, buffer)
      if (!result.ok) {
        if (result.error) {
          ctx.setStatus(t('appSaveFailed', { error: result.error }))
          if (!auto) showToast(t('appSaveFailed', { error: result.error }), 'error')
        }
        return false
      }
      savedPath = result.path!
      if (!doc.filePath) pathlessDocSavedPath = savedPath
    } else {
      const result = await window.desktop.saveDocx(savedPath, buffer, auto)
      if (!result.ok) {
        // external-modified: the main process already prompted (or the autosave
        // deferred to a manual save) — stay dirty, no second dialog/error banner
        if (result.reason !== 'external-modified') {
          ctx.setStatus(t('appSaveFailed', { error: result.error ?? '' }))
          if (!auto) showToast(t('appSaveFailed', { error: result.error ?? '' }), 'error')
        }
        return false
      }
    }
    if (editor.state.doc !== docSnapshot) {
      // The user kept editing while the save was in flight. Replacing the
      // editor content with the reparsed (pre-edit) snapshot would silently
      // drop those keystrokes, so keep the live editor + parsed state as-is
      // and leave the document dirty; the next save persists the newer edits.
      ctx.saveIncompleteRef.current = true
      ctx.setDoc((prev) =>
        prev
          ? {
              ...prev,
              filePath: savedPath,
              fileName: savedPath?.split(/[\\/]/).pop() ?? prev.fileName,
            }
          : prev,
      )
      ctx.setStatus(
        auto ? t('appAutoSavedAt', { time: new Date().toLocaleTimeString() }) : t('appSaved'),
      )
      // No success toast: the doc is still dirty (edits raced the save), and the
      // close-guard's saveUntilPersisted retries would repeat it on every pass —
      // the converging complete save below toasts once.
      return true
    }
    // Reload from saved bytes so docxIndex anchors point at the new file.
    const reparsed = await parseDocx(bytes)
    editor.storage.listNumbering.defs = reparsed.numbering
    const rebasedPm = blocksToPmDoc(reparsed.blocks)
    let unchanged = false
    try {
      unchanged = editor.state.doc.eq(editor.schema.nodeFromJSON(rebasedPm))
    } catch {
      /* unrepresentable → rewrite */
    }
    // Equal doc: skip the rewrite so undo history, caret and scroll survive.
    if (!unchanged) {
      editor.commands.setContent(rebasedPm as never)
      resetEditorHistory(editor)
      const chain = editor
        .chain()
        .setTextSelection(Math.min(selectionPos, editor.state.doc.content.size))
      if (!auto) chain.scrollIntoView()
      chain.run()
    }
    ctx.setDocCss(docStyleCss(reparsed))
    ctx.setDoc((prev) =>
      prev
        ? {
            ...prev,
            parsed: reparsed,
            filePath: savedPath,
            fileName: savedPath?.split(/[\\/]/).pop() ?? prev.fileName,
          }
        : prev,
    )
    ctx.setSection(readSectionSettings(reparsed))
    ctx.setSections(readSections(reparsed))
    ctx.setSectionDirty(false)
    ctx.setSectionsDirty([])
    ctx.setTrailingStartType(null)
    ctx.setPageColor(readPageColor(reparsed))
    ctx.setPageColorDirty(false)
    ctx.setHeader(
      reparsed.headerText || reparsed.headerHasPageNumber || reparsed.headerParas?.length
        ? {
            text: reparsed.headerText ?? '',
            pageNumber: reparsed.headerHasPageNumber,
            paras: reparsed.headerParas ?? undefined,
          }
        : null,
    )
    ctx.setHeaderDirty(false)
    ctx.setFooter(
      reparsed.footerText || reparsed.footerHasPageNumber || reparsed.footerParas?.length
        ? {
            text: reparsed.footerText ?? '',
            pageNumber: reparsed.footerHasPageNumber,
            paras: reparsed.footerParas ?? undefined,
          }
        : null,
    )
    ctx.setFooterDirty(false)
    ctx.setHfVariants(hfVariantsFromParsed(reparsed))
    ctx.setSectionHfEdits({})
    ctx.setPgNumEdit(null)
    ctx.setPgNumDirtySections([])
    ctx.setPendingNumbering({ newDefs: [], restartNums: [] })
    ctx.setStyleUpserts({})
    ctx.setHfVariantsDirty([])
    ctx.setTitlePg(reparsed.titlePg ?? false)
    ctx.setTitlePgDirty(false)
    ctx.setEvenOddHf(reparsed.evenAndOddHeaders ?? false)
    ctx.setEvenOddHfDirty(false)
    ctx.setComments(reparsed.comments)
    ctx.setCommentsDirty(false)
    ctx.setWatermark(reparsed.watermarkText ?? null)
    ctx.setWatermarkDirty(false)
    ctx.setInkAnnotations(annotationsFromParsed(reparsed.inks))
    ctx.setInksDirty(false)
    ctx.setFootnotes(reparsed.footnotes)
    ctx.setEndnotes(reparsed.endnotes)
    ctx.setNotesDirty(false)
    ctx.setSources(reparsed.sources)
    ctx.setSourcesDirty(false)
    ctx.setThemeFonts(reparsed.themeFonts ?? null)
    ctx.setThemeFontsDirty(false)
    ctx.setThemeColors(reparsed.themeColors ?? null)
    ctx.setThemeColorsDirty(false)
    ctx.setProtection(reparsed.protection)
    ctx.setProtectionDirty(false)
    ctx.dirtyRef.current = false
    ctx.setStatus(
      auto ? t('appAutoSavedAt', { time: new Date().toLocaleTimeString() }) : t('appSaved'),
    )
    if (!auto) showToast(t('appSaved'))
    return true
  } catch (err) {
    ctx.setStatus(t('appSaveFailed', { error: String(err) }))
    if (!auto) showToast(t('appSaveFailed', { error: String(err) }), 'error')
    return false
  } finally {
    ctx.saveInFlightRef.current = false
  }
}

export async function exportPdf(ctx: FileActionContext, outPath?: string): Promise<void> {
  const { doc } = ctx
  if (!doc) return
  ctx.setStatus(t('appExportingPdf'))
  // with the pagination preview open: group by actual pv-page size. Mixed paper →
  // print group by group (other pages hidden for printing), then merge in page order
  // with pdf-lib; each page keeps its section's paper size
  const pvPages = [...document.querySelectorAll('.pv-page')] as HTMLElement[]
  if (pvPages.length > 0) {
    const pxToTwips = (px: number) => Math.round((px / 96) * 1440)
    const groups: Array<{ w: number; h: number; from: number; to: number }> = []
    pvPages.forEach((page, i) => {
      const w = pxToTwips(parseFloat(page.style.width))
      const h = pxToTwips(parseFloat(page.style.height))
      const last = groups[groups.length - 1]
      if (last && last.w === w && last.h === h) last.to = i
      else groups.push({ w, h, from: i, to: i })
    })
    if (groups.length > 1) {
      const parts: string[] = []
      try {
        for (const g of groups) {
          pvPages.forEach((page, i) =>
            page.classList.toggle('pv-print-skip', i < g.from || i > g.to),
          )
          const part = await window.desktop.printPdfBuffer(g.w, g.h)
          if (!part.ok || !part.base64) {
            ctx.setStatus(
              t('appExportPdfFailed', { error: part.error ?? t('appPrintGroupFailed') }),
            )
            return
          }
          parts.push(part.base64)
        }
      } finally {
        pvPages.forEach((page) => page.classList.remove('pv-print-skip'))
      }
      const result = await window.desktop.saveMergedPdf(doc.fileName, parts, outPath)
      ctx.setStatus(
        result.ok
          ? t('appExportedPdfMixed', { path: result.path ?? '', n: groups.length })
          : result.error
            ? t('appExportPdfFailed', { error: result.error })
            : t('appExportPdfCanceled'),
      )
      return
    }
    // preview open but uniform paper: single export at the preview size
    const g = groups[0]
    if (g) {
      const result = await window.desktop.exportPdf(doc.fileName, g.w, g.h, outPath)
      ctx.setStatus(
        result.ok
          ? t('appExportedPdf', { path: result.path ?? '' })
          : result.error
            ? t('appExportPdfFailed', { error: result.error })
            : t('appExportPdfCanceled'),
      )
      return
    }
  }
  // preview closed: mixed paper auto-opens the preview and uses the merge path; uniform paper exports directly
  const sizeKey = (w: number, h: number) => `${w}x${h}`
  const counts = new Map<string, { w: number; h: number; n: number }>()
  const list =
    ctx.sections.length > 0
      ? ctx.sections.map((sec) => sec.settings)
      : ctx.section
        ? [ctx.section]
        : []
  for (const st of list) {
    const key = sizeKey(st.pageWidth, st.pageHeight)
    const cur = counts.get(key) ?? { w: st.pageWidth, h: st.pageHeight, n: 0 }
    cur.n += 1
    counts.set(key, cur)
  }
  // headers/footers exist once on the edit canvas (not once per page), so direct
  // print would show them on the last page only — force the preview-merge path too
  const mixedPaper = counts.size > 1
  if (
    mixedPaper ||
    hasPrintableHeaderFooter({
      edited: [
        ctx.header,
        ctx.footer,
        ...(ctx.titlePg ? [ctx.hfVariants.headerFirst, ctx.hfVariants.footerFirst] : []),
        ...(ctx.evenOddHf ? [ctx.hfVariants.headerEven, ctx.hfVariants.footerEven] : []),
        ...Object.values(ctx.sectionHfEdits),
      ],
      sections: ctx.sections,
      hfParts: doc.parsed.hfParts ?? undefined,
      evenOddHf: ctx.evenOddHf,
    })
  ) {
    ctx.setShowPagePreview(true)
    ctx.pendingMixedExportRef.current = outPath ?? true
    if (mixedPaper) ctx.setStatus(t('appMixedExportOpening'))
    return
  }
  const major = [...counts.values()][0]
  const result = await window.desktop.exportPdf(
    doc.fileName,
    major?.w ?? ctx.section?.pageWidth ?? 12240,
    major?.h ?? ctx.section?.pageHeight ?? 15840,
    outPath,
  )
  ctx.setStatus(
    result.ok
      ? t('appExportedPdf', { path: result.path ?? '' })
      : result.error
        ? t('appExportPdfFailed', { error: result.error })
        : t('appExportPdfCanceled'),
  )
}
