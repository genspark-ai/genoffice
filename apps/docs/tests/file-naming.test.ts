import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { TextSelection } from '@tiptap/pm/state'
import { parseDocx } from '@genoffice/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { DocRedaction } from '../src/renderer/editor/redaction'
import {
  dialogNameFor,
  noteDocumentSwapped,
  save,
  suggestedSaveAsName,
  type FileActionContext,
} from '../src/renderer/file-actions'
import { t } from '../src/renderer/i18n/locale'

/**
 * The two ways a document gets a name from the model.
 *
 * The **first save** is the silent one: a document that has never been on disk
 * is named from what the model reads in it, the moment it is saved for the
 * first time. The **manual** trigger is the reader pressing the button on a
 * document that already has a name; the answer is offered in the Save As
 * dialog, because a file that may be open elsewhere or synced is not ours to
 * move behind the reader's back.
 *
 * These go through the app's own functions rather than `nameForSave`, because a
 * test of the shared client still passes if this app stops using the
 * projection — which is the failure that leaks a secret. The file name is the
 * one thing that reaches the reader after the model has seen the rest of the
 * document, so a leak here would outlive the redaction it was meant to.
 *
 * The save path matters as much as the call: a name that never reaches the
 * save is a button that does nothing, and that is exactly what this shipped
 * as before it had a call site.
 */

// jsdom has no CSS.escape, and the document's stylesheet builder needs it
;(globalThis as { CSS?: unknown }).CSS ??= { escape: (s: string) => s }

const LABEL = '客户电话'
const SECRET = '13800138000'
const FILE_NAME = '现有文档.docx'
const ANSWER = '季度营收总结'
const SAVED_PATH = '/tmp/季度营收总结.docx'
const extensions = [...editorExtensions, DocRedaction]
const editors = new Set<Editor>()
afterEach(() => {
  for (const e of editors) e.destroy()
  editors.clear()
})

/** Calls the naming bridge received, in order. */
let sent: Array<{ content: string; trigger: string; filePath: string | null }> = []
/** What the model answers with; '' is a model that returned nothing usable. */
let answer = ANSWER
/** What the main process said about the reader's settings. */
let namingEnabled = true
/** What the model was configured with: a provider, or none at all. */
let providerConfigured = true
/** The names the save path handed to the main process. */
let saved: Array<{ kind: string; name: string }> = []
/** What the save path told the reader; a failure lands here. */
let statuses: string[] = []

function installBridge(): void {
  ;(window as unknown as { aiNamingApi: unknown }).aiNamingApi = {
    suggestFileName: async (input: {
      content: string
      trigger: string
      filePath?: string | null
    }) => {
      sent.push({
        content: input.content,
        trigger: input.trigger,
        filePath: input.filePath ?? null,
      })
      if (!providerConfigured) return { ok: false, reason: 'skip-no-provider' }
      return answer ? { ok: true, name: answer } : { ok: false, reason: 'empty-name' }
    },
    getFileNamingEnabled: async () => namingEnabled,
    setFileNamingEnabled: async (on: boolean) => on,
  }
}

function installDesktop(): void {
  const record = (kind: string) => async (name: string) => {
    saved.push({ kind, name })
    return { ok: true, path: SAVED_PATH }
  }
  ;(window as unknown as { desktop: unknown }).desktop = {
    saveDocxNew: record('new'),
    saveDocxAs: record('as'),
    saveDocx: record('save'),
    discardDocPasswordIntents: async () => {},
    docPasswordIntentRevision: () => 0,
  }
}

interface Opened {
  editor: Editor
  parsed: Awaited<ReturnType<typeof parseDocx>>
}

/** a document with a real parse behind it, so the save path can serialize it */
async function open(text: string): Promise<Opened> {
  const bodyXml = `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`
  const parsed = await parseDocx(await buildDocx({ bodyXml }))
  const editor = new Editor({
    element: document.createElement('div'),
    extensions,
    content: { type: 'doc', content: blocksToPmDoc(parsed.blocks).content as never[] },
  })
  editors.add(editor)
  return { editor, parsed }
}

function selectText(editor: Editor, needle: string): void {
  let start = -1
  editor.state.doc.descendants((node, pos) => {
    if (start >= 0) return
    const at = node.textContent.indexOf(needle)
    if (at >= 0) start = pos + 1 + at
  })
  expect(start).toBeGreaterThan(0)
  editor.view.dispatch(
    editor.state.tr.setSelection(
      TextSelection.create(editor.state.doc, start, start + needle.length),
    ),
  )
}

/** a document that already has a name, holding one withheld span */
async function namedWithSecret(): Promise<Opened> {
  const opened = await open(`联系 ${SECRET} 确认季度营收`)
  selectText(opened.editor, SECRET)
  opened.editor.commands.setRedaction(LABEL)
  return opened
}

/**
 * Only the fields the save path actually reads. `save` takes the whole
 * context, and the rest of it belongs to the window this test is not building.
 */
function ctxFor(opened: Opened, fileName: string, filePath: string | null): FileActionContext {
  const noop = () => {}
  return {
    editor: opened.editor,
    doc: { parsed: opened.parsed, filePath, fileName, hash: '', isBlank: false },
    dirtyRef: { current: true },
    saveInFlightRef: { current: false },
    saveIncompleteRef: { current: false },
    pendingMixedExportRef: { current: false },
    printAutoOpenedPreviewRef: { current: false },
    setStatus: (status: string) => {
      statuses.push(status)
    },
    setRecent: noop,
    setDoc: noop,
    setDocCss: noop,
    setDocLoading: noop,
    setShowAi: noop,
    setAiPanelKey: noop,
    setReadMode: noop,
    setLargeDocSpellOff: noop,
    setShowPrintDialog: noop,
    setShowPagePreview: noop,
    bumpPendingExportTick: noop,
    section: null,
    sectionDirty: false,
    sections: [],
    sectionsDirty: [],
    pgNumDirtySections: [],
    trailingStartType: null,
    setSection: noop,
    setSections: noop,
    setSectionDirty: noop,
    setSectionsDirty: noop,
    setTrailingStartType: noop,
    setPageColor: noop,
    setPageColorDirty: noop,
    setHeader: noop,
    setHeaderDirty: noop,
    setFooter: noop,
    setFooterDirty: noop,
    setHeaderFooter: noop,
    setHfVariants: noop,
    setHfVariantsDirty: noop,
    setTitlePg: noop,
    setTitlePgDirty: noop,
    setEvenOddHf: noop,
    setEvenOddHfDirty: noop,
    setMirrorMargins: noop,
    setMirrorMarginsDirty: noop,
    setComments: noop,
    setCommentsDirty: noop,
    setWatermark: noop,
    setWatermarkDirty: noop,
    setWatermarkStyle: noop,
    setWatermarkPicture: noop,
    setInkAnnotations: noop,
    setInksDirty: noop,
    setFootnotes: noop,
    setEndnotes: noop,
    setNotesDirty: noop,
    setSources: noop,
    setSourcesDirty: noop,
    setZoteroDocumentData: noop,
    setZoteroDocumentDataDirty: noop,
    setThemeFonts: noop,
    setThemeFontsDirty: noop,
    setThemeColors: noop,
    setThemeColorsDirty: noop,
    setProtection: noop,
    setProtectionDirty: noop,
    setWriteProtection: noop,
    setWriteProtectionDirty: noop,
    setRemovePersonalInfo: noop,
    setRemovePersonalInfoDirty: noop,
    setSectionHfEdits: noop,
    setHfLinks: noop,
    setPgNumEdit: noop,
    setPgNumDirtySections: noop,
    setPendingNumbering: noop,
    setStyleUpserts: noop,
    inksDirty: false,
    inkAnnotations: [],
    sectionHfEdits: {},
    hfLinks: {},
    styleUpserts: {},
    comments: [],
    commentsDirty: false,
    protectionDirty: false,
    writeProtectionDirty: false,
    removePersonalInfoDirty: false,
    notesDirty: false,
    sourcesDirty: false,
    zoteroDocumentDataDirty: false,
    themeFontsDirty: false,
    themeColorsDirty: false,
    watermarkDirty: false,
    numberingDirty: false,
    pageColorDirty: false,
    titlePgDirty: false,
    evenOddHfDirty: false,
    mirrorMarginsDirty: false,
    pgNumEdit: null,
  } as unknown as FileActionContext
}

beforeEach(() => {
  sent = []
  answer = ANSWER
  namingEnabled = true
  providerConfigured = true
  saved = []
  statuses = []
  installBridge()
  installDesktop()
  // the latch is module state; each case starts from a freshly swapped document
  noteDocumentSwapped()
})

describe('the manual trigger: the reader asks the model for a name', () => {
  it('hands the model the projected text, never the withheld words', async () => {
    const opened = await namedWithSecret()
    await suggestedSaveAsName(ctxFor(opened, FILE_NAME, '/tmp/existing.docx'))
    expect(sent).toHaveLength(1)
    expect(sent[0].content).not.toContain(SECRET)
    // and the marker is there, so the model can still tell something was hidden
    expect(sent[0].content).toContain(LABEL)
  })

  it('asks on request, so it does not consult the first-save preference', async () => {
    const opened = await namedWithSecret()
    await suggestedSaveAsName(ctxFor(opened, FILE_NAME, '/tmp/existing.docx'))
    expect(sent[0].trigger).toBe('manual')
  })

  it('offers the answer as a complete file name, extension and all', async () => {
    const opened = await namedWithSecret()
    expect(await suggestedSaveAsName(ctxFor(opened, FILE_NAME, '/tmp/existing.docx'))).toBe(
      `${ANSWER}.docx`,
    )
  })

  it('proposes nothing when the model has nothing usable', async () => {
    // '' is the signal that leaves the dialog on the name the document had
    answer = ''
    const opened = await namedWithSecret()
    expect(await suggestedSaveAsName(ctxFor(opened, FILE_NAME, '/tmp/existing.docx'))).toBe('')
  })

  it('proposes nothing when no provider is configured', async () => {
    providerConfigured = false
    const opened = await namedWithSecret()
    expect(await suggestedSaveAsName(ctxFor(opened, FILE_NAME, '/tmp/existing.docx'))).toBe('')
  })
})

describe('the first save of a document that has never been on disk', () => {
  it('asks the model and saves the document under the name it gives', async () => {
    const opened = await open('季度营收总结与明细')
    const ok = await save(ctxFor(opened, t('appUntitledDocx'), null), false)
    expect(ok).toBe(true)
    expect(sent).toHaveLength(1)
    expect(sent[0].trigger).toBe('first-save')
    expect(saved).toEqual([{ kind: 'new', name: `${ANSWER}.docx` }])
  })

  it('hands the model the projected text, never the withheld words', async () => {
    const opened = await namedWithSecret()
    await save(ctxFor(opened, t('appUntitledDocx'), null), false)
    expect(sent).toHaveLength(1)
    expect(sent[0].content).not.toContain(SECRET)
    expect(sent[0].content).toContain(LABEL)
  })

  it('never renames a document that already has a name', async () => {
    const opened = await open('季度营收总结与明细')
    await save(ctxFor(opened, FILE_NAME, '/tmp/existing.docx'), false)
    // the model is not asked at all: a file on disk is not ours to rename
    expect(sent).toHaveLength(0)
    expect(saved).toEqual([{ kind: 'save', name: '/tmp/existing.docx' }])
  })

  it('never names a file that is on disk, even one still called Untitled', async () => {
    // The sharpest case: the document is called "Untitled.docx" — so the
    // first-save branch would take it — but it is already a file on disk, and
    // a file on disk is never renamed behind the reader's back. The model is
    // not asked at all, and the dialog opens on the name the file already has.
    const opened = await open('季度营收总结与明细')
    await save(ctxFor(opened, t('appUntitledDocx'), '/tmp/existing.docx'), true)
    expect(sent).toHaveLength(0)
    expect(saved).toEqual([{ kind: 'as', name: t('appUntitledDocx') }])
  })

  it('never renames a document that already has a name, not even through Save As', async () => {
    // Save As is the branch that reaches the naming, so it is where a
    // first-save naming would wrongly rename a file that is already on disk.
    // The manual button is the only thing that may rename it.
    const opened = await open('季度营收总结与明细')
    await save(ctxFor(opened, FILE_NAME, '/tmp/existing.docx'), true)
    expect(sent).toHaveLength(0)
    expect(saved).toEqual([{ kind: 'as', name: FILE_NAME }])
  })

  it('keeps the current name when no provider is configured', async () => {
    providerConfigured = false
    const opened = await open('季度营收总结与明细')
    await save(ctxFor(opened, t('appUntitledDocx'), null), false)
    expect(saved).toEqual([{ kind: 'new', name: t('appUntitledDocx') }])
  })

  it('keeps the current name when the document is empty', async () => {
    const opened = await open('')
    await save(ctxFor(opened, t('appUntitledDocx'), null), false)
    expect(sent).toHaveLength(0)
    expect(saved).toEqual([{ kind: 'new', name: t('appUntitledDocx') }])
  })

  it('does not ask the model when the caller already carries a name', async () => {
    // The AI's create_document writes a document under the title it was given.
    // That name is settled, so the save path asks nobody: a second opinion
    // would be a model call nobody asked for, spent on every document the
    // model itself wrote.
    const opened = await open('季度营收总结与明细')
    await save(ctxFor(opened, t('appUntitledDocx'), null), false, false, 'AI 写好的标题.docx')
    expect(sent).toHaveLength(0)
    expect(saved).toEqual([{ kind: 'new', name: 'AI 写好的标题.docx' }])
  })

  it('asks once per document, however many times the save is retried', async () => {
    const opened = await open('季度营收总结与明细')
    const ctx = ctxFor(opened, t('appUntitledDocx'), null)
    await save(ctx, true)
    await save(ctx, true)
    expect(sent).toHaveLength(1)
  })
})

describe('the name the Save As dialog opens on', () => {
  it('leads with what the caller proposed', () => {
    expect(dialogNameFor(`${ANSWER}.docx`, null, FILE_NAME)).toBe(`${ANSWER}.docx`)
  })

  it('falls back to the name derived from the document, then to its own', () => {
    expect(dialogNameFor(undefined, '季度营收.docx', FILE_NAME)).toBe('季度营收.docx')
    expect(dialogNameFor(undefined, null, FILE_NAME)).toBe(FILE_NAME)
  })

  it('lets a proposal win over the derived name', () => {
    expect(dialogNameFor('手动命名.docx', '季度营收.docx', FILE_NAME)).toBe('手动命名.docx')
  })
})
