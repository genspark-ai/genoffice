/**
 * Univer side of the Function Arguments dialog: where the formula goes,
 * what the cell editor holds when the dialog opens, live sub-expression
 * evaluation, grid range picks and the formula bar's fx button.
 */
import { basename } from 'node:path'
import { DOCS_FORMULA_BAR_EDITOR_UNIT_ID_KEY, DOCS_NORMAL_EDITOR_UNIT_ID_KEY } from '@univerjs/core'
import { IEditorService } from '@univerjs/docs-ui'
import { IFormulaEditorManagerService } from '@univerjs/sheets-ui'

import { columnLabel, pickedRangeRef, type PreviewValue } from './function-arguments'
import type { UniverRuntime } from './univer-state'

export interface FunctionTarget {
  readonly unitId: string
  readonly sheetId: string
  readonly sheetName: string
  readonly row: number
  readonly col: number
  readonly label: string
  /// Formula text the dialog edits (`=` included); '' for a cell without one.
  readonly formula: string
  readonly caret: number
  /// The cell editor was open with this text; Cancel reopens it.
  readonly editorText: string | null
}

export interface ExternalPickTarget {
  /// The link's stored `[n]` index.
  readonly index: number
  readonly name: string
  readonly source: string
  /// True when the source workbook is open in another tab: a grid pick can
  /// relay the range. Closed sources take a typed range instead.
  readonly open: boolean
  readonly wcId: number | null
  readonly sheetNames: readonly string[]
}

export interface FunctionArgumentsHost {
  readTarget(): Promise<FunctionTarget | null>
  write(target: FunctionTarget, formula: string): string | null
  evaluate(target: FunctionTarget, expressions: readonly string[]): Promise<PreviewValue[]>
  observeRangePick(target: FunctionTarget, onPick: (ref: string) => void): () => void
  /** Linked workbooks of the file being edited, flagged by whether their
      source is open in another tab. */
  listExternalPickTargets(): Promise<ExternalPickTarget[]>
  /** Arms a cross-workbook pick against one open linked workbook; the next
      range selected on that workbook's grid arrives through onPick as the
      stored `'[n]Sheet'!A1:B2` spelling. */
  pickExternalRange(target: ExternalPickTarget, onPick: (ref: string) => void): () => void
  cancel(target: FunctionTarget): void
}

function editorTextAndCaret(runtime: UniverRuntime): { text: string; caret: number } | null {
  const editorService = runtime.univer.__getInjector().get(IEditorService)
  const editor = editorService.getEditor(DOCS_NORMAL_EDITOR_UNIT_ID_KEY)
  if (!editor) return null
  const text = (editor.getDocumentData().body?.dataStream ?? '').replace(/\r?\n$/, '')
  const ranges =
    editor.getSelectionRanges().length > 0
      ? editor.getSelectionRanges()
      : (editorService.getEditor(DOCS_FORMULA_BAR_EDITOR_UNIT_ID_KEY)?.getSelectionRanges() ?? [])
  const caret = Math.min(text.length, ranges[0]?.startOffset ?? text.length)
  return { text, caret }
}

let evaluationQueue: Promise<void> = Promise.resolve()

/// The stored spelling of one external sheet reference: `[1]Sheet1` when the
/// name reads as a bare identifier (Excel leaves it unquoted), the quoted
/// `'[1]Sheet Name'` otherwise — the two shapes the gateway's break/renumber
/// scan and the sidecar's usage scan both match.
export function formatExternalRef(index: number, sheetName: string, body: string): string {
  const bare = /^[A-Za-z_][A-Za-z0-9_.]*$/.test(sheetName) && !/^[A-Za-z]{1,3}\d+$/.test(sheetName)
  const part = bare ? `[${index}]${sheetName}` : `'[${index}]${sheetName.replace(/'/g, "''")}'`
  return `${part}!${body}`
}

/// The stored spelling for a range picked on another workbook's grid: the
/// sheet name comes from the link's stored list when the picked tab's sheet
/// matches one (case-insensitive), the live name otherwise. Whole-row/column
/// picks keep their `1:3` / `A:C` bodies.
export function formatExternalPickRef(
  target: ExternalPickTarget,
  selection: {
    sheetName: string
    startRow: number
    endRow: number
    startColumn: number
    endColumn: number
    rangeType?: number | undefined
  },
): string {
  const stored =
    target.sheetNames.find((name) => name.toLowerCase() === selection.sheetName.toLowerCase()) ??
    selection.sheetName
  // equal names keep pickedRangeRef sheet-unqualified, leaving the A1 body
  const body = pickedRangeRef(selection, stored, stored)
  return formatExternalRef(target.index, stored, body)
}

export function createFunctionArgumentsHost(
  getRuntime: () => UniverRuntime | null,
  notReady: () => string,
  getFilePath: () => string | null = () => null,
): FunctionArgumentsHost {
  return {
    async readTarget() {
      const runtime = getRuntime()
      const workbook = runtime?.univerAPI.getActiveWorkbook()
      const worksheet = workbook?.getActiveSheet()
      const range = workbook?.getActiveRange()
      if (!runtime || !workbook || !worksheet || !range) return null
      const row = range.getRow()
      const col = range.getColumn()
      let editorText: string | null = null
      let caret = 0
      if (workbook.isCellEditing()) {
        const editing = editorTextAndCaret(runtime)
        editorText = editing?.text ?? ''
        caret = editing?.caret ?? 0
        await workbook.endEditingAsync(false)
      }
      const cellFormula = worksheet.getRange(row, col, 1, 1).getFormula()
      const formula = editorText ?? cellFormula
      const isFormula = formula.startsWith('=')
      return {
        unitId: workbook.getId(),
        sheetId: worksheet.getSheetId(),
        sheetName: worksheet.getSheetName(),
        row,
        col,
        label: `${columnLabel(col)}${row + 1}`,
        formula: isFormula ? formula : '',
        caret: isFormula ? (editorText === null ? formula.length : caret) : 0,
        editorText,
      }
    },
    write(target, formula) {
      const workbook = getRuntime()?.univerAPI.getActiveWorkbook()
      const sheet = workbook?.getSheetBySheetId(target.sheetId)
      if (!workbook || !sheet || workbook.getId() !== target.unitId) return notReady()
      try {
        const cell = sheet.getRange(target.row, target.col, 1, 1)
        cell.setValue({ f: formula })
        workbook.setActiveSheet(sheet)
        workbook.setActiveRange(cell)
      } catch (error: unknown) {
        return error instanceof Error ? error.message : notReady()
      }
      return null
    },
    async evaluate(target, expressions) {
      const runtime = getRuntime()
      if (!runtime || expressions.length === 0) return expressions.map(() => undefined)
      const formulas = expressions.map((expr) => (expr.startsWith('=') ? expr : `=${expr}`))
      // The facade resolves every pending executeFormulas with whichever batch
      // result lands first, so batches run one at a time.
      const run = evaluationQueue.then(() =>
        runtime.univerAPI
          .getFormula()
          .executeFormulas(
            { [target.unitId]: { [target.sheetId]: { [target.row]: { [target.col]: formulas } } } },
            3000,
          ),
      )
      evaluationQueue = run.then(
        () => undefined,
        () => undefined,
      )
      const result = await run
      const items = result[target.unitId]?.[target.sheetId]?.[target.row]?.[target.col] ?? []
      const byFormula = new Map(items.map((item) => [item.formula, item.value as PreviewValue]))
      return formulas.map((formula) => byFormula.get(formula))
    },
    observeRangePick(target, onPick) {
      const runtime = getRuntime()
      if (!runtime) return () => {}
      // Ending the cell editor re-emits the selection; only a pointer on the
      // grid after the dialog opened counts as a pick.
      let armed = false
      const arm = (event: PointerEvent): void => {
        if (event.target instanceof Element && event.target.closest('#univer-container')) {
          armed = true
        }
      }
      document.addEventListener('pointerdown', arm, true)
      const disposable = runtime.univerAPI.addEvent(
        runtime.univerAPI.Event.SelectionMoveEnd,
        (params) => {
          const range = params.selections[0]
          if (!armed || !range || params.workbook.getId() !== target.unitId) return
          onPick(pickedRangeRef(range, params.worksheet.getSheetName(), target.sheetName))
        },
      )
      return () => {
        document.removeEventListener('pointerdown', arm, true)
        disposable.dispose()
      }
    },
    async listExternalPickTargets() {
      const path = getFilePath()
      if (!path) return []
      const [links, tabs] = await Promise.all([
        window.desktopApi.listExternalLinks({ path }),
        window.desktopApi.listWorkbookTabs(),
      ])
      const normalize = (value: string): string => value.replaceAll('\\', '/').toLowerCase()
      return links.links.map((link) => {
        const tab = link.source
          ? tabs.tabs.find((entry) => normalize(entry.path) === normalize(link.source!))
          : undefined
        return {
          index: link.index,
          name: basename(link.source ?? link.rawTarget),
          source: link.source ?? link.rawTarget,
          open: tab !== undefined,
          wcId: tab?.wcId ?? null,
          sheetNames: link.sheetNames,
        }
      })
    },
    pickExternalRange(target, onPick) {
      let disposed = false
      let offSelected: (() => void) | null = null
      let pickId: string | null = null
      void (async () => {
        const armed = await window.desktopApi.armWorkbookPick()
        if (disposed) {
          void window.desktopApi.cancelWorkbookPick(armed.pickId)
          return
        }
        pickId = armed.pickId
        offSelected = window.desktopApi.onWorkbookPickSelected((selection) => {
          if (selection.pickId !== pickId) return
          onPick(formatExternalPickRef(target, selection))
          // the pick is consumed; release the other tabs
          void window.desktopApi.cancelWorkbookPick(pickId!)
        })
      })()
      return () => {
        disposed = true
        offSelected?.()
        offSelected = null
        if (pickId !== null) void window.desktopApi.cancelWorkbookPick(pickId)
      }
    },
    cancel(target) {
      if (target.editorText === null) return
      const runtime = getRuntime()
      const workbook = runtime?.univerAPI.getActiveWorkbook()
      const sheet = workbook?.getSheetBySheetId(target.sheetId)
      if (!runtime || !workbook || !sheet) return
      try {
        const cell = sheet.getRange(target.row, target.col, 1, 1)
        workbook.setActiveSheet(sheet)
        workbook.setActiveRange(cell)
        if (!workbook.startEditing()) return
        runtime.univer
          .__getInjector()
          .get(IEditorService)
          .getEditor(DOCS_NORMAL_EDITOR_UNIT_ID_KEY)
          ?.replaceText(target.editorText, true)
      } catch {
        // Best effort: the typed text is lost only if the editor refuses to reopen
      }
    },
  }
}

let fxOpener: (() => void) | null = null

/// The shell registers how an fx click opens Insert Function.
export function registerFxButtonOpener(open: () => void): () => void {
  fxOpener = open
  return () => {
    if (fxOpener === open) fxOpener = null
  }
}

/**
 * The picking side of the cross-workbook range picker: while another tab's
 * Function Arguments dialog has a pick armed, the next selection on THIS
 * workbook's grid is relayed to it through the main process, which then
 * re-activates the editing tab.
 */
export function installWorkbookPickRelay(runtime: UniverRuntime): { dispose(): void } {
  let armed: string | null = null
  const offArmed = window.desktopApi.onWorkbookPickArmed(({ pickId }) => {
    armed = pickId
  })
  const offCancelled = window.desktopApi.onWorkbookPickCancelled(({ pickId }) => {
    if (armed === pickId) armed = null
  })
  const disposable = runtime.univerAPI.addEvent(
    runtime.univerAPI.Event.SelectionMoveEnd,
    (params) => {
      if (armed === null) return
      const range = params.selections[0]
      if (!range) return
      const pickId = armed
      armed = null
      void window.desktopApi.sendWorkbookPickSelection({
        pickId,
        sheetName: params.worksheet.getSheetName(),
        startRow: range.startRow,
        endRow: range.endRow,
        startColumn: range.startColumn,
        endColumn: range.endColumn,
        ...(range.rangeType === undefined ? {} : { rangeType: range.rangeType }),
      })
    },
  )
  return {
    dispose: () => {
      offArmed()
      offCancelled()
      disposable.dispose()
    },
  }
}

/// The formula bar's fx button: Univer opens the cell editor on pointer down
/// and emits on click; nothing in Univer listens, so the click is ours.
export function installFxButtonHook(runtime: UniverRuntime): { dispose(): void } {
  const subscription = runtime.univer
    .__getInjector()
    .get(IFormulaEditorManagerService)
    .fxBtnClick$.subscribe(() => fxOpener?.())
  return { dispose: () => subscription.unsubscribe() }
}
