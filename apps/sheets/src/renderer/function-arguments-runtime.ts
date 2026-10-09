/**
 * Univer side of the Function Arguments dialog: where the formula goes,
 * what the cell editor holds when the dialog opens, live sub-expression
 * evaluation, grid range picks and the formula bar's fx button.
 */
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

export interface FunctionArgumentsHost {
  readTarget(): Promise<FunctionTarget | null>
  write(target: FunctionTarget, formula: string): string | null
  evaluate(target: FunctionTarget, expressions: readonly string[]): Promise<PreviewValue[]>
  observeRangePick(target: FunctionTarget, onPick: (ref: string) => void): () => void
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

export function createFunctionArgumentsHost(
  getRuntime: () => UniverRuntime | null,
  notReady: () => string,
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

/// The formula bar's fx button: Univer opens the cell editor on pointer down
/// and emits on click; nothing in Univer listens, so the click is ours.
export function installFxButtonHook(runtime: UniverRuntime): { dispose(): void } {
  const subscription = runtime.univer
    .__getInjector()
    .get(IFormulaEditorManagerService)
    .fxBtnClick$.subscribe(() => fxOpener?.())
  return { dispose: () => subscription.unsubscribe() }
}
