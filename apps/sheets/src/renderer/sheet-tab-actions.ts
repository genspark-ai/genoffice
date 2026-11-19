/// Univer glue for the sheet-tab context menu. Every structural change goes
/// through the stock sheet commands so the mutation listener in App journals
/// it (insert/remove/order/hidden/tab-color) exactly like the ribbon paths.
import { IConfirmService } from '@univerjs/core'

import { t } from './i18n/locale'
import { finalOrder, planReorder, type SheetTabEntry } from './sheet-tab-menu'
import type { LazyWorkbookState, UniverRuntime } from './univer-state'

const INSERT_SHEET = 'sheet.command.insert-sheet'
const REMOVE_SHEET = 'sheet.command.remove-sheet'
const COPY_SHEET = 'sheet.command.copy-sheet'
const RENAME_SHEET = 'sheet.operation.rename-sheet'
const SET_ORDER = 'sheet.command.set-worksheet-order'
const SET_TAB_COLOR = 'sheet.command.set-tab-color'
const HIDE_SHEET = 'sheet.command.set-worksheet-hidden'
const SHOW_SHEET = 'sheet.command.set-worksheet-show'
const ACTIVATE_SHEET = 'sheet.operation.set-worksheet-active'

export interface SheetTabInfo extends SheetTabEntry {
  readonly tabColor: string | null
}

export interface SheetTabActions {
  listSheets(): SheetTabInfo[]
  activeSheetId(): string | null
  activate(id: string): Promise<void>
  /// Excel inserts in front of the active sheet and activates the new one.
  insertBefore(id: string): Promise<void>
  /// Resolves true only when the user confirmed and the sheets were removed.
  remove(ids: readonly string[]): Promise<boolean>
  rename(id: string): void
  moveOrCopy(ids: readonly string[], beforeId: string | null, copy: boolean): Promise<void>
  setTabColor(ids: readonly string[], hex: string | null): Promise<void>
  hide(ids: readonly string[]): Promise<void>
  show(ids: readonly string[]): Promise<void>
  notify(message: string): void
}

export function createSheetTabActions(deps: {
  readonly univerRef: { readonly current: UniverRuntime | null }
  readonly lazyWorkbookRef: { readonly current: LazyWorkbookState | null }
  readonly notify: (message: string) => void
}): SheetTabActions {
  const workbook = () => deps.univerRef.current?.univerAPI.getActiveWorkbook() ?? null
  const run = async (id: string, params: object): Promise<boolean> => {
    const runtime = deps.univerRef.current
    if (!runtime) return false
    try {
      return Boolean(await runtime.univerAPI.executeCommand(id, params))
    } catch {
      return false
    }
  }
  const orderIds = (): string[] =>
    workbook()
      ?.getSheets()
      .map((s) => s.getSheetId()) ?? []

  const activate = async (id: string): Promise<void> => {
    const unitId = workbook()?.getId()
    if (unitId) await run(ACTIVATE_SHEET, { unitId, subUnitId: id })
  }

  const reorder = async (ids: readonly string[], beforeId: string | null): Promise<void> => {
    const order = orderIds()
    const target = finalOrder(order, new Set(ids), beforeId)
    for (const step of planReorder(order, target, new Set(ids))) {
      await run(SET_ORDER, { subUnitId: step.id, order: step.to })
    }
  }

  return {
    listSheets() {
      const file = deps.lazyWorkbookRef.current?.file
      return (
        workbook()
          ?.getSheets()
          .map((sheet) => {
            const id = sheet.getSheetId()
            return {
              id,
              name: sheet.getSheetName(),
              hidden: sheet.isSheetHidden(),
              veryHidden: file?.sheets.find((s) => s.id === id)?.veryHidden === true,
              tabColor: sheet.getTabColor() || null,
            }
          }) ?? []
      )
    },
    activeSheetId() {
      // Univer throws while a workbook is between snapshots (no active sheet).
      try {
        return workbook()?.getActiveSheet().getSheetId() ?? null
      } catch {
        return null
      }
    },
    activate,
    async insertBefore(id) {
      const wb = workbook()
      if (!wb) return
      const before = new Set(orderIds())
      const index = Math.max(0, orderIds().indexOf(id))
      if (!(await run(INSERT_SHEET, { unitId: wb.getId(), index }))) return
      const added = orderIds().find((sheetId) => !before.has(sheetId))
      if (added) await activate(added)
    },
    async remove(ids) {
      const runtime = deps.univerRef.current
      if (!runtime || ids.length === 0) return false
      const confirmed = await runtime.univer
        .__getInjector()
        .get(IConfirmService)
        .confirm({
          id: 'genoffice.sheet-tab.delete',
          title: { title: t('dlgSheetDeleteTitle') },
          children: { title: t('dlgSheetDeleteBody', { count: ids.length }) },
          cancelText: t('dlgCancel'),
          confirmText: t('dlgSheetDeleteConfirm'),
        })
      if (!confirmed) return false
      let removed = false
      for (const id of ids) removed = (await run(REMOVE_SHEET, { subUnitId: id })) || removed
      return removed
    },
    rename(id) {
      void run(RENAME_SHEET, { subUnitId: id })
    },
    async moveOrCopy(ids, beforeId, copy) {
      if (!copy) {
        await reorder(ids, beforeId)
        return
      }
      const clones: string[] = []
      for (const id of ids) {
        const before = new Set(orderIds())
        await run(COPY_SHEET, { subUnitId: id })
        const clone = orderIds().find((sheetId) => !before.has(sheetId))
        if (clone) clones.push(clone)
      }
      if (clones.length === 0) return
      await reorder(clones, beforeId)
      await activate(clones[clones.length - 1]!)
    },
    async setTabColor(ids, hex) {
      for (const id of ids) await run(SET_TAB_COLOR, { subUnitId: id, value: hex ?? '' })
    },
    async hide(ids) {
      for (const id of ids) await run(HIDE_SHEET, { subUnitId: id })
    },
    async show(ids) {
      for (const id of ids) await run(SHOW_SHEET, { subUnitId: id })
    },
    notify: deps.notify,
  }
}
