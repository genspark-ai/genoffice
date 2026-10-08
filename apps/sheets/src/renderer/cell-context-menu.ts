/**
 * Excel's cell / row-header / column-header right-click entries that Univer's
 * context menu does not ship: Paste Special…, Filter ▸, Number Format ▸,
 * Format Cells…, Pick From Drop-down List…, Define Name…, Link…, Insert
 * Function…. Dialog entries hand their ribbon command to the shell through
 * the sink ExcelShell registers; the rest run Univer commands directly.
 */
import {
  ColorKit,
  CommandType,
  ICommandService,
  IUniverInstanceService,
  LocaleService,
  UniverInstanceType,
  extractPureTextFromCell,
  getDisplayValueFromCell,
} from '@univerjs/core'
import type { IAccessor, ICommand, IDisposable, ILocales } from '@univerjs/core'
import type { Lang } from '@genoffice/i18n'
import { SheetsSelectionsService, getSheetCommandTarget } from '@univerjs/sheets'
import {
  ClearSheetsFilterCriteriaCommand,
  SetSheetsFilterCriteriaCommand,
  SheetsFilterService,
  SmartToggleSheetsFilterCommand,
} from '@univerjs/sheets-filter'
import type { IFilterColumn } from '@univerjs/sheets-filter'
import { ISheetCellDropdownManagerService } from '@univerjs/sheets-ui'
import {
  ContextMenuGroup,
  ContextMenuPosition,
  IMenuManagerService,
  MenuItemType,
  getMenuHiddenObservable,
} from '@univerjs/ui'
import type { IMenuButtonItem, IMenuSelectorItem, MenuSchemaType } from '@univerjs/ui'

import { getLang, t, tFor, type StringKey } from './i18n/locale'
import { NUMBER_FORMAT_LABEL, numberFormatCategories } from './number-format'
import { collectPickListValues, filterCriteriaForText } from './pick-from-list'
import type { UniverRuntime } from './univer-state'
import { visualPasteDisabled$ } from './visual-arrange-actions'
import type { Observable } from 'rxjs'

export const SHELL_COMMAND_ID = 'genoffice.ctx.shell-command'
export const PASTE_SPECIAL_OPEN_COMMAND_ID = 'sheets.paste-special.open'
export const FILTER_BY_COMMAND_ID = 'genoffice.ctx.filter-by'
export const PICK_FROM_LIST_COMMAND_ID = 'genoffice.ctx.pick-from-list'
export const FILTER_MENU_ID = 'genoffice.ctx.filter'
export const NUMBER_FORMAT_MENU_ID = 'genoffice.ctx.number-format'
export const CLEAR_FILTER_MENU_ID = 'genoffice.ctx.clear-filter'

export type FilterBy = 'value' | 'fill' | 'font'

export interface ContextMenuLeaf {
  readonly id: string
  readonly title: string
  readonly titleKey: StringKey
  readonly commandId: string
  readonly params?: Record<string, unknown>
}

export interface ContextMenuNode {
  readonly id: string
  readonly title: string
  readonly titleKey: StringKey
  readonly group: ContextMenuGroup
  readonly order: number
  readonly commandId?: string
  readonly params?: Record<string, unknown>
  readonly children?: readonly ContextMenuLeaf[]
  readonly disabled$?: Observable<boolean>
}

const shell = (
  id: string,
  titleKey: StringKey,
  command: string,
  group: ContextMenuGroup,
  order: number,
): ContextMenuNode => ({
  id,
  title: t(titleKey),
  titleKey,
  group,
  order,
  commandId: SHELL_COMMAND_ID,
  params: { command },
})

/** Cell-menu additions in Excel's order, merged around Univer's own groups. */
export function cellContextMenuLayout(): readonly ContextMenuNode[] {
  const filterLeaf = (by: FilterBy, titleKey: StringKey): ContextMenuLeaf => ({
    id: `${FILTER_BY_COMMAND_ID}.${by}`,
    title: t(titleKey),
    titleKey,
    commandId: FILTER_BY_COMMAND_ID,
    params: { by },
  })
  return [
    {
      id: PASTE_SPECIAL_OPEN_COMMAND_ID,
      title: t('appCtxPasteSpecial'),
      titleKey: 'appCtxPasteSpecial',
      group: ContextMenuGroup.FORMAT,
      order: 1.5,
      commandId: PASTE_SPECIAL_OPEN_COMMAND_ID,
    },
    {
      ...shell(
        'genoffice.ctx.paste-visual',
        'appCtxPasteVisual',
        'visual:paste',
        ContextMenuGroup.FORMAT,
        1.6,
      ),
      disabled$: visualPasteDisabled$,
    },
    {
      id: FILTER_MENU_ID,
      title: t('appCtxFilter'),
      titleKey: 'appCtxFilter',
      group: ContextMenuGroup.DATA,
      order: -1,
      children: [
        filterLeaf('value', 'appCtxFilterByValue'),
        filterLeaf('fill', 'appCtxFilterByFill'),
        filterLeaf('font', 'appCtxFilterByFont'),
        {
          id: CLEAR_FILTER_MENU_ID,
          title: t('appCtxClearFilter'),
          titleKey: 'appCtxClearFilter',
          commandId: ClearSheetsFilterCriteriaCommand.id,
        },
      ],
    },
    // Threaded comments sit beside Univer's note entries (OTHERS, order 0).
    shell(
      'genoffice.ctx.new-comment',
      'appCtxNewComment',
      'comment-new',
      ContextMenuGroup.OTHERS,
      0.1,
    ),
    shell(
      'genoffice.ctx.reply-comment',
      'appCtxReplyComment',
      'comment-reply',
      ContextMenuGroup.OTHERS,
      0.2,
    ),
    shell(
      'genoffice.ctx.delete-comment',
      'appCtxDeleteComment',
      'comment-delete',
      ContextMenuGroup.OTHERS,
      0.3,
    ),
    shell(
      'genoffice.ctx.resolve-comment',
      'appCtxResolveComment',
      'comment-resolve',
      ContextMenuGroup.OTHERS,
      0.4,
    ),
    {
      id: NUMBER_FORMAT_MENU_ID,
      title: t('appCtxNumberFormat'),
      titleKey: 'appCtxNumberFormat',
      group: ContextMenuGroup.OTHERS,
      order: 1,
      children: [
        ...numberFormatCategories().map((category): ContextMenuLeaf => {
          const titleKey = NUMBER_FORMAT_LABEL[category.label] ?? 'dlgFcNumGeneral'
          return {
            id: `${NUMBER_FORMAT_MENU_ID}.${category.label}`,
            title: t(titleKey),
            titleKey,
            commandId: SHELL_COMMAND_ID,
            params: { command: `format:${category.pattern}` },
          }
        }),
        {
          id: `${NUMBER_FORMAT_MENU_ID}.more`,
          title: t('appCtxMoreNumberFormats'),
          titleKey: 'appCtxMoreNumberFormats',
          commandId: SHELL_COMMAND_ID,
          params: { command: 'format-cells' },
        },
      ],
    },
    shell(
      'genoffice.ctx.format-cells',
      'appCtxFormatCells',
      'format-cells',
      ContextMenuGroup.OTHERS,
      2,
    ),
    {
      id: PICK_FROM_LIST_COMMAND_ID,
      title: t('appCtxPickFromList'),
      titleKey: 'appCtxPickFromList',
      group: ContextMenuGroup.OTHERS,
      order: 3,
      commandId: PICK_FROM_LIST_COMMAND_ID,
    },
    shell(
      'genoffice.ctx.define-name',
      'appCtxDefineName',
      'name-manager-open',
      ContextMenuGroup.OTHERS,
      4,
    ),
    shell('genoffice.ctx.link', 'appCtxLink', 'link-open', ContextMenuGroup.OTHERS, 5),
    shell(
      'genoffice.ctx.insert-function',
      'appCtxInsertFunction',
      'insert-function-open',
      ContextMenuGroup.OTHERS,
      6,
    ),
  ]
}

/** Row/column header menus only gain Format Cells…, between Clear and Row Height / Column Width. */
export function headerContextMenuLayout(): readonly ContextMenuNode[] {
  return [
    shell(
      'genoffice.ctx.format-cells',
      'appCtxFormatCells',
      'format-cells',
      ContextMenuGroup.LAYOUT,
      2.5,
    ),
  ]
}

let commandSink: ((command: string) => void) | null = null
let pasteSpecialOpener: (() => void) | null = null

/** ExcelShell owns the dialogs; it registers its dispatcher here. */
export function setContextMenuCommandSink(sink: ((command: string) => void) | null): void {
  commandSink = sink
}

export function registerPasteSpecialOpener(open: () => void): () => void {
  pasteSpecialOpener = open
  return () => {
    if (pasteSpecialOpener === open) pasteSpecialOpener = null
  }
}

function activeCell(accessor: IAccessor) {
  const target = getSheetCommandTarget(accessor.get(IUniverInstanceService))
  const selection = accessor.get(SheetsSelectionsService).getCurrentLastSelection()
  if (!target || !selection) return null
  const row = selection.primary?.actualRow ?? selection.range.startRow
  const col = selection.primary?.actualColumn ?? selection.range.startColumn
  return { ...target, row, col }
}

async function filterBySelectedCell(accessor: IAccessor, by: FilterBy): Promise<boolean> {
  const cell = activeCell(accessor)
  if (!cell) return false
  const { unitId, subUnitId, worksheet, row, col } = cell
  const filterService = accessor.get(SheetsFilterService)
  const commandService = accessor.get(ICommandService)
  let model = filterService.getFilterModel(unitId, subUnitId)
  if (!model) {
    await commandService.executeCommand(SmartToggleSheetsFilterCommand.id)
    model = filterService.getFilterModel(unitId, subUnitId)
  }
  if (!model) return false
  const range = model.getRange()
  if (col < range.startColumn || col > range.endColumn) return false
  let criteria: IFilterColumn
  if (by === 'value') {
    criteria = {
      colId: col,
      filters: filterCriteriaForText(extractPureTextFromCell(worksheet.getCell(row, col))),
    }
  } else {
    const style = worksheet.getComposedCellStyle(row, col)
    const rgb = by === 'fill' ? style.bg?.rgb : style.cl?.rgb
    const color = rgb ? new ColorKit(rgb).toRgbString() : null
    // the filter model's default font colour is black; its default fill is "no fill"
    criteria =
      by === 'fill'
        ? { colId: col, colorFilters: { cellFillColors: [color] } }
        : { colId: col, colorFilters: { cellTextColors: [color ?? 'rgb(0,0,0)'] } }
  }
  return commandService.executeCommand(SetSheetsFilterCriteriaCommand.id, {
    unitId,
    subUnitId,
    col,
    criteria,
  })
}

function pickFromList(accessor: IAccessor, runtime: UniverRuntime): boolean {
  const cell = activeCell(accessor)
  if (!cell) return false
  const { unitId, subUnitId, workbook, worksheet, row, col } = cell
  const entries = collectPickListValues(
    (r) => {
      const data = worksheet.getCell(r, col)
      return data ? { display: getDisplayValueFromCell(data), cell: data } : null
    },
    row,
    worksheet.getRowCount(),
  )
  if (entries.length === 0) return false
  const byDisplay = new Map(entries.map((entry) => [entry.display, entry]))
  const dropdowns = accessor.get(ISheetCellDropdownManagerService)
  // the menu click that got us here is still settling; a dropdown opened now closes with it
  setTimeout(() => {
    dropdowns.showDropdown({
      location: { unitId, subUnitId, workbook, worksheet, row, col },
      closeOnOutSide: true,
      type: 'list',
      props: {
        options: entries.map((entry) => ({ label: entry.display, value: entry.display })),
        defaultValue: getDisplayValueFromCell(worksheet.getCell(row, col)),
        showEdit: false,
        showSearch: entries.length > 8,
        onChange: async ([picked]) => {
          const entry = picked === undefined ? undefined : byDisplay.get(picked)
          if (!entry) return true
          const range = runtime.univerAPI
            .getUniverSheet(unitId)
            ?.getSheetBySheetId(subUnitId)
            ?.getRange(row, col)
          if (!range) return true
          const source = entry.cell
          const plain = source?.p || source?.v === undefined || source.v === null
          // a formula in the target cell must not survive the pick
          range.setValue({
            ...(plain ? { v: entry.display } : { v: source.v, t: source.t }),
            p: null,
            f: null,
            si: null,
          })
          return true
        },
      },
    })
  }, 0)
  return true
}

const LOCALE_NAMESPACE = 'genofficeCtx'
const localeKey = (key: StringKey): string => `${LOCALE_NAMESPACE}.${key}`

/**
 * Titles are Univer locale keys so a live language switch re-labels the
 * entries: the strings live under our own namespace of whichever Univer
 * locale is active and are reloaded on every switch.
 */
function loadMenuLocale(localeService: LocaleService, lang: Lang): void {
  const strings: Record<string, string> = {}
  for (const node of [...cellContextMenuLayout(), ...headerContextMenuLayout()]) {
    strings[node.titleKey] = tFor(lang, node.titleKey)
    for (const child of node.children ?? []) strings[child.titleKey] = tFor(lang, child.titleKey)
  }
  localeService.load({
    [localeService.getCurrentLocale()]: { [LOCALE_NAMESPACE]: strings },
  } as unknown as ILocales)
}

function toMenuItem(
  accessor: IAccessor,
  node: ContextMenuNode | ContextMenuLeaf,
): IMenuButtonItem | IMenuSelectorItem {
  const hidden$ = getMenuHiddenObservable(accessor, UniverInstanceType.UNIVER_SHEET)
  if ('children' in node && node.children) {
    return { id: node.id, type: MenuItemType.SUBITEMS, title: localeKey(node.titleKey), hidden$ }
  }
  return {
    id: node.id,
    ...(node.commandId ? { commandId: node.commandId } : {}),
    type: MenuItemType.BUTTON,
    title: localeKey(node.titleKey),
    ...(node.params ? { params: node.params } : {}),
    ...('disabled$' in node && node.disabled$ ? { disabled$: node.disabled$ } : {}),
    hidden$,
  }
}

function schemaFor(nodes: readonly ContextMenuNode[]): Record<string, MenuSchemaType> {
  const groups: Record<string, Record<string, MenuSchemaType>> = {}
  for (const node of nodes) {
    const entry: Record<string, MenuSchemaType> = {}
    for (const child of node.children ?? []) {
      entry[child.id] = { menuItemFactory: (accessor) => toMenuItem(accessor, child) }
    }
    ;(groups[node.group] ??= {})[node.id] = {
      order: node.order,
      menuItemFactory: (accessor: IAccessor) => toMenuItem(accessor, node),
      ...entry,
    }
  }
  return groups
}

export function installCellContextMenu(runtime: UniverRuntime): IDisposable {
  const injector = runtime.univer.__getInjector()
  const commandService = injector.get(ICommandService)
  const commands: ICommand[] = [
    {
      id: SHELL_COMMAND_ID,
      type: CommandType.COMMAND,
      handler: (_accessor, params?: { command?: string }) => {
        if (!params?.command || !commandSink) return false
        commandSink(params.command)
        return true
      },
    },
    {
      id: PASTE_SPECIAL_OPEN_COMMAND_ID,
      type: CommandType.COMMAND,
      handler: () => {
        if (!pasteSpecialOpener) {
          console.warn(`${PASTE_SPECIAL_OPEN_COMMAND_ID}: no Paste Special dialog registered`)
          return false
        }
        pasteSpecialOpener()
        return true
      },
    },
    {
      id: FILTER_BY_COMMAND_ID,
      type: CommandType.COMMAND,
      handler: (accessor, params?: { by?: FilterBy }) =>
        params?.by ? filterBySelectedCell(accessor, params.by) : false,
    },
    {
      id: PICK_FROM_LIST_COMMAND_ID,
      type: CommandType.COMMAND,
      handler: (accessor) => pickFromList(accessor, runtime),
    },
  ]
  const disposables = commands.map((command) => commandService.registerCommand(command))
  const localeService = injector.get(LocaleService)
  loadMenuLocale(localeService, getLang())
  // applyUniverLocale switches Univer's locale type after boot; reload under the new one too
  const localeSubscription = localeService.currentLocale$.subscribe(() =>
    loadMenuLocale(localeService, getLang()),
  )
  disposables.push(
    { dispose: () => localeSubscription.unsubscribe() },
    { dispose: window.desktopApi.onLanguageChanged((lang) => loadMenuLocale(localeService, lang)) },
  )
  const header = schemaFor(headerContextMenuLayout())
  injector.get(IMenuManagerService).mergeMenu({
    [ContextMenuPosition.MAIN_AREA]: schemaFor(cellContextMenuLayout()),
    [ContextMenuPosition.ROW_HEADER]: header,
    [ContextMenuPosition.COL_HEADER]: header,
  })
  return {
    dispose() {
      for (const disposable of disposables) disposable.dispose()
    },
  }
}
