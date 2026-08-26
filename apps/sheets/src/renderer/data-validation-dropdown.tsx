import type { PointerEvent as ReactPointerEvent } from 'react'

import type { UniverRuntime, UniverWorksheet } from './univer-state'

const DROPDOWN_COMPONENT_KEY = 'genoffice-active-data-validation-dropdown'
const INPUT_MESSAGE_COMPONENT_KEY = 'genoffice-active-data-validation-input-message'
const SHOW_DROPDOWN_COMMAND = 'sheet.operation.show-data-validation-dropdown'
const DROPDOWN_SIZE = 14
const INPUT_MESSAGE_WIDTH = 108
const INPUT_MESSAGE_LINE_HEIGHT = 12
const INPUT_MESSAGE_PADDING_HEIGHT = 16
const INPUT_MESSAGE_MAX_HEIGHT = 180
const INPUT_MESSAGE_UNITS_PER_LINE = 20

interface ValidationLike {
  readonly rule?: {
    readonly showDropDown?: boolean
    readonly showInputMessage?: boolean
    readonly promptTitle?: string
    readonly prompt?: string
  }
  getCriteriaType(): string
}

interface ActiveValidationTarget {
  readonly unitId: string
  readonly sheetId: string
  readonly row: number
  readonly column: number
  readonly validation: ValidationLike
  readonly range: ReturnType<UniverWorksheet['getRange']>
  readonly worksheet: UniverWorksheet
}

export interface DataValidationInputMessage {
  readonly title: string
  readonly prompt: string
}

function rangePixelHeight(target: ActiveValidationTarget): number {
  let height = 0
  for (let row = target.range.getRow(); row <= target.range.getLastRow(); row += 1) {
    height += target.worksheet.getRowHeight(row)
  }
  return height
}

export function shouldShowDataValidationDropdown(
  validation: ValidationLike | null | void,
): boolean {
  return validation?.getCriteriaType() === 'list' && validation.rule?.showDropDown !== false
}

export function getDataValidationInputMessage(
  validation: ValidationLike | null | void,
): DataValidationInputMessage | null {
  if (validation?.rule?.showInputMessage !== true) return null
  const title = validation.rule.promptTitle ?? ''
  const prompt = validation.rule.prompt ?? ''
  return title || prompt ? { title, prompt } : null
}

function textDisplayUnits(value: string): number {
  let units = 0
  for (const character of value) units += character.codePointAt(0)! > 0xff ? 2 : 1
  return units
}

export function dataValidationInputMessageHeight(message: DataValidationInputMessage): number {
  const linesFor = (value: string): number =>
    value
      ? value
          .split(/\r\n?|\n/)
          .reduce(
            (total, line) =>
              total + Math.max(1, Math.ceil(textDisplayUnits(line) / INPUT_MESSAGE_UNITS_PER_LINE)),
            0,
          )
      : 0
  const lines = Math.max(1, linesFor(message.title) + linesFor(message.prompt))
  return Math.min(
    INPUT_MESSAGE_MAX_HEIGHT,
    INPUT_MESSAGE_PADDING_HEIGHT + lines * INPUT_MESSAGE_LINE_HEIGHT,
  )
}

function resolveActiveValidationTarget(runtime: UniverRuntime): ActiveValidationTarget | null {
  try {
    const workbook = runtime.univerAPI.getActiveWorkbook()
    const worksheet = workbook?.getActiveSheet()
    const cell = workbook?.getActiveCell()
    if (!workbook || !worksheet || !cell) return null
    const row = cell.getRow()
    const column = cell.getColumn()
    if (!Number.isInteger(row) || !Number.isInteger(column) || row < 0 || column < 0) return null
    const range = worksheet.getCellMergeData(row, column) ?? worksheet.getRange(row, column, 1, 1)
    const validation = range.getDataValidation()
    if (!validation) return null
    return {
      unitId: workbook.getId(),
      sheetId: worksheet.getSheetId(),
      row,
      column,
      validation,
      range,
      worksheet,
    }
  } catch {
    return null
  }
}

function resolveActiveDropdownTarget(runtime: UniverRuntime): ActiveValidationTarget | null {
  const target = resolveActiveValidationTarget(runtime)
  return target && shouldShowDataValidationDropdown(target.validation) ? target : null
}

export function openActiveDataValidationDropdown(
  runtime: UniverRuntime,
  event?: Pick<ReactPointerEvent<HTMLButtonElement>, 'preventDefault' | 'stopPropagation'>,
): void {
  event?.preventDefault()
  event?.stopPropagation()
  const target = resolveActiveDropdownTarget(runtime)
  if (!target) return
  void runtime.univerAPI.executeCommand(SHOW_DROPDOWN_COMMAND, {
    unitId: target.unitId,
    subUnitId: target.sheetId,
    row: target.row,
    column: target.column,
  })
}

function DataValidationDropdownButton({
  runtime,
}: {
  readonly runtime: UniverRuntime
}): React.JSX.Element {
  return (
    <button
      type="button"
      className="data-validation-dropdown-button"
      aria-label="Open dropdown"
      tabIndex={-1}
      onPointerDown={(event) => openActiveDataValidationDropdown(runtime, event)}
    >
      <svg aria-hidden="true" viewBox="0 0 10 10">
        <path d="M1.5 3.25 5 6.75l3.5-3.5Z" />
      </svg>
    </button>
  )
}

function DataValidationInputMessage({
  runtime,
}: {
  readonly runtime: UniverRuntime
}): React.JSX.Element | null {
  const target = resolveActiveValidationTarget(runtime)
  const message = getDataValidationInputMessage(target?.validation)
  if (!message) return null
  return (
    <div className="data-validation-input-message" role="note">
      {message.title ? (
        <div className="data-validation-input-message-title">{message.title}</div>
      ) : null}
      {message.prompt ? <div>{message.prompt}</div> : null}
    </div>
  )
}

/**
 * Excel renders validation UI as active-cell chrome: the dropdown is attached
 * outside the right edge and the optional input message starts at the cell's
 * horizontal midpoint below it. Anchored DOM layers follow scroll and zoom.
 */
export function installActiveCellDataValidationChrome(runtime: UniverRuntime): {
  dispose(): void
} {
  let dropdownFloating: { dispose(): void } | null = null
  let inputMessageFloating: { dispose(): void } | null = null
  let targetKey = ''
  let disposed = false

  const componentDisposables = [
    runtime.univerAPI.registerComponent(DROPDOWN_COMPONENT_KEY, () => (
      <DataValidationDropdownButton runtime={runtime} />
    )),
    runtime.univerAPI.registerComponent(INPUT_MESSAGE_COMPONENT_KEY, () => (
      <DataValidationInputMessage runtime={runtime} />
    )),
  ]

  const clear = (): void => {
    dropdownFloating?.dispose()
    inputMessageFloating?.dispose()
    dropdownFloating = null
    inputMessageFloating = null
    targetKey = ''
  }

  const sync = (force = false): void => {
    if (disposed) return
    const target = resolveActiveValidationTarget(runtime)
    const inputMessage = getDataValidationInputMessage(target?.validation)
    const showDropdown = target ? shouldShowDataValidationDropdown(target.validation) : false
    const nextKey = target
      ? JSON.stringify([
          target.unitId,
          target.sheetId,
          target.row,
          target.column,
          showDropdown,
          inputMessage?.title ?? '',
          inputMessage?.prompt ?? '',
        ])
      : ''
    if (!force && (dropdownFloating || inputMessageFloating) && nextKey === targetKey) return
    clear()
    if (!target) return
    if (showDropdown) {
      const cellHeight = rangePixelHeight(target)
      const height = Math.min(DROPDOWN_SIZE, Math.max(1, cellHeight))
      dropdownFloating =
        target.worksheet.addFloatDomToRange(
          target.range,
          {
            componentKey: DROPDOWN_COMPONENT_KEY,
            allowTransform: false,
            eventPassThrough: false,
          },
          {
            width: DROPDOWN_SIZE,
            height,
            marginX: '100%',
            marginY: 0,
            verticalOffsetAlign: 'bottom',
          },
          DROPDOWN_COMPONENT_KEY,
        ) ?? null
    }
    if (inputMessage) {
      inputMessageFloating =
        target.worksheet.addFloatDomToRange(
          target.range,
          {
            componentKey: INPUT_MESSAGE_COMPONENT_KEY,
            allowTransform: false,
            eventPassThrough: true,
          },
          {
            width: INPUT_MESSAGE_WIDTH,
            height: dataValidationInputMessageHeight(inputMessage),
            marginX: '50%',
            marginY: '100%',
          },
          INPUT_MESSAGE_COMPONENT_KEY,
        ) ?? null
    }
    if (dropdownFloating || inputMessageFloating) targetKey = nextKey
  }

  const disposables = [
    runtime.univerAPI.addEvent(runtime.univerAPI.Event.SelectionChanged, () => sync()),
    runtime.univerAPI.addEvent(runtime.univerAPI.Event.ActiveSheetChanged, () => sync()),
    runtime.univerAPI.addEvent(runtime.univerAPI.Event.CommandExecuted, ({ id }) => {
      if (id.includes('data-validation')) sync(true)
    }),
  ]
  sync()

  return {
    dispose: () => {
      disposed = true
      clear()
      for (const disposable of disposables) disposable.dispose()
      for (const disposable of componentDisposables) disposable.dispose()
    },
  }
}

export function installActiveCellDataValidationDropdown(runtime: UniverRuntime): {
  dispose(): void
} {
  return installActiveCellDataValidationChrome(runtime)
}
