import { useState } from 'react'
import type { SectionSettings } from '@genoffice/docx-engine'
import type { SectionScope } from '../commands/menu-commands'
import { useI18n, type StringKey } from '../i18n/locale'
import { ApplyToSelect, defaultScope } from './ApplyToSelect'
import { LengthInput } from './LengthInput'
import { useModalKeys } from './modal-keys'

export const MAX_COLUMNS = 12

export function clampColumns(count: unknown): number {
  const n = Math.round(Number(count))
  if (!Number.isFinite(n)) return 1
  return Math.min(MAX_COLUMNS, Math.max(1, n))
}

export interface ColumnsDialogValue {
  columns: number
  /** unequal columns, twips; absent for equal ones */
  colWidths?: number[]
  /** the gap after each unequal column but the last, twips */
  colSpaces?: number[]
  /** the gap between equal columns, twips; absent when it stays the section's */
  colSpace?: number
  lineBetween: boolean
  /** with This point forward: what follows the caret starts in the next column */
  startNewColumn?: boolean
  scope?: SectionScope
}

type Unequal = 'left' | 'right'
type PageBox = Pick<SectionSettings, 'pageWidth' | 'marginLeft' | 'marginRight' | 'colSpace'>

const MIN_COLUMN_TWIPS = 720
const DEFAULT_SPACE_TWIPS = 720

const sum = (list: number[]): number => list.reduce((total, w) => total + w, 0)
const textWidth = (s: PageBox): number => s.pageWidth - s.marginLeft - s.marginRight

/**
 * Word's Left / Right preset: two columns, the narrow one on that side holding
 * a third of the text width less its share of the gap. Null when the page is
 * too narrow for it.
 */
export function unequalPresetWidths(kind: Unequal, section: PageBox): [number, number] | null {
  const content = textWidth(section)
  const space = section.colSpace ?? DEFAULT_SPACE_TWIPS
  const narrow = Math.round((content - 2 * space) / 3)
  const wide = content - space - narrow
  if (narrow < MIN_COLUMN_TWIPS) return null
  return kind === 'left' ? [narrow, wide] : [wide, narrow]
}

/** the width equal columns come to with `space` between them */
export function equalColumnWidth(section: PageBox, columns: number, space: number): number {
  return Math.round((textWidth(section) - space * (columns - 1)) / columns)
}

/**
 * Unequal columns after one width or gap was typed: the last column (the first,
 * when the last is the one typed) takes up the difference so the columns still
 * fill the text width, and the typed value stops where that column would get
 * narrower than Word allows.
 */
export function fitColumns(
  section: PageBox,
  widths: number[],
  spaces: number[],
  edit: { width: number; value: number } | { space: number; value: number },
): { widths: number[]; spaces: number[] } {
  const nextWidths = [...widths]
  const nextSpaces = [...spaces]
  const absorber = 'width' in edit && edit.width === widths.length - 1 ? 0 : widths.length - 1
  if ('width' in edit) nextWidths[edit.width] = 0
  else nextSpaces[edit.space] = 0
  const fixed = sum(nextWidths) - nextWidths[absorber] + sum(nextSpaces)
  const room = textWidth(section) - fixed - MIN_COLUMN_TWIPS
  const floor = 'width' in edit ? MIN_COLUMN_TWIPS : 0
  const typed = Math.max(floor, Math.min(Math.round(edit.value), room))
  if ('width' in edit) nextWidths[edit.width] = typed
  else nextSpaces[edit.space] = typed
  nextWidths[absorber] = textWidth(section) - fixed - typed
  return { widths: nextWidths, spaces: nextSpaces }
}

interface Shape {
  columns: number
  equal: boolean
  widths: number[]
  spaces: number[]
  /** the gap between equal columns */
  space: number
  lineBetween: boolean
}

function shapeOf(section: SectionSettings): Shape {
  const space = section.colSpace ?? DEFAULT_SPACE_TWIPS
  const widths = section.colWidths
  const equal = section.colEqualWidth !== false || widths?.length !== section.columns
  return {
    columns: section.columns,
    equal,
    widths: equal ? [] : widths!,
    spaces: equal ? [] : widths!.slice(0, -1).map((_, i) => section.colSpaces?.[i] ?? space),
    space,
    lineBetween: !!section.colSep,
  }
}

function valueOf(shape: Shape, section: SectionSettings): ColumnsDialogValue {
  const lineBetween = shape.lineBetween && shape.columns > 1
  if (!shape.equal && shape.columns > 1)
    return { columns: shape.columns, colWidths: shape.widths, colSpaces: shape.spaces, lineBetween }
  const kept = shape.space === (section.colSpace ?? DEFAULT_SPACE_TWIPS)
  return { columns: shape.columns, ...(kept ? {} : { colSpace: shape.space }), lineBetween }
}

const PRESETS: Array<[number, StringKey]> = [
  [1, 'ribbonOneColumn'],
  [2, 'ribbonTwoColumns'],
  [3, 'ribbonThreeColumns'],
]

const UNEQUAL: Array<[Unequal, StringKey]> = [
  ['left', 'ribbonMarginLeft'],
  ['right', 'ribbonMarginRight'],
]

/** Word's Format > Columns: a preset or a number of columns for the section, each column's width and gap, and the line between them */
export function ColumnsDialog({
  section,
  multiSection,
  canApplyForward,
  onApply,
  onClose,
}: {
  section: SectionSettings
  /** the document has several sections and the dialog acts on the one with the caret */
  multiSection: boolean
  /** a section break can go in at the caret and be acted on at once */
  canApplyForward: boolean
  onApply: (value: ColumnsDialogValue) => void
  onClose: () => void
}) {
  const { t } = useI18n()
  const keys = useModalKeys(onClose)
  const [count, setCount] = useState(String(section.columns))
  const [shape, setShape] = useState<Shape>(() => shapeOf(section))
  const [scope, setScope] = useState<SectionScope>(defaultScope(multiSection))
  const [startNew, setStartNew] = useState(false)
  const picked = clampColumns(count)
  const unequal = !shape.equal && picked > 1
  const preset = (kind: Unequal) => unequalPresetWidths(kind, { ...section, colSpace: shape.space })
  const isPreset = (kind: Unequal): boolean => {
    const widths = preset(kind)
    return (
      unequal &&
      picked === 2 &&
      shape.spaces[0] === shape.space &&
      !!widths?.every((w, i) => Math.abs(w - shape.widths[i]) <= 1)
    )
  }
  const pickCount = (next: string) => {
    setCount(next)
    const columns = clampColumns(next)
    // a width list made for another number of columns does not carry over
    setShape((s) => ({ ...s, columns, equal: s.equal || columns !== s.widths.length }))
  }
  const maxEqualSpace =
    picked > 1
      ? Math.max(0, Math.floor((textWidth(section) - MIN_COLUMN_TWIPS * picked) / (picked - 1)))
      : 0
  const setEqual = (equal: boolean) =>
    setShape((s) =>
      equal
        ? { ...s, equal }
        : {
            ...s,
            equal,
            widths: Array.from({ length: picked }, () =>
              equalColumnWidth(section, picked, s.space),
            ),
            spaces: Array.from({ length: picked - 1 }, () => s.space),
          },
    )
  const widthLabel = (i: number) => `${t('ribbonPbWidth')} ${i + 1}`
  const spaceLabel = (i: number) => `${t('appCharSpacing')} ${i + 1}`
  return (
    <div
      className="modal-backdrop"
      ref={keys.ref}
      onKeyDown={keys.onKeyDown}
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
    >
      <form
        className="modal gs-form tco-dialog hlc-dialog columns-dialog"
        noValidate
        onSubmit={(e) => {
          e.preventDefault()
          onClose()
          const value = valueOf({ ...shape, columns: picked }, section)
          const wider = scope !== defaultScope(multiSection)
          const forward = scope === 'forward' && startNew ? { startNewColumn: true } : {}
          if (wider || JSON.stringify(value) !== JSON.stringify(valueOf(shapeOf(section), section)))
            onApply(wider || multiSection ? { ...value, ...forward, scope } : value)
        }}
      >
        <h2>{t('ribbonColumns')}</h2>
        <div className="doc-dialog-caption">{t('ribbonColPresets')}</div>
        {PRESETS.map(([n, labelKey]) => (
          <label key={n} className="tco-check">
            <input
              type="radio"
              name="columns"
              checked={!unequal && picked === n}
              onChange={() => {
                setCount(String(n))
                setShape((s) => ({ ...s, columns: n, equal: true }))
              }}
            />
            {t(labelKey)}
          </label>
        ))}
        {UNEQUAL.map(([kind, labelKey]) => (
          <label key={kind} className="tco-check">
            <input
              type="radio"
              name="columns"
              value={kind}
              checked={isPreset(kind)}
              disabled={!preset(kind)}
              onChange={() => {
                setCount('2')
                setShape((s) => ({
                  ...s,
                  columns: 2,
                  equal: false,
                  widths: preset(kind)!,
                  spaces: [s.space],
                }))
              }}
            />
            {t(labelKey)}
          </label>
        ))}
        <div className="modal-row margin-row">
          <label>
            {t('ribbonColNumber')}
            <input
              type="number"
              min={1}
              max={MAX_COLUMNS}
              value={count}
              onChange={(e) => pickCount(e.target.value)}
            />
          </label>
        </div>
        <label className="tco-check">
          <input
            type="checkbox"
            name="lineBetween"
            checked={shape.lineBetween && picked > 1}
            disabled={picked < 2}
            onChange={(e) => setShape((s) => ({ ...s, lineBetween: e.target.checked }))}
          />
          {t('ribbonColLineBetween')}
        </label>
        <div className="doc-dialog-caption">{t('ribbonColWidthSpacing')}</div>
        <div className="columns-dialog-cols">
          <span>{t('ribbonColNo')}</span>
          <span>{t('ribbonPbWidth')}</span>
          <span>{t('appCharSpacing')}</span>
          {Array.from({ length: picked }, (_, i) => [
            <span key={`n${i}`}>{i + 1}:</span>,
            <LengthInput
              key={`w${i}`}
              value={unequal ? shape.widths[i] : equalColumnWidth(section, picked, shape.space)}
              min={MIN_COLUMN_TWIPS}
              ariaLabel={widthLabel(i)}
              disabled={!unequal}
              live
              onCommit={(twips) =>
                setShape((s) => ({
                  ...s,
                  ...fitColumns(section, s.widths, s.spaces, { width: i, value: twips ?? 0 }),
                }))
              }
            />,
            i < picked - 1 ? (
              <LengthInput
                key={`s${i}`}
                value={unequal ? shape.spaces[i] : shape.space}
                min={0}
                max={unequal ? undefined : maxEqualSpace}
                ariaLabel={spaceLabel(i)}
                // equal columns share one gap, set on the first row
                disabled={!unequal && i > 0}
                live
                onCommit={(twips) =>
                  setShape((s) =>
                    unequal
                      ? {
                          ...s,
                          ...fitColumns(section, s.widths, s.spaces, {
                            space: i,
                            value: twips ?? 0,
                          }),
                        }
                      : { ...s, space: Math.min(twips ?? 0, maxEqualSpace) },
                  )
                }
              />
            ) : (
              <span key={`s${i}`} />
            ),
          ])}
        </div>
        <label className="tco-check">
          <input
            type="checkbox"
            name="equalWidth"
            checked={!unequal}
            disabled={picked < 2}
            onChange={(e) => setEqual(e.target.checked)}
          />
          {t('ribbonColEqualWidth')}
        </label>
        <ApplyToSelect
          scope={scope}
          multiSection={multiSection}
          canApplyForward={canApplyForward}
          onChange={setScope}
        />
        <label className="tco-check">
          <input
            type="checkbox"
            name="startNewColumn"
            checked={startNew && scope === 'forward'}
            disabled={scope !== 'forward'}
            onChange={(e) => setStartNew(e.target.checked)}
          />
          {t('ribbonColStartNew')}
        </label>
        <div className="modal-actions">
          <span />
          <div className="tco-actions-right">
            <button type="button" onClick={onClose}>
              {t('ribbonCancel')}
            </button>
            <button type="submit" className="primary">
              {t('ribbonOk')}
            </button>
          </div>
        </div>
      </form>
    </div>
  )
}
