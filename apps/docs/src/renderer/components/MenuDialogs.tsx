import { useEffect, useRef, useState, type ReactNode } from 'react'
import { DEFAULT_MATH_BREAKS, type LineNumbering, type MathPr } from '@genoffice/docx-engine'
import type { DocMoveResult, TemplateEntry } from '../../shared/ipc'
import { originalPictureSizeOf } from '../editor/picture-bytes'
import type { PictureGeometry, PicturePreset } from '../editor/picture-props'
import {
  MAX_SORT_KEYS,
  sortFields,
  type SortBy,
  type SortOptions,
  type SortScope,
} from '../editor/sort-paragraphs'
import { TEXT_DIRECTION_APPLY_TO, type TextDirectionContext } from '../commands/text-direction'
import { useI18n, type StringKey } from '../i18n/locale'
import { formatNoteNumber } from '../note-format'
import type { ReduceResult } from '../reduce-file-size'
import { REDUCE_PPI, type ReducePpi } from '../reduce-pictures'
import { MATH_FONT_CANDIDATES, WORD_MATH_FONT } from '../editor/math-defaults'
import { useSystemFontFamilies } from '../system-fonts'
import { pxToTwips, twipsToPx } from '../units'
import { IconDoc } from './icons'
import { LengthInput } from './LengthInput'
import { useModalKeys } from './modal-keys'

/** the small one-form dialogs the menus open: a title, the fields, Cancel and OK */
export function FormDialog({
  title,
  name,
  okDisabled,
  onSubmit,
  onClose,
  children,
}: {
  title: string
  name: string
  okDisabled?: boolean
  onSubmit: () => void
  onClose: () => void
  children: ReactNode
}) {
  const { t } = useI18n()
  const keys = useModalKeys(onClose)
  return (
    <div
      className="modal-backdrop"
      ref={keys.ref}
      onKeyDown={keys.onKeyDown}
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
    >
      <form
        className={`modal gs-form tco-dialog hlc-dialog ${name}`}
        noValidate
        onSubmit={(e) => {
          e.preventDefault()
          onSubmit()
        }}
      >
        <h2>{title}</h2>
        {children}
        <div className="modal-actions">
          <span />
          <div className="tco-actions-right">
            <button type="button" onClick={onClose}>
              {t('ribbonCancel')}
            </button>
            <button type="submit" className="primary" disabled={okDisabled}>
              {t('ribbonOk')}
            </button>
          </div>
        </div>
      </form>
    </div>
  )
}

const RENAME_ERRORS: Record<Extract<DocMoveResult, { ok: false }>['error'], StringKey> = {
  badName: 'ribbonRenameBadName',
  exists: 'ribbonRenameExists',
  missing: 'ribbonMoveFailed',
  failed: 'ribbonMoveFailed',
}

export const moveErrorKey = (result: Extract<DocMoveResult, { ok: false }>) =>
  RENAME_ERRORS[result.error]

/** Word's File > Rename…: the file's name without its extension; stays open on a name that cannot be used */
export function RenameDialog({
  fileName,
  onRename,
  onClose,
}: {
  fileName: string
  onRename: (name: string) => Promise<DocMoveResult>
  onClose: () => void
}) {
  const { t } = useI18n()
  const [name, setName] = useState(fileName.replace(/\.[^.]+$/, ''))
  const [error, setError] = useState<StringKey | null>(null)
  return (
    <FormDialog
      title={t('ribbonRenameTitle')}
      name="rename-dialog"
      okDisabled={!name.trim()}
      onClose={onClose}
      onSubmit={() =>
        void onRename(name).then((result) =>
          result.ok ? onClose() : setError(moveErrorKey(result)),
        )
      }
    >
      <label>
        {t('ribbonStyleDialogName')}
        <input name="name" value={name} onChange={(e) => setName(e.target.value)} />
      </label>
      {error && <div className="modal-error">{t(error)}</div>}
    </FormDialog>
  )
}

const NUMBER_TYPES = ['decimal', 'lowerLetter', 'upperLetter', 'lowerRoman', 'upperRoman'] as const

/** Word's Insert > Number: a number written out in the chosen numbering type */
export function NumberDialog({
  onInsert,
  onClose,
}: {
  onInsert: (text: string) => void
  onClose: () => void
}) {
  const { t } = useI18n()
  const [value, setValue] = useState('1')
  const [type, setType] = useState<(typeof NUMBER_TYPES)[number]>('decimal')
  const number = Number(value)
  const valid = Number.isInteger(number) && number >= 1 && number <= 3999
  return (
    <FormDialog
      title={t('ribbonNumber')}
      name="number-dialog"
      okDisabled={!valid}
      onClose={onClose}
      onSubmit={() => {
        onClose()
        onInsert(formatNoteNumber(type, number))
      }}
    >
      <label>
        {t('ribbonNumber')}
        <input
          name="number"
          inputMode="numeric"
          value={value}
          onChange={(e) => setValue(e.target.value)}
        />
      </label>
      <label>
        {t('ribbonNumberType')}
        <select
          name="type"
          value={type}
          onChange={(e) => setType(e.target.value as (typeof NUMBER_TYPES)[number])}
        >
          {NUMBER_TYPES.map((code) => (
            <option key={code} value={code}>
              {[1, 2, 3].map((n) => formatNoteNumber(code, n)).join(', ')}, …
            </option>
          ))}
        </select>
      </label>
    </FormDialog>
  )
}

const SORT_TYPES: ReadonlyArray<{ by: SortBy; label: StringKey }> = [
  { by: 'text', label: 'ribbonGroupText' },
  { by: 'number', label: 'ribbonSortNumber' },
  { by: 'date', label: 'ribbonFieldDate' },
]

const SEPARATORS = [
  { id: 'tab', label: 'ribbonSortTabs', text: '\t' },
  { id: 'comma', label: 'ribbonSortCommas', text: ',' },
  { id: 'other', label: 'ribbonSortOther', text: null },
] as const satisfies ReadonlyArray<{ id: string; label: StringKey; text: string | null }>

type SeparatorId = (typeof SEPARATORS)[number]['id']

const NO_KEY = -1

/** the rows of fields the keys choose from: a table's cells, or each paragraph's parts */
export function sortFieldRows(scope: SortScope | null, separator: string): string[][] {
  if (!scope) return []
  if (scope.kind === 'table') return scope.rows
  return scope.texts.map((text) => sortFields(text, separator).slice(1))
}

/**
 * The keys the dialog lists, by field number. A table's are its columns; a
 * paragraph is one key itself and, once the separator parts it, a key per
 * field. With a header row the fields go by the names it gives them.
 */
export function sortKeyChoices(
  scope: SortScope | null,
  separator: string,
  header: boolean,
): Array<{ field: number; name: string | null }> {
  const rows = sortFieldRows(scope, separator)
  const count = rows.reduce((most, row) => Math.max(most, row.length), 0)
  const table = scope?.kind === 'table'
  const names = header ? (rows[0] ?? []) : []
  const fields =
    table || count > 1
      ? Array.from({ length: count }, (_, index) => ({
          field: index + 1,
          name: names[index]?.trim() || null,
        }))
      : []
  return table ? fields : [{ field: 0, name: null }, ...fields]
}

interface SortKeyDraft {
  field: number
  by: SortBy
  descending: boolean
}

/**
 * Word's Sort: up to three keys, each a column of the table at the selection
 * or a field of the selected paragraphs, with a type and a direction
 */
export function SortDialog({
  loadScope,
  onSort,
  onClose,
}: {
  /** what the selection holds as the dialog opens; null for nothing Sort can order: OK then says so */
  loadScope: () => SortScope | null
  onSort: (options: SortOptions) => void
  onClose: () => void
}) {
  const { t } = useI18n()
  const [scope] = useState(loadScope)
  const table = scope?.kind === 'table'
  const [keys, setKeys] = useState<SortKeyDraft[]>(() =>
    Array.from({ length: MAX_SORT_KEYS }, (_, index) => ({
      field: index > 0 ? NO_KEY : table ? 1 : 0,
      by: 'text' as const,
      descending: false,
    })),
  )
  const [header, setHeader] = useState(false)
  const [separatorId, setSeparatorId] = useState<SeparatorId>('tab')
  const [other, setOther] = useState('-')
  const [caseSensitive, setCaseSensitive] = useState(false)
  const separator = SEPARATORS.find((entry) => entry.id === separatorId)?.text ?? other
  const choices = sortKeyChoices(scope, separator, header)
  const offered = (field: number) => choices.some((choice) => choice.field === field)
  const setKey = (index: number, patch: Partial<SortKeyDraft>) =>
    setKeys((prev) => prev.map((key, i) => (i === index ? { ...key, ...patch } : key)))
  const choiceName = (choice: (typeof choices)[number]) =>
    choice.name ??
    (choice.field === 0
      ? t('appStatParagraphs')
      : t(table ? 'ribbonColumnN' : 'ribbonSortFieldN', { n: choice.field }))
  return (
    <FormDialog
      title={t('ribbonSort')}
      name="sort-dialog"
      onClose={onClose}
      onSubmit={() => {
        onClose()
        onSort({
          // a key whose field the separator no longer makes falls back to the first one offered
          keys: keys
            .filter((key, index) => index === 0 || key.field !== NO_KEY)
            .map((key) => ({ ...key, field: offered(key.field) ? key.field : choices[0].field })),
          header,
          separator,
          caseSensitive,
        })
      }}
    >
      {keys.map((key, index) => {
        const unused = key.field === NO_KEY
        // Word offers a Then by once the key above it is set
        const locked = index > 0 && keys[index - 1].field === NO_KEY
        return (
          <fieldset key={index} disabled={locked}>
            <legend>{t(index === 0 ? 'ribbonSortBy' : 'ribbonSortThenBy')}</legend>
            <div className="sort-key">
              <select
                name={`key${index + 1}`}
                value={offered(key.field) || unused ? key.field : choices[0].field}
                onChange={(e) => {
                  const field = Number(e.target.value)
                  setKeys((prev) =>
                    prev.map((entry, i) =>
                      i === index
                        ? { ...entry, field }
                        : // a key under one switched off has nothing to follow
                          i > index && field === NO_KEY
                          ? { ...entry, field: NO_KEY }
                          : entry,
                    ),
                  )
                }}
              >
                {index > 0 && <option value={NO_KEY}>{t('ribbonSortNoKey')}</option>}
                {choices.map((choice) => (
                  <option key={choice.field} value={choice.field}>
                    {choiceName(choice)}
                  </option>
                ))}
              </select>
              <select
                name={`type${index + 1}`}
                aria-label={t('ribbonSortType')}
                value={key.by}
                disabled={unused}
                onChange={(e) => setKey(index, { by: e.target.value as SortBy })}
              >
                {SORT_TYPES.map((type) => (
                  <option key={type.by} value={type.by}>
                    {t(type.label)}
                  </option>
                ))}
              </select>
              <span className="sort-order">
                {[false, true].map((down) => (
                  <label key={String(down)} className="tco-check">
                    <input
                      type="radio"
                      name={`order${index + 1}`}
                      value={down ? 'descending' : 'ascending'}
                      checked={key.descending === down}
                      disabled={unused}
                      onChange={() => setKey(index, { descending: down })}
                    />
                    {t(down ? 'ribbonSortDescending' : 'ribbonSortAscending')}
                  </label>
                ))}
              </span>
            </div>
          </fieldset>
        )
      })}
      <fieldset>
        <legend>{t('ribbonSortMyListHas')}</legend>
        <div className="sort-choices">
          {[true, false].map((on) => (
            <label key={String(on)} className="tco-check">
              <input
                type="radio"
                name="header"
                value={on ? 'header' : 'none'}
                checked={header === on}
                onChange={() => setHeader(on)}
              />
              {t(on ? 'ribbonSortHeaderRow' : 'ribbonSortNoHeaderRow')}
            </label>
          ))}
        </div>
      </fieldset>
      <fieldset>
        <legend>{t('ribbonGroupHfOptions')}</legend>
        {/* a table's fields are its cells: the separators are Word's greyed ones there */}
        <div className="sort-choices">
          <span className="sort-caption">{t('ribbonSortSeparator')}</span>
          {SEPARATORS.map((entry) => (
            <label key={entry.id} className="tco-check">
              <input
                type="radio"
                name="separator"
                value={entry.id}
                checked={separatorId === entry.id}
                disabled={table}
                onChange={() => setSeparatorId(entry.id)}
              />
              {t(entry.label)}
            </label>
          ))}
          <input
            name="other"
            className="sort-other"
            aria-label={t('ribbonSortOther')}
            maxLength={1}
            value={other}
            disabled={table || separatorId !== 'other'}
            onChange={(e) => setOther(e.target.value)}
          />
        </div>
        <label className="tco-check">
          <input
            type="checkbox"
            name="caseSensitive"
            checked={caseSensitive}
            onChange={(e) => setCaseSensitive(e.target.checked)}
          />
          {t('ribbonSortCaseSensitive')}
        </label>
      </fieldset>
    </FormDialog>
  )
}

/** Word's Format > Text Direction…: how the text of the target at the caret runs */
export function TextDirectionDialog({
  context: { target, choices, current },
  onApply,
  onClose,
}: {
  context: TextDirectionContext
  onApply: (direction: string) => void
  onClose: () => void
}) {
  const { t } = useI18n()
  // the target's own direction; nothing for one the dialog does not offer
  const [direction, setDirection] = useState(current)
  return (
    <FormDialog
      title={t('ribbonTextDirection')}
      name="text-direction-dialog"
      okDisabled={!direction}
      onClose={onClose}
      onSubmit={() => {
        onClose()
        if (direction && direction !== current) onApply(direction)
      }}
    >
      {choices.map((option) => (
        <label key={option.value} className="tco-check">
          <input
            type="radio"
            name="direction"
            value={option.value}
            checked={direction === option.value}
            onChange={() => setDirection(option.value)}
          />
          {t(option.label)}
        </label>
      ))}
      <div className="modal-row margin-row">
        <label>
          {t('ribbonPaperApplyTo')}
          <select name="applyTo">
            <option value={target}>{t(TEXT_DIRECTION_APPLY_TO[target])}</option>
          </select>
        </label>
      </div>
    </FormDialog>
  )
}

/** a menu label as a dialog title: without the ellipsis that says a dialog follows */
const asTitle = (label: string) => label.replace(/(?:\u2026|\.{3})$/, '')

/** Word's View Alt Text…: the description of the selected picture for readers who cannot see it */
export function AltTextDialog({
  current,
  onApply,
  onClose,
}: {
  current: string
  onApply: (alt: string) => void
  onClose: () => void
}) {
  const { t } = useI18n()
  const [alt, setAlt] = useState(current)
  return (
    <FormDialog
      title={t('ribbonTabAltText')}
      name="alt-text-dialog"
      onClose={onClose}
      onSubmit={() => {
        onClose()
        onApply(alt)
      }}
    >
      <label>
        {t('ribbonAltDescription')}
        <textarea
          name="alt"
          rows={4}
          autoFocus
          value={alt}
          onChange={(e) => setAlt(e.target.value)}
        />
      </label>
      <p className="modal-desc">{t('ribbonPictureAltHint')}</p>
    </FormDialog>
  )
}

const PICTURE_TWIPS_MIN = 14
const PICTURE_TWIPS_MAX = 31680
const OFFSET_TWIPS_MAX = 31680

const SCALE_MIN = 1
const SCALE_MAX = 10000
const PRESET_H: Array<{ value: PicturePreset['h']; label: StringKey }> = [
  { value: 'left', label: 'appAlignLeft' },
  { value: 'center', label: 'appAlignCenter' },
  { value: 'right', label: 'appAlignRight' },
]
const PRESET_V: Array<{ value: PicturePreset['v']; label: StringKey }> = [
  { value: 'top', label: 'ribbonMarginTop' },
  { value: 'center', label: 'ribbonPgCentered' },
  { value: 'bottom', label: 'ribbonMarginBottom' },
]

/** a size as a percentage of what it is scaled from, the way the field shows it */
const percentOf = (px: number, base: number) => String(Math.round((px / base) * 1000) / 10)

/** Word's Size and Position…: the selected picture's size, scale, turn and place */
export function SizePositionDialog({
  current,
  onApply,
  onClose,
}: {
  current: PictureGeometry
  onApply: (next: PictureGeometry) => void
  onClose: () => void
}) {
  const { t } = useI18n()
  const [size, setSize] = useState({ w: current.widthPx, h: current.heightPx })
  const [lock, setLock] = useState(true)
  const [rotation, setRotation] = useState(String(current.rotDeg))
  const [offset, setOffset] = useState(current.offsetPx)
  const [original, setOriginal] = useState<{ w: number; h: number } | null>(null)
  const [fromOriginal, setFromOriginal] = useState(true)
  // a percentage being typed stays as typed until it is a number the picture can take
  const [typed, setTyped] = useState<{ side: 'w' | 'h'; text: string } | null>(null)
  const opened = current.preset
  const [preset, setPreset] = useState<PicturePreset | null>(opened)
  useEffect(() => {
    let live = true
    if (current.dataUrl)
      originalPictureSizeOf(current.dataUrl).then(
        (natural) => {
          if (live && natural.width > 0 && natural.height > 0)
            setOriginal({ w: natural.width, h: natural.height })
        },
        () => {},
      )
    return () => {
      live = false
    }
  }, [current.dataUrl])
  const turn = Number(rotation)
  const valid = rotation.trim() !== '' && Number.isFinite(turn) && turn >= -3600 && turn <= 3600
  // Word scales from the picture's own size, or from the size it had when the dialog opened
  const base = fromOriginal && original ? original : { w: current.widthPx, h: current.heightPx }
  const setSide = (side: 'w' | 'h', px: number) =>
    // locked: the other side follows at the ratio the picture has now
    setSize((was) =>
      side === 'w'
        ? { w: px, h: lock ? (px * was.h) / was.w : was.h }
        : { w: lock ? (px * was.w) / was.h : was.w, h: px },
    )
  const resize = (side: 'w' | 'h', twips: number | null) => {
    if (!twips) return
    setTyped(null)
    setSide(side, twipsToPx(twips))
  }
  const scale = (side: 'w' | 'h', text: string) => {
    setTyped({ side, text })
    const pct = Number(text)
    if (text.trim() === '' || !Number.isFinite(pct) || pct < SCALE_MIN || pct > SCALE_MAX) return
    const px = (base[side] * pct) / 100
    if (pxToTwips(px) >= PICTURE_TWIPS_MIN && pxToTwips(px) <= PICTURE_TWIPS_MAX) setSide(side, px)
  }
  const scaleField = (side: 'w' | 'h', label: string) => (
    <>
      <span>{label}</span>
      <span className="size-scale-field">
        <input
          name={side === 'w' ? 'scaleWidth' : 'scaleHeight'}
          inputMode="decimal"
          aria-label={`${t('ribbonSizeScale')}: ${label}`}
          value={typed?.side === side ? typed.text : percentOf(size[side], base[side])}
          onChange={(e) => scale(side, e.target.value)}
          onBlur={() => setTyped(null)}
        />
        %
      </span>
    </>
  )
  return (
    <FormDialog
      title={asTitle(t('appSizeAndPosition'))}
      name="size-position-dialog"
      okDisabled={!valid}
      onClose={onClose}
      onSubmit={() => {
        onClose()
        const moved = preset && (preset.h !== opened?.h || preset.v !== opened?.v)
        onApply({
          ...current,
          widthPx: size.w,
          heightPx: size.h,
          rotDeg: turn,
          offsetPx: preset ? null : offset,
          preset: moved ? preset : null,
        })
      }}
    >
      <fieldset>
        <legend>{t('ribbonGroupSize')}</legend>
        <div className="tco-grid">
          <span>{t('ribbonPicHeight')}</span>
          <LengthInput
            value={pxToTwips(size.h)}
            min={PICTURE_TWIPS_MIN}
            max={PICTURE_TWIPS_MAX}
            ariaLabel={t('ribbonPicHeight')}
            live
            onCommit={(twips) => resize('h', twips)}
          />
          <span>{t('ribbonPicWidth')}</span>
          <LengthInput
            value={pxToTwips(size.w)}
            min={PICTURE_TWIPS_MIN}
            max={PICTURE_TWIPS_MAX}
            ariaLabel={t('ribbonPicWidth')}
            live
            onCommit={(twips) => resize('w', twips)}
          />
          <span>{t('ribbonArrangeRotate')}</span>
          <input
            name="rotation"
            disabled={!current.canRotate}
            inputMode="numeric"
            aria-label={t('ribbonArrangeRotate')}
            value={rotation}
            onChange={(e) => setRotation(e.target.value)}
          />
        </div>
      </fieldset>
      <fieldset>
        <legend>{t('ribbonSizeScale')}</legend>
        <div className="tco-grid">
          {scaleField('h', t('ribbonPicHeight'))}
          {scaleField('w', t('ribbonPicWidth'))}
        </div>
        <label className="tco-check">
          <input
            type="checkbox"
            name="lock"
            checked={lock}
            onChange={(e) => setLock(e.target.checked)}
          />
          {t('ribbonLockAspectRatio')}
        </label>
        <label className="tco-check">
          <input
            type="checkbox"
            name="fromOriginal"
            disabled={!original}
            checked={fromOriginal && !!original}
            onChange={(e) => {
              setTyped(null)
              setFromOriginal(e.target.checked)
            }}
          />
          {t('ribbonSizeScaleFromOriginal')}
        </label>
      </fieldset>
      {(offset || current.canPreset) && (
        <fieldset>
          <legend>{t('ribbonPosition')}</legend>
          {current.canPreset && (
            <>
              <label className="tco-check">
                <input
                  type="radio"
                  name="place"
                  value="preset"
                  checked={!!preset}
                  onChange={() => setPreset(opened ?? { h: 'left', v: 'top' })}
                />
                {t('ribbonSizePosAlignment')}
              </label>
              <div className="tco-grid">
                <span>{t('ribbonWatermarkHorizontal')}</span>
                <select
                  name="presetH"
                  disabled={!preset}
                  aria-label={t('ribbonWatermarkHorizontal')}
                  value={preset?.h ?? 'left'}
                  onChange={(e) =>
                    setPreset((was) => was && { ...was, h: e.target.value as PicturePreset['h'] })
                  }
                >
                  {PRESET_H.map((o) => (
                    <option key={o.value} value={o.value}>
                      {t(o.label)}
                    </option>
                  ))}
                </select>
                <span>{t('ribbonDocVertical')}</span>
                <select
                  name="presetV"
                  disabled={!preset}
                  aria-label={t('ribbonDocVertical')}
                  value={preset?.v ?? 'top'}
                  onChange={(e) =>
                    setPreset((was) => was && { ...was, v: e.target.value as PicturePreset['v'] })
                  }
                >
                  {PRESET_V.map((o) => (
                    <option key={o.value} value={o.value}>
                      {t(o.label)}
                    </option>
                  ))}
                </select>
              </div>
              <label className="tco-check">
                <input
                  type="radio"
                  name="place"
                  value="offset"
                  // a picture at a preset place has no other place to go back to
                  disabled={!!opened && !offset}
                  checked={!preset}
                  onChange={() => setPreset(null)}
                />
                {t(offset || current.floating ? 'ribbonSizePosAbsolute' : 'appWrapInline')}
              </label>
            </>
          )}
          {offset && (
            <div className="tco-grid">
              <span>{t('ribbonWatermarkHorizontal')}</span>
              <LengthInput
                value={pxToTwips(offset.x)}
                min={-OFFSET_TWIPS_MAX}
                max={OFFSET_TWIPS_MAX}
                ariaLabel={t('ribbonWatermarkHorizontal')}
                live
                disabled={!!preset}
                onCommit={(twips) =>
                  twips !== null && setOffset((was) => was && { ...was, x: twipsToPx(twips) })
                }
              />
              <span>{t('ribbonDocVertical')}</span>
              <LengthInput
                value={pxToTwips(offset.y)}
                min={-OFFSET_TWIPS_MAX}
                max={OFFSET_TWIPS_MAX}
                ariaLabel={t('ribbonDocVertical')}
                live
                disabled={!!preset}
                onCommit={(twips) =>
                  twips !== null && setOffset((was) => was && { ...was, y: twipsToPx(twips) })
                }
              />
            </div>
          )}
        </fieldset>
      )}
    </FormDialog>
  )
}

const REDUCE_LABELS: Record<ReducePpi, StringKey> = {
  220: 'ribbonReducePrint',
  150: 'ribbonReduceScreen',
  96: 'ribbonReduceEmail',
}

const megabytes = (bytes: number) =>
  bytes >= 1024 * 1024
    ? `${(bytes / (1024 * 1024)).toFixed(1)} MB`
    : `${Math.max(1, Math.round(bytes / 1024))} KB`

/**
 * Word's File > Reduce File Size: the pictures stored at more pixels than the
 * chosen quality needs are resampled and the document is saved; the dialog
 * stays open to say what happened.
 */
export function ReduceFileSizeDialog({
  selectedPictures,
  onReduce,
  onClose,
}: {
  /** how many pictures the selection held when the dialog opened */
  selectedPictures: number
  onReduce: (
    ppi: ReducePpi | null,
    deleteCropped: boolean,
    selectedOnly: boolean,
  ) => Promise<ReduceResult>
  onClose: () => void
}) {
  const { t } = useI18n()
  const [ppi, setPpi] = useState<ReducePpi | null>(150)
  const [deleteCropped, setDeleteCropped] = useState(false)
  const [selectedOnly, setSelectedOnly] = useState(false)
  const [busy, setBusy] = useState(false)
  const [result, setResult] = useState<ReduceResult | null>(null)
  const close = () => {
    if (!busy) onClose()
  }
  const keys = useModalKeys(close)
  const untouched = result && 'outcome' in result ? result.outcome.untouched : 0
  return (
    <div
      className="modal-backdrop"
      ref={keys.ref}
      onKeyDown={keys.onKeyDown}
      onMouseDown={(e) => e.target === e.currentTarget && close()}
    >
      <form
        className="modal gs-form tco-dialog hlc-dialog reduce-size-dialog"
        noValidate
        onSubmit={(e) => {
          e.preventDefault()
          setBusy(true)
          setResult(null)
          void onReduce(ppi, deleteCropped, selectedOnly)
            .catch((): ReduceResult => ({ status: 'failed' }))
            .then((next) => {
              setResult(next)
              setBusy(false)
            })
        }}
      >
        <h2>{t('ribbonReduceFileSize')}</h2>
        <fieldset>
          <legend>{t('ribbonReduceQuality')}</legend>
          {REDUCE_PPI.map((value) => (
            <label key={value} className="tco-check">
              <input
                type="radio"
                name="ppi"
                value={value}
                disabled={busy}
                checked={ppi === value}
                onChange={() => setPpi(value)}
              />
              {t(REDUCE_LABELS[value])}
            </label>
          ))}
          <label className="tco-check">
            <input
              type="radio"
              name="ppi"
              value="keep"
              disabled={busy}
              checked={ppi === null}
              onChange={() => setPpi(null)}
            />
            {t('ribbonReduceKeepResolution')}
          </label>
        </fieldset>
        <label className="tco-check">
          <input
            type="checkbox"
            name="deleteCropped"
            disabled={busy}
            checked={deleteCropped}
            onChange={(e) => setDeleteCropped(e.target.checked)}
          />
          {t('ribbonReduceDeleteCropped')}
        </label>
        <fieldset>
          <legend>{t('ribbonReduceApplyTo')}</legend>
          <label className="tco-check">
            <input
              type="radio"
              name="scope"
              value="all"
              disabled={busy}
              checked={!selectedOnly}
              onChange={() => setSelectedOnly(false)}
            />
            {t('ribbonReduceAllPictures')}
          </label>
          <label className="tco-check">
            <input
              type="radio"
              name="scope"
              value="selected"
              disabled={busy || selectedPictures === 0}
              checked={selectedOnly}
              onChange={() => setSelectedOnly(true)}
            />
            {t('ribbonReduceSelectedOnly')}
          </label>
        </fieldset>
        <p className="reduce-size-note">{t('ribbonReduceNote')}</p>
        <div className="reduce-size-result" role="status" data-status={result?.status}>
          {busy && t('ribbonReduceWorking')}
          {result?.status === 'done' &&
            t('ribbonReduceDone', {
              n: result.outcome.reduced,
              total: result.outcome.reduced + result.outcome.smallEnough + untouched,
              before: megabytes(result.outcome.bytesBefore),
              after: megabytes(result.outcome.bytesAfter),
            })}
          {result?.status === 'none' &&
            t(
              ppi === null || deleteCropped || selectedOnly
                ? 'ribbonReduceNothing'
                : 'ribbonReduceNone',
            )}
          {result?.status === 'failed' && t('ribbonReduceFailed')}
          {result?.status === 'raced' && t('ribbonReduceRaced')}
          {untouched > 0 && <div>{t('ribbonReduceUntouched', { n: untouched })}</div>}
        </div>
        <div className="modal-actions">
          <span />
          <div className="tco-actions-right">
            <button type="button" disabled={busy} onClick={close}>
              {t('ribbonClose')}
            </button>
            <button
              type="submit"
              className="primary"
              disabled={busy || (ppi === null && !deleteCropped)}
            >
              {t('ribbonReduceAction')}
            </button>
          </div>
        </div>
      </form>
    </div>
  )
}

/**
 * Insert > File for a document that has bookmarks: the whole of it, or only
 * what one bookmark covers (the Range of Word's dialog).
 */
export function InsertFileRangeDialog({
  fileName,
  bookmarks,
  onAnswer,
}: {
  fileName: string
  bookmarks: string[]
  /** a bookmark's name, null for the whole document, undefined to give up */
  onAnswer: (range: string | null | undefined) => void
}) {
  const { t } = useI18n()
  const [range, setRange] = useState('')
  return (
    <FormDialog
      title={t('ribbonInsertFileRangeTitle')}
      name="insert-file-range-dialog"
      onSubmit={() => onAnswer(range === '' ? null : range)}
      onClose={() => onAnswer(undefined)}
    >
      <p className="insert-file-name">{fileName}</p>
      <label className="insert-file-range">
        {t('ribbonInsertFileRange')}
        <select name="range" autoFocus value={range} onChange={(e) => setRange(e.target.value)}>
          <option value="">{t('ribbonInsertFileWhole')}</option>
          {bookmarks.map((name) => (
            <option key={name} value={name}>
              {name}
            </option>
          ))}
        </select>
      </label>
    </FormDialog>
  )
}

const templateTitle = (fileName: string) => fileName.replace(/\.[^.]+$/, '')

/** File > New from Template: the templates folder as a gallery, and a way to any other file */
export function TemplateGalleryDialog({
  onCreate,
  onBrowse,
  onClose,
}: {
  onCreate: (fileName: string) => Promise<unknown>
  onBrowse: () => Promise<unknown>
  onClose: () => void
}) {
  const { t } = useI18n()
  const [templates, setTemplates] = useState<TemplateEntry[] | null>(null)
  const [picked, setPicked] = useState<string | null>(null)
  const [busy, setBusy] = useState(false)
  // the ref turns away a second activation that lands before the re-render
  const busyRef = useRef(false)
  const run = async (start: () => Promise<unknown>) => {
    if (busyRef.current) return
    busyRef.current = true
    setBusy(true)
    try {
      await start()
    } finally {
      busyRef.current = false
      setBusy(false)
    }
  }
  const close = () => {
    if (!busyRef.current) onClose()
  }
  const keys = useModalKeys(close)
  useEffect(() => {
    let live = true
    void window.desktop
      .listTemplates()
      .catch((): TemplateEntry[] => [])
      .then((list) => {
        if (!live) return
        setTemplates(list)
        setPicked(list[0]?.name ?? null)
      })
    return () => {
      live = false
    }
  }, [])
  return (
    <div
      className="modal-backdrop"
      ref={keys.ref}
      onKeyDown={keys.onKeyDown}
      onMouseDown={(e) => e.target === e.currentTarget && close()}
    >
      <form
        className="modal gs-form tco-dialog template-gallery-dialog"
        noValidate
        aria-busy={busy}
        onSubmit={(e) => {
          e.preventDefault()
          if (picked) void run(() => onCreate(picked))
        }}
      >
        <h2>{t('appTemplateGalleryTitle')}</h2>
        <div className="template-gallery" role="listbox" aria-label={t('appTemplateGalleryTitle')}>
          {templates?.length === 0 && (
            <p className="template-gallery-empty">{t('appTemplateGalleryEmpty')}</p>
          )}
          {templates?.map((entry) => (
            <button
              key={entry.name}
              type="button"
              role="option"
              aria-selected={picked === entry.name}
              className="template-tile"
              data-template={entry.name}
              title={entry.name}
              onClick={() => setPicked(entry.name)}
              onDoubleClick={() => void run(() => onCreate(entry.name))}
            >
              <span className="template-thumb">
                {entry.thumbnail ? <img src={entry.thumbnail} alt="" /> : <IconDoc />}
              </span>
              <span className="template-name">{templateTitle(entry.name)}</span>
            </button>
          ))}
        </div>
        <div className="modal-actions">
          <button type="button" name="browse" onClick={() => void run(onBrowse)}>
            {t('appTemplateGalleryBrowse')}
          </button>
          <div className="tco-actions-right">
            <button type="button" onClick={close}>
              {t('ribbonCancel')}
            </button>
            <button type="submit" className="primary" disabled={!picked}>
              {t('appTemplateGalleryCreate')}
            </button>
          </div>
        </div>
      </form>
    </div>
  )
}

const LN_RESTARTS: Array<[LineNumbering['restart'], StringKey]> = [
  ['newPage', 'ribbonLnRestartPage'],
  ['newSection', 'ribbonLnRestartSection'],
  ['continuous', 'ribbonBreakContinuous'],
]

/** Word's Line Numbers dialog: whether the section's lines are numbered, and how */
export function LineNumbersDialog({
  current,
  onApply,
  onClose,
}: {
  /** the section's numbering; undefined when its lines are not numbered */
  current: LineNumbering | undefined
  /** null takes the numbering off */
  onApply: (next: LineNumbering | null) => void
  onClose: () => void
}) {
  const { t } = useI18n()
  const [on, setOn] = useState(!!current)
  const [start, setStart] = useState(String(current?.start ?? 1))
  const [countBy, setCountBy] = useState(String(current?.countBy ?? 1))
  const [distance, setDistance] = useState<number | null>(current?.distance ?? null)
  const [restart, setRestart] = useState<LineNumbering['restart']>(current?.restart ?? 'newPage')
  const whole = (text: string) => (/^\d+$/.test(text.trim()) ? Number(text) : NaN)
  const valid = !on || (whole(start) >= 1 && whole(countBy) >= 1 && whole(countBy) <= 100)
  return (
    <FormDialog
      title={t('ribbonLineNumbers')}
      name="line-numbers-dialog"
      okDisabled={!valid}
      onClose={onClose}
      onSubmit={() => {
        onClose()
        onApply(
          on
            ? {
                countBy: whole(countBy),
                start: whole(start),
                restart,
                ...(distance !== null ? { distance } : {}),
              }
            : null,
        )
      }}
    >
      <label className="tco-check">
        <input type="checkbox" name="on" checked={on} onChange={(e) => setOn(e.target.checked)} />
        {t('ribbonLnAdd')}
      </label>
      <div className="tco-grid">
        <span>{t('ribbonListStartAt')}</span>
        <input
          name="start"
          inputMode="numeric"
          aria-label={t('ribbonListStartAt')}
          disabled={!on}
          value={start}
          onChange={(e) => setStart(e.target.value)}
        />
        <span>{t('ribbonLnFromText')}</span>
        <LengthInput
          value={distance}
          max={31680}
          allowEmpty
          placeholder={t('ribbonAuto')}
          ariaLabel={t('ribbonLnFromText')}
          disabled={!on}
          live
          onCommit={setDistance}
        />
        <span>{t('ribbonLnCountBy')}</span>
        <input
          name="countBy"
          inputMode="numeric"
          aria-label={t('ribbonLnCountBy')}
          disabled={!on}
          value={countBy}
          onChange={(e) => setCountBy(e.target.value)}
        />
      </div>
      <fieldset>
        <legend>{t('ribbonLnNumbering')}</legend>
        {LN_RESTARTS.map(([value, label]) => (
          <label key={value} className="tco-check">
            <input
              type="radio"
              name="restart"
              value={value}
              disabled={!on}
              checked={restart === value}
              onChange={() => setRestart(value)}
            />
            {t(label)}
          </label>
        ))}
      </fieldset>
    </FormDialog>
  )
}

/** Word's Custom Paragraph Spacing: the spacing of the document's default paragraph style, typed by hand */
export function ParagraphSpacingDialog({
  current,
  onApply,
  onClose,
}: {
  current: { beforeTwips: number; afterTwips: number; line: number }
  onApply: (next: { beforeTwips: number; afterTwips: number; line: number }) => void
  onClose: () => void
}) {
  const { t } = useI18n()
  const [before, setBefore] = useState(String(current.beforeTwips / 20))
  const [after, setAfter] = useState(String(current.afterTwips / 20))
  const [line, setLine] = useState(String(current.line))
  const points = (text: string) => (/^\d+(\.\d+)?$/.test(text.trim()) ? Number(text) : NaN)
  const valid =
    points(before) <= 1584 && points(after) <= 1584 && points(line) >= 0.25 && points(line) <= 10
  const field = (name: string, label: string, value: string, set: (v: string) => void) => (
    <>
      <span>{label}</span>
      <input
        name={name}
        inputMode="decimal"
        aria-label={label}
        value={value}
        onChange={(e) => set(e.target.value)}
      />
    </>
  )
  return (
    <FormDialog
      title={asTitle(t('ribbonPsCustom'))}
      name="paragraph-spacing-dialog"
      okDisabled={!valid}
      onClose={onClose}
      onSubmit={() => {
        onClose()
        onApply({
          beforeTwips: Math.round(points(before) * 20),
          afterTwips: Math.round(points(after) * 20),
          line: points(line),
        })
      }}
    >
      <div className="tco-grid">
        {field('before', `${t('ribbonSpaceBefore')} (pt)`, before, setBefore)}
        {field('after', `${t('ribbonSpaceAfter')} (pt)`, after, setAfter)}
        {field('line', t('appLineSpacingLabel'), line, setLine)}
      </div>
    </FormDialog>
  )
}

const EQ_JC: Array<[MathPr['defJc'], StringKey]> = [
  ['centerGroup', 'ribbonEqOptJcGroup'],
  ['left', 'ribbonMarginLeft'],
  ['right', 'ribbonMarginRight'],
  ['center', 'ribbonPgCentered'],
]

const EQ_BRK: Array<[NonNullable<MathPr['brkBin']>, StringKey]> = [
  ['before', 'ribbonEqOptBrkBefore'],
  ['after', 'ribbonEqOptBrkAfter'],
  ['repeat', 'ribbonEqOptBrkRepeat'],
]

const EQ_BREAK_DEFAULTS: Partial<MathPr> = { ...DEFAULT_MATH_BREAKS, wrapRight: false }

const EQ_LIMITS: Array<['intLim' | 'naryLim', StringKey, StringKey]> = [
  ['intLim', 'ribbonEqOptIntSide', 'ribbonEqOptIntCenter'],
  ['naryLim', 'ribbonEqOptNarySide', 'ribbonEqOptNaryCenter'],
]

/** Word's Equation Options: the document's math settings (settings.xml m:mathPr) */
export function EquationOptionsDialog({
  current,
  onApply,
  onClose,
}: {
  current: MathPr
  onApply: (next: MathPr) => void
  onClose: () => void
}) {
  const { t } = useI18n()
  const [next, setNext] = useState(current)
  const set = (patch: Partial<MathPr>) => setNext((prev) => ({ ...prev, ...patch }))
  // an absent break child and Word's value for it are the same setting
  const valueOf = (pr: MathPr, k: keyof MathPr) => pr[k] ?? EQ_BREAK_DEFAULTS[k]
  const changed = (Object.keys(next) as Array<keyof MathPr>).some(
    (k) => valueOf(next, k) !== valueOf(current, k),
  )
  const { families, load: loadFonts } = useSystemFontFamilies()
  const mathFont = next.mathFont ?? WORD_MATH_FONT
  const brkBin = next.brkBin ?? DEFAULT_MATH_BREAKS.brkBin
  const mathFonts = [
    ...new Set([
      WORD_MATH_FONT,
      current.mathFont ?? WORD_MATH_FONT,
      ...MATH_FONT_CANDIDATES.filter((font) => families.includes(font)),
    ]),
  ]
  return (
    <FormDialog
      title={t('ribbonEqOptTitle')}
      name="equation-options-dialog"
      onClose={onClose}
      onSubmit={() => {
        onClose()
        if (changed) onApply(next)
      }}
    >
      <div className="tco-grid">
        <span>{t('ribbonEqOptFont')}</span>
        <select
          name="mathFont"
          value={mathFont}
          aria-label={t('ribbonEqOptFont')}
          onFocus={loadFonts}
          onChange={(e) => set({ mathFont: e.target.value })}
        >
          {mathFonts.map((font) => (
            <option key={font} value={font}>
              {font}
            </option>
          ))}
        </select>
        <span>{t('ribbonEqOptBrkBin')}</span>
        <select
          name="brkBin"
          value={brkBin}
          aria-label={t('ribbonEqOptBrkBin')}
          onChange={(e) => set({ brkBin: e.target.value as MathPr['brkBin'] })}
        >
          {EQ_BRK.map(([value, label]) => (
            <option key={value} value={value}>
              {t(label)}
            </option>
          ))}
        </select>
        <span>{t('ribbonEqOptBrkSub')}</span>
        <select
          name="brkBinSub"
          disabled={brkBin !== 'repeat'}
          value={next.brkBinSub ?? DEFAULT_MATH_BREAKS.brkBinSub}
          aria-label={t('ribbonEqOptBrkSub')}
          onChange={(e) => set({ brkBinSub: e.target.value as MathPr['brkBinSub'] })}
        >
          {['--', '-+', '+-'].map((value) => (
            <option key={value} value={value}>
              {value}
            </option>
          ))}
        </select>
      </div>
      <label className="tco-check">
        <input
          type="checkbox"
          name="smallFrac"
          checked={next.smallFrac === true}
          onChange={(e) => set({ smallFrac: e.target.checked })}
        />
        {t('ribbonEqOptSmallFrac')}
      </label>
      <fieldset>
        <legend>{t('ribbonEqOptDisplay')}</legend>
        {EQ_LIMITS.map(([field, side, centered]) => (
          <div key={field} className="eq-opt-limits">
            {(
              [
                ['subSup', side],
                ['undOvr', centered],
              ] as const
            ).map(([value, label]) => (
              <label key={value} className="tco-check">
                <input
                  type="radio"
                  name={field}
                  value={value}
                  checked={next[field] === value}
                  onChange={() => set({ [field]: value })}
                />
                {t(label)}
              </label>
            ))}
          </div>
        ))}
        <label className="tco-check">
          <input
            type="checkbox"
            name="dispDef"
            checked={next.dispDef}
            onChange={(e) => set({ dispDef: e.target.checked })}
          />
          {t('ribbonEqOptDispDef')}
        </label>
        <div className="tco-grid">
          <span>{t('ribbonEqOptLeftMargin')}</span>
          <LengthInput
            value={next.lMarginTwips}
            max={31680}
            ariaLabel={t('ribbonEqOptLeftMargin')}
            disabled={!next.dispDef}
            live
            onCommit={(twips) => set({ lMarginTwips: twips ?? 0 })}
          />
          <span>{t('ribbonEqOptRightMargin')}</span>
          <LengthInput
            value={next.rMarginTwips}
            max={31680}
            ariaLabel={t('ribbonEqOptRightMargin')}
            disabled={!next.dispDef}
            live
            onCommit={(twips) => set({ rMarginTwips: twips ?? 0 })}
          />
          <span>{t('ribbonEqOptJc')}</span>
          <select
            name="defJc"
            aria-label={t('ribbonEqOptJc')}
            disabled={!next.dispDef}
            value={next.defJc}
            onChange={(e) => set({ defJc: e.target.value as MathPr['defJc'] })}
          >
            {EQ_JC.map(([value, label]) => (
              <option key={value} value={value}>
                {t(label)}
              </option>
            ))}
          </select>
        </div>
        <label className="tco-check">
          <input
            type="radio"
            name="wrap"
            value="indent"
            checked={next.wrapRight !== true}
            onChange={() => set({ wrapRight: false })}
          />
          {t('ribbonEqOptWrapIndent')}
          <LengthInput
            value={next.wrapIndentTwips ?? DEFAULT_MATH_BREAKS.wrapIndentTwips}
            max={31680}
            ariaLabel={t('ribbonEqOptWrapIndent')}
            disabled={next.wrapRight === true}
            live
            onCommit={(twips) => set({ wrapIndentTwips: twips ?? 0 })}
          />
        </label>
        <label className="tco-check">
          <input
            type="radio"
            name="wrap"
            value="right"
            checked={next.wrapRight === true}
            onChange={() => set({ wrapRight: true })}
          />
          {t('ribbonEqOptWrapRight')}
        </label>
      </fieldset>
    </FormDialog>
  )
}
