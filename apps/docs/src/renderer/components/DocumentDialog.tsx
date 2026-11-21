import { useEffect, useRef, useState } from 'react'
import { DEFAULT_LINE_PITCH_TWIPS } from '@genoffice/docx-engine'
import type { DocGrid, SectionSettings } from '@genoffice/docx-engine'
import type { SectionScope } from '../commands/menu-commands'
import type { PageDefault } from '../design-default'
import { useI18n, type StringKey } from '../i18n/locale'
import { ApplyToSelect, defaultScope } from './ApplyToSelect'
import { LengthInput } from './LengthInput'
import { marginsFitPage, type PageMargins } from './MarginDialog'
import { useModalKeys } from './modal-keys'
import { sectionVertical } from '../pagination-sections'
import {
  CHAR_GRIDS,
  LINE_GRIDS,
  charPitchPt,
  charSpaceForPitch,
  charsPerLine,
  maxCharsPerLine,
} from '../char-grid'
import { GridFontDialog, type GridFont } from './GridFontDialog'
import { Unavailable } from './Unavailable'

export type SectionStart = 'nextPage' | 'continuous' | 'evenPage' | 'oddPage' | 'nextColumn'

export interface DocumentDialogValue {
  margins: PageMargins
  mirror: boolean
  titlePg: boolean
  evenOdd: boolean
  headerDist: number
  footerDist: number
  gutter: number
  /** the document's gutters sit at the top of the page (settings.xml, every section) */
  gutterAtTop: boolean
  startType: SectionStart
  vAlign: NonNullable<SectionSettings['vAlign']>
  columns: number
  gridType: DocGrid['type']
  /** w:docGrid w:linePitch, twips: what a line grid snaps lines to */
  linePitch: number
  /** w:docGrid w:charSpace: the character pitch's distance from the default font size, 1/4096 pt */
  charSpace: number
  /** Set Font…: what changes in the default paragraph style's font */
  gridFont?: GridFont
  /** w:textDirection of the section; lrTb when it has none */
  textDirection: string
  /** w:noEndnote of the section */
  suppressEndnotes: boolean
}

/** what OK has to write: only the parts that differ from what the dialog opened on */
export function documentDialogChanges(
  before: DocumentDialogValue,
  after: DocumentDialogValue,
): Partial<DocumentDialogValue> {
  const out: Partial<DocumentDialogValue> = {}
  const sides = ['top', 'right', 'bottom', 'left'] as const
  // the gutter is part of the margin the layout uses, so the two are written together
  if (
    sides.some((s) => before.margins[s] !== after.margins[s]) ||
    before.mirror !== after.mirror ||
    before.gutter !== after.gutter ||
    before.gutterAtTop !== after.gutterAtTop
  ) {
    out.margins = after.margins
    out.mirror = after.mirror
    out.gutter = after.gutter
    out.gutterAtTop = after.gutterAtTop
  }
  if (before.startType !== after.startType) out.startType = after.startType
  if (before.vAlign !== after.vAlign) out.vAlign = after.vAlign
  if (before.columns !== after.columns) out.columns = after.columns
  if (before.gridType !== after.gridType) out.gridType = after.gridType
  if (before.linePitch !== after.linePitch && LINE_GRIDS.includes(after.gridType))
    out.linePitch = after.linePitch
  if (before.charSpace !== after.charSpace && CHAR_GRIDS.includes(after.gridType))
    out.charSpace = after.charSpace
  if (after.gridFont && Object.keys(after.gridFont).length > 0) out.gridFont = after.gridFont
  if (before.textDirection !== after.textDirection) out.textDirection = after.textDirection
  if (before.suppressEndnotes !== after.suppressEndnotes)
    out.suppressEndnotes = after.suppressEndnotes
  if (before.titlePg !== after.titlePg) out.titlePg = after.titlePg
  if (before.evenOdd !== after.evenOdd) out.evenOdd = after.evenOdd
  if (before.headerDist !== after.headerDist) out.headerDist = after.headerDist
  if (before.footerDist !== after.footerDist) out.footerDist = after.footerDist
  return out
}

/**
 * The parsed left margin (top, for a gutter at the top) already holds the
 * gutter. Word lists the two apart, so the dialog shows the margin without it.
 */
export function marginsWithoutGutter(section: SectionSettings): PageMargins {
  const gutter = section.gutter ?? 0
  return {
    top: Math.max(0, section.marginTop - (section.gutterAtTop ? gutter : 0)),
    right: section.marginRight,
    bottom: section.marginBottom,
    left: Math.max(0, section.marginLeft - (section.gutterAtTop ? 0 : gutter)),
  }
}

/** back to the margin the layout uses: the dialog's value plus the gutter */
export function marginsWithGutter(
  margins: PageMargins,
  section: SectionSettings,
  gutter = section.gutter ?? 0,
  atTop = !!section.gutterAtTop,
): PageMargins {
  return {
    ...margins,
    top: margins.top + (atTop ? gutter : 0),
    left: margins.left + (atTop ? 0 : gutter),
  }
}

/** a section after the document's gutter position changed: its gutter moves between the left and the top margin */
export function gutterMoved(section: SectionSettings, atTop: boolean): SectionSettings {
  const gutter = section.gutter ?? 0
  if (gutter <= 0 || !!section.gutterAtTop === atTop) return section
  const { gutterAtTop: _was, ...rest } = section
  return {
    ...rest,
    marginTop: section.marginTop + (atTop ? gutter : -gutter),
    marginLeft: section.marginLeft + (atTop ? -gutter : gutter),
    ...(atTop ? { gutterAtTop: true } : {}),
  }
}

/** Word's Header / Footer from edge range: 0–22" */
const MAX_HF_DISTANCE_TWIPS = 31680
const DEFAULT_HF_DISTANCE_TWIPS = 720

const START_KEYS: Record<SectionStart, StringKey> = {
  continuous: 'ribbonBreakContinuous',
  nextColumn: 'ribbonDocNewColumn',
  nextPage: 'ribbonDocNewPage',
  evenPage: 'ribbonBreakEvenPage',
  oddPage: 'ribbonBreakOddPage',
}
const START_TYPES = Object.keys(START_KEYS) as SectionStart[]

const V_ALIGN_KEYS: Record<NonNullable<SectionSettings['vAlign']>, StringKey> = {
  top: 'ribbonMarginTop',
  center: 'appTabCenter',
  both: 'appAlignJustify',
  bottom: 'ribbonMarginBottom',
}

const V_ALIGNS = ['top', 'center', 'both', 'bottom'] as const

/** Word's limit on the columns of a section */
const MAX_COLUMNS = 45

/** one point: the smallest line pitch the dialog writes */
export const MIN_LINE_PITCH_TWIPS = 20

export const GRID_TYPES: ReadonlyArray<DocGrid['type']> = [
  'default',
  'lines',
  'linesAndChars',
  'snapToChars',
]

const GRID_KEYS: Array<[DocGrid['type'], StringKey]> = [
  ['default', 'ribbonDocNoGrid'],
  ['lines', 'ribbonDocLineGrid'],
  ['linesAndChars', 'ribbonDocLineCharGrid'],
  ['snapToChars', 'ribbonDocSnapGrid'],
]

type Tab = 'margins' | 'layout' | 'grid'
const TABS: Array<[Tab, StringKey]> = [
  ['margins', 'ribbonMargins'],
  ['layout', 'ribbonTabLayout'],
  ['grid', 'ribbonDocTabGrid'],
]

interface Props {
  section: SectionSettings
  /** which section of the document `section` is */
  sectionIndex: number
  startType: SectionStart
  /** the caret's section ends in a break that is not saved yet: its start cannot be rewritten from here */
  startLocked: boolean
  /** settings.xml w:gutterAtTop as the document has it now */
  gutterAtTop: boolean
  /** the document has several sections and the dialog acts on the one with the caret */
  multiSection: boolean
  mirror: boolean
  titlePg: boolean
  /** the caret is in a section whose first-page option cannot be changed from here yet */
  titlePgLocked: boolean
  evenOdd: boolean
  /** the default paragraph style's font: what the character grid counts its cells from */
  gridFont: Required<GridFont>
  /** a section break can go in at the caret and be acted on at once */
  canApplyForward: boolean
  /** endnotes sit at section ends and a later section can take this one's */
  canSuppressEndnotes: boolean
  /** `scope` comes along when Apply to reaches beyond what the dialog acts on by default */
  onApply: (changes: Partial<DocumentDialogValue> & { scope?: SectionScope }) => void
  /** Default…: the dialog's page setup becomes what new blank documents start with */
  onDefault: (page: PageDefault) => void
  onPageSetup: () => void
  /** Borders…: the Borders and Shading dialog */
  onBorders: () => void
  /** Line Numbers…: the Line Numbers dialog */
  onLineNumbers: () => void
  onClose: () => void
}

/** Word's Format > Document: page margins, section layout and the document grid */
export function DocumentDialog({
  section,
  sectionIndex,
  startType,
  startLocked,
  gutterAtTop,
  multiSection,
  mirror,
  titlePg,
  titlePgLocked,
  evenOdd,
  gridFont,
  canApplyForward,
  canSuppressEndnotes,
  onApply,
  onDefault,
  onPageSetup,
  onBorders,
  onLineNumbers,
  onClose,
}: Props) {
  const { t } = useI18n()
  const keys = useModalKeys(onClose)
  const [tab, setTab] = useState<Tab>('margins')
  const [opened] = useState<DocumentDialogValue>(() => ({
    margins: marginsWithoutGutter(section),
    mirror,
    titlePg,
    evenOdd,
    headerDist: section.headerDist ?? DEFAULT_HF_DISTANCE_TWIPS,
    footerDist: section.footerDist ?? DEFAULT_HF_DISTANCE_TWIPS,
    gutter: section.gutter ?? 0,
    gutterAtTop,
    startType,
    vAlign: section.vAlign ?? 'top',
    columns: section.columns,
    gridType: section.docGrid?.type ?? 'default',
    linePitch: section.docGrid?.linePitch || DEFAULT_LINE_PITCH_TWIPS,
    charSpace: section.docGrid?.charSpace ?? 0,
    textDirection: section.textDirection ?? 'lrTb',
    suppressEndnotes: !!section.noEndnote,
  }))
  const [value, setValue] = useState(opened)
  const [scope, setScope] = useState<SectionScope>(defaultScope(multiSection))
  // the pitch as it is being typed, so a half-typed number is not rewritten under the caret
  const [pitchText, setPitchText] = useState<string | null>(null)
  const [charPitchText, setCharPitchText] = useState<string | null>(null)
  const [defaultCharPitch, setDefaultCharPitch] = useState(opened.charSpace === 0)
  const [fontOpen, setFontOpen] = useState(false)
  const columnsValid =
    Number.isInteger(value.columns) && value.columns >= 1 && value.columns <= MAX_COLUMNS
  const tooLarge = !marginsFitPage(
    marginsWithGutter(value.margins, section, value.gutter, value.gutterAtTop),
    section.pageWidth,
    section.pageHeight,
  )
  const lineGrid = LINE_GRIDS.includes(value.gridType)
  const gridMargins = marginsWithGutter(value.margins, section, value.gutter, value.gutterAtTop)
  const textHeight = Math.max(0, section.pageHeight - gridMargins.top - gridMargins.bottom)
  const pitchOf = (twips: number) =>
    Math.min(Math.max(Math.round(twips), MIN_LINE_PITCH_TWIPS), Math.max(textHeight, 20))
  const charGrid = CHAR_GRIDS.includes(value.gridType)
  const basePt = (value.gridFont?.sizeHalfPoints ?? gridFont.sizeHalfPoints) / 2
  const textWidthPt = Math.max(0, section.pageWidth - gridMargins.left - gridMargins.right) / 20
  const maxChars = maxCharsPerLine(textWidthPt, basePt)
  const charPitch = charPitchPt(basePt, value.charSpace)
  const charSpaceOf = (pitchPt: number) =>
    charSpaceForPitch(basePt, Math.min(Math.max(pitchPt, textWidthPt / maxChars), textWidthPt))
  const submit = () => {
    if (tooLarge || !columnsValid) return
    const wider = scope !== defaultScope(multiSection)
    // the other sections take everything the dialog shows, not only what was changed here
    const {
      startType: _start,
      titlePg: _titlePg,
      evenOdd: _evenOdd,
      linePitch: shownPitch,
      charSpace: shownSpace,
      gridFont: _font,
      ...shown
    } = value
    const changes = {
      ...(wider ? shown : {}),
      ...(wider && lineGrid ? { linePitch: shownPitch } : {}),
      ...(wider && charGrid ? { charSpace: shownSpace } : {}),
      ...documentDialogChanges(opened, value),
    }
    if (changes.margins)
      changes.margins = marginsWithGutter(changes.margins, section, value.gutter, value.gutterAtTop)
    onApply(wider ? { ...changes, scope } : changes)
    onClose()
  }
  // Word's Default… applies the dialog as OK does and keeps it for new documents
  const keepAsDefault = () => {
    if (tooLarge || !columnsValid) return
    const margins = marginsWithGutter(value.margins, section, value.gutter, value.gutterAtTop)
    onDefault({
      section: {
        marginTop: margins.top,
        marginRight: margins.right,
        marginBottom: margins.bottom,
        marginLeft: margins.left,
        headerDist: value.headerDist,
        footerDist: value.footerDist,
        gutter: value.gutter,
        gutterAtTop: (value.gutter > 0 && value.gutterAtTop) || undefined,
        vAlign: value.vAlign === 'top' ? undefined : value.vAlign,
        columns: value.columns,
      },
      mirror: value.mirror,
      titlePg: value.titlePg,
      evenOdd: value.evenOdd,
    })
    submit()
  }
  // Enter inside a field commits its value first; submit runs on the flushed state
  const [submitTick, setSubmitTick] = useState(0)
  const submitRef = useRef(submit)
  submitRef.current = submit
  useEffect(() => {
    if (submitTick) submitRef.current()
  }, [submitTick])
  const onEnter = () => setSubmitTick((n) => n + 1)
  const vertical = !!sectionVertical({ ...section, textDirection: value.textDirection })
  // a section that came vertical keeps its own flow (rotated or bottom-up) when Vertical is picked again
  const verticalFlow = sectionVertical(section) ?? 'tbRl'

  const leaveFor = (next: () => void) => () => {
    onClose()
    next()
  }
  const margin = (side: keyof PageMargins, labelKey: StringKey) => {
    const label = t(
      value.mirror && side === 'left'
        ? 'ribbonMarginInside'
        : value.mirror && side === 'right'
          ? 'ribbonMarginOutside'
          : labelKey,
    )
    return (
      <label>
        {label}
        <LengthInput
          value={value.margins[side]}
          min={0}
          max={side === 'left' || side === 'right' ? section.pageWidth : section.pageHeight}
          ariaLabel={label}
          live
          onCommit={(twips) =>
            setValue((v) => ({ ...v, margins: { ...v.margins, [side]: twips ?? 0 } }))
          }
          onEnter={onEnter}
        />
      </label>
    )
  }
  const distance = (kind: 'headerDist' | 'footerDist', labelKey: StringKey) => (
    <label>
      {t(labelKey)}
      <LengthInput
        value={value[kind]}
        min={0}
        max={MAX_HF_DISTANCE_TWIPS}
        ariaLabel={t(labelKey)}
        live
        onCommit={(twips) => setValue((v) => ({ ...v, [kind]: twips ?? 0 }))}
        onEnter={onEnter}
      />
    </label>
  )
  const applyTo = (
    <ApplyToSelect
      scope={scope}
      multiSection={multiSection}
      canApplyForward={canApplyForward}
      onChange={setScope}
    />
  )

  return (
    <div
      className="modal-backdrop"
      ref={keys.ref}
      onKeyDown={keys.onKeyDown}
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
    >
      <form
        className="modal gs-form tco-dialog doc-dialog"
        data-section={sectionIndex}
        noValidate
        onSubmit={(e) => {
          e.preventDefault()
          submit()
        }}
      >
        <h2>{t('ribbonDocTitle')}</h2>
        <div className="dlg-tabs" role="tablist">
          {TABS.map(([k, labelKey]) => (
            <button
              key={k}
              type="button"
              role="tab"
              aria-selected={tab === k}
              className={tab === k ? 'active' : ''}
              onClick={() => setTab(k)}
            >
              {t(labelKey)}
            </button>
          ))}
        </div>

        {tab === 'margins' && (
          <>
            <div className="modal-row margin-row">
              {margin('top', 'ribbonMarginTop')}
              {margin('bottom', 'ribbonMarginBottom')}
            </div>
            <div className="modal-row margin-row">
              {margin('left', 'ribbonMarginLeft')}
              {margin('right', 'ribbonMarginRight')}
            </div>
            <div className="modal-row margin-row">
              <label>
                {t('ribbonDocGutter')}
                <LengthInput
                  value={value.gutter}
                  min={0}
                  max={MAX_HF_DISTANCE_TWIPS}
                  ariaLabel={t('ribbonDocGutter')}
                  live
                  onCommit={(twips) => setValue((v) => ({ ...v, gutter: twips ?? 0 }))}
                  onEnter={onEnter}
                />
              </label>
              <label>
                {t('ribbonDocGutterPos')}
                <select
                  name="gutterPosition"
                  value={value.gutterAtTop ? 'top' : 'left'}
                  onChange={(e) =>
                    setValue((v) => ({ ...v, gutterAtTop: e.target.value === 'top' }))
                  }
                >
                  <option value="left">{t('ribbonMarginLeft')}</option>
                  <option value="top">{t('ribbonMarginTop')}</option>
                </select>
              </label>
            </div>
            <div className="modal-row margin-row">
              <label>
                {t('ribbonMarginPages')}
                <select
                  value={value.mirror ? 'mirror' : 'normal'}
                  onChange={(e) => setValue((v) => ({ ...v, mirror: e.target.value === 'mirror' }))}
                >
                  <option value="normal">{t('ribbonMarginPagesNormal')}</option>
                  <option value="mirror">{t('ribbonMarginPagesMirror')}</option>
                </select>
              </label>
            </div>
            {applyTo}
            {tooLarge && <div className="modal-error">{t('ribbonMarginTooLarge')}</div>}
          </>
        )}

        {tab === 'layout' && (
          <>
            <fieldset>
              <legend>{t('ribbonDocSection')}</legend>
              <div className="modal-row margin-row">
                <label>
                  {t('ribbonDocSectionStart')}
                  <select
                    name="sectionStart"
                    disabled={startLocked}
                    value={value.startType}
                    onChange={(e) =>
                      setValue((v) => ({ ...v, startType: e.target.value as SectionStart }))
                    }
                  >
                    {START_TYPES.map((type) => (
                      <option key={type} value={type}>
                        {t(START_KEYS[type])}
                      </option>
                    ))}
                  </select>
                </label>
              </div>
              <label className="tco-check">
                <input
                  type="checkbox"
                  name="suppressEndnotes"
                  checked={value.suppressEndnotes}
                  // only endnotes placed at section ends can be handed on, and only to a later section
                  disabled={!opened.suppressEndnotes && !canSuppressEndnotes}
                  onChange={(e) => setValue((v) => ({ ...v, suppressEndnotes: e.target.checked }))}
                />
                {t('ribbonDocSuppressEndnotes')}
              </label>
            </fieldset>
            <fieldset>
              <legend>{t('ribbonDocHeadersFooters')}</legend>
              <label className="tco-check">
                <input
                  type="checkbox"
                  checked={value.evenOdd}
                  onChange={(e) => setValue((v) => ({ ...v, evenOdd: e.target.checked }))}
                />
                {t('ribbonDiffOddEven')}
              </label>
              <label className="tco-check">
                <input
                  type="checkbox"
                  checked={value.titlePg}
                  disabled={titlePgLocked}
                  onChange={(e) => setValue((v) => ({ ...v, titlePg: e.target.checked }))}
                />
                {t('ribbonDiffFirstPage')}
              </label>
              <div className="doc-dialog-caption">{t('ribbonDocFromEdge')}</div>
              <div className="modal-row margin-row">
                {distance('headerDist', 'ribbonHeader')}
                {distance('footerDist', 'ribbonFooter')}
              </div>
            </fieldset>
            <fieldset>
              <legend>{t('ribbonDocPage')}</legend>
              <div className="modal-row margin-row">
                <label>
                  {t('ribbonCellVAlign')}
                  <select
                    name="vAlign"
                    value={value.vAlign}
                    onChange={(e) =>
                      setValue((v) => ({ ...v, vAlign: e.target.value as typeof v.vAlign }))
                    }
                  >
                    {V_ALIGNS.map((align) => (
                      <option key={align} value={align}>
                        {t(V_ALIGN_KEYS[align])}
                      </option>
                    ))}
                  </select>
                </label>
              </div>
            </fieldset>
            {applyTo}
            <div className="modal-actions doc-dialog-links">
              <button type="button" onClick={leaveFor(onLineNumbers)}>
                {t('ribbonDocLineNumbers')}
              </button>
              <button type="button" onClick={leaveFor(onBorders)}>
                {t('ribbonDocBorders')}
              </button>
            </div>
          </>
        )}

        {tab === 'grid' && (
          <>
            <fieldset>
              <legend>{t('ribbonDocTextFlow')}</legend>
              <div className="doc-dialog-caption">{t('ribbonDocDirection')}</div>
              <label className="tco-check">
                <input
                  type="radio"
                  name="textDirection"
                  value="horizontal"
                  checked={!vertical}
                  onChange={() => setValue((v) => ({ ...v, textDirection: 'lrTb' }))}
                />
                {t('ribbonTextDirectionHorizontal')}
              </label>
              <label className="tco-check">
                <input
                  type="radio"
                  name="textDirection"
                  value="vertical"
                  checked={vertical}
                  onChange={() => setValue((v) => ({ ...v, textDirection: verticalFlow }))}
                />
                {t('ribbonDocVertical')}
              </label>
              <div className="modal-row margin-row">
                <label>
                  {t('ribbonColumns')}
                  <input
                    type="number"
                    name="columns"
                    min={1}
                    max={MAX_COLUMNS}
                    value={Number.isNaN(value.columns) ? '' : value.columns}
                    onChange={(e) => setValue((v) => ({ ...v, columns: e.target.valueAsNumber }))}
                  />
                </label>
              </div>
            </fieldset>
            <fieldset>
              <legend>{t('ribbonDocGrid')}</legend>
              {GRID_KEYS.map(([type, labelKey]) => (
                <label key={type} className="tco-check">
                  <input
                    type="radio"
                    name="gridType"
                    value={type}
                    checked={value.gridType === type}
                    onChange={() => setValue((v) => ({ ...v, gridType: type }))}
                  />
                  {t(labelKey)}
                </label>
              ))}
            </fieldset>
            <fieldset>
              <legend>{t('appParaUnitChars')}</legend>
              <div className="modal-row margin-row">
                <label>
                  {t('ribbonDocGridPerLine')}
                  <input
                    type="number"
                    name="charsPerLine"
                    min={1}
                    max={maxChars}
                    title={`1-${maxChars}`}
                    disabled={!charGrid || defaultCharPitch}
                    value={charsPerLine(textWidthPt, charPitch)}
                    onChange={(e) => {
                      const count = Math.floor(e.target.valueAsNumber)
                      setCharPitchText(null)
                      if (count >= 1)
                        setValue((v) => ({ ...v, charSpace: charSpaceOf(textWidthPt / count) }))
                    }}
                  />
                </label>
                <label>
                  {t('ribbonDocGridPitch')}
                  <span className="doc-dialog-pitch">
                    <input
                      type="number"
                      name="charPitch"
                      step={0.05}
                      disabled={!charGrid || defaultCharPitch}
                      value={charPitchText ?? Math.round(charPitch * 100) / 100}
                      onChange={(e) => {
                        const pt = e.target.valueAsNumber
                        setCharPitchText(e.target.value)
                        if (pt > 0) setValue((v) => ({ ...v, charSpace: charSpaceOf(pt) }))
                      }}
                      onBlur={() => setCharPitchText(null)}
                    />
                    <span className="para-unit">{t('appUnitPtAbbr')}</span>
                  </span>
                </label>
              </div>
              <label className="tco-check">
                <input
                  type="checkbox"
                  name="defaultCharPitch"
                  disabled={!charGrid}
                  checked={defaultCharPitch}
                  onChange={(e) => {
                    setDefaultCharPitch(e.target.checked)
                    setCharPitchText(null)
                    if (e.target.checked) setValue((v) => ({ ...v, charSpace: 0 }))
                  }}
                />
                {t('ribbonDocGridDefaultPitch')}
              </label>
            </fieldset>
            <fieldset>
              <legend>{t('appStatLines')}</legend>
              <div className="modal-row margin-row">
                <label>
                  {t('ribbonDocGridPerPage')}
                  <input
                    type="number"
                    name="linesPerPage"
                    min={1}
                    disabled={!lineGrid}
                    value={Math.max(1, Math.floor(textHeight / value.linePitch))}
                    onChange={(e) => {
                      const lines = Math.floor(e.target.valueAsNumber)
                      setPitchText(null)
                      if (lines >= 1)
                        setValue((v) => ({
                          ...v,
                          linePitch: pitchOf(Math.floor(textHeight / lines)),
                        }))
                    }}
                  />
                </label>
                <label>
                  {t('ribbonDocGridPitch')}
                  <span className="doc-dialog-pitch">
                    <input
                      type="number"
                      name="linePitch"
                      min={MIN_LINE_PITCH_TWIPS / 20}
                      step={0.05}
                      disabled={!lineGrid}
                      value={pitchText ?? value.linePitch / 20}
                      onChange={(e) => {
                        const pt = e.target.valueAsNumber
                        setPitchText(e.target.value)
                        if (pt > 0) setValue((v) => ({ ...v, linePitch: pitchOf(pt * 20) }))
                      }}
                      onBlur={() => setPitchText(null)}
                    />
                    <span className="para-unit">{t('appUnitPtAbbr')}</span>
                  </span>
                </label>
              </div>
            </fieldset>
            {applyTo}
            <div className="modal-actions doc-dialog-links">
              <button type="button" name="gridSetFont" onClick={() => setFontOpen(true)}>
                {t('ribbonDocGridSetFont')}
              </button>
              <Unavailable id="document.drawing-grid">
                <button type="button" disabled>
                  {t('ribbonDocGridDrawing')}
                </button>
              </Unavailable>
            </div>
          </>
        )}

        <div className="modal-actions">
          <div className="tco-actions-right">
            <button
              type="button"
              name="default"
              disabled={tooLarge || !columnsValid}
              onClick={keepAsDefault}
            >
              {t('ribbonDocDefault')}
            </button>
            <button type="button" onClick={leaveFor(onPageSetup)}>
              {t('ribbonDocPageSetup')}
            </button>
          </div>
          <div className="tco-actions-right">
            <button type="button" onClick={onClose}>
              {t('ribbonCancel')}
            </button>
            <button type="submit" className="primary" disabled={tooLarge || !columnsValid}>
              {t('ribbonOk')}
            </button>
          </div>
        </div>
      </form>
      {fontOpen && (
        <GridFontDialog
          font={{ ...gridFont, ...value.gridFont }}
          onApply={(changed) =>
            setValue((v) => ({ ...v, gridFont: { ...v.gridFont, ...changed } }))
          }
          onClose={() => setFontOpen(false)}
        />
      )}
    </div>
  )
}
