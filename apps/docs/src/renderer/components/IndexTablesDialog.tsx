import { useState, type ReactNode } from 'react'
import type { FieldTableLeader, StyleUpsert } from '@genoffice/docx-engine'
import { CITATION_CATEGORIES } from '../editor/citations'
import {
  authoritiesOptionsOf,
  authoritiesSpec,
  DEFAULT_SPECS,
  figuresOptionsOf,
  figuresSpec,
  indexOptionsOf,
  indexSpec,
  TABLE_LEADERS,
  tocOptionsOf,
  tocSpec,
  type IndexTablesKind,
  type TableSpec,
  withIndexHeadings,
} from '../editor/index-tables'
import {
  formatIndexHeading,
  formatStyleUpserts,
  TABLE_FORMATS,
  type TableFormat,
} from '../editor/table-formats'
import { useI18n, type StringKey } from '../i18n/locale'
import { defaultParagraphStyleId, type StyleMap } from '../style-gallery'
import { EntryStylesDialog, FiguresOptionsDialog, TocOptionsDialog } from './IndexTablesExtras'
import { CITATION_CATEGORY_LABELS } from './MarkCitationDialog'
import { useModalKeys } from './modal-keys'

export type { IndexTablesKind } from '../editor/index-tables'

const TABS: Array<{ kind: IndexTablesKind; labelKey: StringKey; hintKey: StringKey }> = [
  { kind: 'index', labelKey: 'ribbonIndex', hintKey: 'ribbonItIndexHint' },
  { kind: 'toc', labelKey: 'ribbonToc', hintKey: 'ribbonItTocHint' },
  { kind: 'figures', labelKey: 'ribbonItFigures', hintKey: 'ribbonItFiguresHint' },
  { kind: 'authorities', labelKey: 'ribbonItAuthorities', hintKey: 'ribbonItAuthoritiesHint' },
]

const FORMAT_LABELS: Record<TableFormat, StringKey> = {
  template: 'ribbonItFromTemplate',
  classic: 'ribbonItFormatClassic',
  distinctive: 'ribbonItFormatDistinctive',
  fancy: 'ribbonItFormatFancy',
  modern: 'ribbonItFormatModern',
  bulleted: 'ribbonBnBulleted',
  centered: 'ribbonPgCentered',
  formal: 'ribbonItFormatFormal',
  simple: 'ribbonItFormatSimple',
}

const NO_STYLES: StyleMap = new Map()

// Word's own limit for an index
const MAX_INDEX_COLUMNS = 4

const LEADER_SAMPLE: Record<Exclude<FieldTableLeader, 'none'>, string> = {
  dot: '.......',
  hyphen: '-------',
  underscore: '_______',
}

function Check({
  name,
  checked,
  disabled,
  onChange,
  children,
}: {
  name: string
  checked: boolean
  disabled?: boolean
  onChange?: (on: boolean) => void
  children: ReactNode
}) {
  return (
    <label className="tco-check">
      <input
        type="checkbox"
        data-it={name}
        checked={checked}
        disabled={disabled}
        readOnly={!onChange}
        onChange={(e) => onChange?.(e.target.checked)}
      />
      {children}
    </label>
  )
}

/**
 * Word's Insert > Index and Tables: one place to insert an index, a table of
 * contents, a table of figures or a table of authorities, and to change the
 * one the selection is in. Each option is a switch of the table's field
 * instruction or the tab leader of its entries.
 */
export function IndexTablesDialog({
  captionLabels,
  existing,
  styles = NO_STYLES,
  pendingStyles,
  onApply,
  onUpsertStyle,
  onAutoMark,
  onMarkCitation,
  onMarkEntry,
  onClose,
}: {
  /** the caption labels the document uses; a table of figures lists the captions of one */
  captionLabels: string[]
  /** the table of each kind the selection is in: its options show, and OK rebuilds it */
  existing: Partial<Record<IndexTablesKind, TableSpec>>
  /** the document's styles, pending definitions included */
  styles?: StyleMap
  pendingStyles: Readonly<Record<string, StyleUpsert>>
  /** `formatStyles`: the entry style definitions a format other than From template brings */
  onApply: (kind: IndexTablesKind, spec: TableSpec, formatStyles: StyleUpsert[]) => void
  onUpsertStyle: (up: StyleUpsert) => void
  /** asks for a concordance file and marks its entries in the document */
  onAutoMark: () => void
  onMarkCitation: () => void
  onMarkEntry: () => void
  onClose: () => void
}) {
  const { t } = useI18n()
  const keys = useModalKeys(onClose)
  const [tab, setTab] = useState(() =>
    Math.max(
      0,
      TABS.findIndex((entry) => existing[entry.kind]),
    ),
  )
  const base = (kind: IndexTablesKind): TableSpec => existing[kind] ?? DEFAULT_SPECS[kind]
  const [toc, setToc] = useState(() => tocOptionsOf(base('toc')))
  const [figures, setFigures] = useState(() => {
    const options = figuresOptionsOf(base('figures'))
    return existing.figures ? options : { ...options, label: captionLabels[0] ?? '' }
  })
  const [index, setIndex] = useState(() => indexOptionsOf(base('index')))
  const [authorities, setAuthorities] = useState(() => authoritiesOptionsOf(base('authorities')))
  // a format is not kept in the file: the dialog opens on From template, as Word's does
  const [formats, setFormats] = useState<Partial<Record<IndexTablesKind, TableFormat>>>({})
  const [sub, setSub] = useState<'options' | 'styles' | null>(null)
  const { kind } = TABS[tab]
  const format = formats[kind] ?? 'template'

  const specOf = (): TableSpec => {
    if (kind === 'toc') return tocSpec(toc, base('toc'))
    if (kind === 'figures') return figuresSpec(figures, base('figures'))
    if (kind === 'authorities') return authoritiesSpec(authorities, base('authorities'))
    const spec = indexSpec(index, base('index'))
    const heading = formatIndexHeading(format)
    return heading === undefined ? spec : { ...spec, instr: withIndexHeadings(spec.instr, heading) }
  }

  const leaderRow = (
    value: FieldTableLeader,
    disabled: boolean,
    onChange: (leader: FieldTableLeader) => void,
  ) => (
    <div className="modal-row margin-row">
      <label>
        {t('ribbonItTabLeader')}
        <select
          data-it="leader"
          value={value}
          disabled={disabled}
          onChange={(e) => onChange(e.target.value as FieldTableLeader)}
        >
          {TABLE_LEADERS.map((leader) => (
            <option key={leader} value={leader}>
              {leader === 'none' ? t('ribbonPbArtNone') : LEADER_SAMPLE[leader]}
            </option>
          ))}
        </select>
      </label>
    </div>
  )
  const formatsRow = (
    <div className="modal-row margin-row">
      <label>
        {t('ribbonItFormats')}
        <select
          data-it="formats"
          value={format}
          onChange={(e) => setFormats({ ...formats, [kind]: e.target.value as TableFormat })}
        >
          {TABLE_FORMATS[kind].map((name) => (
            <option key={name} value={name}>
              {t(FORMAT_LABELS[name])}
            </option>
          ))}
        </select>
      </label>
    </div>
  )
  // a format brings its own definitions of the entry styles: Modify edits the template's
  const modifyButton = (
    <button
      type="button"
      data-it="modify"
      disabled={format !== 'template'}
      onClick={() => setSub('styles')}
    >
      {t('ribbonStyleModify')}
    </button>
  )
  const optionsButton = (
    <button type="button" data-it="options" onClick={() => setSub('options')}>
      {t('ribbonPbOptions')}
    </button>
  )
  const figureLabels =
    figures.label && !captionLabels.includes(figures.label)
      ? [figures.label, ...captionLabels]
      : captionLabels

  if (sub === 'styles')
    return (
      <EntryStylesDialog
        kind={kind}
        styles={styles}
        pending={pendingStyles}
        onUpsert={onUpsertStyle}
        onClose={() => setSub(null)}
      />
    )
  if (sub === 'options' && kind === 'toc')
    return (
      <TocOptionsDialog
        values={toc}
        styles={styles}
        onApply={(next) => {
          setToc({ ...toc, ...next })
          setSub(null)
        }}
        onClose={() => setSub(null)}
      />
    )
  if (sub === 'options' && kind === 'figures')
    return (
      <FiguresOptionsDialog
        values={figures}
        styles={styles}
        onApply={(next) => {
          setFigures({ ...figures, ...next })
          setSub(null)
        }}
        onClose={() => setSub(null)}
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
        className="modal gs-form tco-dialog index-tables-dialog"
        noValidate
        onSubmit={(e) => {
          e.preventDefault()
          onClose()
          onApply(
            kind,
            specOf(),
            formatStyleUpserts(kind, format, styles, defaultParagraphStyleId(styles)),
          )
        }}
      >
        <h2>{t('ribbonItTitle')}</h2>
        <div className="dlg-tabs" role="tablist">
          {TABS.map((entry, i) => (
            <button
              key={entry.kind}
              type="button"
              role="tab"
              aria-selected={tab === i}
              className={tab === i ? 'active' : ''}
              onClick={() => setTab(i)}
            >
              {t(entry.labelKey)}
            </button>
          ))}
        </div>
        <p className="index-tables-hint">{t(TABS[tab].hintKey)}</p>
        {kind === 'index' && (
          <>
            <div className="index-tables-radios" role="radiogroup" aria-label={t('ribbonSortType')}>
              <span>{t('ribbonSortType')}</span>
              <label className="tco-check">
                <input
                  type="radio"
                  name="index-type"
                  data-it="indented"
                  checked={!index.runIn}
                  onChange={() => setIndex({ ...index, runIn: false })}
                />
                {t('ribbonItIndented')}
              </label>
              <label className="tco-check">
                <input
                  type="radio"
                  name="index-type"
                  data-it="run-in"
                  checked={index.runIn}
                  onChange={() => setIndex({ ...index, runIn: true })}
                />
                {t('ribbonItRunIn')}
              </label>
            </div>
            {formatsRow}
            <div className="modal-row margin-row">
              <label>
                {t('ribbonColumns')}
                <input
                  data-it="columns"
                  type="number"
                  min={1}
                  max={MAX_INDEX_COLUMNS}
                  value={index.columns}
                  onChange={(e) =>
                    setIndex({
                      ...index,
                      columns: Math.min(
                        MAX_INDEX_COLUMNS,
                        Math.max(1, Math.round(Number(e.target.value)) || 1),
                      ),
                    })
                  }
                />
              </label>
            </div>
            <Check
              name="accented"
              checked={index.accented}
              onChange={(accented) => setIndex({ ...index, accented })}
            >
              {t('ribbonItAccented')}
            </Check>
            <Check
              name="right-align"
              checked={index.rightAlign}
              onChange={(rightAlign) => setIndex({ ...index, rightAlign })}
            >
              {t('ribbonItRightAlign')}
            </Check>
            {leaderRow(index.leader, !index.rightAlign, (leader) => setIndex({ ...index, leader }))}
          </>
        )}
        {kind === 'toc' && (
          <>
            {formatsRow}
            <div className="modal-row margin-row">
              <label>
                {t('ribbonItShowLevels')}
                <input
                  data-it="levels"
                  type="number"
                  min={1}
                  max={9}
                  value={toc.levels || ''}
                  onChange={(e) =>
                    setToc({
                      ...toc,
                      levels: Math.min(9, Math.max(1, Math.round(Number(e.target.value)) || 1)),
                    })
                  }
                />
              </label>
            </div>
            <Check
              name="page-numbers"
              checked={toc.pageNumbers}
              onChange={(pageNumbers) => setToc({ ...toc, pageNumbers })}
            >
              {t('ribbonItShowPageNumbers')}
            </Check>
            <Check
              name="right-align"
              checked={toc.rightAlign}
              disabled={!toc.pageNumbers}
              onChange={(rightAlign) => setToc({ ...toc, rightAlign })}
            >
              {t('ribbonItRightAlign')}
            </Check>
            {leaderRow(toc.leader, !toc.pageNumbers || !toc.rightAlign, (leader) =>
              setToc({ ...toc, leader }),
            )}
            <Check
              name="hyperlinks"
              checked={toc.hyperlinks}
              onChange={(hyperlinks) => setToc({ ...toc, hyperlinks })}
            >
              {t('ribbonItHyperlinks')}
            </Check>
          </>
        )}
        {kind === 'figures' && (
          <>
            <div className="modal-row margin-row">
              <label>
                {t('ribbonCaptionLabel')}
                <select
                  data-it="label"
                  value={figures.label}
                  disabled={figureLabels.length === 0}
                  onChange={(e) => setFigures({ ...figures, label: e.target.value })}
                >
                  {figureLabels.map((name) => (
                    <option key={name} value={name}>
                      {name}
                    </option>
                  ))}
                </select>
              </label>
            </div>
            {formatsRow}
            <Check
              name="include-label"
              checked={figures.includeLabel}
              onChange={(includeLabel) => setFigures({ ...figures, includeLabel })}
            >
              {t('ribbonItIncludeLabel')}
            </Check>
            <Check
              name="page-numbers"
              checked={figures.pageNumbers}
              onChange={(pageNumbers) => setFigures({ ...figures, pageNumbers })}
            >
              {t('ribbonItShowPageNumbers')}
            </Check>
            <Check
              name="right-align"
              checked={figures.rightAlign}
              disabled={!figures.pageNumbers}
              onChange={(rightAlign) => setFigures({ ...figures, rightAlign })}
            >
              {t('ribbonItRightAlign')}
            </Check>
            {leaderRow(figures.leader, !figures.pageNumbers || !figures.rightAlign, (leader) =>
              setFigures({ ...figures, leader }),
            )}
            <Check
              name="hyperlinks"
              checked={figures.hyperlinks}
              onChange={(hyperlinks) => setFigures({ ...figures, hyperlinks })}
            >
              {t('ribbonItHyperlinks')}
            </Check>
          </>
        )}
        {kind === 'authorities' && (
          <>
            {formatsRow}
            <div className="modal-row margin-row">
              <label>
                {t('ribbonToaCategory')}
                <select
                  data-it="category"
                  value={authorities.category}
                  onChange={(e) =>
                    setAuthorities({ ...authorities, category: Number(e.target.value) })
                  }
                >
                  <option value={0}>{t('appFontLigAll')}</option>
                  {CITATION_CATEGORIES.map((n) => (
                    <option key={n} value={n}>
                      {t(CITATION_CATEGORY_LABELS[n])}
                    </option>
                  ))}
                </select>
              </label>
            </div>
            <Check
              name="passim"
              checked={authorities.passim}
              onChange={(passim) => setAuthorities({ ...authorities, passim })}
            >
              {t('ribbonItUsePassim')}
            </Check>
            <Check
              name="keep-formatting"
              checked={authorities.keepFormatting}
              onChange={(keepFormatting) => setAuthorities({ ...authorities, keepFormatting })}
            >
              {t('ribbonItKeepFormatting')}
            </Check>
            {leaderRow(authorities.leader, false, (leader) =>
              setAuthorities({ ...authorities, leader }),
            )}
          </>
        )}
        <div className="modal-actions">
          <div className="tco-actions-right">
            {modifyButton}
            {(kind === 'toc' || kind === 'figures') && optionsButton}
            {kind === 'index' && (
              <>
                <button
                  type="button"
                  data-it="automark"
                  onClick={() => {
                    onClose()
                    onAutoMark()
                  }}
                >
                  {t('ribbonItAutoMark')}
                </button>
                <button
                  type="button"
                  data-it="mark-entry"
                  onClick={() => {
                    onClose()
                    onMarkEntry()
                  }}
                >
                  {t('ribbonItMarkEntry')}
                </button>
              </>
            )}
            {kind === 'authorities' && (
              <button
                type="button"
                data-it="mark-citation"
                onClick={() => {
                  onClose()
                  onMarkCitation()
                }}
              >
                {t('ribbonItMarkCitation')}
              </button>
            )}
          </div>
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
