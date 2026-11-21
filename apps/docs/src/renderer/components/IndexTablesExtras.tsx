import { useMemo, useState } from 'react'
import type { StyleInfo, StyleUpsert } from '@genoffice/docx-engine'
import type { IndexTablesKind, TocSources, TocStyleLevel } from '../editor/index-tables'
import { ENTRY_STYLES, entryStyleId, entryStyleUpsert } from '../editor/table-formats'
import { useI18n } from '../i18n/locale'
import { defaultParagraphStyleId, styleLabel, type StyleMap } from '../style-gallery'
import { useModalKeys } from './modal-keys'
import { StyleDialog, formValuesOf } from './StyleDialog'

const HEADING_NAME = /^heading ([1-9])$/i

/** the paragraph style names Options can list, the nine built-in headings first */
function paragraphStyleNames(styles: StyleMap): string[] {
  const names = [...styles.values()]
    .filter((s) => s.type === 'paragraph' && !s.linkedCharShell && !HEADING_NAME.test(s.name))
    .map((s) => s.name)
    .sort((a, b) => a.localeCompare(b))
  return [...Array.from({ length: 9 }, (_, i) => `heading ${i + 1}`), ...names]
}

export interface TocSourceValues extends TocSources {
  /** deepest built-in heading level listed; 0 when the headings are not a source */
  levels: number
}

/** the style table as Options opens it: the headings the field lists at their own level, then its named styles */
function levelRows(values: TocSourceValues): Record<string, number> {
  const rows: Record<string, number> = {}
  for (let n = 1; n <= values.levels; n++) rows[`heading ${n}`] = n
  for (const s of values.styles) rows[s.name.toLowerCase()] = s.level
  return rows
}

/** the table back as the field's switches: headings 1..N at their own level are \o, every other row \t */
function sourcesOfRows(
  rows: Record<string, number>,
  names: string[],
): Pick<TocSourceValues, 'levels' | 'styles'> {
  let levels = 0
  while (levels < 9 && rows[`heading ${levels + 1}`] === levels + 1) levels++
  const styles: TocStyleLevel[] = []
  for (const name of names) {
    const level = rows[name.toLowerCase()]
    const heading = Number(HEADING_NAME.exec(name)?.[1] ?? 0)
    if (level && !(heading && heading <= levels))
      styles.push({ name: heading ? `Heading ${heading}` : name, level })
  }
  return { levels, styles }
}

/**
 * Word's Table of Contents Options: what the table is built from. Styles with
 * a level (TOC \o for the headings, \t for the rest), paragraphs with an
 * outline level (\u), TC fields (\f).
 */
export function TocOptionsDialog({
  values,
  styles,
  onApply,
  onClose,
}: {
  values: TocSourceValues
  styles: StyleMap
  onApply: (next: TocSourceValues) => void
  onClose: () => void
}) {
  const { t } = useI18n()
  const keys = useModalKeys(onClose)
  const names = useMemo(() => {
    const listed = paragraphStyleNames(styles)
    const known = new Set(listed.map((n) => n.toLowerCase()))
    // a style the field names stays listed when the document no longer has it
    return [
      ...listed,
      ...values.styles.map((s) => s.name).filter((n) => !known.has(n.toLowerCase())),
    ]
  }, [styles, values.styles])
  const [fromStyles, setFromStyles] = useState(values.levels > 0 || values.styles.length > 0)
  const [rows, setRows] = useState(() => levelRows(values))
  const [outline, setOutline] = useState(values.outline)
  const [entryFields, setEntryFields] = useState(values.entryFields)
  const reset = () => {
    setFromStyles(true)
    setRows({ 'heading 1': 1, 'heading 2': 2, 'heading 3': 3 })
    setOutline(true)
    setEntryFields(false)
  }
  const sources = fromStyles ? sourcesOfRows(rows, names) : { levels: 0, styles: [] }
  const any = sources.levels > 0 || sources.styles.length > 0 || outline || entryFields
  return (
    <div
      className="modal-backdrop"
      ref={keys.ref}
      onKeyDown={keys.onKeyDown}
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
    >
      <form
        className="modal gs-form tco-dialog it-options-dialog"
        data-it-dialog="toc-options"
        noValidate
        onSubmit={(e) => {
          e.preventDefault()
          e.stopPropagation()
          if (any) onApply({ ...sources, outline, entryFields })
        }}
      >
        <h2>{t('ribbonItTocOptionsTitle')}</h2>
        <p className="index-tables-hint">{t('ribbonItBuildTocFrom')}</p>
        <label className="tco-check">
          <input
            type="checkbox"
            data-it="from-styles"
            autoFocus
            checked={fromStyles}
            onChange={(e) => setFromStyles(e.target.checked)}
          />
          {t('ribbonGroupStyles')}
        </label>
        <div className="it-style-levels" role="table">
          <div className="it-style-levels-head" role="row">
            <span role="columnheader">{t('ribbonItAvailableStyles')}</span>
            <span role="columnheader">{t('ribbonItTocLevel')}</span>
          </div>
          {names.map((name) => (
            <label key={name} role="row" data-it-style={name.toLowerCase()}>
              <span role="cell">{styleLabel({ name }, t)}</span>
              <input
                type="number"
                min={1}
                max={9}
                disabled={!fromStyles}
                value={rows[name.toLowerCase()] ?? ''}
                onChange={(e) => {
                  const level = Math.round(Number(e.target.value))
                  setRows((prev) => {
                    const { [name.toLowerCase()]: _dropped, ...rest } = prev
                    return level >= 1 && level <= 9
                      ? { ...rest, [name.toLowerCase()]: level }
                      : rest
                  })
                }}
              />
            </label>
          ))}
        </div>
        <label className="tco-check">
          <input
            type="checkbox"
            data-it="outline-levels"
            checked={outline}
            onChange={(e) => setOutline(e.target.checked)}
          />
          {t('ribbonItOutlineLevels')}
        </label>
        <label className="tco-check">
          <input
            type="checkbox"
            data-it="entry-fields"
            checked={entryFields}
            onChange={(e) => setEntryFields(e.target.checked)}
          />
          {t('ribbonItEntryFields')}
        </label>
        <div className="modal-actions">
          <div className="tco-actions-right">
            <button type="button" data-it="reset" onClick={reset}>
              {t('ribbonBnReset')}
            </button>
          </div>
          <div className="tco-actions-right">
            <button type="button" onClick={onClose}>
              {t('ribbonCancel')}
            </button>
            <button type="submit" className="primary" disabled={!any}>
              {t('ribbonOk')}
            </button>
          </div>
        </div>
      </form>
    </div>
  )
}

const IDENTIFIERS = Array.from({ length: 26 }, (_, i) => String.fromCharCode(65 + i))

export interface FiguresSourceValues {
  /** the paragraph style listed in place of the label's captions; '' for the captions */
  style: string
  /** identifier of the TC fields listed; '' for none */
  entryId: string
}

/** Word's Table of Figures Options: a paragraph style (TOC \t) and TC fields of one identifier (\f) as the source */
export function FiguresOptionsDialog({
  values,
  styles,
  onApply,
  onClose,
}: {
  values: FiguresSourceValues
  styles: StyleMap
  onApply: (next: FiguresSourceValues) => void
  onClose: () => void
}) {
  const { t } = useI18n()
  const keys = useModalKeys(onClose)
  const names = useMemo(() => {
    const listed = paragraphStyleNames(styles)
    return values.style && !listed.some((n) => n.toLowerCase() === values.style.toLowerCase())
      ? [values.style, ...listed]
      : listed
  }, [styles, values.style])
  const caption = names.find((n) => n.toLowerCase() === 'caption') ?? names[0]
  const [byStyle, setByStyle] = useState(!!values.style)
  const [style, setStyle] = useState(
    names.find((n) => n.toLowerCase() === values.style.toLowerCase()) ?? caption,
  )
  const [byFields, setByFields] = useState(!!values.entryId)
  const [entryId, setEntryId] = useState(values.entryId.toUpperCase() || 'F')
  return (
    <div
      className="modal-backdrop"
      ref={keys.ref}
      onKeyDown={keys.onKeyDown}
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
    >
      <form
        className="modal gs-form tco-dialog it-options-dialog"
        data-it-dialog="figures-options"
        noValidate
        onSubmit={(e) => {
          e.preventDefault()
          e.stopPropagation()
          onApply({ style: byStyle ? style : '', entryId: byFields ? entryId : '' })
        }}
      >
        <h2>{t('ribbonItFiguresOptionsTitle')}</h2>
        <p className="index-tables-hint">{t('ribbonItBuildFiguresFrom')}</p>
        <div className="it-options-row">
          <label className="tco-check">
            <input
              type="checkbox"
              data-it="by-style"
              autoFocus
              checked={byStyle}
              onChange={(e) => setByStyle(e.target.checked)}
            />
            {t('ribbonPbStyle')}
          </label>
          <select
            data-it="style"
            aria-label={t('ribbonPbStyle')}
            disabled={!byStyle}
            value={style}
            onChange={(e) => setStyle(e.target.value)}
          >
            {names.map((name) => (
              <option key={name} value={name}>
                {styleLabel({ name }, t)}
              </option>
            ))}
          </select>
        </div>
        <div className="it-options-row">
          <label className="tco-check">
            <input
              type="checkbox"
              data-it="entry-fields"
              checked={byFields}
              onChange={(e) => setByFields(e.target.checked)}
            />
            {t('ribbonItEntryFields')}
          </label>
          <label className="it-options-id">
            {t('ribbonItTableIdentifier')}
            <select
              data-it="identifier"
              disabled={!byFields}
              value={entryId}
              onChange={(e) => setEntryId(e.target.value)}
            >
              {IDENTIFIERS.map((id) => (
                <option key={id} value={id}>
                  {id}
                </option>
              ))}
            </select>
          </label>
        </div>
        <div className="modal-actions">
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

/**
 * Word's Style list behind Modify: the built-in styles the entries of this
 * kind of table are written in. Picking one opens the Styles pane's Modify
 * Style dialog on it; a style the document does not define yet is created.
 */
export function EntryStylesDialog({
  kind,
  styles,
  pending,
  onUpsert,
  onClose,
}: {
  kind: IndexTablesKind
  styles: StyleMap
  pending: Readonly<Record<string, StyleUpsert>>
  onUpsert: (up: StyleUpsert) => void
  onClose: () => void
}) {
  const { t } = useI18n()
  const keys = useModalKeys(onClose)
  const entries = ENTRY_STYLES[kind]
  const [picked, setPicked] = useState(entries[0].styleId)
  const [editing, setEditing] = useState<StyleInfo | null>(null)
  const infoOf = (styleId: string): StyleInfo => {
    const entry = entries.find((e) => e.styleId === styleId) ?? entries[0]
    return (
      styles.get(entryStyleId(entry, styles)) ?? {
        styleId: entry.styleId,
        name: entry.name,
        type: 'paragraph',
        basedOn: defaultParagraphStyleId(styles),
      }
    )
  }
  if (editing)
    return (
      <StyleDialog
        target={editing}
        initial={formValuesOf(editing)}
        styles={styles}
        pending={pending}
        onSubmit={(up) => {
          const entry = entries.find((e) => entryStyleId(e, styles) === up.styleId)
          setEditing(null)
          onUpsert(
            entry && !styles.has(up.styleId) && !(up.styleId in pending)
              ? { ...entryStyleUpsert(entry, {}, styles, defaultParagraphStyleId(styles)), ...up }
              : up,
          )
        }}
        onClose={() => setEditing(null)}
      />
    )
  return (
    <div
      className="modal-backdrop"
      ref={keys.ref}
      onKeyDown={keys.onKeyDown}
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
    >
      <div className="modal gs-form tco-dialog it-options-dialog" data-it-dialog="entry-styles">
        <h2>{t('ribbonPbStyle')}</h2>
        <label className="it-entry-styles">
          {t('ribbonGroupStyles')}
          <select
            data-it="entry-style"
            size={Math.min(9, Math.max(3, entries.length))}
            autoFocus
            value={picked}
            onChange={(e) => setPicked(e.target.value)}
            onDoubleClick={() => setEditing(infoOf(picked))}
          >
            {entries.map((entry) => (
              <option key={entry.styleId} value={entry.styleId}>
                {styleLabel(infoOf(entry.styleId), t)}
              </option>
            ))}
          </select>
        </label>
        <div className="modal-actions">
          <div className="tco-actions-right">
            <button type="button" data-it="modify-style" onClick={() => setEditing(infoOf(picked))}>
              {t('ribbonStyleModify')}
            </button>
          </div>
          <div className="tco-actions-right">
            <button type="button" className="primary" onClick={onClose}>
              {t('ribbonClose')}
            </button>
          </div>
        </div>
      </div>
    </div>
  )
}
