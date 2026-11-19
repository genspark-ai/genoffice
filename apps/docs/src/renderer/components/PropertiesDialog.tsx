import { useEffect, useState } from 'react'
import {
  customPropertyKind,
  type CustomPropertyKind,
  type DocumentProperties,
  type PropertyChanges,
} from '@genoffice/docx-engine'
import type { DocFileInfo } from '../../shared/ipc'
import type { HeadingRef } from '../editor/headings'
import { useI18n, type StringKey } from '../i18n/locale'
import { useModalKeys } from './modal-keys'
import type { DocStats } from './WordCountDialog'

/** the Summary fields of Word's Properties, in its order */
const FIELDS: ReadonlyArray<{ part: 'core' | 'app'; key: string; label: StringKey }> = [
  { part: 'core', key: 'title', label: 'ribbonSourceTitle' },
  { part: 'core', key: 'subject', label: 'ribbonPropSubject' },
  { part: 'core', key: 'creator', label: 'ribbonSourceAuthor' },
  { part: 'app', key: 'manager', label: 'ribbonPropManager' },
  { part: 'app', key: 'company', label: 'ribbonPropCompany' },
  { part: 'core', key: 'category', label: 'ribbonPropCategory' },
  { part: 'core', key: 'keywords', label: 'ribbonPropKeywords' },
  { part: 'core', key: 'description', label: 'ribbonPropComments' },
]

const fieldId = (field: (typeof FIELDS)[number]) => `${field.part}.${field.key}`

/** the saved properties with the edits waiting for the next save on top */
export function summaryValues(
  saved: Pick<DocumentProperties, 'core' | 'app'>,
  pending: PropertyChanges | null,
): Record<string, string> {
  return Object.fromEntries(
    FIELDS.map((field) => {
      const edit = pending?.[field.part]?.[field.key]
      return [
        fieldId(field),
        edit === undefined ? (saved[field.part][field.key] ?? '') : (edit ?? ''),
      ]
    }),
  )
}

/** what turns `before` into `after`: an emptied field removes its property; null when nothing changed */
export function summaryChanges(
  before: Record<string, string>,
  after: Record<string, string>,
): PropertyChanges | null {
  const changes: Required<Pick<PropertyChanges, 'core' | 'app'>> = { core: {}, app: {} }
  for (const field of FIELDS) {
    const value = after[fieldId(field)].trim()
    if (value !== before[fieldId(field)]) changes[field.part][field.key] = value || null
  }
  const core = Object.keys(changes.core).length > 0
  const app = Object.keys(changes.app).length > 0
  if (!core && !app) return null
  return { ...(core ? { core: changes.core } : {}), ...(app ? { app: changes.app } : {}) }
}

/** a row of the Custom tab; a date is its day (YYYY-MM-DD), a yes or no is true / false */
export interface CustomEntry {
  name: string
  /** null for a stored type the tab has no type for: listed, and only Delete acts on it */
  kind: CustomPropertyKind | null
  value: string
}

const KINDS: ReadonlyArray<{ kind: CustomPropertyKind; label: StringKey }> = [
  { kind: 'text', label: 'ribbonGroupText' },
  { kind: 'date', label: 'ribbonFieldDate' },
  { kind: 'number', label: 'ribbonSortNumber' },
  { kind: 'bool', label: 'ribbonPropYesNo' },
]

const kindOfValue = (value: string | number | boolean): CustomPropertyKind =>
  typeof value === 'boolean' ? 'bool' : typeof value === 'number' ? 'number' : 'text'

const shown = (kind: CustomPropertyKind | null, value: string) =>
  kind === 'date' ? value.slice(0, 10) : value

/** the saved custom properties, in the file's order, with the edits waiting for the next save on top */
export function customEntries(
  saved: Pick<DocumentProperties, 'custom' | 'customTypes'>,
  pending: PropertyChanges | null,
): CustomEntry[] {
  const entries = new Map<string, CustomEntry>()
  for (const [name, value] of Object.entries(saved.custom)) {
    const kind = customPropertyKind(saved.customTypes[name])
    entries.set(name, { name, kind, value: shown(kind, value) })
  }
  for (const [name, value] of Object.entries(pending?.custom ?? {})) {
    if (value === null) {
      entries.delete(name)
      continue
    }
    const kind = pending?.customKinds?.[name] ?? kindOfValue(value)
    entries.set(name, { name, kind, value: shown(kind, String(value)) })
  }
  return [...entries.values()]
}

/** whether the tab can store the text as the type: a number, a day, some text */
export function customValueFits(kind: CustomPropertyKind, value: string): boolean {
  const text = value.trim()
  if (kind === 'bool') return text === 'true' || text === 'false'
  if (kind === 'number') return /^[+-]?(?:\d+\.?\d*|\.\d+)$/.test(text)
  if (kind === 'date')
    return /^\d{4}-\d{2}-\d{2}$/.test(text) && !Number.isNaN(Date.parse(text + 'T00:00:00Z'))
  return text !== ''
}

/** what turns `before` into `after`: a row gone removes its property; null when nothing changed */
export function customChanges(
  before: readonly CustomEntry[],
  after: readonly CustomEntry[],
): Pick<PropertyChanges, 'custom' | 'customKinds'> | null {
  const custom: NonNullable<PropertyChanges['custom']> = {}
  const customKinds: NonNullable<PropertyChanges['customKinds']> = {}
  const kept = new Map(after.map((entry) => [entry.name, entry]))
  const had = new Map(before.map((entry) => [entry.name, entry]))
  for (const entry of before) if (!kept.has(entry.name)) custom[entry.name] = null
  for (const entry of after) {
    const old = had.get(entry.name)
    if (!entry.kind || (old && old.kind === entry.kind && old.value === entry.value)) continue
    custom[entry.name] = entry.value
    customKinds[entry.name] = entry.kind
  }
  return Object.keys(custom).length > 0 ? { custom, customKinds } : null
}

const TABS = [
  { id: 'general', label: 'ribbonPropGeneral' },
  { id: 'summary', label: 'ribbonPropSummary' },
  { id: 'statistics', label: 'ribbonPropStatistics' },
  { id: 'contents', label: 'ribbonPropContents' },
  { id: 'custom', label: 'ribbonPbCustom' },
] as const satisfies ReadonlyArray<{ id: string; label: StringKey }>

type TabId = (typeof TABS)[number]['id']

const LIVE_COUNTS = [
  ['appStatPages', 'pages'],
  ['appStatParagraphs', 'paragraphs'],
  ['appStatLines', 'lines'],
  ['appStatWords', 'words'],
  ['appStatCharsNoSpaces', 'charsNoSpace'],
  ['appStatCharsWithSpaces', 'charsWithSpace'],
] as const satisfies ReadonlyArray<readonly [StringKey, keyof DocStats]>

const BLANK_CUSTOM: CustomEntry = { name: '', kind: 'text', value: '' }

/**
 * Word's File > Properties. Summary and Custom are written to docProps on
 * save; General reads the file on disk, Statistics the file's own properties
 * and the open document's counts, Contents its headings.
 */
export function PropertiesDialog({
  fileName,
  properties,
  pending,
  loadStats,
  loadHeadings,
  editable,
  loadFileInfo,
  onApply,
  onClose,
}: {
  fileName: string
  properties: DocumentProperties
  pending: PropertyChanges | null
  /** the open document's counts as they are now; null when they cannot be taken */
  loadStats: () => DocStats | null
  loadHeadings: () => readonly HeadingRef[]
  editable: boolean
  /** resolves to null for a document that has no file yet */
  loadFileInfo: () => Promise<DocFileInfo | null>
  onApply: (changes: PropertyChanges) => void
  onClose: () => void
}) {
  const { t, dateLocale } = useI18n()
  const keys = useModalKeys(onClose)
  const [tab, setTab] = useState<TabId>('summary')
  const [values] = useState(() => summaryValues(properties, pending))
  const [draft, setDraft] = useState(values)
  const [savedCustom] = useState(() => customEntries(properties, pending))
  const [custom, setCustom] = useState(savedCustom)
  const [entry, setEntry] = useState(BLANK_CUSTOM)
  const [file, setFile] = useState<DocFileInfo | null>(null)
  const [stats] = useState(loadStats)
  const [headings] = useState(loadHeadings)
  useEffect(() => {
    let live = true
    void loadFileInfo().then((info) => {
      if (live) setFile(info)
    })
    return () => {
      live = false
    }
    // read once, when the dialog opens
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])

  const when = (value: number | string | undefined) => {
    if (value === undefined || value === '') return ''
    const date = new Date(value)
    return Number.isNaN(date.getTime()) ? String(value) : date.toLocaleString(dateLocale)
  }
  const count = (n: number) => n.toLocaleString(dateLocale)
  const rows = (items: ReadonlyArray<readonly [string, string, string]>) => (
    <table className="stats-table">
      <tbody>
        {items.map(([id, label, value]) => (
          <tr key={id} data-prop={id}>
            <td>{label}</td>
            <td>{value}</td>
          </tr>
        ))}
      </tbody>
    </table>
  )

  const separator = file ? (file.path.includes('\\') ? '\\' : '/') : '/'
  const general = rows([
    ['name', t('ribbonStyleDialogName'), fileName],
    [
      'location',
      t('ribbonPropLocation'),
      file ? file.path.slice(0, file.path.lastIndexOf(separator)) : '',
    ],
    [
      'size',
      t('ribbonPaperSize'),
      file
        ? t('ribbonPropSizeValue', {
            kb: count(Math.max(1, Math.round(file.sizeBytes / 1024))),
            n: count(file.sizeBytes),
          })
        : '',
    ],
    ['created', t('ribbonPropCreated'), when(file?.createdMs)],
    ['modified', t('ribbonPropModified'), when(file?.modifiedMs)],
  ])

  const minutes = Number(properties.app.totalTime)
  const statistics = (
    <>
      {rows([
        ['created', t('ribbonPropCreated'), when(properties.core.created)],
        ['modified', t('ribbonPropModified'), when(properties.core.modified)],
        ['lastModifiedBy', t('ribbonPropLastSavedBy'), properties.core.lastModifiedBy ?? ''],
        ['revision', t('ribbonPropRevision'), properties.core.revision ?? ''],
        [
          'totalTime',
          t('ribbonPropEditingTime'),
          Number.isFinite(minutes) && properties.app.totalTime
            ? t('ribbonPropMinutes', { n: count(minutes) })
            : '',
        ],
      ])}
      {stats &&
        rows(LIVE_COUNTS.map(([label, field]) => [field, t(label), count(stats[field])] as const))}
    </>
  )

  const contents = (
    <ul className="properties-contents">
      {headings.map((heading, index) => (
        <li key={index} style={{ paddingInlineStart: (heading.level - 1) * 16 }}>
          {heading.text}
        </li>
      ))}
    </ul>
  )

  const name = entry.name.trim()
  const existing = custom.find((item) => item.name === name)
  const kind = entry.kind
  const fits = kind !== null && name !== '' && customValueFits(kind, entry.value)
  const unchanged = existing?.kind === kind && existing.value === entry.value.trim()
  const put = () => {
    if (!fits) return
    const next = { name, kind, value: entry.value.trim() }
    setCustom((prev) =>
      existing ? prev.map((item) => (item.name === name ? next : item)) : [...prev, next],
    )
    setEntry(BLANK_CUSTOM)
  }
  const shownValue = (item: CustomEntry) =>
    item.kind === 'bool'
      ? t(item.value === 'true' ? 'ribbonPropYes' : 'ribbonPropNo')
      : item.kind === 'date'
        ? new Date(item.value + 'T00:00:00').toLocaleDateString(dateLocale)
        : item.value
  const customTab = (
    <>
      <div
        className="properties-custom-form"
        // Enter adds the property, as Word's default button there does, instead of closing the dialog
        onKeyDown={(e) => {
          if (e.key !== 'Enter') return
          e.preventDefault()
          put()
        }}
      >
        <label>
          {t('ribbonStyleDialogName')}
          <input
            name="customName"
            value={entry.name}
            readOnly={!editable}
            onChange={(e) => setEntry((prev) => ({ ...prev, name: e.target.value }))}
          />
        </label>
        <label>
          {t('ribbonSortType')}
          <select
            name="customType"
            value={kind ?? ''}
            disabled={!editable || kind === null}
            onChange={(e) => {
              const next = e.target.value as CustomPropertyKind
              setEntry((prev) => ({ ...prev, kind: next, value: next === 'bool' ? 'true' : '' }))
            }}
          >
            {kind === null && <option value="" />}
            {KINDS.map((option) => (
              <option key={option.kind} value={option.kind}>
                {t(option.label)}
              </option>
            ))}
          </select>
        </label>
        {kind === 'bool' ? (
          <div className="properties-yes-no">
            {(['true', 'false'] as const).map((value) => (
              <label key={value} className="tco-check">
                <input
                  type="radio"
                  name="customValue"
                  value={value}
                  checked={entry.value === value}
                  disabled={!editable}
                  onChange={() => setEntry((prev) => ({ ...prev, value }))}
                />
                {t(value === 'true' ? 'ribbonPropYes' : 'ribbonPropNo')}
              </label>
            ))}
          </div>
        ) : (
          <label>
            {t('ribbonPropValue')}
            <input
              name="customValue"
              type={kind === 'date' ? 'date' : 'text'}
              inputMode={kind === 'number' ? 'decimal' : undefined}
              value={entry.value}
              readOnly={!editable || kind === null}
              onChange={(e) => setEntry((prev) => ({ ...prev, value: e.target.value }))}
            />
          </label>
        )}
        <div className="modal-actions properties-custom-actions">
          <button
            type="button"
            name="customAdd"
            disabled={!editable || !fits || unchanged}
            onClick={put}
          >
            {t(existing ? 'ribbonPropModify' : 'ribbonAdd')}
          </button>
          <button
            type="button"
            name="customDelete"
            disabled={!editable || !existing}
            onClick={() => {
              setCustom((prev) => prev.filter((item) => item.name !== name))
              setEntry(BLANK_CUSTOM)
            }}
          >
            {t('ribbonGroupDelete')}
          </button>
        </div>
      </div>
      <div className="recipients-table properties-custom-list">
        <table>
          <thead>
            <tr>
              <th>{t('ribbonStyleDialogName')}</th>
              <th>{t('ribbonPropValue')}</th>
              <th>{t('ribbonSortType')}</th>
            </tr>
          </thead>
          <tbody>
            {custom.map((item) => (
              <tr
                key={item.name}
                className={item.name === name ? 'selected' : ''}
                onClick={() => setEntry(item)}
              >
                <td>{item.name}</td>
                <td>{shownValue(item)}</td>
                <td>
                  {t(KINDS.find((option) => option.kind === item.kind)?.label ?? 'ribbonSortOther')}
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </>
  )

  const summary = FIELDS.map((field) => {
    const id = fieldId(field)
    const set = (value: string) => setDraft((d) => ({ ...d, [id]: value }))
    return (
      <label key={id}>
        {t(field.label)}
        {field.key === 'description' ? (
          <textarea
            rows={3}
            name={id}
            value={draft[id]}
            readOnly={!editable}
            onChange={(e) => set(e.target.value)}
          />
        ) : (
          <input
            name={id}
            value={draft[id]}
            readOnly={!editable}
            onChange={(e) => set(e.target.value)}
          />
        )}
      </label>
    )
  })

  const panels: Record<TabId, React.ReactNode> = {
    general,
    summary,
    statistics,
    contents,
    custom: customTab,
  }
  return (
    <div
      className="modal-backdrop"
      ref={keys.ref}
      onKeyDown={keys.onKeyDown}
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
    >
      <form
        className="modal gs-form tco-dialog doc-dialog properties-dialog"
        noValidate
        onSubmit={(e) => {
          e.preventDefault()
          onClose()
          const changes = {
            ...summaryChanges(values, draft),
            ...customChanges(savedCustom, custom),
          }
          if (editable && Object.keys(changes).length > 0) onApply(changes)
        }}
      >
        <h2>{t('ribbonPropertiesTitle')}</h2>
        <div className="bs-tabs" role="tablist">
          {TABS.map(({ id, label }) => (
            <button
              key={id}
              type="button"
              role="tab"
              aria-selected={tab === id}
              className={tab === id ? 'active' : ''}
              onClick={() => setTab(id)}
            >
              {t(label)}
            </button>
          ))}
        </div>
        <div className="properties-panel" role="tabpanel" data-tab={tab}>
          {panels[tab]}
        </div>
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
