import { Fragment, useState } from 'react'
import type { MergeDataSource } from '../../shared/ipc'
import { useI18n, type StringKey } from '../i18n/locale'
import {
  EMPTY_RECIPIENT_QUERY,
  MAX_RECIPIENT_CONDITIONS,
  MAX_RECIPIENT_SORT_KEYS,
  patchQueryRow,
  RECIPIENT_OPS,
  opTakesValue,
  recipientConditions,
  type RecipientJoin,
  type RecipientOp,
  type RecipientQuery,
  type RecipientSortKey,
} from '../mail-merge/recipients'
import {
  RULE_OPS,
  askInstruction,
  fillInInstruction,
  ifInstruction,
  nextIfInstruction,
  ruleNeedsList,
  setInstruction,
  skipIfInstruction,
  type MergeRule,
  type RuleOp,
} from '../mail-merge/rules'
import { validBookmarkName } from '../editor/cross-reference'
import { useModalKeys } from './modal-keys'

const RULE_OP_LABELS: Record<RuleOp, StringKey> = {
  '=': 'ribbonFilterEquals',
  '<>': 'ribbonFilterNotEquals',
  '<': 'ribbonFilterLess',
  '>': 'ribbonFilterGreater',
  '<=': 'ribbonFilterLessEqual',
  '>=': 'ribbonFilterGreaterEqual',
}

const OP_LABELS: Record<RecipientOp, StringKey> = {
  equals: 'ribbonFilterEquals',
  notEquals: 'ribbonFilterNotEquals',
  contains: 'ribbonFilterContains',
  notContains: 'ribbonFilterNotContains',
  blank: 'ribbonFilterBlank',
  notBlank: 'ribbonFilterNotBlank',
}

function Backdrop({ onClose, children }: { onClose: () => void; children: React.ReactNode }) {
  const keys = useModalKeys(onClose)
  return (
    <div
      className="modal-backdrop"
      ref={keys.ref}
      onKeyDown={keys.onKeyDown}
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
    >
      {children}
    </div>
  )
}

/** Word's Mail Merge Recipients: untick a row to leave that recipient out of the merge */
export function RecipientsDialog({
  source,
  excluded,
  onApply,
  onClose,
}: {
  source: MergeDataSource
  excluded: ReadonlySet<number>
  onApply: (excluded: Set<number>) => void
  onClose: () => void
}) {
  const { t } = useI18n()
  const [off, setOff] = useState(() => new Set(excluded))
  const total = source.records.length
  const toggle = (index: number) =>
    setOff((prev) => {
      const next = new Set(prev)
      if (!next.delete(index)) next.add(index)
      return next
    })
  return (
    <Backdrop onClose={onClose}>
      <form
        className="modal gs-form tco-dialog recipients-dialog"
        noValidate
        onSubmit={(e) => {
          e.preventDefault()
          onClose()
          onApply(off)
        }}
      >
        <h2>{t('ribbonEditRecipientList')}</h2>
        <div className="recipients-table">
          <table>
            <thead>
              <tr>
                <th />
                {source.fields.map((name) => (
                  <th key={name}>{name}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {source.records.map((record, index) => (
                <tr key={index}>
                  <td>
                    <input
                      type="checkbox"
                      checked={!off.has(index)}
                      onChange={() => toggle(index)}
                    />
                  </td>
                  {source.fields.map((name) => (
                    <td key={name}>{record[name]}</td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        <div className="modal-actions">
          <span className="recipients-count">
            {t('ribbonRecipientsCount', { n: total - off.size, total })}
          </span>
          <div className="tco-actions-right">
            <button type="button" onClick={() => setOff(new Set())}>
              {t('ribbonSelectAll')}
            </button>
            <button
              type="button"
              onClick={() => setOff(new Set(source.records.map((_, index) => index)))}
            >
              {t('ribbonClearAll')}
            </button>
            <button type="button" onClick={onClose}>
              {t('ribbonCancel')}
            </button>
            <button type="submit" className="primary">
              {t('ribbonOk')}
            </button>
          </div>
        </div>
      </form>
    </Backdrop>
  )
}

interface ConditionRow {
  join: RecipientJoin
  field: string
  op: RecipientOp
  value: string
}

const BLANK_CONDITION: ConditionRow = { join: 'and', field: '', op: 'equals', value: '' }

const QUERY_TABS = [
  { id: 'filter', label: 'ribbonFilterRecords' },
  { id: 'sort', label: 'ribbonSortRecords' },
] as const satisfies ReadonlyArray<{ id: string; label: StringKey }>

const JOIN_LABELS: Record<RecipientJoin, StringKey> = {
  and: 'ribbonFilterAnd',
  or: 'ribbonFilterOr',
}

/**
 * Word's Query Options: Filter Records takes the recipients that meet up to six
 * conditions joined by And / Or, Sort Records orders them by up to three fields
 */
export function FilterRecipientsDialog({
  fields,
  query,
  onApply,
  onClose,
}: {
  fields: string[]
  query: RecipientQuery
  onApply: (query: RecipientQuery) => void
  onClose: () => void
}) {
  const { t } = useI18n()
  const [tab, setTab] = useState<(typeof QUERY_TABS)[number]['id']>('filter')
  const [rows, setRows] = useState<ConditionRow[]>(() =>
    Array.from({ length: MAX_RECIPIENT_CONDITIONS }, (_, index) => {
      const condition = query.conditions[index]
      return condition ? { ...condition } : { ...BLANK_CONDITION }
    }),
  )
  const [sort, setSort] = useState<RecipientSortKey[]>(() =>
    Array.from(
      { length: MAX_RECIPIENT_SORT_KEYS },
      (_, index) => query.sort[index] ?? { field: '', descending: false },
    ),
  )
  const setRow = (index: number, patch: Partial<ConditionRow>) =>
    setRows((prev) => patchQueryRow(prev, index, patch, BLANK_CONDITION))
  const setSortKey = (index: number, patch: Partial<RecipientSortKey>) =>
    setSort((prev) => patchQueryRow(prev, index, patch, { field: '', descending: false }))
  const fieldOptions = (
    <>
      <option value="">{t('ribbonSortNoKey')}</option>
      {fields.map((name) => (
        <option key={name} value={name}>
          {name}
        </option>
      ))}
    </>
  )
  // the first row keeps the names the single-condition form had
  const named = (name: string, index: number) => (index === 0 ? name : `${name}${index + 1}`)
  return (
    <Backdrop onClose={onClose}>
      <form
        className="modal gs-form tco-dialog hlc-dialog filter-recipients-dialog"
        noValidate
        onSubmit={(e) => {
          e.preventDefault()
          onClose()
          onApply({
            conditions: recipientConditions(rows),
            sort: sort.filter((key) => key.field),
          })
        }}
      >
        <h2>{t('ribbonFilterRecipients')}</h2>
        <div className="bs-tabs" role="tablist">
          {QUERY_TABS.map(({ id, label }) => (
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
        {tab === 'filter' ? (
          <div className="filter-grid" role="tabpanel">
            <span />
            <span>{t('ribbonField')}</span>
            <span>{t('ribbonFilterComparison')}</span>
            <span>{t('ribbonFilterCompareTo')}</span>
            {rows.map((row, index) => {
              // Word opens a row once the one above it names a field
              const locked = index > 0 && !rows[index - 1].field
              const needsValue = opTakesValue(row.op)
              return (
                <Fragment key={index}>
                  {index === 0 ? (
                    <span />
                  ) : (
                    <select
                      name={named('join', index)}
                      aria-label={`${t('ribbonFilterAnd')} / ${t('ribbonFilterOr')}`}
                      value={row.join}
                      disabled={locked}
                      onChange={(e) => setRow(index, { join: e.target.value as RecipientJoin })}
                    >
                      {(['and', 'or'] as const).map((join) => (
                        <option key={join} value={join}>
                          {t(JOIN_LABELS[join])}
                        </option>
                      ))}
                    </select>
                  )}
                  <select
                    name={named('field', index)}
                    aria-label={t('ribbonField')}
                    value={row.field}
                    disabled={locked}
                    onChange={(e) => setRow(index, { field: e.target.value })}
                  >
                    {fieldOptions}
                  </select>
                  <select
                    name={named('op', index)}
                    aria-label={t('ribbonFilterComparison')}
                    value={row.op}
                    disabled={locked || !row.field}
                    onChange={(e) => setRow(index, { op: e.target.value as RecipientOp })}
                  >
                    {RECIPIENT_OPS.map((code) => (
                      <option key={code} value={code}>
                        {t(OP_LABELS[code])}
                      </option>
                    ))}
                  </select>
                  <input
                    name={named('value', index)}
                    aria-label={t('ribbonFilterCompareTo')}
                    value={needsValue ? row.value : ''}
                    disabled={locked || !row.field || !needsValue}
                    onChange={(e) => setRow(index, { value: e.target.value })}
                  />
                </Fragment>
              )
            })}
          </div>
        ) : (
          <div role="tabpanel">
            {sort.map((key, index) => {
              const locked = index > 0 && !sort[index - 1].field
              return (
                <fieldset key={index} disabled={locked}>
                  <legend>{t(index === 0 ? 'ribbonSortBy' : 'ribbonSortThenBy')}</legend>
                  <div className="sort-key">
                    <select
                      name={`sort${index + 1}`}
                      value={key.field}
                      onChange={(e) => setSortKey(index, { field: e.target.value })}
                    >
                      {fieldOptions}
                    </select>
                    <span className="sort-order">
                      {[false, true].map((down) => (
                        <label key={String(down)} className="tco-check">
                          <input
                            type="radio"
                            name={`sortOrder${index + 1}`}
                            value={down ? 'descending' : 'ascending'}
                            checked={key.descending === down}
                            disabled={!key.field}
                            onChange={() => setSortKey(index, { descending: down })}
                          />
                          {t(down ? 'ribbonSortDescending' : 'ribbonSortAscending')}
                        </label>
                      ))}
                    </span>
                  </div>
                </fieldset>
              )
            })}
          </div>
        )}
        <div className="modal-actions">
          <button
            type="button"
            onClick={() => {
              onClose()
              onApply(EMPTY_RECIPIENT_QUERY)
            }}
          >
            {t('ribbonClearAll')}
          </button>
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
    </Backdrop>
  )
}

/** Word's Find in Field: moves the previewed record to the next one holding the text; stays open */
export function FindRecipientDialog({
  fields,
  onFind,
  onClose,
}: {
  fields: string[]
  /** false when no recipient holds the text */
  onFind: (text: string, field: string | null) => boolean
  onClose: () => void
}) {
  const { t } = useI18n()
  const [text, setText] = useState('')
  const [field, setField] = useState('')
  const [missed, setMissed] = useState('')
  return (
    <Backdrop onClose={onClose}>
      <form
        className="modal gs-form tco-dialog hlc-dialog find-recipient-dialog"
        noValidate
        onSubmit={(e) => {
          e.preventDefault()
          setMissed(onFind(text, field || null) ? '' : text)
        }}
      >
        <h2>{t('ribbonFindRecipient')}</h2>
        <label>
          {t('ribbonFind')}
          <input name="text" value={text} onChange={(e) => setText(e.target.value)} />
        </label>
        <label>
          {t('ribbonFindLookIn')}
          <select name="field" value={field} onChange={(e) => setField(e.target.value)}>
            <option value="">{t('ribbonFindAllFields')}</option>
            {fields.map((name) => (
              <option key={name} value={name}>
                {name}
              </option>
            ))}
          </select>
        </label>
        {missed && (
          <div className="modal-error">{t('ribbonRecipientNotFound', { text: missed })}</div>
        )}
        <div className="modal-actions">
          <span />
          <div className="tco-actions-right">
            <button type="button" onClick={onClose}>
              {t('ribbonClose')}
            </button>
            <button type="submit" className="primary" disabled={!text.trim()}>
              {t('ribbonFindNext')}
            </button>
          </div>
        </div>
      </form>
    </Backdrop>
  )
}

const RULE_TITLES: Record<MergeRule, StringKey> = {
  if: 'ribbonRulesIf',
  skipIf: 'ribbonRulesSkipIf',
  nextIf: 'ribbonRulesNextIf',
  set: 'ribbonRulesSet',
  ask: 'ribbonRulesAsk',
  fillIn: 'ribbonRulesFillIn',
}

/**
 * The dialogs of Word's Rules menu: a condition on one of the list's fields
 * (If…Then…Else, Skip Record If, Next Record If), a bookmark and its text
 * (Set Bookmark), or a question the merge asks (Ask, Fill-in).
 */
export function RuleDialog({
  rule,
  fields,
  onInsert,
  onClose,
}: {
  rule: MergeRule
  fields: string[]
  /** takes the field instruction to insert at the caret */
  onInsert: (instr: string) => void
  onClose: () => void
}) {
  const { t } = useI18n()
  const [field, setField] = useState(fields[0] ?? '')
  const [op, setOp] = useState<RuleOp>('=')
  const [value, setValue] = useState('')
  const [then, setThen] = useState('')
  const [otherwise, setOtherwise] = useState('')
  const [bookmark, setBookmark] = useState('')
  const [prompt, setPrompt] = useState('')
  const [once, setOnce] = useState(false)
  const conditional = ruleNeedsList(rule)
  const named = rule === 'set' || rule === 'ask'
  const asks = rule === 'ask' || rule === 'fillIn'
  const ready =
    (!conditional || field !== '') &&
    (!named || validBookmarkName(bookmark.trim())) &&
    (!asks || prompt.trim() !== '')
  const instruction = (): string => {
    switch (rule) {
      case 'if':
        return ifInstruction(field, op, value, then, otherwise)
      case 'skipIf':
        return skipIfInstruction(field, op, value)
      case 'nextIf':
        return nextIfInstruction(field, op, value)
      case 'set':
        return setInstruction(bookmark.trim(), value)
      case 'ask':
        return askInstruction(bookmark.trim(), prompt.trim(), value, once)
      case 'fillIn':
        return fillInInstruction(prompt.trim(), value, once)
    }
  }
  return (
    <Backdrop onClose={onClose}>
      <form
        className="modal gs-form tco-dialog hlc-dialog rule-dialog"
        noValidate
        onSubmit={(e) => {
          e.preventDefault()
          if (!ready) return
          onClose()
          onInsert(instruction())
        }}
      >
        <h2>{t(RULE_TITLES[rule])}</h2>
        {named && (
          <label>
            {t('ribbonBookmark')}
            <input
              name="bookmark"
              value={bookmark}
              autoFocus
              spellCheck={false}
              onChange={(e) => setBookmark(e.target.value)}
            />
          </label>
        )}
        {asks && (
          <label>
            {t('ribbonRulesPrompt')}
            <input
              name="prompt"
              value={prompt}
              autoFocus={!named}
              onChange={(e) => setPrompt(e.target.value)}
            />
          </label>
        )}
        {conditional && (
          <>
            <label>
              {t('ribbonField')}
              <select name="field" value={field} onChange={(e) => setField(e.target.value)}>
                {fields.map((name) => (
                  <option key={name} value={name}>
                    {name}
                  </option>
                ))}
              </select>
            </label>
            <label>
              {t('ribbonFilterComparison')}
              <select name="op" value={op} onChange={(e) => setOp(e.target.value as RuleOp)}>
                {RULE_OPS.map((code) => (
                  <option key={code} value={code}>
                    {t(RULE_OP_LABELS[code])}
                  </option>
                ))}
              </select>
            </label>
          </>
        )}
        <label>
          {t(
            conditional
              ? 'ribbonFilterCompareTo'
              : asks
                ? 'ribbonRulesDefaultText'
                : 'ribbonPropValue',
          )}
          <input name="value" value={value} onChange={(e) => setValue(e.target.value)} />
        </label>
        {asks && (
          <label className="tco-check">
            <input
              type="checkbox"
              name="once"
              checked={once}
              onChange={(e) => setOnce(e.target.checked)}
            />
            {t('ribbonRulesAskOnce')}
          </label>
        )}
        {rule === 'if' && (
          <>
            <label>
              {t('ribbonRulesThen')}
              <input name="then" value={then} onChange={(e) => setThen(e.target.value)} />
            </label>
            <label>
              {t('ribbonRulesElse')}
              <input
                name="otherwise"
                value={otherwise}
                onChange={(e) => setOtherwise(e.target.value)}
              />
            </label>
          </>
        )}
        <div className="modal-actions">
          <span />
          <div className="tco-actions-right">
            <button type="button" onClick={onClose}>
              {t('ribbonCancel')}
            </button>
            <button type="submit" className="primary" disabled={!ready}>
              {t('ribbonOk')}
            </button>
          </div>
        </div>
      </form>
    </Backdrop>
  )
}
