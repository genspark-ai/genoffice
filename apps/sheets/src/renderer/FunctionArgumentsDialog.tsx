import { useEffect, useMemo, useRef, useState } from 'react'

import {
  buildFunctionCall,
  formatPreviewValue,
  parameterRows,
  type PreviewValue,
} from './function-arguments'
import {
  formatExternalRef,
  type ExternalPickTarget,
  type FunctionArgumentsHost,
  type FunctionTarget,
} from './function-arguments-runtime'
import type { FunctionSpec } from './function-catalog'
import { useI18n } from './i18n/locale'
import { useModalDialog } from './modal-dialog'

/// Excel's Function Arguments: one input per parameter, each previewing its
/// value, the whole call's result underneath. The card floats without a
/// backdrop so a click or drag on the grid fills the active argument.

const PREVIEW_DEBOUNCE_MS = 150

export function FunctionArgumentsDialog({
  spec,
  target,
  initialArgs,
  host,
  error,
  onOk,
  onCancel,
}: {
  readonly spec: FunctionSpec
  readonly target: FunctionTarget
  readonly initialArgs: readonly string[]
  readonly host: Pick<
    FunctionArgumentsHost,
    'evaluate' | 'observeRangePick' | 'listExternalPickTargets' | 'pickExternalRange'
  >
  readonly error: string | null
  readonly onOk: (args: readonly string[]) => void
  readonly onCancel: () => void
}): React.JSX.Element {
  const { t } = useI18n()
  const [args, setArgs] = useState<string[]>([...initialArgs])
  const [active, setActive] = useState(0)
  const [collapsed, setCollapsed] = useState(false)
  const [previews, setPreviews] = useState<ReadonlyMap<string, string>>(new Map())
  /// Linked workbooks of the edited file: open ones take a grid pick through
  /// another tab, closed ones a typed range (#2109).
  const [pickTargets, setPickTargets] = useState<readonly ExternalPickTarget[]>([])
  const [pickingName, setPickingName] = useState<string | null>(null)
  const [typedSheets, setTypedSheets] = useState<ReadonlyMap<number, string>>(new Map())
  const [typedRange, setTypedRange] = useState('')
  const inputRefs = useRef<(HTMLInputElement | null)[]>([])
  const activeRef = useRef(0)
  const disposePickRef = useRef<(() => void) | null>(null)
  activeRef.current = active

  const rows = useMemo(() => parameterRows(spec.params, args.length), [spec, args.length])
  const call = buildFunctionCall(spec.name, args)
  const expressions = [...new Set([...args.filter((arg) => arg.trim() !== ''), call])]
  const expressionsKey = expressions.join('\u0000')

  useEffect(() => {
    const expressions = expressionsKey.split('\u0000')
    const timer = setTimeout(() => {
      void host.evaluate(target, expressions).then(
        (values: PreviewValue[]) => {
          setPreviews((previous) => {
            const next = new Map(previous)
            expressions.forEach((expr, index) => next.set(expr, formatPreviewValue(values[index])))
            return next
          })
        },
        () => undefined,
      )
    }, PREVIEW_DEBOUNCE_MS)
    return () => clearTimeout(timer)
  }, [expressionsKey, host, target])

  const insertAtActive = (ref: string): void => {
    const index = activeRef.current
    setArgs((previous) => {
      const next = [...previous]
      while (next.length <= index) next.push('')
      next[index] = ref
      return next
    })
    inputRefs.current[index]?.focus()
  }

  useEffect(
    () =>
      host.observeRangePick(target, (ref) => {
        insertAtActive(ref)
      }),
    [host, target],
  )

  // the links list is read once per dialog open; a stale list only means a
  // book opened mid-edit shows up on the next open
  useEffect(() => {
    let disposed = false
    void host.listExternalPickTargets().then(
      (targets) => {
        if (!disposed) setPickTargets(targets)
      },
      () => undefined,
    )
    return () => {
      disposed = true
    }
  }, [host])

  useEffect(
    () => () => {
      disposePickRef.current?.()
      disposePickRef.current = null
    },
    [],
  )

  const armExternalPick = (target: ExternalPickTarget): void => {
    disposePickRef.current?.()
    setPickingName(target.name)
    disposePickRef.current = host.pickExternalRange(target, (ref) => {
      disposePickRef.current?.()
      disposePickRef.current = null
      setPickingName(null)
      insertAtActive(ref)
    })
  }

  const insertTypedRange = (target: ExternalPickTarget): void => {
    const sheet = typedSheets.get(target.index) ?? target.sheetNames[0]
    const range = typedRange.trim()
    if (!sheet || range === '') return
    insertAtActive(formatExternalRef(target.index, sheet, range))
  }

  const setArg = (index: number, value: string): void => {
    setArgs((previous) => {
      const next = [...previous]
      while (next.length <= index) next.push('')
      next[index] = value
      return next
    })
  }

  const argPreview = (index: number): string => {
    const expr = (args[index] ?? '').trim()
    return expr === '' ? '' : (previews.get(expr) ?? '')
  }
  const result = previews.get(call) ?? ''
  const activeRow = rows[Math.min(active, rows.length - 1)]
  const complete = rows.every((row) => !row.require || (args[row.index] ?? '').trim() !== '')

  const modal = useModalDialog(onCancel)
  return (
    <div className="dialog-floating">
      <div
        className={`format-cells-dialog function-args-dialog${collapsed ? ' collapsed' : ''}`}
        role="dialog"
        {...modal}
        aria-label={t('dlgFnArgsTitle')}
        onKeyDown={(event) => {
          if (event.key === 'Enter' && event.target instanceof HTMLInputElement) {
            event.preventDefault()
            if (collapsed) setCollapsed(false)
            else if (complete) onOk(args)
            return
          }
          modal.onKeyDown(event)
        }}
      >
        <header>{collapsed ? `${t('dlgFnArgsTitle')} — ${spec.name}` : t('dlgFnArgsTitle')}</header>
        {!collapsed && <p className="fn-args-name">{spec.name}</p>}
        <div className="fn-args-rows">
          {rows.map((row) =>
            collapsed && row.index !== active ? null : (
              <div className="fn-args-row" key={row.index}>
                <label
                  className={row.require ? 'required' : undefined}
                  htmlFor={`fn-arg-${row.index}`}
                >
                  {row.name}
                </label>
                <span className="input-row">
                  <input
                    id={`fn-arg-${row.index}`}
                    ref={(node) => {
                      inputRefs.current[row.index] = node
                    }}
                    autoFocus={row.index === 0}
                    value={args[row.index] ?? ''}
                    onChange={(event) => setArg(row.index, event.target.value)}
                    onFocus={() => setActive(row.index)}
                  />
                  <button
                    type="button"
                    className="secondary fn-args-pick"
                    aria-label={collapsed ? t('dlgFnArgsExpand') : t('dlgFnArgsPickRange')}
                    title={collapsed ? t('dlgFnArgsExpand') : t('dlgFnArgsPickRange')}
                    onClick={() => {
                      setActive(row.index)
                      setCollapsed((value) => !value)
                      inputRefs.current[row.index]?.focus()
                    }}
                  >
                    {collapsed ? '⤢' : '⊞'}
                  </button>
                </span>
                {!collapsed && (
                  <output>{argPreview(row.index) && `= ${argPreview(row.index)}`}</output>
                )}
              </div>
            ),
          )}
        </div>
        {!collapsed && pickTargets.length > 0 && (
          <div className="fn-args-external">
            <p className="dialog-note">{t('dlgFnArgsLinked')}</p>
            {pickTargets.map((pickTarget) => (
              <div className="fn-args-external-row" key={pickTarget.index}>
                <span className="watch-formula">{pickTarget.name}</span>
                {pickTarget.open ? (
                  pickingName === pickTarget.name ? (
                    <span className="dialog-note">{t('dlgFnArgsPicking')}</span>
                  ) : (
                    <button
                      type="button"
                      className="secondary"
                      onClick={() => armExternalPick(pickTarget)}
                    >
                      {t('dlgFnArgsPick')}
                    </button>
                  )
                ) : (
                  <span className="input-row">
                    <select
                      aria-label={t('appLinksSheets')}
                      value={typedSheets.get(pickTarget.index) ?? pickTarget.sheetNames[0] ?? ''}
                      onChange={(event) =>
                        setTypedSheets((previous) => {
                          const next = new Map(previous)
                          next.set(pickTarget.index, event.target.value)
                          return next
                        })
                      }
                    >
                      {pickTarget.sheetNames.map((name) => (
                        <option key={name} value={name}>
                          {name}
                        </option>
                      ))}
                    </select>
                    <input
                      placeholder="A1:B2"
                      value={typedRange}
                      onChange={(event) => setTypedRange(event.target.value)}
                      onKeyDown={(event) => {
                        if (event.key === 'Enter') {
                          event.preventDefault()
                          insertTypedRange(pickTarget)
                        }
                      }}
                    />
                    <button
                      type="button"
                      className="secondary"
                      onClick={() => insertTypedRange(pickTarget)}
                    >
                      {t('dlgFnArgsInsert')}
                    </button>
                  </span>
                )}
              </div>
            ))}
          </div>
        )}
        {collapsed ? (
          <p className="dialog-note">{t('dlgFnArgsPickHint')}</p>
        ) : (
          <>
            <p className="fn-args-call-result">{result && `= ${result}`}</p>
            <p className="dialog-note">{spec.description}</p>
            {activeRow && (
              <p className="dialog-note fn-args-detail">
                <strong>{activeRow.name}:</strong> {activeRow.detail || t('dlgFnArgsPickHint')}
              </p>
            )}
            <p className="fn-args-result">
              {t('dlgFnArgsResult')} <strong>{result}</strong>
            </p>
            {error && (
              <p className="dialog-note" role="alert">
                {error}
              </p>
            )}
            <div className="dialog-actions">
              <button className="secondary" onClick={onCancel}>
                {t('dlgCancel')}
              </button>
              <button className="primary-action" disabled={!complete} onClick={() => onOk(args)}>
                {t('dlgOk')}
              </button>
            </div>
          </>
        )}
      </div>
    </div>
  )
}
