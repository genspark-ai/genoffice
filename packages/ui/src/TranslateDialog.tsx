/**
 * Modal dialog for a one-shot translate action.
 *
 * Two-column preview (source vs. translated) lets the user verify the
 * result before applying. Defaults are derived from the host app:
 *   - `defaultTargetLang` typically the current UI language
 *   - `defaultSourceLang` "auto" lets the model detect
 *
 * The dialog does NOT call the model itself; it produces a `ChatChangePlan`
 * the host app applies via `useChatRuntime().applyChangePlan(plan)`. This
 * keeps the dialog framework-free and lets the host pick where the run
 * happens (side panel for long documents, ephemeral channel for one-shots).
 */

import React, { useEffect, useMemo, useState } from 'react'
import type { ChatChangePlan } from '@genoffice/chat-runtime/types'

export interface TranslateLanguageOption {
  /** BCP-47 tag (e.g. "zh-CN") or a free-form label used for the dropdown */
  value: string
  /** Display label (e.g. "简体中文") */
  label: string
}

export interface TranslateDialogStrings {
  title: string
  targetLang: string
  sourceLang: string
  preserveFormat: string
  swapLanguages: string
  start: string
  original: string
  translated: string
  previewTitle: string
  previewLoading: string
  cancel: string
  unsupported: string
}

export interface TranslateDialogProps {
  open: boolean
  sourceText: string
  /**
   * Editor range captured at the moment the user picked the chip. The dialog
   * stamps it on the resulting `ChatChangePlan` so the host can re-target
   * the selection when the run finishes (apply/undo both need the original
   * range because the live selection may have wandered).
   */
  sourceRange?: { from: number; to: number; scope?: string } | null
  defaultSourceLang?: string | undefined
  defaultTargetLang: string
  /** Full target-language list the host wants to expose (e.g. 12 common). */
  languages: TranslateLanguageOption[]
  strings: TranslateDialogStrings
  /**
   * Render the translation. The host wires this to either:
   *   - a sync provider call returning the translated string, or
   *   - an async call that yields when the run completes.
   * Returning `null` aborts with an error chip.
   */
  onTranslate: (args: {
    sourceText: string
    sourceLang: string
    targetLang: string
    preserveFormat: boolean
  }) => Promise<string | null>
  /** Apply handler — receives a fully-formed `translate` change plan. */
  onApply: (plan: ChatChangePlan, targetText: string) => void
  /** Optional cancel handler (Esc / X button). */
  onCancel: () => void
  /** Optional app id stamped on the change plan (defaults to 'unknown'). */
  app?: 'docs' | 'sheets' | 'slides' | 'pdf' | 'unknown' | undefined
}

let counter = 0
function nextPlanId(): string {
  counter += 1
  return `translate-${Date.now().toString(36)}-${counter.toString(36)}`
}

export function TranslateDialog(props: TranslateDialogProps): React.JSX.Element | null {
  const { open, sourceText, sourceRange, defaultSourceLang, defaultTargetLang, languages, strings, onTranslate, onApply, onCancel, app } = props
  const [sourceLang, setSourceLang] = useState<string>(defaultSourceLang ?? 'auto')
  const [targetLang, setTargetLang] = useState<string>(defaultTargetLang)
  const [preserveFormat, setPreserveFormat] = useState<boolean>(true)
  const [translated, setTranslated] = useState<string | null>(null)
  const [busy, setBusy] = useState<boolean>(false)
  const [error, setError] = useState<string | null>(null)

  useEffect(() => {
    if (open) {
      setSourceLang(defaultSourceLang ?? 'auto')
      setTargetLang(defaultTargetLang)
      setPreserveFormat(true)
      setTranslated(null)
      setBusy(false)
      setError(null)
    }
  }, [open, defaultSourceLang, defaultTargetLang])

  const targetLabel = useMemo(() => {
    const m = languages.find((l) => l.value === targetLang)
    return m?.label ?? targetLang
  }, [languages, targetLang])

  if (!open) return null

  const handleSwap = () => {
    setSourceLang(targetLang)
    setTargetLang(sourceLang === 'auto' ? defaultTargetLang : sourceLang)
  }

  const handleTranslate = async () => {
    if (busy) return
    setBusy(true)
    setError(null)
    setTranslated(null)
    try {
      const result = await onTranslate({ sourceText, sourceLang, targetLang, preserveFormat })
      if (result === null) {
        // Host chose not to throw — fall back to a generic message so the
        // user still sees *something*; real provider errors come through
        // the catch branch when the host re-throws.
        setError(strings.unsupported)
      } else {
        setTranslated(result)
      }
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e))
    } finally {
      setBusy(false)
    }
  }

  const handleApply = () => {
    if (translated === null) return
    const plan: ChatChangePlan = {
      id: nextPlanId(),
      app: app ?? 'unknown',
      title: `${strings.title} → ${targetLabel}`,
      summary: `${strings.translated}: ${translated.slice(0, 80)}${translated.length > 80 ? '…' : ''}`,
      ops: [
        {
          kind: 'translate',
          ops: [
            {
              sourceText,
              targetText: translated,
              targetLang,
              preserveFormat,
              range: sourceRange ?? null,
            },
          ],
          description: `Translate ${sourceLang} → ${targetLang}`,
        },
      ],
      warnings: preserveFormat ? [] : ['formatNotPreserved'],
      requireConfirm: false,
      createdAt: Date.now(),
    }
    onApply(plan, translated)
  }

  return (
    <div className="ai-translate-dialog-backdrop" role="dialog" aria-modal="true" aria-label={strings.title}>
      <div className="ai-translate-dialog">
        <header className="ai-translate-dialog-header">
          <h2 className="ai-translate-dialog-title">{strings.title}</h2>
          <button type="button" className="ai-translate-dialog-close" aria-label="Close" onClick={onCancel}>
            ×
          </button>
        </header>

        <div className="ai-translate-dialog-langs">
          <label className="ai-translate-dialog-field">
            <span>{strings.sourceLang}</span>
            <select value={sourceLang} onChange={(e) => setSourceLang(e.target.value)} disabled={busy}>
              <option value="auto">Auto-detect</option>
              {languages.map((l) => (
                <option key={`src-${l.value}`} value={l.value}>
                  {l.label}
                </option>
              ))}
            </select>
          </label>
          <button
            type="button"
            className="ai-translate-dialog-swap"
            onClick={handleSwap}
            disabled={busy}
            aria-label={strings.swapLanguages}
            title={strings.swapLanguages}
          >
            ⇄
          </button>
          <label className="ai-translate-dialog-field">
            <span>{strings.targetLang}</span>
            <select value={targetLang} onChange={(e) => setTargetLang(e.target.value)} disabled={busy}>
              {languages.map((l) => (
                <option key={`dst-${l.value}`} value={l.value}>
                  {l.label}
                </option>
              ))}
            </select>
          </label>
        </div>

        <label className="ai-translate-dialog-toggle">
          <input
            type="checkbox"
            checked={preserveFormat}
            onChange={(e) => setPreserveFormat(e.target.checked)}
            disabled={busy}
          />
          <span>{strings.preserveFormat}</span>
        </label>

        <div className="ai-translate-dialog-preview" data-busy={busy}>
          <section className="ai-translate-dialog-side">
            <h3>{strings.original}</h3>
            <div className="ai-translate-dialog-text">{sourceText}</div>
          </section>
          <section className="ai-translate-dialog-side">
            <h3>{strings.translated}</h3>
            <div className="ai-translate-dialog-text">
              {error ? (
                <span className="ai-translate-dialog-error">{error}</span>
              ) : translated === null ? (
                <span className="ai-translate-dialog-placeholder">{busy ? strings.previewLoading : strings.previewTitle}</span>
              ) : (
                translated
              )}
            </div>
          </section>
        </div>

        <footer className="ai-translate-dialog-footer">
          <button type="button" className="ai-translate-dialog-btn" onClick={onCancel} disabled={busy}>
            {strings.cancel}
          </button>
          {translated === null ? (
            <button
              type="button"
              className="ai-translate-dialog-btn ai-translate-dialog-btn--primary"
              onClick={handleTranslate}
              disabled={busy || sourceText.trim().length === 0}
            >
              {strings.start}
            </button>
          ) : (
            <button
              type="button"
              className="ai-translate-dialog-btn ai-translate-dialog-btn--primary"
              onClick={handleApply}
              disabled={busy}
            >
              {strings.start}
            </button>
          )}
        </footer>
      </div>
    </div>
  )
}
