import { useState } from 'react'

import { useI18n, type StringKey } from './i18n/locale'

/// Excel's Insert → Symbol, minimal: a tabbed grid of common special
/// characters. Clicking a symbol emits it to App (which appends it to the
/// active cell) and keeps the dialog open, so several symbols can be
/// inserted in a row — Close dismisses it.

/// Appends a picked symbol to whatever the cell already holds. Numbers and
/// booleans coerce to text: Excel's Symbol dialog also turns the cell into
/// text when characters are appended.
export function appendSymbol(
  existing: string | number | boolean | null | undefined,
  char: string,
): string {
  if (existing === null || existing === undefined || existing === '') return char
  return `${String(existing)}${char}`
}

const SYMBOL_CATEGORIES: readonly {
  readonly labelKey: StringKey
  readonly symbols: readonly string[]
}[] = [
  {
    labelKey: 'dlgSymbolCatCurrency',
    symbols: [
      '€',
      '£',
      '¥',
      '¢',
      '₩',
      '₽',
      '₹',
      '©',
      '®',
      '™',
      '§',
      '¶',
      '†',
      '‡',
      '•',
      '…',
      '°',
      '‰',
      '№',
      '¿',
    ],
  },
  {
    labelKey: 'dlgSymbolCatMath',
    symbols: [
      '±',
      '×',
      '÷',
      '≠',
      '≈',
      '≤',
      '≥',
      '∞',
      '√',
      '∑',
      '∏',
      '∫',
      '∂',
      '∆',
      'π',
      'µ',
      '∈',
      '∉',
      '⊂',
      '⊃',
      '∩',
      '∪',
      '∀',
      '∃',
      '∅',
      '∝',
      '∴',
      '∵',
      '⊥',
      '∥',
    ],
  },
  {
    labelKey: 'dlgSymbolCatArrows',
    symbols: [
      '←',
      '→',
      '↑',
      '↓',
      '↔',
      '↕',
      '⇐',
      '⇒',
      '⇑',
      '⇓',
      '⇔',
      '↗',
      '↘',
      '↙',
      '↖',
      '↩',
      '↪',
      '⇄',
      '⇅',
      '↦',
    ],
  },
  {
    labelKey: 'dlgSymbolCatGreek',
    symbols: [
      'α',
      'β',
      'γ',
      'δ',
      'ε',
      'ζ',
      'η',
      'θ',
      'λ',
      'μ',
      'π',
      'ρ',
      'σ',
      'τ',
      'φ',
      'χ',
      'ψ',
      'ω',
      'Α',
      'Β',
      'Γ',
      'Δ',
      'Θ',
      'Λ',
      'Σ',
      'Φ',
      'Ω',
    ],
  },
  {
    labelKey: 'dlgSymbolCatMisc',
    symbols: [
      '★',
      '☆',
      '●',
      '○',
      '◆',
      '◇',
      '■',
      '□',
      '▲',
      '△',
      '▼',
      '▽',
      '✓',
      '✗',
      '☑',
      '☐',
      '♠',
      '♣',
      '♥',
      '♦',
      '♪',
      '♫',
      '☀',
      '☁',
      '☂',
      '☎',
    ],
  },
]

export function SymbolDialog({
  onInsert,
  onClose,
}: {
  readonly onInsert: (char: string) => void
  readonly onClose: () => void
}): React.JSX.Element {
  const { t } = useI18n()
  const [categoryIndex, setCategoryIndex] = useState(0)
  const active = SYMBOL_CATEGORIES[categoryIndex]

  return (
    <div className="dialog-backdrop" onClick={onClose}>
      <div
        className="format-cells-dialog"
        role="dialog"
        aria-label={t('dlgSymbolTitle')}
        onClick={(event) => event.stopPropagation()}
      >
        <header>{t('dlgSymbolTitle')}</header>
        <nav className="dialog-tabs">
          {SYMBOL_CATEGORIES.map((candidate, index) => (
            <button
              key={candidate.labelKey}
              className={index === categoryIndex ? 'active' : ''}
              onClick={() => setCategoryIndex(index)}
            >
              {t(candidate.labelKey)}
            </button>
          ))}
        </nav>
        <section className="dialog-body">
          <div className="symbol-grid">
            {(active?.symbols ?? []).map((symbol) => (
              <button
                key={symbol}
                type="button"
                data-tip={t('dlgSymbolInsertHint', { symbol })}
                onClick={() => onInsert(symbol)}
              >
                {symbol}
              </button>
            ))}
          </div>
          <p className="dialog-note">{t('dlgSymbolNote')}</p>
        </section>
        <div className="dialog-actions">
          <button className="secondary" onClick={onClose}>
            {t('dlgClose')}
          </button>
        </div>
      </div>
    </div>
  )
}
