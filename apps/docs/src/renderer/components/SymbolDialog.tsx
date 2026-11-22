import { useMemo, useState } from 'react'
import { SYMBOL_FONTS, fontSymbolChars, fontSymbolDisplay } from '../editor/font-symbol'
import { useI18n, type StringKey } from '../i18n/locale'
import { useModalKeys } from './modal-keys'

const RANGES: Array<[number, number]> = [
  [0x00a1, 0x00ff],
  [0x0391, 0x03a9],
  [0x03b1, 0x03c9],
  [0x2010, 0x2027],
  [0x2030, 0x203e],
  [0x20a0, 0x20bf],
  [0x2150, 0x215e],
  [0x2190, 0x21d4],
  [0x2200, 0x226f],
  [0x25a0, 0x25cf],
  [0x2600, 0x2606],
  [0x2660, 0x266f],
]
// the soft hyphen draws nothing; U+03A2 is unassigned
const SKIPPED = new Set([0x00ad, 0x03a2])

export const SYMBOL_CHARS: string[] = RANGES.flatMap(([from, to]) =>
  Array.from({ length: to - from + 1 }, (_, i) => from + i)
    .filter((code) => !SKIPPED.has(code))
    .map((code) => String.fromCodePoint(code)),
)

export const symbolCode = (symbol: string): string =>
  (symbol.codePointAt(0) ?? 0).toString(16).toUpperCase().padStart(4, '0')

/** the character a typed hexadecimal code stands for; null for anything that is not a printable code point */
export function symbolFromCode(hex: string): string | null {
  const text = hex.trim().replace(/^u\+/i, '')
  if (!/^[0-9a-f]{2,6}$/i.test(text)) return null
  const code = parseInt(text, 16)
  const control = code < 0x20 || (code >= 0x7f && code <= 0x9f)
  const surrogate = code >= 0xd800 && code <= 0xdfff
  if (control || surrogate || code > 0x10ffff) return null
  return String.fromCodePoint(code)
}

/** Word's Special Characters list: the named dashes, spaces and marks */
export const SPECIAL_CHARS: Array<[StringKey, string]> = [
  ['ribbonSpEmDash', '\u2014'],
  ['ribbonSpEnDash', '\u2013'],
  ['ribbonSpNbHyphen', '\u2011'],
  ['ribbonSpOptHyphen', '\u00ad'],
  ['ribbonSpEmSpace', '\u2003'],
  ['ribbonSpEnSpace', '\u2002'],
  ['ribbonSpQuarterEmSpace', '\u2005'],
  ['ribbonSpNbSpace', '\u00a0'],
  ['ribbonSpCopyright', '\u00a9'],
  ['ribbonSpRegistered', '\u00ae'],
  ['ribbonSpTrademark', '\u2122'],
  ['ribbonSpSection', '\u00a7'],
  ['ribbonSpParagraph', '\u00b6'],
  ['ribbonSpEllipsis', '\u2026'],
  ['ribbonSpSingleOpenQuote', '\u2018'],
  ['ribbonSpSingleCloseQuote', '\u2019'],
  ['ribbonSpDoubleOpenQuote', '\u201c'],
  ['ribbonSpDoubleCloseQuote', '\u201d'],
]

/** Word's Insert > Advanced Symbol: pick a character, Insert leaves the dialog open for the next one */
export function SymbolDialog({
  onInsert,
  onClose,
}: {
  /** `font` names the symbol font `symbol` is a w:sym character of; absent for plain text */
  onInsert: (symbol: string, font?: string) => void
  onClose: () => void
}) {
  const { t } = useI18n()
  const keys = useModalKeys(onClose)
  const [symbol, setSymbol] = useState(SYMBOL_CHARS[0])
  const [code, setCode] = useState(symbolCode(SYMBOL_CHARS[0]))
  const [tab, setTab] = useState<'symbols' | 'special'>('symbols')
  const [special, setSpecial] = useState(SPECIAL_CHARS[0][1])
  const [font, setFont] = useState('')
  const fontChars = useMemo(() => (font ? fontSymbolChars(font) : []), [font])
  const [fontChar, setFontChar] = useState('')
  const pick = (next: string) => {
    setSymbol(next)
    setCode(symbolCode(next))
  }
  const pickFont = (next: string) => {
    setFont(next)
    const first = next ? (fontSymbolChars(next)[0] ?? '') : ''
    setFontChar(first)
    setCode(next ? first : symbolCode(symbol))
  }
  const pickFontChar = (next: string) => {
    setFontChar(next)
    setCode(next)
  }
  const insertPicked = () => {
    if (!font) onInsert(symbol)
    else if (fontChar) onInsert(fontChar, font)
  }
  const fontCell = (char: string) => {
    const shown = fontSymbolDisplay(font, char)
    return (
      shown && (
        <span style={shown.ownFont ? { fontFamily: `"${font}"` } : undefined}>{shown.text}</span>
      )
    )
  }
  return (
    <div
      className="modal-backdrop"
      ref={keys.ref}
      onKeyDown={keys.onKeyDown}
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
    >
      <form
        className="modal gs-form tco-dialog symbol-dialog"
        noValidate
        onSubmit={(e) => {
          e.preventDefault()
          if (tab === 'special') onInsert(special)
          else insertPicked()
        }}
      >
        <h2>{t('ribbonBulletSymbol')}</h2>
        <div className="dlg-tabs" role="tablist">
          <button
            type="button"
            role="tab"
            aria-selected={tab === 'symbols'}
            className={tab === 'symbols' ? 'active' : ''}
            onClick={() => setTab('symbols')}
          >
            {t('ribbonSymTabSymbols')}
          </button>
          <button
            type="button"
            role="tab"
            aria-selected={tab === 'special'}
            className={tab === 'special' ? 'active' : ''}
            onClick={() => setTab('special')}
          >
            {t('ribbonSymTabSpecial')}
          </button>
        </div>
        {tab === 'special' && (
          <div className="symbol-special-list" role="listbox" aria-label={t('ribbonSymTabSpecial')}>
            {SPECIAL_CHARS.map(([name, char]) => (
              <button
                key={name}
                type="button"
                role="option"
                aria-selected={char === special}
                className={char === special ? 'active' : ''}
                onClick={() => setSpecial(char)}
                onDoubleClick={() => onInsert(char)}
              >
                <span className="symbol-special-char">{char}</span>
                {t(name)}
              </button>
            ))}
          </div>
        )}
        {tab === 'symbols' && (
          <>
            <div className="modal-row margin-row">
              <label>
                {t('ribbonBulletFont')}
                <select name="font" value={font} onChange={(e) => pickFont(e.target.value)}>
                  <option value="">{t('ribbonSymNormalText')}</option>
                  {SYMBOL_FONTS.map((name) => (
                    <option key={name} value={name}>
                      {name}
                    </option>
                  ))}
                </select>
              </label>
            </div>
            <div className="symbol-grid" role="listbox" aria-label={t('ribbonSymTabSymbols')}>
              {font
                ? fontChars.map((char) => (
                    <button
                      key={char}
                      type="button"
                      role="option"
                      aria-selected={char === fontChar}
                      data-sym-char={char}
                      className={char === fontChar ? 'symbol-cell active' : 'symbol-cell'}
                      onClick={() => pickFontChar(char)}
                      onDoubleClick={() => onInsert(char, font)}
                    >
                      {fontCell(char)}
                    </button>
                  ))
                : SYMBOL_CHARS.map((s) => (
                    <button
                      key={s}
                      type="button"
                      role="option"
                      aria-selected={s === symbol}
                      className={s === symbol ? 'symbol-cell active' : 'symbol-cell'}
                      onClick={() => pick(s)}
                      onDoubleClick={() => onInsert(s)}
                    >
                      {s}
                    </button>
                  ))}
            </div>
            <div className="symbol-dialog-picked">
              <span className="symbol-dialog-preview">{font ? fontCell(fontChar) : symbol}</span>
              <label>
                {t('ribbonSymCharCode')}
                <input
                  type="text"
                  value={code}
                  spellCheck={false}
                  onChange={(e) => {
                    setCode(e.target.value)
                    if (font) {
                      const typed = e.target.value.trim().toUpperCase()
                      if (fontChars.includes(typed)) setFontChar(typed)
                      return
                    }
                    const typed = symbolFromCode(e.target.value)
                    if (typed) setSymbol(typed)
                  }}
                />
              </label>
            </div>
          </>
        )}
        <div className="modal-actions">
          <span />
          <div className="tco-actions-right">
            <button type="button" onClick={onClose}>
              {t('ribbonClose')}
            </button>
            <button type="submit" className="primary">
              {t('ribbonInsert')}
            </button>
          </div>
        </div>
      </form>
    </div>
  )
}
