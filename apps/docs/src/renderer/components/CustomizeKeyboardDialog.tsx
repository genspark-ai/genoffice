/**
 * Tools ▸ Customize Keyboard…, laid out as Word's: pick a command, press a
 * chord, Assign. Assignments are kept per machine (there is no Normal
 * template here), which is all "Save changes in" offers.
 */
import { useEffect, useState } from 'react'
import type { KeyboardEvent as ReactKeyboardEvent } from 'react'
import type { StyleInfo } from '@genoffice/docx-engine'
import { NONE } from '../../shared/commands'
import { bindingId, isAssignableChord } from '../../shared/keybindings'
import { HIGHLIGHT_CSS } from '../editor/marks'
import { fontFamiliesFor, systemFamiliesBesidesCandidates } from '../font-list'
import { fontSizeOptions } from '../font-sizes'
import { useI18n } from '../i18n/locale'
import { styleLabel } from '../style-gallery'
import { useSystemFontFamilies } from '../system-fonts'
import { saveKeyOverrides, useKeyOverrides } from '../key-layer'
import {
  IS_MAC,
  assignChord,
  chordFromEvent,
  chordOwner,
  effectiveChords,
  formatChord,
  listKeyCommands,
  removeChord,
  valueBindingLabel,
  VALUE_COMMANDS,
  type KeyCommand,
  type MenuCatalogEntry,
} from '../keybindings'
import { useModalKeys } from './modal-keys'
import { STANDARD_COLORS, THEME_COLORS } from './ribbon-color-palette'

const ALL = ''
// no menu label holds a control character, so these never collide with one
const FONTS = '\u0001fonts'
const STYLES = '\u0001styles'

interface ValueOption {
  payload: string
  label: string
}

export function CustomizeKeyboardDialog({
  styles,
  onClose,
}: {
  /** the styles of the open document, by style id */
  styles: ReadonlyMap<string, StyleInfo>
  onClose: () => void
}) {
  const { t, lang } = useI18n()
  const modalKeys = useModalKeys(onClose)
  const overrides = useKeyOverrides()
  const [catalog, setCatalog] = useState<MenuCatalogEntry[]>([])
  const [category, setCategory] = useState(ALL)
  const [query, setQuery] = useState('')
  const [selected, setSelected] = useState<string | null>(null)
  const [value, setValue] = useState<string | null>(null)
  const { families: systemFonts, load: loadSystemFonts } = useSystemFontFamilies()
  useEffect(loadSystemFonts, [loadSystemFonts])
  const [currentKey, setCurrentKey] = useState<string | null>(null)
  const [pressed, setPressed] = useState<string | null>(null)
  const [confirmReset, setConfirmReset] = useState(false)

  useEffect(() => {
    let live = true
    void window.desktop
      ?.keyCatalog?.()
      .then((entries) => live && setCatalog(entries))
      .catch(() => undefined)
    return () => {
      live = false
    }
  }, [])

  const commands = listKeyCommands(catalog, t)
  const categories = [...new Set(commands.flatMap((item) => item.menu ?? []))]
  const valued = (id: string, label: string): KeyCommand => ({ id, label, menu: null, path: '' })
  const builtInFonts = fontFamiliesFor(lang)
  const listed =
    category === FONTS
      ? [...builtInFonts, ...systemFamiliesBesidesCandidates(builtInFonts, systemFonts)].map(
          (name) => valued(bindingId('font-family', name), name),
        )
      : category === STYLES
        ? [...styles.values()]
            .filter((style) => style.type !== 'table' && !style.semiHidden)
            .map((style) => valued(bindingId('apply-style', style.styleId), styleLabel(style, t)))
        : [
            ...commands.filter((command) => category === ALL || command.menu === category),
            ...(category === ALL
              ? VALUE_COMMANDS.map((entry) => valued(entry.cmd, `${t(entry.labelKey)}:`))
              : []),
          ]
  const needle = query.trim().toLowerCase()
  const shown = listed
    .filter((command) => !needle || command.label.toLowerCase().includes(needle))
    .sort((a, b) => a.label.localeCompare(b.label))
  const picked = shown.find((c) => c.id === selected) ?? null
  const valueOptions: ValueOption[] | null =
    picked?.id === 'font-size'
      ? fontSizeOptions(lang).map((size) => ({ payload: String(size.pt), label: size.name }))
      : picked?.id === 'font-color'
        ? [
            { payload: NONE, label: t('ribbonAutomatic') },
            ...[...THEME_COLORS, ...STANDARD_COLORS].map((color) => ({
              payload: color.hex,
              label: `${t(color.nameKey)} (#${color.hex})`,
            })),
          ]
        : picked?.id === 'highlight'
          ? [
              { payload: NONE, label: t('ribbonNoColor') },
              ...Object.keys(HIGHLIGHT_CSS).map((name) => ({ payload: name, label: name })),
            ]
          : null
  const pickedValue = valueOptions?.find((option) => option.payload === value) ?? valueOptions?.[0]
  /** the binding the keys below belong to: the command, or the command with the value picked for it */
  const command =
    picked && valueOptions
      ? pickedValue
        ? { ...picked, id: bindingId(picked.id, pickedValue.payload) }
        : null
      : picked
  const keys = command ? effectiveChords(overrides, command.id) : []
  const nameOf = (id: string) =>
    commands.find((c) => c.id === id)?.label ??
    valueBindingLabel(id, t, (styleId) => {
      const style = styles.get(styleId)
      return style && styleLabel(style, t)
    }) ??
    id

  const owner = pressed ? chordOwner(overrides, pressed) : null
  const assignable = !!pressed && isAssignableChord(pressed) && owner?.kind !== 'fixed'
  const pressedNote = !pressed
    ? ''
    : !isAssignableChord(pressed)
      ? t('appCkNeedsModifier', { key: IS_MAC ? '⌘' : 'Ctrl' })
      : !owner
        ? t('appCkUnassigned')
        : owner.kind === 'binding'
          ? t('appCkAssignedTo', { name: nameOf(owner.id) })
          : owner.row
            ? t('appCkFixed', { name: t(owner.row.labelKey) + (owner.row.labelSuffix ?? '') })
            : t('appCkReserved')

  const pick = (id: string) => {
    setSelected(id)
    setValue(null)
    setCurrentKey(null)
  }
  const onShortcutKey = (event: ReactKeyboardEvent) => {
    // Escape closes and Tab moves on, as in any dialog
    if (event.key === 'Escape' || event.key === 'Tab') return
    event.preventDefault()
    event.stopPropagation()
    if (event.key === 'Backspace' && !event.ctrlKey && !event.metaKey && !event.altKey) {
      setPressed(null)
      return
    }
    const chord = chordFromEvent(event.nativeEvent)
    if (chord) setPressed(chord)
  }
  const assign = () => {
    if (!command || !pressed) return
    const next = assignChord(overrides, command.id, pressed)
    if (!next) return
    saveKeyOverrides(next)
    setPressed(null)
  }
  const remove = () => {
    if (!command || !currentKey) return
    saveKeyOverrides(removeChord(overrides, command.id, currentKey))
    setCurrentKey(null)
  }

  return (
    <div
      className="modal-backdrop"
      ref={modalKeys.ref}
      onKeyDown={modalKeys.onKeyDown}
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
    >
      <div className="modal modal-customize-keyboard" role="dialog" aria-label={t('appCkTitle')}>
        <h2>{t('appCkTitle')}</h2>
        <div className="ck-lists">
          <div className="ck-column ck-categories">
            <span id="ck-categories-label">{t('appCkCategories')}</span>
            <div className="ck-list" role="listbox" aria-labelledby="ck-categories-label">
              {[...categories, ALL, FONTS, STYLES].map((name) => (
                <button
                  key={name}
                  type="button"
                  role="option"
                  aria-selected={category === name}
                  className={
                    (category === name ? 'is-selected' : '') + (name === FONTS ? ' ck-sep' : '')
                  }
                  data-category={name === FONTS ? 'fonts' : name === STYLES ? 'styles' : undefined}
                  onClick={() => setCategory(name)}
                >
                  {name === ALL
                    ? t('appCkAllCommands')
                    : name === FONTS
                      ? t('ribbonFonts')
                      : name === STYLES
                        ? t('ribbonGroupStyles')
                        : name}
                </button>
              ))}
            </div>
          </div>
          <div className="ck-column ck-commands">
            <span id="ck-commands-label">{t('appCkCommands')}</span>
            <input
              type="search"
              className="ck-search"
              placeholder={t('appCkSearch')}
              aria-label={t('appCkSearch')}
              value={query}
              onChange={(e) => setQuery(e.target.value)}
            />
            <div className="ck-list" role="listbox" aria-labelledby="ck-commands-label">
              {shown.map((item) => (
                <button
                  key={item.id}
                  type="button"
                  role="option"
                  aria-selected={item.id === picked?.id}
                  className={item.id === picked?.id ? 'is-selected' : ''}
                  data-command={item.id}
                  onClick={() => pick(item.id)}
                >
                  <span>{item.label}</span>
                  <kbd>
                    {effectiveChords(overrides, item.id)
                      .map((c) => formatChord(c))
                      .join(' / ')}
                  </kbd>
                </button>
              ))}
            </div>
            {picked && valueOptions && (
              <select
                className="ck-value"
                aria-label={picked.label}
                value={pickedValue?.payload ?? ''}
                onChange={(e) => {
                  setValue(e.target.value)
                  setCurrentKey(null)
                }}
              >
                {valueOptions.map((option) => (
                  <option key={option.payload} value={option.payload}>
                    {option.label}
                  </option>
                ))}
              </select>
            )}
          </div>
        </div>
        <div className="ck-keys">
          <div className="ck-column">
            <span id="ck-current-label">{t('appCkCurrentKeys')}</span>
            <div className="ck-list ck-current" role="listbox" aria-labelledby="ck-current-label">
              {keys.map((chord) => (
                <button
                  key={chord}
                  type="button"
                  role="option"
                  aria-selected={chord === currentKey}
                  className={chord === currentKey ? 'is-selected' : ''}
                  onClick={() => setCurrentKey(chord)}
                >
                  {formatChord(chord)}
                </button>
              ))}
            </div>
            <button
              type="button"
              className="ck-btn"
              disabled={!command || !currentKey}
              onClick={remove}
            >
              {t('ribbonFitRemove')}
            </button>
          </div>
          <div className="ck-column">
            <label htmlFor="ck-press">{t('appCkPressNew')}</label>
            <input
              id="ck-press"
              className="ck-press"
              readOnly
              value={pressed ? formatChord(pressed) : ''}
              onKeyDown={onShortcutKey}
            />
            <p className="ck-note" role="status">
              {pressedNote}
            </p>
            <button
              type="button"
              className="ck-btn"
              disabled={!command || !assignable}
              onClick={assign}
            >
              {t('appCkAssign')}
            </button>
          </div>
        </div>
        <div className="ck-description">
          <span>{t('appCkDescription')}</span>
          <p>{command ? command.path || t('appCkNoMenu') : ''}</p>
        </div>
        <label className="ck-save-in">
          {t('appCkSaveIn')}
          <select aria-label={t('appCkSaveIn')} value="machine" onChange={() => undefined}>
            <option value="machine">{t('appCkThisComputer')}</option>
          </select>
        </label>
        {confirmReset ? (
          <div className="ck-confirm" role="alertdialog" aria-label={t('appCkResetAll')}>
            <p>{t('appCkResetConfirm')}</p>
            <div className="modal-actions">
              <button type="button" onClick={() => setConfirmReset(false)}>
                {t('appCancel')}
              </button>
              <button
                type="button"
                className="btn-primary"
                onClick={() => {
                  saveKeyOverrides({})
                  setCurrentKey(null)
                  setConfirmReset(false)
                }}
              >
                {t('appOk')}
              </button>
            </div>
          </div>
        ) : (
          <div className="modal-actions">
            <button
              type="button"
              className="ck-reset"
              disabled={Object.keys(overrides).length === 0}
              onClick={() => setConfirmReset(true)}
            >
              {t('appCkResetAll')}
            </button>
            <button type="button" className="btn-primary" onClick={onClose}>
              {t('appOk')}
            </button>
          </div>
        )}
      </div>
    </div>
  )
}
