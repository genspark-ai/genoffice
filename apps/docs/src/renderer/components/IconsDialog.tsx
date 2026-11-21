import { useEffect, useMemo, useState } from 'react'
import { useI18n } from '../i18n/locale'
import {
  ICON_CATEGORIES,
  ICON_CATEGORY_LABEL,
  ICON_VIEW_BOX,
  loadIconLibrary,
  searchIcons,
  type IconCategory,
  type LibraryIcon,
} from '../icon-library'
import { useModalKeys } from './modal-keys'

/** Word's Insert > Icons: pick any number of icons by category or by search, Insert places them */
export function IconsDialog({
  onInsert,
  onClose,
}: {
  /** resolves once the icons are in; the dialog stays up, and inert, until then */
  onInsert: (names: string[]) => Promise<void>
  onClose: () => void
}) {
  const { t } = useI18n()
  const [icons, setIcons] = useState<LibraryIcon[] | null>(null)
  const [category, setCategory] = useState<IconCategory>(ICON_CATEGORIES[0])
  const [query, setQuery] = useState('')
  const [picked, setPicked] = useState<string[]>([])
  const [inserting, setInserting] = useState(false)
  // the dialog stays between the user and the caret until the icons are in
  const close = () => {
    if (!inserting) onClose()
  }
  const keys = useModalKeys(close)

  useEffect(() => {
    let open = true
    void loadIconLibrary().then((loaded) => {
      if (open) setIcons(loaded)
    })
    return () => {
      open = false
    }
  }, [])

  const searching = query.trim() !== ''
  const shown = useMemo(() => {
    if (!icons) return []
    return searching
      ? searchIcons(icons, query)
      : icons.filter((icon) => icon.category === category)
  }, [icons, searching, query, category])

  const toggle = (name: string) =>
    setPicked((now) => (now.includes(name) ? now.filter((n) => n !== name) : [...now, name]))

  return (
    <div
      className="modal-backdrop"
      ref={keys.ref}
      onKeyDown={keys.onKeyDown}
      onMouseDown={(e) => e.target === e.currentTarget && close()}
    >
      <form
        className="modal gs-form tco-dialog icons-dialog"
        noValidate
        onSubmit={(e) => {
          e.preventDefault()
          if (picked.length === 0 || inserting) return
          setInserting(true)
          void onInsert(picked).catch(() => setInserting(false))
        }}
      >
        <h2>{t('ribbonIcons')}</h2>
        <input
          type="search"
          className="icons-dialog-search"
          value={query}
          autoFocus
          spellCheck={false}
          placeholder={t('ribbonIconsSearch')}
          aria-label={t('ribbonIconsSearch')}
          onChange={(e) => setQuery(e.target.value)}
        />
        <div className="icons-dialog-body">
          <div className="icons-dialog-categories" role="tablist" aria-orientation="vertical">
            {ICON_CATEGORIES.map((id) => (
              <button
                key={id}
                type="button"
                role="tab"
                data-category={id}
                aria-selected={!searching && id === category}
                className={!searching && id === category ? 'active' : ''}
                onClick={() => {
                  setCategory(id)
                  setQuery('')
                }}
              >
                {t(ICON_CATEGORY_LABEL[id])}
              </button>
            ))}
          </div>
          <div
            className="icons-dialog-grid"
            role="listbox"
            aria-multiselectable="true"
            aria-label={t('ribbonIcons')}
          >
            {shown.map((icon) => {
              const on = picked.includes(icon.name)
              return (
                <button
                  key={icon.name}
                  type="button"
                  role="option"
                  data-icon={icon.name}
                  aria-selected={on}
                  aria-label={icon.label}
                  title={icon.label}
                  className={on ? 'icons-dialog-cell active' : 'icons-dialog-cell'}
                  onClick={() => toggle(icon.name)}
                >
                  <svg
                    viewBox={`0 0 ${ICON_VIEW_BOX} ${ICON_VIEW_BOX}`}
                    fill="currentColor"
                    aria-hidden="true"
                  >
                    <path d={icon.path} />
                  </svg>
                </button>
              )
            })}
            {icons && shown.length === 0 && (
              <div className="icons-dialog-empty">{t('ribbonIconsNone')}</div>
            )}
          </div>
        </div>
        <div className="modal-actions">
          <span className="icons-dialog-count">
            {picked.length > 0 ? t('ribbonIconsPicked', { n: picked.length }) : ''}
          </span>
          <div className="tco-actions-right">
            <button type="button" disabled={inserting} onClick={close}>
              {t('ribbonCancel')}
            </button>
            <button type="submit" className="primary" disabled={picked.length === 0 || inserting}>
              {t('ribbonInsert')}
            </button>
          </div>
        </div>
      </form>
    </div>
  )
}
