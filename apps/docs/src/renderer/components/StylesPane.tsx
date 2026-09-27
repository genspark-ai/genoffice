import { useMemo, useState, type MouseEvent as ReactMouseEvent } from 'react'
import type { Editor } from '@tiptap/core'
import {
  pendingHeadingLevel,
  type DocDefaults,
  type StyleInfo,
  type StyleUpsert,
} from '@genoffice/docx-engine'
import { useI18n } from '../i18n/locale'
import {
  activeStyleKey,
  allStyleEntries,
  collectUsedStyleIds,
  defaultParagraphStyleId,
  quickStyleEntries,
  styleKeyOf,
  styleLabel,
  stylePreviewCss,
  type StyleMap,
} from '../style-gallery'
import { IconClose } from './icons'
import type { RibbonFormatState } from './ribbon-format-state'
import { activeParaAttrs, applyGalleryStyle, applyParagraphStyleId } from './ribbon-tabs'
import { StyleDialog, formValuesOf, upsertFromForm, type StyleFormValues } from './StyleDialog'

type Show = 'recommended' | 'all'

const FORMAT_FIELDS = new Set<keyof StyleFormValues>([
  'font',
  'sizePt',
  'bold',
  'italic',
  'color',
  'align',
  'spaceBeforePt',
  'spaceAfterPt',
  'lineSpacing',
])

/** what the selection looks like (direct formatting over its styles), as dialog values */
function selectionFormValues(
  editor: Editor,
  fs: RibbonFormatState,
  styles: StyleMap,
  type: StyleFormValues['type'],
): StyleFormValues {
  const para = activeParaAttrs(fs.sub ?? editor)
  const paraStyle = styles.get(
    typeof para.styleId === 'string' ? para.styleId : (defaultParagraphStyleId(styles) ?? ''),
  )
  const charStyle = fs.charStyleId ? styles.get(fs.charStyleId) : undefined
  const pd = paraStyle?.display ?? {}
  const cd = charStyle?.display ?? {}
  const twips = (v: unknown, fallback: number | undefined): number | null => {
    const n = typeof v === 'number' ? v : fallback
    return n === undefined ? null : Math.round((n / 20) * 10) / 10
  }
  const color = fs.textColor ?? cd.color ?? pd.color
  const direct = para.align
  const align =
    (direct === 'left' || direct === 'center' || direct === 'right' || direct === 'justify'
      ? direct
      : pd.align) ?? null
  return {
    name: '',
    type,
    basedOn: type === 'paragraph' ? (paraStyle?.styleId ?? null) : null,
    font: fs.fontLatin ?? '',
    sizePt: fs.fontSizePt,
    bold: fs.bold || !!(cd.bold ?? pd.bold),
    italic: fs.italic || !!(cd.italic ?? pd.italic),
    color: color && color !== 'auto' ? color.toUpperCase() : '',
    align: align === 'distribute' ? 'justify' : align,
    spaceBeforePt: twips(para.spaceBefore, pd.spaceBeforeTwips),
    spaceAfterPt: twips(para.spaceAfter, pd.spaceAfterTwips),
    lineSpacing: typeof para.lineSpacing === 'number' ? para.lineSpacing : (pd.lineSpacing ?? null),
  }
}

interface Props {
  editor: Editor
  styles: StyleMap | undefined
  docDefaults?: DocDefaults
  fs: RibbonFormatState
  canEdit: boolean
  /** definitions already queued for the save (the live map catches up asynchronously) */
  pending: Readonly<Record<string, StyleUpsert>>
  /** pending style definition, written to styles.xml on save; returns an error message or null */
  onUpsert: (up: StyleUpsert) => string | null
  onNotice: (message: string) => void
  onClose: () => void
}

/** Word for Mac's Styles pane: every style as a card, search, Recommended / All, right-click to modify */
export function StylesPane({
  editor,
  styles,
  docDefaults,
  fs,
  canEdit,
  pending,
  onUpsert,
  onNotice,
  onClose,
}: Props) {
  const { t } = useI18n()
  const [query, setQuery] = useState('')
  const [show, setShow] = useState<Show>('recommended')
  const [menu, setMenu] = useState<{ x: number; y: number; info: StyleInfo } | null>(null)
  const [dialog, setDialog] = useState<{
    target: StyleInfo | null
    initial: StyleFormValues
  } | null>(null)
  const map = useMemo<StyleMap>(() => styles ?? new Map(), [styles])
  const used = useMemo(() => collectUsedStyleIds(fs.doc, map), [fs.doc, map])
  const entries = useMemo(
    () => (show === 'all' ? allStyleEntries(map) : quickStyleEntries(map, used)),
    [map, show, used],
  )
  const needle = query.trim().toLowerCase()
  const shown = needle
    ? entries.filter(
        (s) =>
          styleLabel(s, t).toLowerCase().includes(needle) ||
          s.name.toLowerCase().includes(needle) ||
          s.styleId.toLowerCase().includes(needle),
      )
    : entries
  const active = activeStyleKey(fs, styles)

  const openMenu = (e: ReactMouseEvent, info: StyleInfo) => {
    e.preventDefault()
    // a textbox has no paragraph styles: its selection can only feed character styles
    if (!canEdit || (fs.sub && info.type !== 'character')) return
    setMenu({ x: e.clientX, y: e.clientY, info })
  }
  const updateFromSelection = (info: StyleInfo) => {
    const values = selectionFormValues(
      editor,
      fs,
      map,
      info.type === 'character' ? 'character' : 'paragraph',
    )
    const err = onUpsert(upsertFromForm(info.styleId, values, FORMAT_FIELDS, false))
    if (!err) onNotice(t('appStyleUpdated', { name: styleLabel(info, t) }))
  }
  const submitDialog = (up: StyleUpsert) => {
    const target = dialog?.target ?? null
    setDialog(null)
    const err = onUpsert(up)
    if (err) return
    if (target) {
      onNotice(t('appStyleUpdated', { name: up.name ?? styleLabel(target, t) }))
      return
    }
    if (up.type === 'character')
      (fs.sub ?? editor).chain().focus().setMark('docTextStyle', { styleId: up.styleId }).run()
    else if (!fs.sub) {
      // a style based on a heading is a heading (inherited outline level), as after save/reopen
      const level = pendingHeadingLevel(
        up.styleId,
        (id) => (id === up.styleId ? up : pending[id]),
        (id) => map.get(id),
      )
      applyParagraphStyleId(editor, up.styleId, level ?? null)
    }
    onNotice(t('appStyleCreated', { name: up.name ?? up.styleId }))
  }

  return (
    <aside className="styles-pane">
      <div className="nav-pane-head">
        <span className="nav-pane-title">{t('appStylesTitle')}</span>
        <button
          className="nav-pane-close"
          aria-label={t('appClose')}
          data-tip={t('appClose')}
          onClick={onClose}
        >
          <IconClose size={14} />
        </button>
      </div>
      <input
        className="styles-pane-search"
        type="search"
        placeholder={t('ribbonStylesSearch')}
        aria-label={t('ribbonStylesSearch')}
        value={query}
        onChange={(e) => setQuery(e.target.value)}
      />
      <div className="styles-pane-list" role="listbox" aria-label={t('appStylesTitle')}>
        {shown.map((info) => {
          const key = styleKeyOf(info)
          const label = styleLabel(info, t)
          return (
            <button
              key={key}
              role="option"
              aria-selected={active === key}
              className={`style-card style-pane-card${active === key ? ' active' : ''}`}
              disabled={!canEdit || (info.type !== 'character' && !!fs.sub)}
              data-style-id={info.styleId}
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => applyGalleryStyle(editor, fs.sub, info, styles, fs.charStyleId)}
              onContextMenu={(e) => openMenu(e, info)}
            >
              <span className="style-card-preview" style={stylePreviewCss(info, docDefaults)}>
                {label}
              </span>
              <span className="style-pane-kind" aria-hidden="true">
                {info.type === 'character' ? 'a' : '¶'}
              </span>
            </button>
          )
        })}
        {shown.length === 0 && <div className="nav-empty">{t('ribbonStylesNoMatch')}</div>}
      </div>
      <div className="styles-pane-foot">
        <label>
          {t('ribbonStylesShow')}
          <select value={show} onChange={(e) => setShow(e.target.value as Show)}>
            <option value="recommended">{t('ribbonStylesShowRecommended')}</option>
            <option value="all">{t('ribbonStylesShowAll')}</option>
          </select>
        </label>
        <button
          className="styles-pane-new"
          disabled={!canEdit || !!fs.sub}
          onClick={() =>
            setDialog({ target: null, initial: selectionFormValues(editor, fs, map, 'paragraph') })
          }
        >
          {t('ribbonStyleNew')}
        </button>
      </div>
      {menu && (
        <div
          className="styles-pane-menu-backdrop"
          onMouseDown={() => setMenu(null)}
          onContextMenu={(e) => {
            e.preventDefault()
            setMenu(null)
          }}
        >
          <div
            className="ctx-menu"
            style={{ left: Math.min(menu.x, window.innerWidth - 240), top: menu.y }}
            onMouseDown={(e) => e.stopPropagation()}
          >
            <button
              className="ctx-item"
              onClick={() => {
                setMenu(null)
                setDialog({ target: menu.info, initial: formValuesOf(menu.info) })
              }}
            >
              <span className="ctx-label">{t('ribbonStyleModify')}</span>
            </button>
            <button
              className="ctx-item"
              onClick={() => {
                setMenu(null)
                updateFromSelection(menu.info)
              }}
            >
              <span className="ctx-label">{t('ribbonStyleUpdateFromSelection')}</span>
            </button>
            <button
              className="ctx-item"
              onClick={() => {
                setMenu(null)
                setDialog({
                  target: null,
                  initial: selectionFormValues(
                    editor,
                    fs,
                    map,
                    menu.info.type === 'character' ? 'character' : 'paragraph',
                  ),
                })
              }}
            >
              <span className="ctx-label">{t('ribbonStyleNew')}</span>
            </button>
          </div>
        </div>
      )}
      {dialog && (
        <StyleDialog
          target={dialog.target}
          initial={dialog.initial}
          styles={map}
          pending={pending}
          onSubmit={submitDialog}
          onClose={() => setDialog(null)}
        />
      )}
    </aside>
  )
}
