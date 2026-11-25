import { useEffect, useLayoutEffect, useMemo, useRef, useState, type ReactNode } from 'react'
import type { Editor } from '@tiptap/core'
import { NodeSelection } from '@tiptap/pm/state'
import { CellSelection, mergeCells } from '@tiptap/pm/tables'
import { platformShortcuts } from '@genoffice/i18n'
import { redactTextBetween, type RedactableNode } from '../ai/redact-view'
import { fetchSynonyms } from '../ai/synonyms'
import type { TextboxDisplay } from '@genoffice/docx-engine'

import { useI18n, type StringKey } from '../i18n/locale'
import { NONE, type ActionCommandId } from '../../shared/commands'
import { canCommentAt, wordRangeAtCaret } from '../editor/comments'
import { SHAPE_INLINE_WRAP, shapeWrapOf } from '../editor/floating-z-order'
import { canGroupObjects, canUngroupObjects } from '../editor/object-group'
import { clipboardHasPicture, pasteFromClipboard } from '../editor/paste-actions'
import { selectedPictureAlt } from '../editor/picture-props'
import { selectedShapeSize } from '../editor/shape-size'
import {
  inTableOrSelected,
  tableCellsSelection,
  type CellHAlign,
  type CellVAlign,
} from '../editor/table-ops'
import { IconSparkle } from './icons'
import { IconCellAlign } from './table-icons'
import { spellcheckEnabled } from '../spellcheck-pref'
import { misspellingAt, spellWordAt, spellingDismissed, wordSetAside } from '../editor/spell-menu'
import {
  commentRows,
  headingFoldRows,
  noteRows,
  revisionLabelKeys,
  spellRows,
  type CommentTarget,
  type CtxEntry,
  type CtxRow,
  type NoteTarget,
  type SpellTarget,
} from './context-menu-model'
import { headingFoldAt } from '../editor/heading-collapse'
import { subPictureEditor } from '../editor/picture-bytes'
import { linkRangeAt, linkTarget, removeLink } from '../editor/link-actions'
import { fieldRangeAt, toggleFieldCodes, type FieldRange } from '../editor/field-codes'
import { hasRedactionIn } from '../editor/redaction'
import { grammarFindingAt } from '../editor/grammar-marks'
import type { SpellLanguages } from '../../shared/ipc'
import type { RevisionRange } from '../editor/revisions'

export { FontDialog } from './FontDialog'

/**
 * Editor context menu (right click in the document body):
 * Cut/Copy/Paste · Font… · Paragraph… · Synonyms · Translate… · Hyperlink (insert or Edit/Open/Copy/Remove) · New Comment.
 */

export interface ContextMenuState {
  x: number
  y: number
  /** src of the picture under the pointer, when the click landed on one */
  imageSrc?: string | null
  /** claim sequence of the right-click this menu answers (pairs it with Chromium's data) */
  seq?: number
  /** document position under the click (the caret for a keyboard-invoked menu) */
  pos?: number | null
  /** Chromium's misspelling under the pointer, relayed by the main process after the menu opened */
  spell?: { word: string; suggestions: string[] } | null
  /** hyperlink under the pointer as rendered (a TOC line counts: Word treats its entries as links) */
  link?: { href: string; toc?: boolean; tocTitle?: string } | null
  /** thread of the comment balloon the click landed on */
  commentId?: string | null
  /** opened from a margin balloon: `pos` is that item's anchor, not a pointer hit */
  balloon?: boolean
  /** footnote / endnote whose text or reference mark the click landed on */
  note?: NoteTarget | null
}

interface EditorContextMenuProps {
  editor: Editor
  menu: ContextMenuState
  onClose: () => void
  /** runs a command of the shared command table, as the menu bar and the ribbon do */
  onCommand: (id: ActionCommandId, payload?: string) => void
  onViewImage: (src: string) => void
  onSaveImageAs: (src: string) => void
  /** List items: restart numbering / continue numbering (shown when the cursor is on a docListItem) */
  onRestartNumbering?: () => void
  onContinueNumbering?: () => void
  onChangeListLevel?: (ilvl: number) => void
  /** F9 update fields (shown when the cursor is on an inline field) */
  onUpdateFields?: () => void
  /** Edit Field…: open the field-code dialog for that field */
  onEditField?: (field: FieldRange) => void
  /** Open Hyperlink: browser for http(s), in-document jump for #bookmark */
  onOpenLink?: (href: string) => void
  /** Ask for a label, then withhold the selection from the model */
  onRedact?: () => void
  /** Stop withholding a selection that already carries the mark, keeping the text */
  onUnredact?: () => void
  /** re-mark existing text after the session dictionary or languages changed */
  onRespell?: () => void
  /** tracked changes under the pointer; null while protection disables accept/reject */
  revisions?: {
    at: RevisionRange[]
    any: boolean
    onRange: (action: 'accept' | 'reject', target: RevisionRange) => void
    onAll: (action: 'accept' | 'reject') => void
  } | null
  /** comment thread under the pointer (Word: Reply / Delete / Resolve stay listed, greyed off a comment) */
  comment?: CommentTarget | null
}

const MENU_WIDTH = 240

const CTX_CELL_ALIGN: Array<[CellVAlign, CellHAlign, StringKey]> = [
  ['top', 'left', 'ribbonAlignTopLeft'],
  ['top', 'center', 'ribbonAlignTopCenter'],
  ['top', 'right', 'ribbonAlignTopRight'],
  ['center', 'left', 'ribbonAlignMiddleLeft'],
  ['center', 'center', 'ribbonAlignMiddleCenter'],
  ['center', 'right', 'ribbonAlignMiddleRight'],
  ['bottom', 'left', 'ribbonAlignBottomLeft'],
  ['bottom', 'center', 'ribbonAlignBottomCenter'],
  ['bottom', 'right', 'ribbonAlignBottomRight'],
]

/** A context-menu submenu that stays inside the window: opened from an item near
 * the bottom edge it shifts up; one too tall for the whole window anchors to the
 * top edge and scrolls the rest (the Translate menu's twenty languages were
 * unreachable on a laptop). */
function CtxSubmenu({ children, grown }: { children: ReactNode; grown?: unknown }) {
  const ref = useRef<HTMLDivElement>(null)
  const [fit, setFit] = useState<{ top: number; maxHeight?: number } | null>(null)
  const applied = useRef<number | null>(null)
  useLayoutEffect(() => {
    const el = ref.current
    if (!el) return
    const rect = el.getBoundingClientRect()
    const limit = window.innerHeight - 8
    if (rect.bottom <= limit && el.scrollHeight <= el.clientHeight) return
    // the item wrap's top edge in the viewport: the panel sits at -5px until shifted
    const origin = rect.top - (applied.current ?? -5)
    const height = Math.max(rect.height, el.scrollHeight)
    const shifted = limit - height - origin
    const raised = 8 - origin // do not rise above the window's top margin
    const next = shifted >= raised ? { top: shifted } : { top: raised, maxHeight: limit - 8 }
    applied.current = next.top
    setFit(next)
    // measured again when rows arrive after the panel opened
  }, [grown])
  return (
    <div
      ref={ref}
      className="ctx-submenu ctx-submenu-scroll"
      style={fit ? { top: fit.top, maxHeight: fit.maxHeight } : undefined}
    >
      {children}
    </div>
  )
}

export function EditorContextMenu({
  editor,
  menu,
  onClose,
  onCommand,
  onViewImage,
  onSaveImageAs,
  onRestartNumbering,
  onContinueNumbering,
  onChangeListLevel,
  onUpdateFields,
  onEditField,
  onOpenLink,
  onRedact,
  onUnredact,
  onRespell,
  revisions,
  comment,
}: EditorContextMenuProps) {
  const { t, lang } = useI18n()
  const ref = useRef<HTMLDivElement>(null)
  const [submenu, setSubmenu] = useState<string | null>(null)

  const clickPos = menu.pos ?? editor.state.selection.from
  const headingFold = menu.balloon || menu.note ? null : headingFoldAt(editor, clickPos)
  // ---- spelling section (Word: suggestions · Ignore · Add to Dictionary · AutoCorrect · Spelling…) ----
  const spellOn = spellcheckEnabled() && editor.isEditable && !menu.balloon && !menu.note
  const spell = useMemo((): SpellTarget | null => {
    if (!spellOn) return null
    const api = window.desktop
    const own = misspellingAt(editor.state, clickPos, (w) => api.spellIsMisspelled?.(w) ?? false)
    if (own)
      return {
        word: own.word,
        pos: own.from,
        suggestions: api.spellSuggestions?.(own.word) ?? [],
        marked: menu.spell?.word === own.word,
      }
    if (!menu.spell) return null
    // the browser's marker can outlive Ignore Once and Ignore All by a moment
    const at = spellWordAt(editor.state.doc, clickPos)
    if (wordSetAside(menu.spell.word) || (at && spellingDismissed(editor.state, at))) return null
    return { ...menu.spell, pos: menu.pos ?? null, marked: true }
  }, [editor, spellOn, clickPos, menu.spell, menu.pos])
  const [spellLangs, setSpellLangs] = useState<SpellLanguages | null>(null)
  useEffect(() => {
    if (!spell) return
    let alive = true
    void window.desktop.spellLanguages?.().then((langs) => {
      if (alive) setSpellLangs(langs)
    })
    return () => {
      alive = false
    }
  }, [spell])
  const langNames = useMemo(() => {
    try {
      return new Intl.DisplayNames([lang], { type: 'language' })
    } catch {
      return null
    }
  }, [lang])
  // active dictionaries first, the rest by display name
  const spellLangOrder = useMemo(() => {
    if (!spellLangs) return []
    const name = (code: string) => langNames?.of(code) ?? code
    const rest = spellLangs.available
      .filter((code) => !spellLangs.active.includes(code))
      .sort((a, b) => name(a).localeCompare(name(b), lang))
    return [...spellLangs.active, ...rest]
  }, [spellLangs, langNames, lang])
  const toggleLanguage = (code: string) => {
    if (!spellLangs) return
    const active = spellLangs.active.includes(code)
      ? spellLangs.active.filter((l) => l !== code)
      : [...spellLangs.active, code]
    if (!active.length) return
    // the respell kick that follows must not run under an open menu (Word closes it too)
    onClose()
    void window.desktop.spellSetLanguages?.(active).then(() => onRespell?.())
  }

  const { from, to } = editor.state.selection
  const hasSelection = from !== to
  // ---- hyperlink / field under the pointer (Word acts on the run, not the selection) ----
  const linkRange = linkRangeAt(editor.state, clickPos)
  const linkHref = menu.link?.href ?? linkRange?.href ?? null
  const onLinkRun = linkHref !== null
  const linkEditable = !!linkRange && !menu.link?.toc && editor.isEditable
  const editLink = () => {
    if (linkRange) editor.commands.setTextSelection({ from: linkRange.from, to: linkRange.to })
    onCommand('insert-link')
  }
  const copyLink = () => {
    if (linkHref) void navigator.clipboard.writeText(linkHref)
  }
  const field = fieldRangeAt(editor.state, clickPos)
  const canComment = canCommentAt(editor)
  const canEdit = editor.isEditable
  // a misspelling under the pointer is the spelling checker's; a balloon's anchor is not a click on the text
  const grammar = spell || menu.balloon ? null : grammarFindingAt(editor.state, clickPos)
  // Whether stopping the withholding is an option here. Decided from the live
  // selection rather than from a list kept alongside the document, so it cannot
  // drift from the marks the editor is actually carrying. A bare caret answers
  // false on its own, so no selection test is needed here either.
  const isRedactedSelection = hasRedactionIn(editor, from, to)
  // Word looks up and translates the word the caret is in when nothing is selected
  const lookup = hasSelection ? { from, to } : wordRangeAtCaret(editor)
  // the presets quote the text back to the model, so it reads the projection
  const lookupText = lookup
    ? redactTextBetween(editor.getJSON() as RedactableNode, lookup.from, lookup.to, ' ').trim()
    : ''
  // Synonyms targets a word / short phrase, not long selections
  const synonymText = lookupText.length > 0 && lookupText.length <= 20 ? lookupText : ''
  // asked for once, when the submenu first opens: a right-click alone spends no request
  const [synonymReply, setSynonymReply] = useState<{ text: string; words: string[] } | null>(null)
  const synonymsAsked = useRef('')
  const synonymsOpen = submenu === 'synonyms' && !!synonymText
  useEffect(() => {
    if (!synonymsOpen || synonymsAsked.current === synonymText) return
    synonymsAsked.current = synonymText
    void fetchSynonyms(synonymText).then((words) => setSynonymReply({ text: synonymText, words }))
  }, [synonymsOpen, synonymText])
  const synonyms = synonymReply?.text === synonymText ? synonymReply.words : null

  // keep the menu inside the viewport (flip up / clamp left near the edges)
  const [pos, setPos] = useState({ left: menu.x, top: menu.y })
  useEffect(() => {
    const el = ref.current
    if (!el) return
    const rect = el.getBoundingClientRect()
    let left = menu.x
    let top = menu.y
    if (left + rect.width > window.innerWidth - 8) left = window.innerWidth - rect.width - 8
    if (top + rect.height > window.innerHeight - 8) top = Math.max(8, menu.y - rect.height)
    setPos({ left, top })
  }, [menu])

  useEffect(() => {
    const close = (e: MouseEvent) => {
      if (!ref.current?.contains(e.target as Node)) onClose()
    }
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') onClose()
    }
    window.addEventListener('mousedown', close)
    window.addEventListener('keydown', onKey)
    window.addEventListener('blur', onClose)
    // shell tab-strip presses never reach this document; the preload relays
    // them (app:chrome-pressed) so the menu still dismisses
    const offChrome = window.desktop?.onChromePressed?.(onClose)
    return () => {
      window.removeEventListener('mousedown', close)
      window.removeEventListener('keydown', onKey)
      window.removeEventListener('blur', onClose)
      offChrome?.()
    }
  }, [onClose])

  const run = (action: () => void) => () => {
    onClose()
    action()
  }

  // ---- table section (shown when the cursor is inside a table, Word parity) ----
  const inTable = inTableOrSelected(editor.state)
  // several cells or the whole table selected: Word swaps Delete Cells… / Split Cells… for the items that act on the block
  const tableCells = tableCellsSelection(editor.state)
  const cells =
    tableCells ??
    (editor.state.selection instanceof CellSelection && editor.state.selection.ranges.length > 1
      ? editor.state.selection
      : null)
  const wholeRows = !!tableCells || !!cells?.isRowSelection()
  const wholeColumns = !!tableCells || !!cells?.isColSelection()
  // the table commands act on every cell of a selected table
  const cellsState = tableCells
    ? editor.state.apply(editor.state.tr.setSelection(tableCells))
    : editor.state
  const [deleteLabel, deleteCommand]: [string, ActionCommandId] =
    wholeRows && wholeColumns
      ? [t('ribbonDeleteTable'), 'table-delete-table']
      : wholeRows
        ? [t('ribbonDeleteRow'), 'table-delete-rows']
        : wholeColumns
          ? [t('ribbonDeleteColumn'), 'table-delete-columns']
          : [t('ribbonDeleteCells'), 'table-delete-cells']
  // only the object itself (right-click selects it as a node): a text range
  // that merely spans a picture keeps the text menu
  const protAttrs =
    editor.state.selection instanceof NodeSelection &&
    editor.state.selection.node.type.name === 'docProtected'
      ? editor.state.selection.node.attrs
      : null
  const isImage = protAttrs?.blockType === 'image'
  const hasBoxes =
    Array.isArray(protAttrs?.textboxes) && (protAttrs.textboxes as unknown[]).length > 0
  const isFloating = isImage || hasBoxes
  // a picture in a line of text has no place in the stacking order (Word greys the four)
  const imageWrap = (protAttrs?.imageWrap as string | null | undefined) ?? null
  const imageCanOrder = isImage && imageWrap !== null && imageWrap !== SHAPE_INLINE_WRAP
  const boxes = protAttrs?.textboxes as TextboxDisplay[] | undefined
  // a run picture (in line or floating) wraps through its own node
  const runPicture =
    editor.state.selection instanceof NodeSelection &&
    editor.state.selection.node.type.name === 'docInlineImage' &&
    !editor.state.selection.node.attrs.rule
      ? editor.state.selection.node
      : null
  const canOrder = hasBoxes || imageCanOrder || !!runPicture?.attrs.wrap
  // a picture selected in the header, footer or text box being edited changes where it stands
  const subPicture = menu.imageSrc ? subPictureEditor(editor) : null
  const canChange = isImage || (!!runPicture && !runPicture.attrs.layers) || !!subPicture
  // Word's picture menu: clipboard, the picture, wrap / arrange, link, comment
  const onPicture = !!menu.imageSrc || !!runPicture || isImage
  const [clipboardPicture, setClipboardPicture] = useState(false)
  useEffect(() => {
    if (!canChange) return
    let alive = true
    void clipboardHasPicture().then((has) => {
      if (alive) setClipboardPicture(has)
    })
    return () => {
      alive = false
    }
  }, [canChange])
  const currentWrap = runPicture
    ? ((runPicture.attrs.wrap as string | null) ?? null)
    : boxes?.length
      ? shapeWrapOf(boxes[0])
      : ((protAttrs?.imageWrap as string | null) ?? null)

  const clipboard = (action: 'cut' | 'copy') => {
    editor.commands.focus()
    document.execCommand(action)
  }

  const item = (
    label: string,
    opts: {
      key?: string
      disabled?: boolean
      onClick?: () => void
      submenuKey?: string
      ai?: boolean
    },
  ) => (
    <button
      className="ctx-item"
      disabled={opts.disabled}
      data-tip={opts.ai ? t('appAiBadgeTip') : undefined}
      onMouseEnter={() => setSubmenu(opts.submenuKey ?? null)}
      onClick={opts.submenuKey ? undefined : opts.onClick}
    >
      <span className="ctx-label">{label}</span>
      {opts.ai && (
        <span className="copilot-badge copilot-badge-menu">
          <IconSparkle size={10} />
        </span>
      )}
      {opts.key && <span className="ctx-key">{platformShortcuts(opts.key)}</span>}
      {opts.submenuKey && <span className="ctx-arrow">›</span>}
    </button>
  )

  const entry = (row: CtxEntry, key: number, submenuKey?: string | null) => (
    <button
      key={key}
      className={row.strong ? 'ctx-item ctx-item-strong' : 'ctx-item'}
      data-ctx={row.id}
      disabled={row.disabled}
      onMouseEnter={submenuKey === undefined ? undefined : () => setSubmenu(submenuKey)}
      onClick={row.command ? run(() => onCommand(row.command!, row.payload)) : undefined}
    >
      <span className="ctx-label">
        {row.label ?? (row.labelKey ? t(row.labelKey) : '')}
        {row.ellipsis ? '…' : ''}
      </span>
      {submenuKey && <span className="ctx-arrow">›</span>}
    </button>
  )
  /** a section of the menu as its model lists it (context-menu-model.ts) */
  const rows = (list: CtxRow[]) =>
    list.map((row, i) =>
      row === 'sep' ? (
        <div key={i} className="ctx-sep" />
      ) : row.children ? (
        <div key={i} className="ctx-item-wrap" onMouseLeave={() => setSubmenu(null)}>
          {entry(row, i, row.id)}
          {submenu === row.id && !row.disabled && (
            <CtxSubmenu>{row.children.map((child, j) => entry(child, j))}</CtxSubmenu>
          )}
        </div>
      ) : (
        entry(row, i, null)
      ),
    )
  const { selection } = editor.state
  const note =
    menu.note ??
    (selection instanceof NodeSelection && selection.node.type.name === 'docNoteRef'
      ? ({
          kind: selection.node.attrs.kind === 'endnote' ? 'endnote' : 'footnote',
          id: String(selection.node.attrs.id),
          at: 'ref',
        } satisfies NoteTarget)
      : null)
  // Word's short menus: a comment balloon and a note's text carry their own items only
  const ownRows =
    menu.balloon && menu.commentId && comment
      ? commentRows(comment)
      : note?.at === 'text'
        ? noteRows(note, canEdit)
        : null
  const revisionAt = revisions?.at[0]
  const revisionKeys = revisionAt ? revisionLabelKeys(revisionAt.kind) : null

  const wrapItem = (
    <div className="ctx-item-wrap" onMouseLeave={() => setSubmenu(null)}>
      {item(t('appWrapTextMenu'), { submenuKey: 'wrap', disabled: !isFloating && !runPicture })}
      {submenu === 'wrap' && (isFloating || !!runPicture) && (
        <CtxSubmenu>
          {WRAP_OPTIONS.map((opt) => (
            <button
              key={String(opt.value)}
              className="ctx-item"
              onClick={run(() => onCommand('picture-wrap', opt.value ?? NONE))}
            >
              <span className="ctx-label">
                {currentWrap === opt.value ? '✓ ' : ''}
                {t(opt.labelKey)}
              </span>
            </button>
          ))}
        </CtxSubmenu>
      )}
    </div>
  )
  const canGroup = canEdit && canGroupObjects(editor.state)
  const canUngroup = canEdit && canUngroupObjects(editor.state)
  const groupItem = (
    <div className="ctx-item-wrap" onMouseLeave={() => setSubmenu(null)}>
      {item(t('appGroupMenu'), { submenuKey: 'group', disabled: !canGroup && !canUngroup })}
      {submenu === 'group' && (canGroup || canUngroup) && (
        <CtxSubmenu>
          <button
            className="ctx-item"
            disabled={!canGroup}
            onClick={run(() => onCommand('arrange-group', 'group'))}
          >
            <span className="ctx-label">{t('appGroupMenu')}</span>
          </button>
          <button
            className="ctx-item"
            disabled={!canUngroup}
            onClick={run(() => onCommand('arrange-group', 'ungroup'))}
          >
            <span className="ctx-label">{t('ribbonArrangeUngroup')}</span>
          </button>
        </CtxSubmenu>
      )}
    </div>
  )
  const orderItems = (
    <>
      <div className="ctx-item-wrap" onMouseLeave={() => setSubmenu(null)}>
        {item(t('appBringToFront'), { submenuKey: 'toFront', disabled: !canOrder })}
        {submenu === 'toFront' && canOrder && (
          <CtxSubmenu>
            <button className="ctx-item" onClick={run(() => onCommand('arrange-order', 'front'))}>
              <span className="ctx-label">{t('appBringToFront')}</span>
            </button>
            <button className="ctx-item" onClick={run(() => onCommand('arrange-order', 'forward'))}>
              <span className="ctx-label">{t('appBringForward')}</span>
            </button>
          </CtxSubmenu>
        )}
      </div>
      <div className="ctx-item-wrap" onMouseLeave={() => setSubmenu(null)}>
        {item(t('appSendToBack'), { submenuKey: 'toBack', disabled: !canOrder })}
        {submenu === 'toBack' && canOrder && (
          <CtxSubmenu>
            <button className="ctx-item" onClick={run(() => onCommand('arrange-order', 'back'))}>
              <span className="ctx-label">{t('appSendToBack')}</span>
            </button>
            <button
              className="ctx-item"
              onClick={run(() => onCommand('arrange-order', 'backward'))}
            >
              <span className="ctx-label">{t('appSendBackward')}</span>
            </button>
          </CtxSubmenu>
        )}
      </div>
    </>
  )
  const linkItems = onLinkRun ? (
    <>
      {item(t('appEditHyperlink'), {
        disabled: !linkEditable,
        // command-exempt: acts on the link under the pointer
        onClick: run(editLink),
      })}
      {item(t('appOpenHyperlink'), {
        disabled: !onOpenLink || !(linkTarget(linkHref) || menu.link?.toc),
        // command-exempt: acts on the link under the pointer
        onClick: run(() => onOpenLink?.(linkHref)),
      })}
      {item(t('appCopyHyperlink'), { onClick: run(copyLink) })}
      {item(t('appRemoveHyperlink'), {
        disabled: !linkEditable,
        // command-exempt: acts on the link under the pointer
        onClick: run(() => linkRange && removeLink(editor, linkRange)),
      })}
    </>
  ) : (
    item(t('appHyperlinkMenu'), {
      key: '⌘K',
      disabled: !canEdit,
      onClick: run(() => onCommand('insert-link')),
    })
  )
  const commentSection = comment && (
    <>
      <div className="ctx-sep" />
      {rows(commentRows(comment))}
    </>
  )
  // No AI badge on either: they move a mark and spend no request. Text only: a
  // picture takes no mark that the docx save keeps.
  const redactItems = !onPicture && (onRedact || (onUnredact && isRedactedSelection)) && (
    <>
      {onRedact &&
        item(t('redactMenuLabel'), {
          disabled: !canEdit || !hasSelection,
          // command-exempt: asks for a label first, through the app's own dialog
          onClick: run(onRedact),
        })}
      {onUnredact &&
        isRedactedSelection &&
        item(t('redactShowLabel'), {
          disabled: !canEdit,
          // command-exempt: acts on the marks of the live selection
          onClick: run(onUnredact),
        })}
    </>
  )
  const newCommentItem = item(t('appNewComment'), {
    disabled: !canComment,
    onClick: run(() => onCommand('insert-comment')),
  })

  if (ownRows)
    return (
      <div
        ref={ref}
        className="ctx-menu"
        style={{ left: pos.left, top: pos.top, minWidth: MENU_WIDTH }}
        onContextMenu={(e) => e.preventDefault()}
      >
        {rows(ownRows)}
      </div>
    )

  return (
    <div
      ref={ref}
      className="ctx-menu"
      style={{ left: pos.left, top: pos.top, minWidth: MENU_WIDTH }}
      onContextMenu={(e) => e.preventDefault()}
    >
      {spell && (
        <>
          {rows(spellRows(spell, canEdit))}
          {spellLangs && spellLangs.available.length > 0 && (
            <div className="ctx-item-wrap" onMouseLeave={() => setSubmenu(null)}>
              {item(t('appSpellLanguage'), { submenuKey: 'spellLang' })}
              {submenu === 'spellLang' && (
                <CtxSubmenu>
                  {spellLangOrder.map((code) => (
                    <button key={code} className="ctx-item" onClick={() => toggleLanguage(code)}>
                      <span className="ctx-label">
                        {spellLangs.active.includes(code) ? '✓ ' : ''}
                        {langNames?.of(code) ?? code}
                      </span>
                    </button>
                  ))}
                </CtxSubmenu>
              )}
            </div>
          )}
          <div className="ctx-sep" />
        </>
      )}
      {grammar && (
        <>
          {/* always-disabled: says what is wrong with the marked words; not an action */}
          <button className="ctx-item ctx-item-note" disabled>
            <span className="ctx-label">{grammar.message}</span>
          </button>
          {grammar.suggestions.map((text) => (
            <button
              key={text}
              className="ctx-item ctx-item-strong"
              disabled={!canEdit}
              onClick={run(() =>
                onCommand('grammar-replace', JSON.stringify({ id: grammar.id, text })),
              )}
            >
              <span className="ctx-label">{text}</span>
            </button>
          ))}
          <div className="ctx-sep" />
          {item(t('appSpellingIgnoreOnce'), {
            onClick: run(() => onCommand('grammar-ignore', grammar.id)),
          })}
          {item(t('appGrammarMenu'), {
            onClick: run(() => onCommand('grammar-dialog', grammar.id)),
          })}
          <div className="ctx-sep" />
        </>
      )}
      {!onPicture && revisions && revisionAt && revisionKeys && (
        <>
          {item(t(revisionKeys.accept), {
            // command-exempt: acts on the change under the pointer
            onClick: run(() => revisions.onRange('accept', revisionAt)),
          })}
          {item(t(revisionKeys.reject), {
            // command-exempt: acts on the change under the pointer
            onClick: run(() => revisions.onRange('reject', revisionAt)),
          })}
          <div className="ctx-sep" />
        </>
      )}
      {item(t('appCut'), {
        key: '⌘X',
        disabled: !hasSelection || !canEdit,
        onClick: run(() => clipboard('cut')),
      })}
      {item(t('appCopy'), {
        key: '⌘C',
        disabled: !hasSelection,
        onClick: run(() => clipboard('copy')),
      })}
      {item(t('appPaste'), {
        key: '⌘V',
        disabled: !canEdit,
        onClick: run(() => void pasteFromClipboard(editor)),
      })}
      {!onPicture &&
        !inTable &&
        !headingFold &&
        item(t('appPasteSpecial'), {
          disabled: !canEdit,
          onClick: run(() => onCommand('edit-paste-special')),
        })}
      {!onPicture && headingFold && (
        <>
          <div className="ctx-sep" />
          {rows(headingFoldRows(headingFold))}
        </>
      )}
      {onPicture && (
        <>
          <div className="ctx-sep" />
          {item(t('appSaveImageAs'), {
            disabled: !menu.imageSrc,
            // command-exempt: acts on the picture under the pointer
            onClick: run(() => menu.imageSrc && onSaveImageAs(menu.imageSrc)),
          })}
          <div className="ctx-item-wrap" onMouseLeave={() => setSubmenu(null)}>
            {item(t('appChangePicture'), {
              submenuKey: 'changePicture',
              disabled: !canEdit || !canChange,
            })}
            {submenu === 'changePicture' && canEdit && canChange && (
              <CtxSubmenu>
                <button className="ctx-item" onClick={run(() => onCommand('picture-replace'))}>
                  <span className="ctx-label">{t('appChangePictureFromFile')}</span>
                </button>
                <button className="ctx-item" onClick={run(() => onCommand('picture-replace-icon'))}>
                  <span className="ctx-label">{t('appChangePictureFromIcons')}</span>
                </button>
                <button
                  className="ctx-item"
                  disabled={!clipboardPicture}
                  onClick={run(() => onCommand('picture-replace-clipboard'))}
                >
                  <span className="ctx-label">{t('appChangePictureFromClipboard')}</span>
                </button>
              </CtxSubmenu>
            )}
          </div>
          <div className="ctx-sep" />
          {groupItem}
          {orderItems}
          <div className="ctx-sep" />
          {linkItems}
          <div className="ctx-sep" />
          {item(t('appInsertCaption'), {
            disabled: !canEdit,
            onClick: run(() => onCommand('insert-caption')),
          })}
          {wrapItem}
          <div className="ctx-sep" />
          {item(t('appViewAltText'), {
            disabled: selectedPictureAlt(editor) === null,
            onClick: run(() => onCommand('picture-alt-text')),
          })}
          {item(t('appCropMenu'), {
            disabled: !canEdit || !isImage,
            onClick: run(() => onCommand('picture-edit', 'crop')),
          })}
          {item(t('appAutoCrop'), {
            disabled: !canEdit || (!isImage && !runPicture),
            onClick: run(() => onCommand('picture-auto-crop')),
          })}
          {item(t('appSizeAndPosition'), {
            disabled: !canEdit || (!isImage && !runPicture),
            onClick: run(() => onCommand('picture-size-position')),
          })}
          {item(t('appFormatPicture'), {
            disabled: !isImage && !runPicture,
            onClick: run(() => onCommand('format-object')),
          })}
          {menu.imageSrc && (
            <>
              <div className="ctx-sep" />
              {item(t('appViewImage'), {
                // command-exempt: acts on the picture under the pointer
                onClick: run(() => onViewImage(menu.imageSrc!)),
              })}
            </>
          )}
        </>
      )}
      {comment && !onPicture && !cells && commentSection}
      {!onPicture && !inTable && (
        <>
          <div className="ctx-sep" />
          {item(t('appFontMenu'), { key: '⌘D', onClick: run(() => onCommand('font-dialog')) })}
          {item(t('appParagraphMenu'), {
            key: '⌥⌘M',
            onClick: run(() => onCommand('paragraph-dialog')),
          })}
        </>
      )}
      {!onPicture && menu.link?.toc && (
        <>
          <div className="ctx-sep" />
          {item(t('ribbonTocUpdate'), {
            disabled: !canEdit,
            onClick: run(() => onCommand('update-toc')),
          })}
        </>
      )}
      {!onPicture && note && (
        <>
          <div className="ctx-sep" />
          {rows(noteRows(note, canEdit))}
        </>
      )}
      {inTable && !onPicture && (
        <>
          <div className="ctx-sep" />
          <div className="ctx-item-wrap" onMouseLeave={() => setSubmenu(null)}>
            {item(t('ribbonInsert'), { submenuKey: 'tableInsert', disabled: !canEdit })}
            {submenu === 'tableInsert' && canEdit && (
              <CtxSubmenu>
                {(
                  [
                    ['ribbonInsertAbove', 'table-insert-rows-above'],
                    ['ribbonInsertBelow', 'table-insert-rows-below'],
                    ['ribbonInsertLeft', 'table-insert-cols-left'],
                    ['ribbonInsertRight', 'table-insert-cols-right'],
                  ] as Array<[StringKey, ActionCommandId]>
                ).map(([labelKey, command]) => (
                  <button
                    key={labelKey}
                    className="ctx-item"
                    onClick={run(() => onCommand(command))}
                  >
                    <span className="ctx-label">{t(labelKey)}</span>
                  </button>
                ))}
                <button className="ctx-item" onClick={run(() => onCommand('table-insert-cells'))}>
                  <span className="ctx-label">{t('ribbonInsertCells')}</span>
                </button>
              </CtxSubmenu>
            )}
          </div>
          {item(deleteLabel, {
            disabled: !canEdit,
            onClick: run(() => onCommand(deleteCommand)),
          })}
          {cells ? (
            <>
              {item(t('ribbonMergeCells'), {
                disabled: !canEdit || !mergeCells(cellsState),
                onClick: run(() => onCommand('table-merge-cells')),
              })}
              {item(t('ribbonDistributeRowsEvenly'), {
                disabled: !canEdit,
                onClick: run(() => onCommand('table-distribute-rows')),
              })}
              {item(t('ribbonDistributeColumnsEvenly'), {
                disabled: !canEdit,
                onClick: run(() => onCommand('table-distribute-columns')),
              })}
            </>
          ) : (
            item(t('ribbonSplitCellsDialog'), {
              disabled: !canEdit,
              onClick: run(() => onCommand('table-split-cells')),
            })
          )}
          <div className="ctx-sep" />
          {item(t('appBordersAndShading'), {
            disabled: !canEdit,
            onClick: run(() => onCommand('format-borders-shading')),
          })}
          {item(`${t('ribbonTextDirection')}…`, {
            disabled: !canEdit,
            onClick: run(() => onCommand('format-text-direction')),
          })}
          <div className="ctx-item-wrap" onMouseLeave={() => setSubmenu(null)}>
            {item(t('ribbonCellAlignment'), { submenuKey: 'tableAlign', disabled: !canEdit })}
            {submenu === 'tableAlign' && canEdit && (
              <CtxSubmenu>
                {CTX_CELL_ALIGN.map(([v, h, labelKey]) => (
                  <button
                    key={`${v}-${h}`}
                    className="ctx-item"
                    onClick={run(() => onCommand('table-cell-align', `${v},${h}`))}
                  >
                    <span className="ctx-ico">
                      <IconCellAlign v={v} h={h} size={14} />
                    </span>
                    <span className="ctx-label">{t(labelKey)}</span>
                  </button>
                ))}
              </CtxSubmenu>
            )}
          </div>
          {item(`${t('ribbonTableProperties')}…`, {
            disabled: !canEdit,
            onClick: run(() => onCommand('table-properties')),
          })}
        </>
      )}
      {comment && !onPicture && cells && commentSection}
      {!onPicture && (field || editor.isActive('instrField')) && onUpdateFields && (
        <>
          <div className="ctx-sep" />
          {item(t('appUpdateField'), { key: 'F9', onClick: run(() => onUpdateFields()) })}
          {field && (
            <>
              {item(t('appToggleFieldCodes'), {
                onClick: run(() => toggleFieldCodes(editor.view, field)),
              })}
              {item(t('appEditField'), {
                disabled: !canEdit || !onEditField,
                onClick: run(() => onEditField?.(field)),
              })}
            </>
          )}
        </>
      )}
      {!onPicture &&
        editor.isActive('docListItem') &&
        !!editor.getAttributes('docListItem').numId && (
          <>
            <div className="ctx-sep" />
            {item(t('appRestartNumbering'), {
              disabled: !onRestartNumbering,
              onClick: run(() => onRestartNumbering?.()),
            })}
            {item(t('appContinueNumbering'), {
              disabled: !onContinueNumbering,
              onClick: run(() => onContinueNumbering?.()),
            })}
            {item(t('appSetNumberingValue'), {
              disabled: !canEdit,
              onClick: run(() => onCommand('list-dialog', 'value')),
            })}
            {item(t('appAdjustListIndents'), {
              disabled: !canEdit,
              onClick: run(() => onCommand('list-dialog', 'indents')),
            })}
            <div className="ctx-item-wrap" onMouseLeave={() => setSubmenu(null)}>
              {item(t('appChangeListLevel'), {
                disabled: !onChangeListLevel || !canEdit,
                submenuKey: 'listLevel',
              })}
              {submenu === 'listLevel' && onChangeListLevel && canEdit && (
                <CtxSubmenu>
                  {Array.from({ length: 9 }, (_, i) => (
                    <button
                      key={i}
                      className="ctx-item"
                      aria-current={
                        (Number(editor.getAttributes('docListItem').ilvl) || 0) === i
                          ? 'true'
                          : undefined
                      }
                      onClick={run(() => onChangeListLevel(i))}
                    >
                      <span className="ctx-label" style={{ paddingLeft: i * 8 }}>
                        {t('appParaOutlineLevelN', { n: String(i + 1) })}
                      </span>
                    </button>
                  ))}
                </CtxSubmenu>
              )}
            </div>
          </>
        )}
      {!onPicture && !inTable && (
        <>
          <div className="ctx-sep" />
          <div className="ctx-item-wrap" onMouseLeave={() => setSubmenu(null)}>
            {item(t('appSynonyms'), {
              disabled: !synonymText,
              submenuKey: 'synonyms',
              ai: true,
            })}
            {submenu === 'synonyms' && synonymText && (
              <CtxSubmenu grown={synonyms}>
                {/* always-disabled: a status line, not an action */}
                {synonyms === null && item(t('ribbonLoading'), { disabled: true })}
                {synonyms?.length === 0 &&
                  item(`(${t('appSpellNoSuggestions')})`, { disabled: true })}
                {synonyms?.map((word) => (
                  <button
                    key={word}
                    className="ctx-item"
                    disabled={!canEdit}
                    onClick={run(() => onCommand('synonym-replace', word))}
                  >
                    <span className="ctx-label">{word}</span>
                  </button>
                ))}
                <div className="ctx-sep" />
                <button className="ctx-item" onClick={run(() => onCommand('tools-thesaurus'))}>
                  <span className="ctx-label">{t('ribbonThesaurus')}…</span>
                </button>
              </CtxSubmenu>
            )}
          </div>
          {item(`${t('appTranslate')}…`, {
            disabled: !lookupText,
            ai: true,
            onClick: run(() => onCommand('tools-translate-selection')),
          })}
        </>
      )}
      {redactItems && inTable && <div className="ctx-sep" />}
      {redactItems}
      {!onPicture && isFloating && (
        <>
          <div className="ctx-sep" />
          {groupItem}
          {wrapItem}
          {orderItems}
          {item(t('appSizeAndPosition'), {
            disabled: !canEdit || !selectedShapeSize(editor),
            onClick: run(() => onCommand('picture-size-position')),
          })}
        </>
      )}
      {!onPicture && (
        <>
          <div className="ctx-sep" />
          {(!inTable || onLinkRun) && linkItems}
          {newCommentItem}
        </>
      )}
      {!onPicture && revisions?.any && (
        <>
          <div className="ctx-sep" />
          {/* command-exempt: the review callbacks the items above use */}
          {item(t('ribbonAcceptAll'), { onClick: run(() => revisions.onAll('accept')) })}
          {item(t('ribbonRejectAll'), { onClick: run(() => revisions.onAll('reject')) })}
        </>
      )}
    </div>
  )
}

/** Wrap options from Word's image context menu (inline/square/top-bottom/behind text/in front of text); shared with the Picture Format tab */
export const WRAP_OPTIONS: Array<{ labelKey: StringKey; value: string | null }> = [
  { labelKey: 'appWrapInline', value: null },
  { labelKey: 'appWrapSquareLeft', value: 'square-left' },
  { labelKey: 'appWrapSquareRight', value: 'square-right' },
  { labelKey: 'appWrapTopBottom', value: 'topBottom' },
  { labelKey: 'appWrapBehind', value: 'behind' },
  { labelKey: 'appWrapFront', value: 'front' },
]

export { ParagraphDialog } from './ParagraphDialog'
