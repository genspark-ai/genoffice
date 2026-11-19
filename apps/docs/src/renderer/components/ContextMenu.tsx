import { useEffect, useLayoutEffect, useMemo, useRef, useState, type ReactNode } from 'react'
import type { Editor } from '@tiptap/core'
import { NodeSelection } from '@tiptap/pm/state'
import { mergeCells } from '@tiptap/pm/tables'
import { platformShortcuts } from '@genoffice/i18n'
import { redactTextBetween, type RedactableNode } from '../ai/redact-view'
import type { TextboxDisplay } from '@genoffice/docx-engine'

import { useI18n, type StringKey } from '../i18n/locale'
import { NONE, type ActionCommandId } from '../../shared/commands'
import { canCommentAt } from '../editor/comments'
import { SHAPE_INLINE_WRAP, shapeWrapOf } from '../editor/floating-z-order'
import { canGroupObjects, canUngroupObjects } from '../editor/object-group'
import { pasteFromClipboard } from '../editor/paste-actions'
import { selectedPictureAlt } from '../editor/picture-props'
import {
  inTableOrSelected,
  splitTableAtSelection,
  type CellHAlign,
  type CellTextDirection,
  type CellVAlign,
  type TableSelectKind,
} from '../editor/table-ops'
import { IconSparkle } from './icons'
import { IconCellAlign } from './table-icons'
import { spellcheckEnabled } from '../spellcheck-pref'
import { applySpellingSuggestion } from '../editor/spell-replace'
import { linkRangeAt, linkTarget, removeLink } from '../editor/link-actions'
import { fieldRangeAt, toggleFieldCodes, type FieldRange } from '../editor/field-codes'
import { hasRedactionIn } from '../editor/redaction'
import type { SpellLanguages } from '../../shared/ipc'
import { TRANSLATE_LANGS, appLangKey } from './translate-langs'
import type { RevisionRange } from '../editor/revisions'

export { FontDialog } from './FontDialog'

/**
 * Editor context menu (right click in the document body):
 * Cut/Copy/Paste · Font… · Paragraph… · Synonyms · Translate · Hyperlink (insert or Edit/Open/Copy/Remove) · New Comment.
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
  comment?: {
    id: string | null
  } | null
  /** Track Changes toggle; null while protection or a forced lock removes the choice */
  trackChanges?: { on: boolean } | null
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

const CTX_TEXT_DIRECTIONS: Array<[CellTextDirection, StringKey]> = [
  ['lrTb', 'ribbonTextDirectionHorizontal'],
  ['tbRl', 'ribbonTextDirectionRotate90'],
  ['btLr', 'ribbonTextDirectionRotate270'],
]

/** A context-menu submenu that stays inside the window: opened from an item near
 * the bottom edge it shifts up; one too tall for the whole window anchors to the
 * top edge and scrolls the rest (the Translate menu's twenty languages were
 * unreachable on a laptop). */
function CtxSubmenu({ children }: { children: ReactNode }) {
  const ref = useRef<HTMLDivElement>(null)
  const [fit, setFit] = useState<{ top: number; maxHeight?: number } | null>(null)
  useLayoutEffect(() => {
    const el = ref.current
    if (!el) return
    const rect = el.getBoundingClientRect()
    const limit = window.innerHeight - 8
    if (rect.bottom <= limit) return
    // current top sits at -5px inside the item wrap: viewport top = rect.top
    const shifted = limit - rect.height - rect.top - 5
    const raised = 3 - rect.top // do not rise above the window's top margin
    if (shifted >= raised) setFit({ top: shifted })
    else setFit({ top: raised, maxHeight: limit - 8 })
  }, [])
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
  trackChanges,
}: EditorContextMenuProps) {
  const { t, lang } = useI18n()
  const ref = useRef<HTMLDivElement>(null)
  const [submenu, setSubmenu] = useState<string | null>(null)

  // ---- spelling section (Word: suggestions · Add to Dictionary · Language) ----
  const spell = spellcheckEnabled() && editor.isEditable ? menu.spell : null
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
  const applySuggestion = (replacement: string) => {
    if (!spell) return
    applySpellingSuggestion(
      editor,
      menu.pos ?? null,
      spell.word,
      replacement,
      window.desktop.spellReplace,
    )
  }
  const addToDictionary = () => {
    if (spell) void window.desktop.spellAddWord?.(spell.word).then(() => onRespell?.())
  }
  const ignoreAll = () => {
    if (spell) void window.desktop.spellIgnoreWord?.(spell.word).then(() => onRespell?.())
  }
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
  const clickPos = menu.pos ?? from
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
  // Whether stopping the withholding is an option here. Decided from the live
  // selection rather than from a list kept alongside the document, so it cannot
  // drift from the marks the editor is actually carrying. A bare caret answers
  // false on its own, so no selection test is needed here either.
  const isRedactedSelection = hasRedactionIn(editor, from, to)
  // the presets quote the selection back to the model, so it reads the projection
  const selectedText = hasSelection
    ? redactTextBetween(editor.getJSON() as RedactableNode, from, to, ' ').trim()
    : ''
  // Synonyms targets a word / short phrase, not long selections
  const synonymText = selectedText.length > 0 && selectedText.length <= 20 ? selectedText : ''

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
  const canOrder = hasBoxes || (isImage && imageWrap !== null && imageWrap !== SHAPE_INLINE_WRAP)
  const boxes = protAttrs?.textboxes as TextboxDisplay[] | undefined
  // a run picture (in line or floating) wraps through its own node
  const runPicture =
    editor.state.selection instanceof NodeSelection &&
    editor.state.selection.node.type.name === 'docInlineImage' &&
    !editor.state.selection.node.attrs.rule
      ? editor.state.selection.node
      : null
  // Word's picture menu: clipboard, the picture, wrap / arrange, link, comment
  const onPicture = !!menu.imageSrc || !!runPicture || isImage
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
  // stacking ranks live on block anchors and shapes, not on run pictures
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
  const newCommentItem = item(t('appNewComment'), {
    disabled: !canComment,
    onClick: run(() => onCommand('insert-comment')),
  })

  return (
    <div
      ref={ref}
      className="ctx-menu"
      style={{ left: pos.left, top: pos.top, minWidth: MENU_WIDTH }}
      onContextMenu={(e) => e.preventDefault()}
    >
      {spell && (
        <>
          {/* always-disabled: says there is nothing to pick; not an action */}
          {spell.suggestions.length === 0 && item(t('appSpellNoSuggestions'), { disabled: true })}
          {spell.suggestions.slice(0, 5).map((word) => (
            <button
              key={word}
              className="ctx-item ctx-item-strong"
              onClick={run(() => applySuggestion(word))}
            >
              <span className="ctx-label">{word}</span>
            </button>
          ))}
          <div className="ctx-sep" />
          {item(t('appSpellIgnoreAll'), { onClick: run(ignoreAll) })}
          {item(t('appSpellAddToDictionary'), { onClick: run(addToDictionary) })}
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
        item(t('appPasteSpecial'), {
          disabled: !canEdit,
          onClick: run(() => onCommand('edit-paste-special')),
        })}
      {!onPicture &&
        item(t('appPastePlain'), {
          disabled: !canEdit,
          onClick: run(() => void pasteFromClipboard(editor, 'text')),
        })}
      {onPicture && (
        <>
          <div className="ctx-sep" />
          {item(t('appSaveImageAs'), {
            disabled: !menu.imageSrc,
            // command-exempt: acts on the picture under the pointer
            onClick: run(() => menu.imageSrc && onSaveImageAs(menu.imageSrc)),
          })}
          {item(t('appChangePicture'), {
            disabled: !canEdit || !isImage,
            onClick: run(() => onCommand('picture-replace')),
          })}
          {groupItem}
          {orderItems}
          <div className="ctx-sep" />
          {linkItems}
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
            disabled: !canEdit || !isImage,
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
          <div className="ctx-sep" />
          {menu.imageSrc &&
            item(t('appViewImage'), {
              // command-exempt: acts on the picture under the pointer
              onClick: run(() => onViewImage(menu.imageSrc!)),
            })}
          {newCommentItem}
        </>
      )}
      {comment && !onPicture && (
        <>
          <div className="ctx-sep" />
          {item(t('appCtxReplyComment'), {
            disabled: !comment.id,
            onClick: run(() => comment.id && onCommand('comment-reply', comment.id)),
          })}
          {item(t('appCtxDeleteComment'), {
            disabled: !comment.id,
            onClick: run(() => comment.id && onCommand('delete-comment', comment.id)),
          })}
          {item(t('appCtxResolveComment'), {
            disabled: !comment.id,
            onClick: run(() => comment.id && onCommand('comment-resolve', comment.id)),
          })}
        </>
      )}
      {!onPicture && (
        <>
          <div className="ctx-sep" />
          {item(t('appFontMenu'), { key: '⌘D', onClick: run(() => onCommand('font-dialog')) })}
          {item(t('appParagraphMenu'), {
            key: '⌥⌘M',
            onClick: run(() => onCommand('paragraph-dialog')),
          })}
          {menu.link?.toc && (
            <>
              <div className="ctx-sep" />
              {item(t('ribbonTocUpdate'), {
                disabled: !canEdit,
                onClick: run(() => onCommand('update-toc')),
              })}
            </>
          )}
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
          <div className="ctx-item-wrap" onMouseLeave={() => setSubmenu(null)}>
            {item(t('ribbonTableDeleteMenu'), { submenuKey: 'tableDelete', disabled: !canEdit })}
            {submenu === 'tableDelete' && canEdit && (
              <CtxSubmenu>
                {(
                  [
                    ['ribbonDeleteRow', 'table-delete-rows'],
                    ['ribbonDeleteColumn', 'table-delete-columns'],
                    ['ribbonDeleteTable', 'table-delete-table'],
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
                <button className="ctx-item" onClick={run(() => onCommand('table-delete-cells'))}>
                  <span className="ctx-label">{t('ribbonDeleteCells')}</span>
                </button>
              </CtxSubmenu>
            )}
          </div>
          <div className="ctx-item-wrap" onMouseLeave={() => setSubmenu(null)}>
            {item(t('ribbonSelect'), { submenuKey: 'tableSelect' })}
            {submenu === 'tableSelect' && (
              <CtxSubmenu>
                {(
                  [
                    ['row', 'ribbonSelectRow'],
                    ['column', 'ribbonSelectColumn'],
                    ['table', 'ribbonSelectTable'],
                    ['cell', 'ribbonSelectCell'],
                  ] as Array<[TableSelectKind, StringKey]>
                ).map(([kind, labelKey]) => (
                  <button
                    key={kind}
                    className="ctx-item"
                    onClick={run(() => onCommand(`table-select-${kind}`))}
                  >
                    <span className="ctx-label">{t(labelKey)}</span>
                  </button>
                ))}
              </CtxSubmenu>
            )}
          </div>
          {item(t('ribbonMergeCells'), {
            disabled: !canEdit || !mergeCells(editor.state),
            onClick: run(() => onCommand('table-merge-cells')),
          })}
          {item(t('ribbonSplitCellsDialog'), {
            disabled: !canEdit,
            onClick: run(() => onCommand('table-split-cells')),
          })}
          {item(t('ribbonSplitTable'), {
            disabled: !canEdit || !splitTableAtSelection()(editor.state),
            onClick: run(() => onCommand('table-split-table')),
          })}
          <div className="ctx-item-wrap" onMouseLeave={() => setSubmenu(null)}>
            {item(t('ribbonDistribute'), { submenuKey: 'tableDistribute', disabled: !canEdit })}
            {submenu === 'tableDistribute' && canEdit && (
              <CtxSubmenu>
                <button
                  className="ctx-item"
                  onClick={run(() => onCommand('table-distribute-rows'))}
                >
                  <span className="ctx-label">{t('ribbonDistributeRows')}</span>
                </button>
                <button
                  className="ctx-item"
                  onClick={run(() => onCommand('table-distribute-columns'))}
                >
                  <span className="ctx-label">{t('ribbonDistributeColumns')}</span>
                </button>
              </CtxSubmenu>
            )}
          </div>
          {item(t('appBordersAndShading'), {
            disabled: !canEdit,
            onClick: run(() => onCommand('format-borders-shading')),
          })}
          <div className="ctx-item-wrap" onMouseLeave={() => setSubmenu(null)}>
            {item(t('ribbonTextDirection'), { submenuKey: 'tableTextDir', disabled: !canEdit })}
            {submenu === 'tableTextDir' && canEdit && (
              <CtxSubmenu>
                {CTX_TEXT_DIRECTIONS.map(([dir, labelKey]) => (
                  <button
                    key={dir}
                    className="ctx-item"
                    onClick={run(() => onCommand('table-text-direction', dir))}
                  >
                    <span className="ctx-label">{t(labelKey)}</span>
                  </button>
                ))}
              </CtxSubmenu>
            )}
          </div>
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
          <div className="ctx-item-wrap" onMouseLeave={() => setSubmenu(null)}>
            {item(t('ribbonAutoFit'), { submenuKey: 'tableAutoFit', disabled: !canEdit })}
            {submenu === 'tableAutoFit' && canEdit && (
              <CtxSubmenu>
                {(
                  [
                    ['contents', 'ribbonAutoFitContents'],
                    ['window', 'ribbonAutoFitWindow'],
                    ['fixed', 'ribbonFixedColumnWidth'],
                  ] as Array<['contents' | 'window' | 'fixed', StringKey]>
                ).map(([mode, labelKey]) => (
                  <button
                    key={mode}
                    className="ctx-item"
                    onClick={run(() => onCommand(`table-autofit-${mode}`))}
                  >
                    <span className="ctx-label">{t(labelKey)}</span>
                  </button>
                ))}
              </CtxSubmenu>
            )}
          </div>
          {item(t('ribbonTableProperties'), {
            disabled: !canEdit,
            onClick: run(() => onCommand('table-properties')),
          })}
        </>
      )}
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
      {!onPicture && (
        <>
          <div className="ctx-sep" />
          {item(t('appSynonyms'), {
            disabled: !synonymText,
            ai: true,
            onClick: run(() =>
              onCommand('ai-preset', t('appSynonymsPrompt', { text: synonymText })),
            ),
          })}
          <div className="ctx-item-wrap" onMouseLeave={() => setSubmenu(null)}>
            {item(t('appTranslate'), {
              disabled: !hasSelection,
              submenuKey: 'translate',
              ai: true,
            })}
            {submenu === 'translate' && hasSelection && (
              <CtxSubmenu>
                {TRANSLATE_LANGS.map(({ code }) => (
                  <button
                    key={code}
                    className="ctx-item"
                    onClick={run(() =>
                      onCommand(
                        'ai-preset',
                        t('appTranslateSelectionPrompt', {
                          lang: t(appLangKey(code)),
                          text: selectedText,
                        }),
                      ),
                    )}
                  >
                    <span className="ctx-label">
                      {t('appTranslateTo', { lang: t(appLangKey(code)) })}
                    </span>
                  </button>
                ))}
              </CtxSubmenu>
            )}
          </div>
          {/* Withholding a span. No AI badge: the item spends the words rather
          than using the model, and the badge would read "Uses AI".
          Text only, deliberately. A picture alone in a paragraph is a *block*
          node (docProtected), which takes no marks at all, and a picture
          beside text is an inline atom whose mark is dropped by the docx save
          (convert.ts reads the run's node attrs, never the node's marks) — so
          offering it for a picture would hide it for one session and then lose
          the decision on the next save, with nothing to tell the reader. */}
          {onRedact &&
            item(t('redactMenuLabel'), {
              disabled: !canEdit || !hasSelection,
              onClick: run(onRedact),
            })}
          {/* The other half of the item above, and its mirror: offered only when the
          selection really carries the mark. Without this the sole way to stop
          withholding a span was to delete the words, which loses the reader's
          own text to undo a decision they made in the editor. Same reason as
          the item beside it for no AI badge — this one only moves a mark. */}
          {onUnredact &&
            isRedactedSelection &&
            item(t('redactShowLabel'), {
              disabled: !canEdit,
              onClick: run(onUnredact),
            })}
        </>
      )}
      {!onPicture && isFloating && (
        <>
          <div className="ctx-sep" />
          {groupItem}
          {wrapItem}
          {orderItems}
        </>
      )}
      {!onPicture && (
        <>
          <div className="ctx-sep" />
          {linkItems}
          {newCommentItem}
        </>
      )}
      {!onPicture && (revisions?.any || trackChanges) && (
        <>
          <div className="ctx-sep" />
          {revisions && revisions.at.length > 0 && (
            <>
              {item(t('appCtxAcceptChange'), {
                onClick: run(() => revisions.onRange('accept', revisions.at[0])),
              })}
              {item(t('appCtxRejectChange'), {
                onClick: run(() => revisions.onRange('reject', revisions.at[0])),
              })}
            </>
          )}
          {revisions?.any && (
            <>
              {item(t('ribbonAcceptAll'), { onClick: run(() => revisions.onAll('accept')) })}
              {item(t('ribbonRejectAll'), { onClick: run(() => revisions.onAll('reject')) })}
            </>
          )}
          {trackChanges &&
            item(`${trackChanges.on ? '✓ ' : ''}${t('ribbonTrackChanges')}`, {
              key: '⇧⌘E',
              onClick: run(() => onCommand('track-changes')),
            })}
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
