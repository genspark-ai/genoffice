import { useEffect, useMemo, useRef, useState } from 'react'
import type { Editor } from '@tiptap/core'
import type { Command } from '@tiptap/pm/state'
import { NodeSelection, TextSelection } from '@tiptap/pm/state'
import {
  addColumnAfter,
  addColumnBefore,
  addRowAfter,
  addRowBefore,
  deleteColumn,
  deleteRow,
  deleteTable,
  mergeCells,
  splitCell,
} from '@tiptap/pm/tables'
import { platformShortcuts } from '@genoffice/i18n'

import { useI18n, type StringKey } from '../i18n/locale'
import { wordRangeAtCaret } from '../editor/comments'
import { pasteFromClipboard } from '../editor/paste-actions'
import { setTableAutoFit } from '../editor/table-properties'
import { distributeSelectedColumns } from '../editor/table-sizing'
import {
  distributeRowsEvenly,
  inTableOrSelected,
  selectTablePart,
  setCellAlignment,
  setCellTextDirection,
  splitTableAtSelection,
  type CellHAlign,
  type CellTextDirection,
  type CellVAlign,
  type TableSelectKind,
} from '../editor/table-ops'
import type { TableDialogKind } from './TableDialogs'
import { IconSparkle } from './icons'
import { spellcheckEnabled } from '../spellcheck-pref'
import { applySpellingSuggestion } from '../editor/spell-replace'
import { linkRangeAt, linkTarget, removeLink } from '../editor/link-actions'
import { fieldRangeAt, toggleFieldCodes, type FieldRange } from '../editor/field-codes'
import type { SpellLanguages } from '../../shared/ipc'

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
}

interface EditorContextMenuProps {
  editor: Editor
  menu: ContextMenuState
  onClose: () => void
  onFontDialog: () => void
  onParagraphDialog: () => void
  onLink: () => void
  onNewComment: () => void
  onViewImage: (src: string) => void
  onSaveImageAs: (src: string) => void
  onAiPreset: (instruction: string) => void
  /** List items: restart numbering / continue numbering (shown when the cursor is on a docListItem) */
  onRestartNumbering?: () => void
  onContinueNumbering?: () => void
  onSetNumberingValue?: () => void
  onAdjustListIndents?: () => void
  onChangeListLevel?: (ilvl: number) => void
  /** F9 update fields (shown when the cursor is on an inline field) */
  onUpdateFields?: () => void
  /** Edit Field…: open the field-code dialog for that field */
  onEditField?: (field: FieldRange) => void
  /** Open Hyperlink: browser for http(s), in-document jump for #bookmark */
  onOpenLink?: (href: string) => void
  /** Word's table dialogs (Split Cells… / Insert Cells… / Delete Cells… / Table Properties…) */
  onTableDialog?: (kind: TableDialogKind) => void
  /** section content width the AutoFit / Distribute commands fit the grid into */
  sectionContentWidthPx?: number
  /** re-mark existing text after the session dictionary or languages changed */
  onRespell?: () => void
}

/** target languages mirrored from the Review → Translate dropdown; the localized label also goes into the LLM prompt */
const TRANSLATE_TARGETS: Array<{ labelKey: StringKey }> = [
  { labelKey: 'appLangEnglish' },
  { labelKey: 'appLangSimplifiedChinese' },
  { labelKey: 'appLangJapanese' },
  { labelKey: 'appLangKorean' },
  { labelKey: 'appLangFrench' },
  { labelKey: 'appLangGerman' },
  { labelKey: 'appLangSpanish' },
]

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

export function EditorContextMenu({
  editor,
  menu,
  onClose,
  onFontDialog,
  onParagraphDialog,
  onLink,
  onNewComment,
  onViewImage,
  onSaveImageAs,
  onAiPreset,
  onRestartNumbering,
  onContinueNumbering,
  onSetNumberingValue,
  onAdjustListIndents,
  onChangeListLevel,
  onUpdateFields,
  onEditField,
  onOpenLink,
  onTableDialog,
  sectionContentWidthPx = 624,
  onRespell,
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
    onLink()
  }
  const copyLink = () => {
    if (linkHref) void navigator.clipboard.writeText(linkHref)
  }
  const field = fieldRangeAt(editor.state, clickPos)
  const canComment = hasSelection || wordRangeAtCaret(editor) !== null
  const canEdit = editor.isEditable
  const selectedText = hasSelection ? editor.state.doc.textBetween(from, to, ' ').trim() : ''
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
  /** cell commands can't run on a whole-table NodeSelection: drop the caret into the first cell */
  const enterFirstCell = () => {
    const sel = editor.state.selection
    if (sel instanceof NodeSelection && sel.node.type.name === 'docTable') {
      editor.view.dispatch(
        editor.state.tr.setSelection(TextSelection.near(editor.state.doc.resolve(sel.from + 1))),
      )
    }
  }
  const runTable = (command: Command) => {
    editor.view.focus()
    enterFirstCell()
    command(editor.state, editor.view.dispatch)
  }

  const protAttrs = editor.getAttributes('docProtected')
  const isImage = protAttrs?.blockType === 'image'
  const isFloating =
    isImage ||
    (Array.isArray(protAttrs?.textboxes) && (protAttrs.textboxes as unknown[]).length > 0)
  const currentWrap = (protAttrs?.imageWrap as string | null) ?? null
  const setWrap = (wrap: string | null) => {
    const clearedPosition =
      wrap === null
        ? { imagePosH: null, imagePosV: null, imageOffsetXEmu: null, imageOffsetYEmu: null }
        : {}
    editor
      .chain()
      .focus()
      .updateAttributes('docProtected', { imageWrap: wrap, ...clearedPosition })
      .run()
  }
  // Stacking order among overlapping floating pictures. z-order only has a
  // visible effect on floating (front/behind) images, so the menu enables it
  // there; a bring-forward on an inline image also floats it (Word parity).
  const currentZOrder = Number((protAttrs?.imageZOrder as number | null) ?? 0)
  const isFloatingWrap = currentWrap === 'front' || currentWrap === 'behind'
  const setZOrder = (z: number) => {
    const attrs: Record<string, unknown> = { imageZOrder: z }
    // an inline image has no paint order; floating it (in front) makes the
    // reorder meaningful, matching Word's "Bring to Front" on an inline picture
    if (!isFloatingWrap) attrs.imageWrap = 'front'
    editor.chain().focus().updateAttributes('docProtected', attrs).run()
  }
  /** z-order of every floating anchor in the document (Word's to-front/to-back are document-global) */
  const floatingZOrders = (): number[] => {
    const zs: number[] = [currentZOrder]
    editor.state.doc.descendants((n) => {
      if (
        n.type.name === 'docProtected' &&
        (n.attrs.imageWrap === 'front' || n.attrs.imageWrap === 'behind')
      )
        zs.push(Number(n.attrs.imageZOrder ?? 0))
    })
    return zs
  }
  const bringToFront = () => setZOrder(Math.max(...floatingZOrders()) + 1)
  const sendToBack = () => setZOrder(Math.min(...floatingZOrders()) - 1)
  const bringForward = () => setZOrder(currentZOrder + 1)
  const sendBackward = () => setZOrder(currentZOrder - 1)

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

  return (
    <div
      ref={ref}
      className="ctx-menu"
      style={{ left: pos.left, top: pos.top, minWidth: MENU_WIDTH }}
      onContextMenu={(e) => e.preventDefault()}
    >
      {spell && (
        <>
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
                <div className="ctx-submenu ctx-submenu-scroll">
                  {spellLangOrder.map((code) => (
                    <button key={code} className="ctx-item" onClick={() => toggleLanguage(code)}>
                      <span className="ctx-label">
                        {spellLangs.active.includes(code) ? '✓ ' : ''}
                        {langNames?.of(code) ?? code}
                      </span>
                    </button>
                  ))}
                </div>
              )}
            </div>
          )}
          <div className="ctx-sep" />
        </>
      )}
      {menu.imageSrc && (
        <>
          {item(t('appViewImage'), { onClick: run(() => onViewImage(menu.imageSrc!)) })}
          {item(t('appSaveImageAs'), { onClick: run(() => onSaveImageAs(menu.imageSrc!)) })}
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
      {item(t('appPastePlain'), {
        disabled: !canEdit,
        onClick: run(() => void pasteFromClipboard(editor, 'text')),
      })}
      <div className="ctx-sep" />
      {item(t('appFontMenu'), { key: '⌘D', onClick: run(onFontDialog) })}
      {item(t('appParagraphMenu'), { key: '⌥⌘M', onClick: run(onParagraphDialog) })}
      {inTable && (
        <>
          <div className="ctx-sep" />
          <div className="ctx-item-wrap" onMouseLeave={() => setSubmenu(null)}>
            {item(t('ribbonInsert'), { submenuKey: 'tableInsert', disabled: !canEdit })}
            {submenu === 'tableInsert' && canEdit && (
              <div className="ctx-submenu">
                {(
                  [
                    ['ribbonInsertAbove', addRowBefore],
                    ['ribbonInsertBelow', addRowAfter],
                    ['ribbonInsertLeft', addColumnBefore],
                    ['ribbonInsertRight', addColumnAfter],
                  ] as Array<[StringKey, Command]>
                ).map(([labelKey, command]) => (
                  <button
                    key={labelKey}
                    className="ctx-item"
                    onClick={run(() => runTable(command))}
                  >
                    <span className="ctx-label">{t(labelKey)}</span>
                  </button>
                ))}
                {onTableDialog && (
                  <button className="ctx-item" onClick={run(() => onTableDialog('insertCells'))}>
                    <span className="ctx-label">{t('ribbonInsertCells')}</span>
                  </button>
                )}
              </div>
            )}
          </div>
          <div className="ctx-item-wrap" onMouseLeave={() => setSubmenu(null)}>
            {item(t('ribbonTableDeleteMenu'), { submenuKey: 'tableDelete', disabled: !canEdit })}
            {submenu === 'tableDelete' && canEdit && (
              <div className="ctx-submenu">
                {(
                  [
                    ['ribbonDeleteRow', deleteRow],
                    ['ribbonDeleteColumn', deleteColumn],
                    ['ribbonDeleteTable', deleteTable],
                  ] as Array<[StringKey, Command]>
                ).map(([labelKey, command]) => (
                  <button
                    key={labelKey}
                    className="ctx-item"
                    onClick={run(() => runTable(command))}
                  >
                    <span className="ctx-label">{t(labelKey)}</span>
                  </button>
                ))}
                {onTableDialog && (
                  <button className="ctx-item" onClick={run(() => onTableDialog('deleteCells'))}>
                    <span className="ctx-label">{t('ribbonDeleteCells')}</span>
                  </button>
                )}
              </div>
            )}
          </div>
          <div className="ctx-item-wrap" onMouseLeave={() => setSubmenu(null)}>
            {item(t('ribbonSelect'), { submenuKey: 'tableSelect' })}
            {submenu === 'tableSelect' && (
              <div className="ctx-submenu">
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
                    onClick={run(() => runTable(selectTablePart(kind)))}
                  >
                    <span className="ctx-label">{t(labelKey)}</span>
                  </button>
                ))}
              </div>
            )}
          </div>
          {item(t('ribbonMergeCells'), {
            disabled: !canEdit || !mergeCells(editor.state),
            onClick: run(() => runTable(mergeCells)),
          })}
          {item(t('ribbonSplitCellsDialog'), {
            disabled: !canEdit,
            onClick: run(() => {
              enterFirstCell()
              if (onTableDialog) onTableDialog('splitCells')
              else runTable(splitCell)
            }),
          })}
          {item(t('ribbonSplitTable'), {
            disabled: !canEdit || !splitTableAtSelection()(editor.state),
            onClick: run(() => runTable(splitTableAtSelection())),
          })}
          <div className="ctx-item-wrap" onMouseLeave={() => setSubmenu(null)}>
            {item(t('ribbonDistribute'), { submenuKey: 'tableDistribute', disabled: !canEdit })}
            {submenu === 'tableDistribute' && canEdit && (
              <div className="ctx-submenu">
                <button
                  className="ctx-item"
                  onClick={run(() => {
                    enterFirstCell()
                    distributeRowsEvenly(editor.view)
                  })}
                >
                  <span className="ctx-label">{t('ribbonDistributeRows')}</span>
                </button>
                <button
                  className="ctx-item"
                  onClick={run(() => runTable(distributeSelectedColumns(sectionContentWidthPx)))}
                >
                  <span className="ctx-label">{t('ribbonDistributeColumns')}</span>
                </button>
              </div>
            )}
          </div>
          <div className="ctx-item-wrap" onMouseLeave={() => setSubmenu(null)}>
            {item(t('ribbonCellAlignment'), { submenuKey: 'tableAlign', disabled: !canEdit })}
            {submenu === 'tableAlign' && canEdit && (
              <div className="ctx-submenu">
                {CTX_CELL_ALIGN.map(([v, h, labelKey]) => (
                  <button
                    key={`${v}-${h}`}
                    className="ctx-item"
                    onClick={run(() => runTable(setCellAlignment(v, h)))}
                  >
                    <span className="ctx-label">{t(labelKey)}</span>
                  </button>
                ))}
              </div>
            )}
          </div>
          <div className="ctx-item-wrap" onMouseLeave={() => setSubmenu(null)}>
            {item(t('ribbonTextDirection'), { submenuKey: 'tableTextDir', disabled: !canEdit })}
            {submenu === 'tableTextDir' && canEdit && (
              <div className="ctx-submenu">
                {CTX_TEXT_DIRECTIONS.map(([dir, labelKey]) => (
                  <button
                    key={dir}
                    className="ctx-item"
                    onClick={run(() => runTable(setCellTextDirection(dir)))}
                  >
                    <span className="ctx-label">{t(labelKey)}</span>
                  </button>
                ))}
              </div>
            )}
          </div>
          <div className="ctx-item-wrap" onMouseLeave={() => setSubmenu(null)}>
            {item(t('ribbonAutoFit'), { submenuKey: 'tableAutoFit', disabled: !canEdit })}
            {submenu === 'tableAutoFit' && canEdit && (
              <div className="ctx-submenu">
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
                    onClick={run(() => runTable(setTableAutoFit(mode, sectionContentWidthPx)))}
                  >
                    <span className="ctx-label">{t(labelKey)}</span>
                  </button>
                ))}
              </div>
            )}
          </div>
          {onTableDialog &&
            item(t('ribbonTableProperties'), {
              disabled: !canEdit,
              onClick: run(() => {
                enterFirstCell()
                onTableDialog('properties')
              }),
            })}
        </>
      )}
      {(field || editor.isActive('instrField')) && onUpdateFields && (
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
      {editor.isActive('docListItem') && !!editor.getAttributes('docListItem').numId && (
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
            disabled: !onSetNumberingValue || !canEdit,
            onClick: run(() => onSetNumberingValue?.()),
          })}
          {item(t('appAdjustListIndents'), {
            disabled: !onAdjustListIndents || !canEdit,
            onClick: run(() => onAdjustListIndents?.()),
          })}
          <div className="ctx-item-wrap" onMouseLeave={() => setSubmenu(null)}>
            {item(t('appChangeListLevel'), {
              disabled: !onChangeListLevel || !canEdit,
              submenuKey: 'listLevel',
            })}
            {submenu === 'listLevel' && onChangeListLevel && canEdit && (
              <div className="ctx-submenu">
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
              </div>
            )}
          </div>
        </>
      )}
      <div className="ctx-sep" />
      {item(t('appSynonyms'), {
        disabled: !synonymText,
        ai: true,
        onClick: run(() => onAiPreset(t('appSynonymsPrompt', { text: synonymText }))),
      })}
      <div className="ctx-item-wrap" onMouseLeave={() => setSubmenu(null)}>
        {item(t('appTranslate'), { disabled: !hasSelection, submenuKey: 'translate', ai: true })}
        {submenu === 'translate' && hasSelection && (
          <div className="ctx-submenu">
            {TRANSLATE_TARGETS.map((target) => (
              <button
                key={target.labelKey}
                className="ctx-item"
                onClick={run(() =>
                  onAiPreset(
                    t('appTranslateSelectionPrompt', {
                      lang: t(target.labelKey),
                      text: selectedText,
                    }),
                  ),
                )}
              >
                <span className="ctx-label">
                  {t('appTranslateTo', { lang: t(target.labelKey) })}
                </span>
              </button>
            ))}
          </div>
        )}
      </div>
      {isFloating && (
        <>
          <div className="ctx-sep" />
          <div className="ctx-item-wrap" onMouseLeave={() => setSubmenu(null)}>
            {item(t('appWrapTextMenu'), { submenuKey: 'wrap' })}
            {submenu === 'wrap' && (
              <div className="ctx-submenu">
                {WRAP_OPTIONS.map((opt) => (
                  <button
                    key={String(opt.value)}
                    className="ctx-item"
                    onClick={run(() => setWrap(opt.value))}
                  >
                    <span className="ctx-label">
                      {currentWrap === opt.value ? '✓ ' : ''}
                      {t(opt.labelKey)}
                    </span>
                  </button>
                ))}
              </div>
            )}
          </div>
          <div className="ctx-item-wrap" onMouseLeave={() => setSubmenu(null)}>
            {item(t('appArrangeMenu'), { submenuKey: 'arrange' })}
            {submenu === 'arrange' && (
              <div className="ctx-submenu">
                <button className="ctx-item" onClick={run(bringToFront)}>
                  <span className="ctx-label">{t('appBringToFront')}</span>
                </button>
                <button className="ctx-item" onClick={run(bringForward)}>
                  <span className="ctx-label">{t('appBringForward')}</span>
                </button>
                <button className="ctx-item" onClick={run(sendBackward)}>
                  <span className="ctx-label">{t('appSendBackward')}</span>
                </button>
                <button className="ctx-item" onClick={run(sendToBack)}>
                  <span className="ctx-label">{t('appSendToBack')}</span>
                </button>
              </div>
            )}
          </div>
        </>
      )}
      <div className="ctx-sep" />
      {onLinkRun ? (
        <>
          {item(t('appEditHyperlink'), { disabled: !linkEditable, onClick: run(editLink) })}
          {item(t('appOpenHyperlink'), {
            disabled: !onOpenLink || !(linkTarget(linkHref) || menu.link?.toc),
            onClick: run(() => onOpenLink?.(linkHref)),
          })}
          {item(t('appCopyHyperlink'), { onClick: run(copyLink) })}
          {item(t('appRemoveHyperlink'), {
            disabled: !linkEditable,
            onClick: run(() => linkRange && removeLink(editor, linkRange)),
          })}
        </>
      ) : (
        item(t('appHyperlinkMenu'), { key: '⌘K', disabled: !canEdit, onClick: run(onLink) })
      )}
      {item(t('appNewComment'), { disabled: !canComment, onClick: run(onNewComment) })}
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
