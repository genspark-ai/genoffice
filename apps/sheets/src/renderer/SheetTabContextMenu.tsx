import { useCallback, useEffect, useRef, useState } from 'react'
import { ColorPicker } from '@genoffice/ui'

import { useI18n } from './i18n/locale'
import { useModalDialog } from './modal-dialog'
import type { SheetTabActions, SheetTabInfo } from './sheet-tab-actions'
import {
  insertIndex,
  moveTargets,
  tabRange,
  unhideCandidates,
  visibleAfterRemoving,
} from './sheet-tab-menu'

const TAB_SELECTOR = '#univer-container [data-u-comp="slide-tab-item"]'
const TAB_BAR_SELECTOR = '#univer-container [data-u-comp="slide-tab-bar"]'
const MENU_WIDTH = 220
// Control-click is the macOS context-menu gesture; only Command groups there.
const IS_MAC = navigator.platform.toLowerCase().includes('mac')

type Dialog = 'insert' | 'move' | 'unhide' | null

const tabIdAt = (target: EventTarget | null): string | null => {
  if (!(target instanceof Element)) return null
  const tab = target.closest<HTMLElement>(TAB_SELECTOR)
  return tab?.dataset.id ?? null
}

/// Excel's sheet-tab right-click menu, replacing Univer's (the stock menu
/// cannot be reordered or given Excel's Insert / Move or Copy / Unhide
/// dialogs). Rename-on-double-click and tab dragging stay Univer's.
/// Ctrl/Cmd+click and Shift+click group tabs like Excel; Delete / Hide /
/// Tab Color then apply to the whole group.
export function SheetTabContextMenu({
  actions,
  getSheetProtected,
  getWorkbookLocked,
  onProtectSheet,
}: {
  readonly actions: SheetTabActions | null
  /// Sampled when the menu opens, for the sheet that was just activated.
  readonly getSheetProtected: () => boolean | null
  /// Workbook structure lock: Excel greys every structural tab command.
  readonly getWorkbookLocked: () => boolean | null
  readonly onProtectSheet: () => void
}): React.JSX.Element | null {
  const { t } = useI18n()
  const [menuAt, setMenuAt] = useState<{ x: number; y: number } | null>(null)
  const [group, setGroup] = useState<readonly string[]>([])
  const [dialog, setDialog] = useState<Dialog>(null)
  const [colorOpen, setColorOpen] = useState(false)
  const [sheets, setSheets] = useState<readonly SheetTabInfo[]>([])
  // Snapshot of the right-clicked (now active) tab: Univer is only queried
  // when the menu opens, never during a render.
  const [activeId, setActiveId] = useState<string | null>(null)
  const [protection, setProtection] = useState<{ sheet: boolean | null; locked: boolean }>({
    sheet: null,
    locked: false,
  })
  const getSheetProtectedRef = useRef(getSheetProtected)
  getSheetProtectedRef.current = getSheetProtected
  const getWorkbookLockedRef = useRef(getWorkbookLocked)
  getWorkbookLockedRef.current = getWorkbookLocked
  const groupRef = useRef(group)
  groupRef.current = group
  const actionsRef = useRef(actions)
  actionsRef.current = actions

  const refresh = useCallback((): SheetTabInfo[] => {
    const list = actionsRef.current?.listSheets() ?? []
    setSheets(list)
    return list
  }, [])

  // Grouped tabs are highlighted through an attribute Univer's React tree
  // does not own; re-applied whenever the strip re-renders its items.
  useEffect(() => {
    const paint = (): void => {
      for (const tab of document.querySelectorAll<HTMLElement>(TAB_SELECTOR)) {
        if (group.includes(tab.dataset.id ?? '')) tab.dataset.grouped = 'true'
        else delete tab.dataset.grouped
      }
    }
    paint()
    const bar = document.querySelector(TAB_BAR_SELECTOR)
    if (!bar) return
    const observer = new MutationObserver(paint)
    observer.observe(bar, { childList: true, subtree: true })
    return () => observer.disconnect()
  }, [group])

  useEffect(() => {
    const onContextMenu = (event: MouseEvent): void => {
      const id = tabIdAt(event.target)
      const current = actionsRef.current
      if (!id || !current) return
      event.preventDefault()
      event.stopPropagation()
      if (!groupRef.current.includes(id)) setGroup([])
      void current.activate(id).then(() => {
        const ids = new Set(refresh().map((sheet) => sheet.id))
        setActiveId(current.activeSheetId() ?? id)
        setGroup((tabs) => tabs.filter((tab) => ids.has(tab)))
        setProtection({
          sheet: getSheetProtectedRef.current(),
          locked: getWorkbookLockedRef.current() === true,
        })
        setColorOpen(false)
        setMenuAt({ x: event.clientX, y: event.clientY })
      })
    }
    const onPointerDown = (event: PointerEvent): void => {
      const id = tabIdAt(event.target)
      const current = actionsRef.current
      if (!id || !current || event.button !== 0) return
      const modifier = IS_MAC ? event.metaKey : event.ctrlKey
      if (IS_MAC && event.ctrlKey && !modifier) return
      if (!modifier && !event.shiftKey) {
        if (!groupRef.current.includes(id)) setGroup([])
        return
      }
      event.preventDefault()
      event.stopPropagation()
      const active = current.activeSheetId()
      if (event.shiftKey) {
        setGroup(active ? tabRange(current.listSheets(), active, id) : [id])
        return
      }
      const base = groupRef.current.length > 0 ? groupRef.current : active ? [active] : []
      setGroup(base.includes(id) ? base.filter((tab) => tab !== id) : [...base, id])
    }
    document.addEventListener('contextmenu', onContextMenu, true)
    document.addEventListener('pointerdown', onPointerDown, true)
    return () => {
      document.removeEventListener('contextmenu', onContextMenu, true)
      document.removeEventListener('pointerdown', onPointerDown, true)
    }
  }, [refresh])

  const menuRef = useRef<HTMLDivElement>(null)
  useEffect(() => {
    if (!menuAt) return
    const onPress = (event: MouseEvent): void => {
      if (!menuRef.current?.contains(event.target as Node)) setMenuAt(null)
    }
    const onKey = (event: KeyboardEvent): void => {
      if (event.key === 'Escape') setMenuAt(null)
    }
    window.addEventListener('mousedown', onPress)
    window.addEventListener('keydown', onKey)
    return () => {
      window.removeEventListener('mousedown', onPress)
      window.removeEventListener('keydown', onKey)
    }
  }, [menuAt])

  if (!actions) return null
  const targets = group.length > 0 ? group : activeId ? [activeId] : []
  const targetSet = new Set(targets)
  const visibleLeft = visibleAfterRemoving(sheets, targetSet)
  const canHideOrDelete = targets.length > 0 && visibleLeft >= 1
  const locked = protection.locked
  const hiddenSheets = unhideCandidates(sheets)
  const activeColor = sheets.find((sheet) => sheet.id === activeId)?.tabColor ?? null
  const close = (): void => setMenuAt(null)
  const guarded = (): boolean => {
    if (canHideOrDelete) return true
    actions.notify(t('dlgSheetKeepVisible'))
    return false
  }

  const item = (
    label: string,
    onClick: () => void,
    options: { disabled?: boolean } = {},
  ): React.JSX.Element => (
    <button type="button" role="menuitem" disabled={options.disabled} onClick={onClick}>
      {label}
    </button>
  )

  const menuStyle: React.CSSProperties | undefined = menuAt
    ? {
        left: Math.min(menuAt.x, window.innerWidth - MENU_WIDTH),
        bottom: Math.max(8, window.innerHeight - menuAt.y),
      }
    : undefined
  const submenuLeft = menuAt ? menuAt.x + MENU_WIDTH + 260 < window.innerWidth : true

  return (
    <>
      {menuAt && (
        <div ref={menuRef} className="sheet-tab-menu" role="menu" style={menuStyle}>
          {item(
            t('dlgSheetTabInsert'),
            () => {
              close()
              setDialog('insert')
            },
            { disabled: locked },
          )}
          {item(
            t('dlgSheetTabDelete'),
            () => {
              close()
              if (!guarded()) return
              void actions.remove(targets).then((removed) => {
                if (removed) setGroup([])
              })
            },
            { disabled: locked || targets.length === 0 },
          )}
          {item(
            t('dlgSheetTabRename'),
            () => {
              close()
              if (activeId) actions.rename(activeId)
            },
            { disabled: locked },
          )}
          {item(
            t('dlgSheetTabMoveCopy'),
            () => {
              close()
              setDialog('move')
            },
            { disabled: locked },
          )}
          <hr />
          {item(t(protection.sheet ? 'dlgSheetTabUnprotect' : 'dlgSheetTabProtect'), () => {
            close()
            onProtectSheet()
          })}
          <div
            className={`sheet-tab-menu-sub ${submenuLeft ? '' : 'flip'}`}
            onMouseEnter={() => !locked && setColorOpen(true)}
          >
            <button
              type="button"
              role="menuitem"
              aria-haspopup="menu"
              aria-expanded={colorOpen}
              disabled={locked}
              onClick={() => setColorOpen((open) => !open)}
            >
              {t('dlgSheetTabColor')}
              <span aria-hidden="true">{submenuLeft ? '▸' : '◂'}</span>
            </button>
            {colorOpen && (
              <div className="sheet-tab-menu-pop" role="menu">
                <ColorPicker
                  value={activeColor}
                  strings={{
                    auto: t('dlgSheetTabNoColor'),
                    themeColors: t('appThemeColors'),
                    standardColors: t('appStandardColors'),
                    moreColors: t('appMoreColors'),
                  }}
                  onPick={(hex) => {
                    close()
                    void actions.setTabColor(targets, hex)
                  }}
                />
              </div>
            )}
          </div>
          <hr />
          {item(
            t('dlgSheetTabHide'),
            () => {
              close()
              if (!guarded()) return
              void actions.hide(targets).then(() => setGroup([]))
            },
            { disabled: locked || targets.length === 0 },
          )}
          {item(
            t('dlgSheetTabUnhide'),
            () => {
              close()
              setDialog('unhide')
            },
            { disabled: locked || hiddenSheets.length === 0 },
          )}
          <hr />
          {item(t('dlgSheetTabSelectAll'), () => {
            close()
            setGroup(sheets.filter((sheet) => !sheet.hidden).map((sheet) => sheet.id))
          })}
          {group.length > 0 &&
            item(t('dlgSheetTabUngroup'), () => {
              close()
              setGroup([])
            })}
        </div>
      )}
      {dialog === 'insert' && activeId && (
        <InsertSheetDialog
          onInsert={() => void actions.insertBefore(activeId)}
          onClose={() => setDialog(null)}
        />
      )}
      {dialog === 'move' && activeId && (
        <MoveCopySheetDialog
          sheets={moveTargets(sheets)}
          activeId={activeId}
          onApply={(beforeId, copy) => {
            void actions.moveOrCopy(targets, beforeId, copy).then(() => {
              if (copy) setGroup([])
            })
          }}
          onClose={() => setDialog(null)}
        />
      )}
      {dialog === 'unhide' && (
        <UnhideSheetDialog
          sheets={hiddenSheets}
          onApply={(ids) => void actions.show(ids)}
          onClose={() => setDialog(null)}
        />
      )}
    </>
  )
}

function DialogFrame({
  title,
  children,
  onClose,
  onOk,
  okDisabled,
}: {
  readonly title: string
  readonly children: React.ReactNode
  readonly onClose: () => void
  readonly onOk: () => void
  readonly okDisabled?: boolean
}): React.JSX.Element {
  const { t } = useI18n()
  const modal = useModalDialog(onClose)
  return (
    <div className="dialog-backdrop" onClick={onClose}>
      <div
        className="format-cells-dialog sheet-tab-dialog"
        role="dialog"
        {...modal}
        aria-label={title}
        onClick={(event) => event.stopPropagation()}
      >
        <header>{title}</header>
        <section className="dialog-body">{children}</section>
        <div className="dialog-actions">
          <button className="secondary" onClick={onClose}>
            {t('dlgCancel')}
          </button>
          <button className="primary-action" disabled={okDisabled} onClick={onOk}>
            {t('dlgOk')}
          </button>
        </div>
      </div>
    </div>
  )
}

function InsertSheetDialog({
  onInsert,
  onClose,
}: {
  readonly onInsert: () => void
  readonly onClose: () => void
}): React.JSX.Element {
  const { t } = useI18n()
  const ok = (): void => {
    onInsert()
    onClose()
  }
  return (
    <DialogFrame title={t('dlgSheetInsertTitle')} onClose={onClose} onOk={ok}>
      <div className="sheet-tab-list" role="listbox" aria-label={t('dlgSheetInsertTitle')}>
        <button
          type="button"
          role="option"
          aria-selected="true"
          className="selected"
          autoFocus
          onDoubleClick={ok}
          onKeyDown={(event) => {
            if (event.key === 'Enter') ok()
          }}
        >
          {t('dlgSheetInsertWorksheet')}
        </button>
      </div>
      <p className="dialog-note">{t('dlgSheetInsertNote')}</p>
    </DialogFrame>
  )
}

const END = '__move-to-end__'

function MoveCopySheetDialog({
  sheets,
  activeId,
  onApply,
  onClose,
}: {
  readonly sheets: readonly SheetTabInfo[]
  readonly activeId: string
  readonly onApply: (beforeId: string | null, copy: boolean) => void
  readonly onClose: () => void
}): React.JSX.Element {
  const { t } = useI18n()
  const order = sheets.map((sheet) => sheet.id)
  const [before, setBefore] = useState<string>(order[insertIndex(order, activeId)] ?? END)
  const [copy, setCopy] = useState(false)
  const ok = (): void => {
    onApply(before === END ? null : before, copy)
    onClose()
  }
  return (
    <DialogFrame title={t('dlgSheetMoveTitle')} onClose={onClose} onOk={ok}>
      <label className="dialog-span sheet-tab-label">
        {t('dlgSheetMoveBefore')}
        <select
          autoFocus
          size={Math.min(10, Math.max(4, sheets.length + 1))}
          className="sheet-tab-select"
          value={before}
          onChange={(event) => setBefore(event.target.value)}
          onDoubleClick={ok}
        >
          {sheets.map((sheet) => (
            <option key={sheet.id} value={sheet.id}>
              {sheet.name}
            </option>
          ))}
          <option value={END}>{t('dlgSheetMoveToEnd')}</option>
        </select>
      </label>
      <label className="dialog-check">
        <input type="checkbox" checked={copy} onChange={(event) => setCopy(event.target.checked)} />
        {t('dlgSheetMoveCreateCopy')}
      </label>
      <p className="dialog-note">{t('dlgSheetMoveNote')}</p>
    </DialogFrame>
  )
}

function UnhideSheetDialog({
  sheets,
  onApply,
  onClose,
}: {
  readonly sheets: readonly SheetTabInfo[]
  readonly onApply: (ids: readonly string[]) => void
  readonly onClose: () => void
}): React.JSX.Element {
  const { t } = useI18n()
  const [picked, setPicked] = useState<readonly string[]>(sheets[0] ? [sheets[0].id] : [])
  const ok = (): void => {
    if (picked.length === 0) return
    onApply(picked)
    onClose()
  }
  return (
    <DialogFrame
      title={t('dlgSheetUnhideTitle')}
      onClose={onClose}
      onOk={ok}
      okDisabled={picked.length === 0}
    >
      <label className="dialog-span sheet-tab-label">
        {t('dlgSheetUnhideList')}
        <select
          autoFocus
          multiple
          size={Math.min(10, Math.max(4, sheets.length))}
          className="sheet-tab-select"
          value={picked as string[]}
          onChange={(event) =>
            setPicked(Array.from(event.target.selectedOptions, (option) => option.value))
          }
          onDoubleClick={ok}
        >
          {sheets.map((sheet) => (
            <option key={sheet.id} value={sheet.id}>
              {sheet.name}
            </option>
          ))}
        </select>
      </label>
    </DialogFrame>
  )
}
