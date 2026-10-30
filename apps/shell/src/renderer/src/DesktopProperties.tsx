import { useEffect, useState } from 'react'
import type { ReactElement } from 'react'
import type { Lang } from '@genoffice/i18n'
import type { HomeApi, RecentEntry } from '../../shared/home-api'
import type { DesktopStrings } from './desktop-strings'

/**
 * The desktop's two properties dialogs, both opened off the desktop:
 * the per-icon one a file or folder icon's own menu carries (name, type,
 * size and modified time for files, the path row and the reveal button), and
 * the desktop's own off the empty-surface menu (the file and folder counts,
 * the default save location, and the settings shortcut).
 *
 * Each dialog loads its own answer off the bridge while it is open — one
 * statPaths call per icon target, one getDefaultSaveDir per open — and falls
 * back to the label the icon already carries until it lands. The view keeps
 * the open state and the icon set, and hands the rest in through the props.
 * The words come from desktop-strings, the look from desktop.css.
 */

/** a file size the way the OS writes it: B, then KB/MB/GB with one decimal */
function formatBytes(bytes: number): string {
  if (!Number.isFinite(bytes) || bytes < 1024) return `${Math.max(0, Math.round(bytes))} B`
  const units = ['KB', 'MB', 'GB', 'TB']
  let value = bytes
  let unit = -1
  do {
    value /= 1024
    unit += 1
  } while (value >= 1024 && unit < units.length - 1)
  const rounded =
    value >= 100 || unit === units.length - 1 ? Math.round(value) : Math.round(value * 10) / 10
  return `${rounded} ${units[unit]}`
}

/** the properties dialog's modified time, localized for the UI language */
function formatDateTime(ms: number, lang: Lang): string {
  const date = new Date(ms)
  try {
    return date.toLocaleString(lang)
  } catch {
    return date.toISOString()
  }
}

/** the icon whose properties dialog is open (file and folder icons only) */
export interface DesktopPropsTarget {
  /** which of the two dialog shapes: a folder has no size or modified rows */
  readonly kind: 'file' | 'folder'
  /** the label the dialog falls back to before statPaths answers */
  readonly label: string
  /** the extension the type row shows when no stat has landed */
  readonly ext: string
  /** the real path: the reveal button's target, the path row's value */
  readonly path: string
}

export interface IconPropsDialogProps {
  /** the icon the dialog is about */
  readonly target: DesktopPropsTarget
  /** the UI language, for the modified time's locale */
  readonly lang: Lang
  /** the desktop's words, already resolved for the UI language */
  readonly strings: DesktopStrings
  /** close the dialog: the scrim, the close button, or Escape */
  readonly onClose: () => void
}

export function IconPropsDialog({
  target,
  lang,
  strings,
  onClose,
}: IconPropsDialogProps): ReactElement {
  const s = strings
  /** statPaths' answer for the dialog; null while it loads */
  const [stat, setStat] = useState<RecentEntry | null>(null)

  // the dialog's stats: one statPaths call per open, for file icons only
  // (folders have no size or mtime worth reporting)
  useEffect(() => {
    if (target.kind !== 'file') return
    let alive = true
    const request = (window.aiOffice as Partial<HomeApi>).statPaths?.([target.path])
    if (!request) return
    void request.then(
      (entries) => {
        if (alive && entries[0] !== undefined) setStat(entries[0])
      },
      () => undefined,
    )
    return () => {
      alive = false
    }
  }, [target])

  return (
    <div
      className="set-overlay"
      onMouseDown={(event) => {
        if (event.target === event.currentTarget) onClose()
      }}
    >
      <div
        className="set-dialog desktop-props-dialog"
        role="dialog"
        aria-modal="true"
        aria-label={s.menuProps}
      >
        <div className="set-header">
          <h2 className="set-title">{stat?.name ?? target.label}</h2>
          <button className="set-close" aria-label={s.wpClose} onClick={() => onClose()}>
            <svg width="14" height="14" viewBox="0 0 14 14" aria-hidden="true">
              <path
                d="M2 2l10 10M12 2L2 12"
                stroke="currentColor"
                strokeWidth="1.5"
                strokeLinecap="round"
              />
            </svg>
          </button>
        </div>
        <div className="desktop-props-body">
          <dl className="desktop-props-rows">
            <div className="desktop-props-row">
              <dt>{s.propsName}</dt>
              <dd>{stat?.name ?? target.label}</dd>
            </div>
            <div className="desktop-props-row">
              <dt>{s.propsType}</dt>
              <dd>
                {target.kind === 'folder'
                  ? s.propsFolderType
                  : (stat?.ext ?? target.ext).toUpperCase() || '—'}
              </dd>
            </div>
            {target.kind === 'file' && (
              <>
                <div className="desktop-props-row">
                  <dt>{s.propsSize}</dt>
                  <dd>{stat !== null ? formatBytes(stat.sizeBytes) : '…'}</dd>
                </div>
                <div className="desktop-props-row">
                  <dt>{s.propsModified}</dt>
                  <dd>{stat !== null ? formatDateTime(stat.mtimeMs, lang) : '…'}</dd>
                </div>
              </>
            )}
            <div className="desktop-props-row">
              <dt>{s.propsPath}</dt>
              <dd>
                <input
                  className="desktop-props-path"
                  readOnly
                  value={target.path}
                  onFocus={(event) => event.currentTarget.select()}
                />
              </dd>
            </div>
          </dl>
          <div className="desktop-props-actions">
            <button
              type="button"
              className="desktop-props-reveal"
              onClick={() => void (window.aiOffice as Partial<HomeApi>).revealPath?.(target.path)}
            >
              {s.menuReveal}
            </button>
          </div>
        </div>
      </div>
    </div>
  )
}

export interface DesktopPropsDialogProps {
  /** the desktop's words, already resolved for the UI language */
  readonly strings: DesktopStrings
  /** how many file icons the desktop is currently showing */
  readonly fileCount: number
  /** how many folder icons the desktop is currently showing */
  readonly folderCount: number
  /** close the dialog: the scrim, the close button, or Escape */
  readonly onClose: () => void
  /** open Home's settings modal; the view picks the section it lands on */
  readonly onOpenSettings: () => void
}

export function DesktopPropsDialog({
  strings,
  fileCount,
  folderCount,
  onClose,
  onOpenSettings,
}: DesktopPropsDialogProps): ReactElement {
  const s = strings
  /** getDefaultSaveDir's answer for the dialog; null while it loads */
  const [saveDir, setSaveDir] = useState<string | null>(null)

  // the dialog's save folder: one call per open
  useEffect(() => {
    const request = (window.aiOffice as Partial<HomeApi>).getDefaultSaveDir?.()
    if (!request) {
      setSaveDir('')
      return
    }
    void request.then(
      (dir) => setSaveDir(dir),
      () => setSaveDir(''),
    )
  }, [])

  return (
    <div
      className="set-overlay"
      onMouseDown={(event) => {
        if (event.target === event.currentTarget) onClose()
      }}
    >
      <div
        className="set-dialog desktop-props-dialog"
        role="dialog"
        aria-modal="true"
        aria-label={s.menuProps}
      >
        <div className="set-header">
          <h2 className="set-title">{s.navDesktop}</h2>
          <button className="set-close" aria-label={s.wpClose} onClick={() => onClose()}>
            <svg width="14" height="14" viewBox="0 0 14 14" aria-hidden="true">
              <path
                d="M2 2l10 10M12 2L2 12"
                stroke="currentColor"
                strokeWidth="1.5"
                strokeLinecap="round"
              />
            </svg>
          </button>
        </div>
        <div className="desktop-props-body">
          <dl className="desktop-props-rows">
            <div className="desktop-props-row">
              <dt>{s.dpFiles}</dt>
              <dd>{fileCount}</dd>
            </div>
            <div className="desktop-props-row">
              <dt>{s.dpFolders}</dt>
              <dd>{folderCount}</dd>
            </div>
            <div className="desktop-props-row">
              <dt>{s.dpSaveDir}</dt>
              <dd>
                {saveDir !== null && saveDir !== '' ? (
                  <input
                    className="desktop-props-path"
                    readOnly
                    value={saveDir}
                    onFocus={(event) => event.currentTarget.select()}
                  />
                ) : (
                  (saveDir ?? '…')
                )}
              </dd>
            </div>
          </dl>
          <div className="desktop-props-actions">
            <button
              type="button"
              className="desktop-props-open-settings"
              onClick={() => {
                onClose()
                onOpenSettings()
              }}
            >
              {s.dpOpenSettings}
            </button>
          </div>
        </div>
      </div>
    </div>
  )
}
