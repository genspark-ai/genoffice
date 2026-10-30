import { useEffect, useMemo, useRef, useState } from 'react'
import type { KeyboardEvent as ReactKeyboardEvent, MouseEvent as ReactMouseEvent } from 'react'
import type { Lang } from '@genoffice/i18n'
import type { HomeApi, RecentEntry, UserSkillEntry } from '../../shared/home-api'
import type { InstalledPlugin } from '../../shared/found-skill'
import { FileBadge } from './file-badge'
import { COMMAND_PALETTE_STRINGS, type CommandPaletteStrings } from './command-palette-strings'
import {
  defaultPaletteProviders,
  mergePaletteRows,
  RECENT_LOAD_LIMIT,
  type GroupId,
  type PaletteContext,
  type PaletteProvider,
  type PaletteRow,
} from './palette-providers'
import './command-palette.css'

// the matching/scoring helpers live with the providers now; re-exported so
// existing imports (and their tests) keep working against this module
export { folderOf, scoreMatch } from './palette-providers'

const noopAskAi = (_query: string): void => {}

/**
 * The Ctrl/Cmd+K launcher: a centered top-of-screen overlay in the Spotlight /
 * PowerToys Run shape. The main process summons it (`home:palette-open`, after
 * focusing the window and switching to Home); the parent mounts this component
 * in response and unmounts it on close.
 *
 * The row sources form a provider registry (`palette-providers.ts`): the
 * create actions, recent files, the file index (debounced, async), skills,
 * settings sections and plugins each contribute, and an Ask AI row trails
 * every non-empty query as the natural-language fallback. Sync providers
 * publish on the keystroke; async rows fold in when they land. Enter runs the
 * row's action; Esc / backdrop click / Enter all hand control back via
 * `onClose`.
 *
 * Degraded picks, because the component contract is `lang` + `onClose` only:
 * attaching a skill to the AI ask view lives in Home's internals (the ask
 * panel's own state) with no bridge to reach it, so picking a skill or a plugin
 * component opens the skills settings page instead. The settings group lists
 * exactly the sections with a renderer-invokable open path (AI model, skills),
 * and picking a plugin toggles it on or off — both noted in the module docs.
 */
export function CommandPalette({
  lang,
  onClose,
  onAskAi = noopAskAi,
  providers,
}: {
  readonly lang: Lang
  readonly onClose: () => void
  /** the Ask AI fallback row's query goes here (Home routes it into the ask view) */
  readonly onAskAi?: (query: string) => void
  /** test seam: a custom registry; production uses the default providers */
  readonly providers?: readonly PaletteProvider[]
}) {
  const strings = COMMAND_PALETTE_STRINGS[lang]
  const [query, setQuery] = useState('')
  const [active, setActive] = useState(0)
  const [recentsList, setRecentsList] = useState<readonly RecentEntry[] | null>(null)
  const [skills, setSkills] = useState<readonly UserSkillEntry[] | null>(null)
  const [plugins, setPlugins] = useState<readonly InstalledPlugin[] | null>(null)
  const inputRef = useRef<HTMLInputElement | null>(null)
  const listRef = useRef<HTMLDivElement | null>(null)

  useEffect(() => {
    inputRef.current?.focus()
    let alive = true
    const load = <T,>(request: Promise<T> | undefined, fallback: T, set: (value: T) => void) => {
      if (!request) {
        set(fallback)
        return
      }
      request.then(
        (value) => {
          if (alive) set(value)
        },
        () => {
          if (alive) set(fallback)
        },
      )
    }
    // an app update can leave the previous preload in place for one launch:
    // absent methods read as "nothing of that kind" rather than crashing
    const api: HomeApi | undefined = window.aiOffice
    load(
      api?.recents?.({ limit: RECENT_LOAD_LIMIT }).then((page) => page.entries),
      [],
      setRecentsList,
    )
    load(api?.skillsList?.(), [], setSkills)
    load(api?.pluginsList?.(), [], setPlugins)
    return () => {
      alive = false
    }
  }, [])

  const providerList = useMemo(() => providers ?? defaultPaletteProviders(), [providers])
  const [rows, setRows] = useState<PaletteRow[]>([])
  const querySeqRef = useRef(0)
  useEffect(() => {
    let alive = true
    const seq = ++querySeqRef.current
    const ctx: PaletteContext = {
      strings,
      recents: recentsList,
      skills,
      plugins,
      onAskAi,
    }
    // sync providers publish on this keystroke; an async provider's slot waits
    // for its promise, and a stale keystroke's late answer is discarded
    const parts: { priority: number; rows: PaletteRow[] }[] = []
    const publish = () => {
      if (alive && querySeqRef.current === seq) setRows(mergePaletteRows(parts))
    }
    for (const provider of providerList) {
      let out: PaletteRow[] | Promise<PaletteRow[]>
      try {
        out = provider.getRows(query, ctx)
      } catch {
        continue
      }
      const part = { priority: provider.priority, rows: Array.isArray(out) ? out : [] }
      parts.push(part)
      if (!Array.isArray(out)) {
        void Promise.resolve(out).then(
          (resolved) => {
            part.rows = resolved
            publish()
          },
          () => undefined,
        )
      }
    }
    publish()
    return () => {
      alive = false
    }
  }, [providerList, strings, query, recentsList, skills, plugins, onAskAi])

  // a new result set starts from the top, whether it came from the query or a
  // source landing late
  useEffect(() => {
    setActive(0)
  }, [rows])

  useEffect(() => {
    const el = listRef.current?.querySelector(`[data-idx='${active}']`)
    if (el instanceof HTMLElement && typeof el.scrollIntoView === 'function') {
      el.scrollIntoView({ block: 'nearest' })
    }
  }, [active])

  const runRow = (row: PaletteRow) => {
    row.run()
    onClose()
  }

  const onKeyDown = (event: ReactKeyboardEvent<HTMLInputElement>) => {
    if (rows.length === 0 && event.key !== 'Escape') return
    if (event.key === 'Escape') {
      event.preventDefault()
      onClose()
    } else if (event.key === 'ArrowDown') {
      event.preventDefault()
      setActive((current) => (current + 1) % rows.length)
    } else if (event.key === 'ArrowUp') {
      event.preventDefault()
      setActive((current) => (current - 1 + rows.length) % rows.length)
    } else if (event.key === 'Enter') {
      event.preventDefault()
      const row = rows[active]
      if (row) runRow(row)
    }
  }

  const onBackdropDown = (event: ReactMouseEvent<HTMLElement>) => {
    if (event.target === event.currentTarget) onClose()
  }

  // the flat list is already in group order; split it back into runs so each
  // group renders one header, while row ids stay positions in the flat list
  const groups = useMemo(() => {
    const out: { id: GroupId; label: string; offset: number; rows: PaletteRow[] }[] = []
    for (const row of rows) {
      const current = out[out.length - 1]
      if (current && current.id === row.group) current.rows.push(row)
      else
        out.push({
          id: row.group,
          label: groupLabel(strings, row.group),
          offset: 0,
          rows: [row],
        })
    }
    let offset = 0
    for (const group of out) {
      group.offset = offset
      offset += group.rows.length
    }
    return out
  }, [rows, strings])

  return (
    <div className="cmd-overlay" onMouseDown={onBackdropDown}>
      <div className="cmd-palette" role="dialog" aria-modal="true" aria-label={strings.inputLabel}>
        <div className="cmd-input-row">
          <span className="cmd-search-icon" aria-hidden="true">
            <SearchIcon />
          </span>
          <input
            ref={inputRef}
            className="cmd-input"
            value={query}
            placeholder={strings.placeholder}
            aria-label={strings.inputLabel}
            role="combobox"
            aria-expanded="true"
            aria-controls="cmd-list"
            aria-autocomplete="list"
            aria-activedescendant={rows.length > 0 ? `cmd-opt-${active}` : undefined}
            spellCheck={false}
            autoComplete="off"
            onChange={(event) => setQuery(event.target.value)}
            onKeyDown={onKeyDown}
          />
          <span className="cmd-esc" aria-hidden="true">
            Esc
          </span>
        </div>
        <div className="cmd-divider" />
        {rows.length === 0 ? (
          <div className="cmd-empty">{strings.empty}</div>
        ) : (
          <div
            id="cmd-list"
            ref={listRef}
            className="cmd-list"
            role="listbox"
            aria-label={strings.resultsLabel}
          >
            {groups.map((group) => (
              <div key={group.id} role="group" aria-label={group.label}>
                <div className="cmd-group-title" aria-hidden="true">
                  {group.label}
                </div>
                {group.rows.map((row, index) => {
                  const flat = group.offset + index
                  return (
                    <button
                      key={row.id}
                      type="button"
                      id={`cmd-opt-${flat}`}
                      data-idx={flat}
                      role="option"
                      aria-selected={flat === active}
                      className={`cmd-row${flat === active ? ' active' : ''}${row.missing ? ' missing' : ''}`}
                      onMouseEnter={() => setActive(flat)}
                      onClick={() => runRow(row)}
                    >
                      <span className="cmd-row-icon" aria-hidden="true">
                        <RowIcon icon={row.icon} />
                      </span>
                      <span className="cmd-row-text">
                        <span className="cmd-row-title">{row.title}</span>
                        {row.detail && <span className="cmd-row-detail">{row.detail}</span>}
                      </span>
                      {row.chip && <span className="cmd-chip">{row.chip}</span>}
                    </button>
                  )
                })}
              </div>
            ))}
          </div>
        )}
        <div className="cmd-footer" aria-hidden="true">
          {strings.kbdHint}
        </div>
      </div>
    </div>
  )
}

function groupLabel(strings: CommandPaletteStrings, group: GroupId): string {
  switch (group) {
    case 'create':
      return strings.groupCreate
    case 'recent':
      return strings.groupRecent
    case 'files':
      return strings.groupFiles
    case 'skills':
      return strings.groupSkills
    case 'settings':
      return strings.groupSettings
    case 'plugins':
      return strings.groupPlugins
    case 'ask':
      return strings.groupAsk
  }
}

// ── icons ──

function RowIcon({ icon }: { icon: PaletteRow['icon'] }) {
  if (icon.kind === 'file') return <FileBadge ext={icon.ext} size={20} />
  if (icon.kind === 'folder') return <FolderIcon />
  if (icon.kind === 'pen') return <PenIcon />
  if (icon.kind === 'gear') return <GearIcon />
  if (icon.kind === 'sparkle') return <SparkleIcon />
  return <PlugIcon />
}

function SearchIcon() {
  return (
    <svg width="16" height="16" viewBox="0 0 16 16" fill="none" aria-hidden="true">
      <circle cx="7" cy="7" r="5" stroke="currentColor" strokeWidth="1.5" />
      <path d="M10.8 10.8L14 14" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" />
    </svg>
  )
}

function PenIcon() {
  return (
    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" aria-hidden="true">
      <path
        d="M12 20h9"
        stroke="currentColor"
        strokeWidth="2"
        strokeLinecap="round"
        strokeLinejoin="round"
      />
      <path
        d="M16.5 3.5a2.1 2.1 0 0 1 3 3L7 19l-4 1 1-4z"
        stroke="currentColor"
        strokeWidth="2"
        strokeLinecap="round"
        strokeLinejoin="round"
      />
    </svg>
  )
}

function GearIcon() {
  return (
    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" aria-hidden="true">
      <circle cx="12" cy="12" r="3" stroke="currentColor" strokeWidth="2" />
      <path
        d="M19.4 15a1.7 1.7 0 0 0 .3 1.8l.1.1a2 2 0 1 1-2.8 2.8l-.1-.1a1.7 1.7 0 0 0-1 1.5V21a2 2 0 1 1-4 0v-.1a1.7 1.7 0 0 0-1.1-1.5 1.7 1.7 0 0 0-1.8.3l-.1.1a2 2 0 1 1-2.8-2.8l.1-.1a1.7 1.7 0 0 0 .3-1.8 1.7 1.7 0 0 0-1.5-1H3a2 2 0 1 1 0-4h.1a1.7 1.7 0 0 0 1.5-1.1 1.7 1.7 0 0 0-.3-1.8l-.1-.1a2 2 0 1 1 2.8-2.8l.1.1a1.7 1.7 0 0 0 1.8.3H9a1.7 1.7 0 0 0 1-1.5V3a2 2 0 1 1 4 0v.1a1.7 1.7 0 0 0 1 1.5 1.7 1.7 0 0 0 1.8-.3l.1-.1a2 2 0 1 1 2.8 2.8l-.1.1a1.7 1.7 0 0 0-.3-1.8V9a1.7 1.7 0 0 0 1.5 1H21a2 2 0 1 1 0 4h-.1a1.7 1.7 0 0 0-1.5-1z"
        stroke="currentColor"
        strokeWidth="2"
        strokeLinecap="round"
        strokeLinejoin="round"
      />
    </svg>
  )
}

function PlugIcon() {
  return (
    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" aria-hidden="true">
      <path
        d="M9 2v6M15 2v6M6 8h12v4a6 6 0 0 1-6 6 6 6 0 0 1-6-6V8zM12 18v4"
        stroke="currentColor"
        strokeWidth="2"
        strokeLinecap="round"
        strokeLinejoin="round"
      />
    </svg>
  )
}

function SparkleIcon() {
  return (
    <svg width="16" height="16" viewBox="0 0 16 16" fill="none" aria-hidden="true">
      <path
        d="M8 1.8l1.55 4.65L14.2 8l-4.65 1.55L8 14.2 6.45 9.55 1.8 8l4.65-1.55z"
        stroke="currentColor"
        strokeWidth="1.3"
        strokeLinejoin="round"
      />
    </svg>
  )
}

function FolderIcon() {
  return (
    <svg width="18" height="18" viewBox="0 0 24 24" fill="none" aria-hidden="true">
      <path
        d="M3 7a2 2 0 0 1 2-2h4l2 3h8a2 2 0 0 1 2 2v8a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V7z"
        stroke="currentColor"
        strokeWidth="1.8"
        strokeLinejoin="round"
      />
    </svg>
  )
}
