import { useEffect, useMemo, useRef, useState } from 'react'
import type { KeyboardEvent as ReactKeyboardEvent, MouseEvent as ReactMouseEvent } from 'react'
import type { Lang } from '@genoffice/i18n'
import type { HomeApi, RecentEntry, UserSkillEntry } from '../../shared/home-api'
import type { InstalledPlugin } from '../../shared/found-skill'
import { FileBadge } from './file-badge'
import { COMMAND_PALETTE_STRINGS, type CommandPaletteStrings } from './command-palette-strings'
import './command-palette.css'

/** recents pulled once on open; the palette filters this page client-side */
const RECENT_LOAD_LIMIT = 100
/** rows shown before the query narrows anything (recent files top 5 + the create actions) */
const RECENT_EMPTY_LIMIT = 5
/** overall row cap, applied after the fixed group order */
const RESULT_CAP = 50

type GroupId = 'create' | 'recent' | 'skills' | 'settings' | 'plugins'

/** one selectable palette row; `run` performs the bridge action */
interface PaletteRow {
  readonly id: string
  readonly group: GroupId
  readonly title: string
  /** secondary line, also searched (folder, description, app name + extension) */
  readonly detail: string
  readonly icon: { kind: 'file'; ext: string } | { kind: 'folder' | 'pen' | 'gear' | 'plug' }
  /** the file failed to stat; the row dims but still tries to open */
  readonly missing: boolean
  /** plugins: the chip naming what Enter does (Disable / Enable) */
  readonly chip: string
  readonly run: () => void
}

/**
 * The Ctrl/Cmd+K launcher: a centered top-of-screen overlay in the Spotlight /
 * PowerToys Run shape. The main process summons it (`home:palette-open`, after
 * focusing the window and switching to Home); the parent mounts this component
 * in response and unmounts it on close.
 *
 * Sources, all read through `window.aiOffice`: recent files, the create
 * actions, installed skills (with plugin components), the settings sections
 * the bridge can open, and installed plugins. Enter runs the row's action;
 * Esc / backdrop click / Enter all hand control back via `onClose`.
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
}: {
  readonly lang: Lang
  readonly onClose: () => void
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

  const rows = useMemo(
    () => buildRows(strings, query, recentsList, skills, plugins),
    [strings, query, recentsList, skills, plugins],
  )

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
        out.push({ id: row.group, label: groupLabel(strings, row.group), offset: 0, rows: [row] })
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

// ── row construction ──

function buildRows(
  strings: CommandPaletteStrings,
  query: string,
  recentsList: readonly RecentEntry[] | null,
  skills: readonly UserSkillEntry[] | null,
  plugins: readonly InstalledPlugin[] | null,
): PaletteRow[] {
  const createRows: PaletteRow[] = [
    commandRow(
      'doc',
      strings.cmdNewDoc,
      strings.hintNewDoc,
      'docx',
      () => void window.aiOffice.newDoc(),
    ),
    commandRow(
      'sheet',
      strings.cmdNewSheet,
      strings.hintNewSheet,
      'xlsx',
      () => void window.aiOffice.newSheet(),
    ),
    commandRow(
      'slide',
      strings.cmdNewSlide,
      strings.hintNewSlide,
      'pptx',
      () => void window.aiOffice.newSlide(),
    ),
    commandRow(
      'markdown',
      strings.cmdNewMarkdown,
      strings.hintNewMarkdown,
      'md',
      () => void window.aiOffice.newMarkdown(),
    ),
    commandRow(
      'html',
      strings.cmdNewHtml,
      strings.hintNewHtml,
      'html',
      () => void window.aiOffice.newHtml(),
    ),
    commandRow(
      'pdf',
      strings.cmdNewPdf,
      strings.hintNewPdf,
      'pdf',
      () => void window.aiOffice.newPdf(),
    ),
    commandRow(
      'browse',
      strings.cmdBrowse,
      strings.hintBrowse,
      'folder',
      () => void window.aiOffice.browse(),
    ),
  ]

  // the ask view keeps its attached skills in its own state with no bridge to
  // it from a self-contained component; the skills settings page is the one
  // skill surface the palette can reach
  const openSkillsSettings = () => void window.aiOffice.openSkillsSettings()
  const skillRows: PaletteRow[] = []
  const seenSkillIds = new Set<string>()
  for (const skill of skills ?? []) {
    if (seenSkillIds.has(skill.name)) continue
    seenSkillIds.add(skill.name)
    skillRows.push(
      plainRow(
        'skills',
        `skill:${skill.name}`,
        skill.name,
        skill.description,
        { kind: 'pen' },
        () => openSkillsSettings(),
      ),
    )
  }
  for (const plugin of plugins ?? []) {
    for (const component of [...plugin.components.skills, ...plugin.components.commands]) {
      const name = `${plugin.name}:${component}`
      if (seenSkillIds.has(name)) continue
      seenSkillIds.add(name)
      skillRows.push(
        plainRow('skills', `skill:${name}`, name, plugin.description, { kind: 'pen' }, () =>
          openSkillsSettings(),
        ),
      )
    }
  }

  const settingsRows: PaletteRow[] = [
    plainRow(
      'settings',
      'settings:ai-model',
      strings.setAiModel,
      strings.hintAiModel,
      { kind: 'gear' },
      () => void window.aiOffice.openAiModelSettings(),
    ),
    plainRow(
      'settings',
      'settings:skills',
      strings.setSkills,
      strings.hintSkills,
      { kind: 'gear' },
      () => openSkillsSettings(),
    ),
  ]

  // the plugins settings section has no renderer-invokable open path; the one
  // per-plugin action the bridge offers is the (reversible) enable toggle
  const pluginRows: PaletteRow[] = (plugins ?? []).map((plugin) =>
    plainRow(
      'plugins',
      `plugin:${plugin.name}`,
      plugin.name,
      plugin.description,
      { kind: 'plug' },
      () => void window.aiOffice.pluginsSetEnabled(plugin.name, !plugin.enabled),
      plugin.enabled ? strings.pluginDisable : strings.pluginEnable,
    ),
  )

  const recentRows: PaletteRow[] = (recentsList ?? []).map((entry) =>
    plainRow(
      'recent',
      `recent:${entry.path}`,
      entry.name,
      entry.missing ? strings.missingFile : folderOf(entry.path),
      { kind: 'file', ext: entry.ext },
      () => void window.aiOffice.openPath(entry.path),
      '',
      entry.missing,
    ),
  )

  const q = query.trim()
  if (!q) {
    return [...createRows, ...recentRows.slice(0, RECENT_EMPTY_LIMIT)]
  }
  return [
    ...rankRows(q, createRows),
    ...rankRows(q, recentRows),
    ...rankRows(q, skillRows),
    ...rankRows(q, settingsRows),
    ...rankRows(q, pluginRows),
  ].slice(0, RESULT_CAP)
}

function commandRow(
  key: string,
  title: string,
  detail: string,
  ext: string,
  run: () => void,
): PaletteRow {
  return {
    id: `create:${key}`,
    group: 'create',
    title,
    detail,
    icon: ext === 'folder' ? { kind: 'folder' } : { kind: 'file', ext },
    missing: false,
    chip: '',
    run,
  }
}

function plainRow(
  group: GroupId,
  id: string,
  title: string,
  detail: string,
  icon: PaletteRow['icon'],
  run: () => void,
  chip = '',
  missing = false,
): PaletteRow {
  return { id, group, title, detail, icon, missing, chip, run }
}

function groupLabel(strings: CommandPaletteStrings, group: GroupId): string {
  switch (group) {
    case 'create':
      return strings.groupCreate
    case 'recent':
      return strings.groupRecent
    case 'skills':
      return strings.groupSkills
    case 'settings':
      return strings.groupSettings
    case 'plugins':
      return strings.groupPlugins
  }
}

/** keep the rows whose title or detail matches, best first, source order on ties */
function rankRows(query: string, rows: readonly PaletteRow[]): PaletteRow[] {
  return rows
    .map((row, index) => ({ row, index, score: scoreMatch(query, row.title, row.detail) }))
    .filter((entry) => entry.score !== null)
    .sort((a, b) => b.score! - a.score! || a.index - b.index)
    .map((entry) => entry.row)
}

/**
 * Case-insensitive substring + subsequence score, higher = better; null means
 * "no match". A substring beats a subsequence, the title beats the detail line,
 * and an earlier hit beats a later one; equal scores keep the source order
 * (`rankRows` sorts stably by construction).
 */
export function scoreMatch(query: string, title: string, detail = ''): number | null {
  const q = query.trim().toLowerCase()
  if (!q) return 0
  const t = title.toLowerCase()
  const d = detail.toLowerCase()
  const inTitle = t.indexOf(q)
  if (inTitle === 0) return 1000
  if (inTitle > 0) return 900 - Math.min(inTitle, 100)
  const inDetail = d.indexOf(q)
  if (inDetail === 0) return 800
  if (inDetail > 0) return 700 - Math.min(inDetail, 100)
  const titleSpread = subsequenceSpread(q, t)
  if (titleSpread !== null) return 600 - Math.min(titleSpread, 200)
  const detailSpread = subsequenceSpread(q, d)
  if (detailSpread !== null) return 500 - Math.min(detailSpread, 200)
  return null
}

/** how far apart the query's characters sit in `text` when matched in order
 *  (0 = contiguous), or null when they do not all appear in order */
function subsequenceSpread(query: string, text: string): number | null {
  let from = 0
  let spread = 0
  for (const ch of query) {
    const at = text.indexOf(ch, from)
    if (at === -1) return null
    spread += at - from
    from = at + 1
  }
  return spread
}

/** the directory part of a path, both separators honored; '' at the root */
export function folderOf(path: string): string {
  const cut = Math.max(path.lastIndexOf('/'), path.lastIndexOf('\\'))
  return cut > 0 ? path.slice(0, cut) : ''
}

// ── icons ──

function RowIcon({ icon }: { icon: PaletteRow['icon'] }) {
  if (icon.kind === 'file') return <FileBadge ext={icon.ext} size={20} />
  if (icon.kind === 'folder') return <FolderIcon />
  if (icon.kind === 'pen') return <PenIcon />
  if (icon.kind === 'gear') return <GearIcon />
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
        d="M19.4 15a1.7 1.7 0 0 0 .3 1.8l.1.1a2 2 0 1 1-2.8 2.8l-.1-.1a1.7 1.7 0 0 0-1.8-.3 1.7 1.7 0 0 0-1 1.5V21a2 2 0 1 1-4 0v-.1a1.7 1.7 0 0 0-1.1-1.5 1.7 1.7 0 0 0-1.8.3l-.1.1a2 2 0 1 1-2.8-2.8l.1-.1a1.7 1.7 0 0 0 .3-1.8 1.7 1.7 0 0 0-1.5-1H3a2 2 0 1 1 0-4h.1a1.7 1.7 0 0 0 1.5-1.1 1.7 1.7 0 0 0-.3-1.8l-.1-.1a2 2 0 1 1 2.8-2.8l.1.1a1.7 1.7 0 0 0 1.8.3H9a1.7 1.7 0 0 0 1-1.5V3a2 2 0 1 1 4 0v.1a1.7 1.7 0 0 0 1 1.5 1.7 1.7 0 0 0 1.8-.3l.1-.1a2 2 0 1 1 2.8 2.8l-.1.1a1.7 1.7 0 0 0-.3 1.8V9a1.7 1.7 0 0 0 1.5 1H21a2 2 0 1 1 0 4h-.1a1.7 1.7 0 0 0-1.5 1z"
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
