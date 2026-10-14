import { useEffect, useId, useRef, useState } from 'react'
import {
  slotsAt,
  viewHost,
  type FeatureRecord,
  type SlotAction,
  type SlotNode,
  type SlotPlace,
} from '@genoffice/agent-core'
import { Markdown } from '../Markdown'
import './mods.css'

/**
 * Renders the slots the registry holds.
 *
 * A slot is data: the host decides what a text, a button or a list looks like,
 * so theme, dark mode, text size and translations stay with the shell. A press
 * runs a declared action — open a file, open a feature's own panel, or hand a
 * prompt to the assistant. Nothing here can touch a document.
 */

/**
 * What a drawn element may do, supplied by the app that hosts the slots. The
 * package owns the rendering; the app owns the bridge, so one renderer serves the
 * shell and every editor without either duplicating the other.
 */
export interface SlotActions {
  /** absent when the host cannot open a file by path: an open action then draws disabled */
  open?(path: string): void
  ask(prompt: string): void
  /** absent when the host has no panel container for a feature's own view */
  panel?(id: string): void
  /** rows for a file-list; absent when the host has no file bridge */
  listSource?(source: string, limit: number): Promise<{ path: string; name: string }[]>
  /** mount a real view over a rectangle; absent when the host has no native view */
  mountView?(request: {
    id: string
    url: string
    rect: { x: number; y: number; width: number; height: number }
  }): void
  unmountView?(id: string): void
}

/** Strings the host supplies, so this package owns no locale. */
export interface SlotLabels {
  close: string
}

/**
 * Whether this host can run a declared action. An app that cannot open a file by
 * path draws the control disabled instead of a button that does nothing.
 */
export function canRunAction(action: SlotAction, actions: SlotActions): boolean {
  switch (action.kind) {
    case 'open':
      return typeof actions.open === 'function'
    case 'ask-assistant':
      return typeof actions.ask === 'function'
    case 'panel':
      return typeof actions.panel === 'function'
    case 'run-feature':
      return typeof actions.ask === 'function'
  }
}

/** Run a declared action. Exported so the command palette runs one the same way. */
export function runSlotAction(action: SlotAction, actions: SlotActions): void {
  switch (action.kind) {
    case 'open':
      actions.open?.(action.path)
      break
    case 'ask-assistant':
      actions.ask(action.prompt)
      break
    case 'panel':
      actions.panel?.(action.id)
      break
    case 'run-feature':
      // no kind executes yet, so the honest reading is "ask the assistant about it"
      actions.ask('Use the feature "' + action.id + '".')
      break
  }
}

/** One row a host source hands the slot. */
interface SourceRow {
  path: string
  name: string
  folder: string
}

function folderOf(path: string): string {
  const at = Math.max(path.lastIndexOf('/'), path.lastIndexOf('\\'))
  if (at <= 0) return ''
  const parent = path.slice(0, at)
  const cut = Math.max(parent.lastIndexOf('/'), parent.lastIndexOf('\\'))
  return parent.slice(cut + 1)
}

/** Rows from a named host source; null while the first read is in flight. */
function useSourceRows(
  source: string,
  limit: number,
  listSource: SlotActions['listSource'],
): SourceRow[] | null {
  const [rows, setRows] = useState<SourceRow[] | null>(null)
  useEffect(() => {
    let live = true
    if (!listSource) {
      setRows([])
      return
    }
    void listSource(source, limit)
      .then((entries) => {
        if (!live) return
        setRows(
          entries.map((entry) => ({
            path: entry.path,
            name: entry.name,
            folder: folderOf(entry.path),
          })),
        )
      })
      .catch(() => {
        if (live) setRows([])
      })
    return () => {
      live = false
    }
  }, [source, limit])
  return rows
}

/** Substitute the three allowed fields into an action's text. */
function fill(text: string, row: SourceRow): string {
  return text.replace(/\$(path|name|folder)/g, (_whole, key: string) => {
    if (key === 'path') return row.path
    if (key === 'name') return row.name
    return row.folder
  })
}

function fillAction(action: SlotAction, row: SourceRow): SlotAction {
  switch (action.kind) {
    case 'open':
      return { kind: 'open', path: fill(action.path, row) }
    case 'ask-assistant':
      return { kind: 'ask-assistant', prompt: fill(action.prompt, row) }
    case 'panel':
      return { kind: 'panel', id: fill(action.id, row) }
    case 'run-feature':
      return { kind: 'run-feature', id: fill(action.id, row) }
  }
}

/** A list the host fills: the slot names a source, the host owns the data. */
function FileListNode({
  node,
  actions,
}: {
  node: Extract<SlotNode, { type: 'file-list' }>
  actions: SlotActions
}): React.JSX.Element | null {
  const limit = node.limit ?? 8
  const rows = useSourceRows(node.source, limit, actions.listSource)
  if (rows === null) return <p className="slot-text slot-text--caption slot-text--muted">…</p>
  if (!rows.length) {
    return node.emptyText ? (
      <p className="slot-text slot-text--caption slot-text--muted">{node.emptyText}</p>
    ) : null
  }
  return (
    <ul className="slot-list">
      {rows.map((row) => (
        <li key={row.path}>
          <button
            className="slot-list-item"
            disabled={!node.action || !canRunAction(node.action, actions)}
            onClick={() => {
              if (node.action) runSlotAction(fillAction(node.action, row), actions)
            }}
          >
            <span className="slot-text">{row.name}</span>
            {row.folder ? (
              <span className="slot-text slot-text--caption slot-text--muted">{row.folder}</span>
            ) : null}
          </button>
        </li>
      ))}
    </ul>
  )
}

/**
 * A real view the host mounts over this rectangle (K5).
 *
 * The rectangle is the body of the frame, not the whole card, so the host's own
 * frame — the address and the site's name — stays visible and consistent with
 * the rest of the shell. The renderer reports where; the shell decides whether
 * the address is allowed and owns the view.
 */
function ViewSlot({
  node,
  actions,
}: {
  node: Extract<SlotNode, { type: 'view' }>
  actions: SlotActions
}): React.JSX.Element {
  const id = useId()
  const bodyRef = useRef<HTMLDivElement>(null)
  const [live, setLive] = useState(false)

  useEffect(() => {
    const body = bodyRef.current
    const mountView = actions.mountView
    if (!body || !mountView) return
    let mounted = false
    const place = (): void => {
      const rect = body.getBoundingClientRect()
      mounted = true
      setLive(true)
      mountView({
        id,
        url: node.url,
        rect: { x: rect.x, y: rect.y, width: rect.width, height: rect.height },
      })
    }
    place()
    // the panel moves when the window, the sidebar or a scroll changes: keep the
    // view over the same rectangle instead of letting it drift
    const observer = new ResizeObserver(place)
    observer.observe(body)
    window.addEventListener('resize', place)
    window.addEventListener('scroll', place, true)
    return () => {
      observer.disconnect()
      window.removeEventListener('resize', place)
      window.removeEventListener('scroll', place, true)
      if (mounted) actions.unmountView?.(id)
    }
  }, [id, node.url, actions])

  return (
    <section
      className="slot-view"
      style={node.minHeight ? { minHeight: node.minHeight } : undefined}
    >
      <header className="slot-view-bar">
        <span className="slot-view-dot" aria-hidden="true" />
        <span className="slot-view-title">{node.title ?? viewHost(node.url)}</span>
        <span className="slot-view-url">{viewHost(node.url)}</span>
      </header>
      <div className="slot-view-body" ref={bodyRef}>
        {live ? null : <span className="slot-view-placeholder" aria-hidden="true" />}
      </div>
    </section>
  )
}

function toneClass(tone: string | undefined): string {
  return tone && tone !== 'default' ? ' slot-text--' + tone : ''
}

function sizeClass(size: string | undefined): string {
  return size === 'caption' ? ' slot-text--caption' : ''
}

function InputNode({
  node,
  actions,
}: {
  node: Extract<SlotNode, { type: 'input' }>
  actions: SlotActions
}): React.JSX.Element {
  const [value, setValue] = useState('')
  const submit = (): void => {
    const action = node.action
    if (action.kind === 'ask-assistant') {
      const prompt = (action.prompt + ' ' + value).trim()
      actions.ask(prompt)
    } else {
      runSlotAction(action, actions)
    }
    setValue('')
  }
  return (
    <div className="slot-input">
      <input
        className="slot-input-field"
        value={value}
        placeholder={node.placeholder ?? ''}
        aria-label={node.name ?? node.placeholder ?? ''}
        onChange={(event) => setValue(event.target.value)}
        onKeyDown={(event) => {
          if (event.key === 'Enter') {
            event.preventDefault()
            submit()
          }
        }}
      />
      <button
        className="slot-button slot-button--primary"
        disabled={!canRunAction(node.action, actions)}
        onClick={submit}
      >
        {node.submitLabel ?? 'Go'}
      </button>
    </div>
  )
}

function NodeView({
  node,
  actions,
}: {
  node: SlotNode
  actions: SlotActions
}): React.JSX.Element | null {
  switch (node.type) {
    case 'text':
      return (
        <p className={'slot-text' + toneClass(node.tone) + sizeClass(node.size)}>{node.value}</p>
      )
    case 'markdown':
      return (
        <div className="slot-text slot-markdown">
          <Markdown text={node.value} />
        </div>
      )
    case 'view':
      return <ViewSlot node={node} actions={actions} />
    case 'divider':
      return <hr className="slot-divider" />
    case 'group':
      return (
        <div className="slot-group">
          {node.children.map((child, index) => (
            <NodeView key={index} node={child} actions={actions} />
          ))}
        </div>
      )
    case 'button':
      return (
        <button
          className={'slot-button' + (node.tone === 'primary' ? ' slot-button--primary' : '')}
          disabled={!canRunAction(node.action, actions)}
          onClick={() => runSlotAction(node.action, actions)}
        >
          {node.label}
        </button>
      )
    case 'input':
      return <InputNode node={node} actions={actions} />
    case 'select':
      return (
        <label className="slot-select">
          {node.name ? <span className="slot-text slot-text--caption">{node.name}</span> : null}
          <select
            className="slot-select-field"
            defaultValue=""
            onChange={(event) => {
              if (!node.action || !event.target.value) return
              const item = node.options.find(
                (option) => (option.value ?? option.label) === event.target.value,
              )
              const action = node.action
              if (action.kind === 'ask-assistant') {
                actions.ask((action.prompt + ' ' + (item?.label ?? '')).trim())
              } else {
                runSlotAction(action, actions)
              }
              event.target.value = ''
            }}
          >
            <option value="" disabled hidden />
            {node.options.map((option, index) => (
              <option key={index} value={option.value ?? option.label}>
                {option.label}
              </option>
            ))}
          </select>
        </label>
      )
    case 'file-list':
      return <FileListNode node={node} actions={actions} />
    case 'list':
      return (
        <ul className="slot-list">
          {node.items.map((item, index) => (
            <li key={index}>
              <button
                className="slot-list-item"
                disabled={!node.action || !canRunAction(node.action, actions)}
                onClick={() => {
                  if (!node.action) return
                  const action = node.action
                  if (action.kind === 'ask-assistant') {
                    actions.ask((action.prompt + ' ' + item.label).trim())
                  } else {
                    runSlotAction(action, actions)
                  }
                }}
              >
                <span className="slot-text">{item.label}</span>
                {item.description ? (
                  <span className="slot-text slot-text--caption slot-text--muted">
                    {item.description}
                  </span>
                ) : null}
              </button>
            </li>
          ))}
        </ul>
      )
    default:
      return null
  }
}

/** Every slot registered for one place, oldest first. Renders nothing when there is none. */
export function SlotHost({
  place,
  features,
  actions,
}: {
  place: SlotPlace
  features: FeatureRecord[]
  actions: SlotActions
}): React.JSX.Element | null {
  const slots = slotsAt(features, place)
  if (!slots.length) return null
  return (
    <>
      {slots.map((feature) => (
        <section key={feature.id} className={'slot-place slot-place--' + place.replace('.', '-')}>
          <NodeView node={feature.slot!.element} actions={actions} />
        </section>
      ))}
    </>
  )
}

/** One feature's home.panel, shown full width in the content area. */
export function SlotPanel({
  featureId,
  features,
  actions,
  labels,
  onClose,
}: {
  featureId: string
  features: FeatureRecord[]
  actions: SlotActions
  labels: SlotLabels
  onClose: () => void
}): React.JSX.Element {
  const feature: FeatureRecord | undefined = features.find((one) => one.id === featureId)
  return (
    <div className="slot-panel">
      <header className="slot-panel-head">
        <div className="slot-panel-titles">
          <h1 className="slot-panel-title">{feature?.name ?? featureId}</h1>
          {feature?.description ? <p className="slot-panel-sub">{feature.description}</p> : null}
        </div>
        <button
          className="slot-close"
          onClick={onClose}
          title={labels.close}
          aria-label={labels.close}
        >
          <svg width="16" height="16" viewBox="0 0 16 16" fill="none" aria-hidden="true">
            <path
              d="M4 4l8 8M12 4l-8 8"
              stroke="currentColor"
              strokeWidth="1.5"
              strokeLinecap="round"
            />
          </svg>
        </button>
      </header>
      <div className="slot-panel-body">
        {feature?.slot ? (
          <NodeView node={feature.slot.element} actions={actions} />
        ) : (
          <p className="slot-text slot-text--muted">
            This feature is not in the registry any more.
          </p>
        )}
      </div>
    </div>
  )
}
