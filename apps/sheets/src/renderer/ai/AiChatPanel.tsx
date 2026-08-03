import React, { useEffect, useRef, useState } from 'react'
import { AiComposer, AiTypingIndicator } from '@genoffice/ui'
import { GensparkMark } from '../ribbon-icons'
import type { ChangePlan } from '../../domain/workbook.types'
import type { AttachmentMeta } from '../../shared/desktop-api'
import { useI18n, type TFunc } from '../i18n/locale'
import { Markdown } from '@genoffice/ui'
import sendEnterOn from '../assets/send-enter-on.png'
import sendEnterOff from '../assets/send-enter-off.png'
import sendStop from '../assets/send-stop.png'
import attachIcon from '../assets/attach-icon.png'

/** Clipboard bitmap MIME → attachment extension (matches the main process's
 * ATTACHMENT_IMAGE_EXTS) */
const PASTE_MIME_EXT: Record<string, string> = {
  'image/png': 'png',
  'image/jpeg': 'jpg',
  'image/gif': 'gif',
  'image/webp': 'webp',
}

/** Resizable panel width: persisted, clamped to min/max, drives the .sheet-body grid column via --copilot-width */
const PANEL_WIDTH_KEY = 'sheets-ai-panel-width'
const PANEL_WIDTH_MIN = 280

function clampPanelWidth(w: number): number {
  return Math.min(Math.max(w, PANEL_WIDTH_MIN), Math.min(720, Math.round(window.innerWidth * 0.6)))
}

function loadPanelWidth(): number | null {
  const saved = Number(localStorage.getItem(PANEL_WIDTH_KEY))
  return Number.isFinite(saved) && saved > 0 ? clampPanelWidth(saved) : null
}

export interface AiToolChip {
  readonly summary: string
  readonly isError: boolean
  /** still executing: rendered as a spinner chip, replaced in place when the tool finishes */
  readonly running?: boolean
  /** Tool name (title tooltip) */
  readonly name?: string
  /** Tool output (truncated UI-side); when present the row expands to details */
  readonly output?: string
}

export interface AiChatMessage {
  readonly role: 'user' | 'assistant'
  readonly text: string
  readonly tools: readonly AiToolChip[]
  readonly streaming?: boolean | undefined
  readonly isError?: boolean | undefined
  /** the run failed and this user message was rolled back out of the model context */
  readonly undelivered?: boolean | undefined
  /** the run failed because Genspark is signed out — render an inline sign-in button */
  readonly loginRequired?: boolean | undefined
  /** Set when this message reflects an auto-applied plan; renders an inline [Undo] button. */
  readonly autoApplied?: { readonly opCount: number } | undefined
}

export function AiChatPanel({
  isOpen,
  hasContent,
  chat,
  historicChat = [],
  attachments,
  attachNotice,
  onPickAttachments,
  onAddAttachmentPaths,
  onAddPastedImage,
  onRemoveAttachment,
  prompt,
  preview,
  aiBusy,
  onPromptChange,
  onSend,
  onStop,
  onNewChat,
  onUndo,
  onExpand,
  onCollapse,
}: {
  readonly isOpen: boolean
  /** the workbook has cells with content — empty workbooks get "build me a sheet" copy instead */
  readonly hasContent: boolean
  readonly chat: readonly AiChatMessage[]
  readonly historicChat?: readonly AiChatMessage[]
  /// Chat attachments (chips + 📎 button + drag onto the panel), same structure
  /// as the docs/slides AI panels.
  readonly attachments: readonly AttachmentMeta[]
  readonly attachNotice: string | null
  readonly onPickAttachments: () => void
  readonly onAddAttachmentPaths: (paths: readonly string[]) => void
  /// Clipboard-pasted bitmaps (screenshots etc. without a local path): bytes +
  /// extension
  readonly onAddPastedImage: (data: ArrayBuffer, ext: string) => void
  readonly onRemoveAttachment: (path: string) => void
  readonly prompt: string
  readonly preview: ChangePlan | null
  readonly aiBusy: boolean
  readonly onPromptChange: (prompt: string) => void
  /** Send the composer text, or the given instruction when provided (used by the failed-run Retry) */
  readonly onSend: (instruction?: string) => void
  readonly onStop: () => void
  readonly onNewChat: () => void
  readonly onUndo: () => void
  readonly onExpand: () => void
  readonly onCollapse: () => void
}): React.JSX.Element {
  const { t } = useI18n()
  const chatRef = useRef<HTMLDivElement | null>(null)
  const inputRef = useRef<HTMLTextAreaElement | null>(null)
  const stickToBottomRef = useRef(true)
  const [dragOver, setDragOver] = useState(false)
  const asideRef = useRef<HTMLElement | null>(null)
  const [resizing, setResizing] = useState(false)
  /** Wall-clock start of the current run (aiBusy false→true), drives the elapsed badge */
  const busyStartRef = useRef(0)
  useEffect(() => {
    if (aiBusy) busyStartRef.current = Date.now()
  }, [aiBusy])

  // Restore the persisted panel width (the grid column tracks --copilot-width on .sheet-body)
  useEffect(() => {
    if (!isOpen) return
    const saved = loadPanelWidth()
    if (saved === null) return
    const area = asideRef.current?.closest('.sheet-body') as HTMLElement | null
    area?.style.setProperty('--copilot-width', `${saved}px`)
  }, [isOpen])

  // Re-clamp the panel width when the window shrinks (max is 60% of the window)
  useEffect(() => {
    if (!isOpen) return
    const onResize = (): void => {
      const area = asideRef.current?.closest('.sheet-body') as HTMLElement | null
      const current = parseFloat(area?.style.getPropertyValue('--copilot-width') ?? '')
      if (!area || !Number.isFinite(current)) return
      area.style.setProperty('--copilot-width', `${clampPanelWidth(current)}px`)
    }
    window.addEventListener('resize', onResize)
    return () => window.removeEventListener('resize', onResize)
  }, [isOpen])

  // follow the stream, but stop yanking once the user scrolls up to read
  useEffect(() => {
    if (stickToBottomRef.current) {
      chatRef.current?.scrollTo({ top: chatRef.current.scrollHeight })
    }
  }, [chat, preview])

  /** Drag the right edge to resize: the panel is flush with the window's left edge, so width = clientX; the grid transition is disabled while dragging */
  const startResize = (e: React.PointerEvent<HTMLDivElement>): void => {
    e.preventDefault()
    const area = asideRef.current?.closest('.sheet-body') as HTMLElement | null
    if (!area) return
    setResizing(true)
    area.style.transition = 'none'
    document.body.style.cursor = 'col-resize'
    document.body.style.userSelect = 'none'
    let width = 0
    const onMove = (ev: PointerEvent): void => {
      width = clampPanelWidth(ev.clientX)
      area.style.setProperty('--copilot-width', `${width}px`)
    }
    const onUp = (): void => {
      window.removeEventListener('pointermove', onMove)
      window.removeEventListener('pointerup', onUp)
      area.style.transition = ''
      document.body.style.cursor = ''
      document.body.style.userSelect = ''
      setResizing(false)
      if (width > 0) localStorage.setItem(PANEL_WIDTH_KEY, String(Math.round(width)))
    }
    window.addEventListener('pointermove', onMove)
    window.addEventListener('pointerup', onUp)
  }

  const onChatScroll = (): void => {
    const el = chatRef.current
    if (!el) return
    stickToBottomRef.current = el.scrollHeight - el.scrollTop - el.clientHeight < 48
  }

  if (!isOpen) {
    return (
      <aside className="copilot collapsed">
        <button className="expand-copilot" onClick={onExpand} title={t('aiOpenAssistant')}>
          <GensparkMark size={22} />
        </button>
      </aside>
    )
  }

  const canSend = prompt.trim().length > 0 && !aiBusy

  const send = (): void => {
    if (!canSend) return
    stickToBottomRef.current = true
    onSend()
  }

  const onDrop = (e: React.DragEvent): void => {
    e.preventDefault()
    e.stopPropagation()
    setDragOver(false)
    const paths = Array.from(e.dataTransfer.files)
      .map((f) => window.desktopApi.getPathForFile(f))
      .filter(Boolean)
    if (paths.length > 0) onAddAttachmentPaths(paths)
  }

  /** Files pasted into the input: ones with a local path go the regular
   * attachment route; pure bitmaps like screenshots are persisted by the host */
  const onPasteFiles = (files: File[]): void => {
    const paths: string[] = []
    for (const f of files) {
      const p = window.desktopApi.getPathForFile(f)
      if (p) {
        paths.push(p)
        continue
      }
      const ext = PASTE_MIME_EXT[f.type] ?? f.name.split('.').pop()?.toLowerCase() ?? 'bin'
      void f.arrayBuffer().then((buf) => onAddPastedImage(buf, ext))
    }
    if (paths.length > 0) onAddAttachmentPaths(paths)
  }

  return (
    <aside
      ref={asideRef}
      className={`copilot${dragOver ? ' ai-panel-dragover' : ''}${resizing ? ' ai-panel-resizing' : ''}`}
      onDragOver={(e) => {
        if (e.dataTransfer.types.includes('Files')) {
          e.preventDefault()
          e.stopPropagation()
          setDragOver(true)
        }
      }}
      onDragLeave={(e) => {
        if (!e.currentTarget.contains(e.relatedTarget as Node | null)) setDragOver(false)
      }}
      onDrop={onDrop}
    >
      <div
        className="ai-panel-resizer"
        onPointerDown={startResize}
        role="separator"
        aria-orientation="vertical"
        aria-label="Genspark"
      />
      <header className="ai-panel-header">
        <span className="ai-panel-title">
          <GensparkMark size={22} />
          Genspark
        </span>
        <div className="ai-panel-header-actions">
          {(chat.length > 0 || historicChat.length > 0) && (
            <button className="ai-header-btn" onClick={onNewChat} title={t('aiNewChat')}>
              <IconNewChat size={15} />
            </button>
          )}
          <button className="ai-header-btn" onClick={onCollapse} title={t('aiCollapsePanel')}>
            <IconCollapse size={15} />
          </button>
        </div>
      </header>

      <div className="ai-chat" ref={chatRef} onScroll={onChatScroll}>
        {/* Past conversation (read-only transcript), shown continuously with the current turn */}
        {historicChat.length > 0 && (
          <>
            {historicChat.map((entry, i) => (
              <div key={`h${i}`} className={`ai-msg ai-msg-${entry.role} ai-msg-historic`}>
                {entry.tools.length > 0 && <ToolChipList tools={entry.tools} />}
                {entry.text && <Markdown text={entry.text} />}
              </div>
            ))}
            <div className="ai-history-sep">{t('aiHistorySep')}</div>
          </>
        )}
        {chat.length === 0 && historicChat.length === 0 && (
          <div className="ai-chat-empty">
            <div className="ai-chat-empty-title">
              {t(hasContent ? 'aiEmptyTitle' : 'aiEmptyBuildTitle')}
            </div>
            <div className="ai-chat-empty-body">
              {t(hasContent ? 'aiEmptyBodyLine1' : 'aiEmptyBuildBody')}
            </div>
          </div>
        )}
        {chat.map((entry, index) => (
          <div
            key={index}
            className={`ai-msg ai-msg-${entry.role}${entry.isError ? ' ai-msg-error' : ''}${entry.role === 'assistant' && entry.streaming ? ' ai-msg-streaming' : ''}`}
          >
            {entry.role === 'user' ? (
              <>
                {entry.text}
                {entry.undelivered && (
                  <div className="ai-msg-undelivered">
                    {t('aiUndelivered')}
                    {!aiBusy && (
                      <button className="ai-retry-btn" onClick={() => onSend(entry.text)}>
                        {t('aiRetry')}
                      </button>
                    )}
                  </div>
                )}
              </>
            ) : (
              <>
                {entry.tools.length > 0 && <ToolChipList tools={entry.tools} />}
                {entry.text ? (
                  <Markdown text={entry.text} />
                ) : (
                  entry.streaming && (
                    <span className="ai-typing-row">
                      <AiTypingIndicator
                        label={entry.tools.length > 0 ? t('aiWorking') : t('aiThinking')}
                      />
                    </span>
                  )
                )}
                {entry.autoApplied && (
                  <div className="ai-auto-applied">
                    <span className="ai-auto-applied-text">
                      {t('aiAutoApplied', { count: entry.autoApplied.opCount })}
                    </span>
                    <button className="ai-undo-btn" onClick={onUndo} title={t('aiUndoTitle')}>
                      {t('aiUndo')}
                    </button>
                  </div>
                )}
                {entry.loginRequired && (
                  <button
                    className="ai-login-btn"
                    onClick={() => void window.desktopApi.aiGskLogin()}
                  >
                    {t('aiGskLoginBtn')}
                  </button>
                )}
              </>
            )}
          </div>
        ))}

        {preview && (
          <section className="preview ai-preview-card" aria-label={t('aiPreviewAria')}>
            <h3>{t('aiProposedChanges')}</h3>
            {preview.structuralChanges.map((change, index) => (
              <div className="change" key={`structural-${index}`}>
                <strong>{t('aiChangeStructure')}</strong>
                <span>{change.label}</span>
              </div>
            ))}
            {preview.formatChanges.map((change, index) => (
              <div className="change" key={`format-${index}`}>
                <strong>{t('aiChangeFormat')}</strong>
                <span>{change.label}</span>
              </div>
            ))}
            {preview.cellChanges.slice(0, MAX_PREVIEW_CELL_ROWS).map((change) => (
              <div className="change" key={`${change.sheetId}-${change.address}`}>
                <strong>{change.address}</strong>
                <span>
                  {formatCell(change.before, t)} → {formatCell(change.after, t)}
                </span>
              </div>
            ))}
            {preview.cellChanges.length > MAX_PREVIEW_CELL_ROWS && (
              <div className="change">
                <strong>…</strong>
                <span>
                  {t('aiMoreCells', { count: preview.cellChanges.length - MAX_PREVIEW_CELL_ROWS })}
                </span>
              </div>
            )}
            {preview.sheetRenames.map((rename) => (
              <div className="change" key={rename.sheetId}>
                <strong>{t('aiChangeSheet')}</strong>
                <span>
                  {rename.before} → {rename.after}
                </span>
              </div>
            ))}
            {preview.warnings.map((warning) => (
              <div className="change" key={warning}>
                <strong>⚠</strong>
                <span>{warning}</span>
              </div>
            ))}
          </section>
        )}
      </div>

      <div className="ai-composer">
        {attachments.length > 0 && (
          <div className="ai-attachments">
            {attachments.map((attachment) => (
              <span key={attachment.path} className="ai-attachment-chip" title={attachment.path}>
                <IconPaperclip size={11} />
                {attachment.name}
                <button
                  className="ai-attachment-remove"
                  onClick={() => onRemoveAttachment(attachment.path)}
                  title={t('aiRemoveAttachment')}
                >
                  ×
                </button>
              </span>
            ))}
          </div>
        )}
        {attachNotice && <div className="ai-attach-notice">{attachNotice}</div>}
        <AiComposer
          value={prompt}
          busy={aiBusy}
          placeholder={t(hasContent ? 'aiComposerPlaceholder' : 'aiComposerPlaceholderBuild')}
          hintIdle={t('aiHintIdle')}
          hintBusy={t('aiHintBusy')}
          hintIdleTitle={t('aiHintIdleTitle')}
          sendLabel={t('aiSend')}
          stopLabel={t('aiStop')}
          ariaLabel={t('aiInstructionAria')}
          iconOnly
          sendIconEnabled={<img src={sendEnterOn} alt="" aria-hidden />}
          sendIconDisabled={<img src={sendEnterOff} alt="" aria-hidden />}
          stopIcon={<img src={sendStop} alt="" aria-hidden />}
          footerStart={
            <button
              className="ai-attach-btn"
              onClick={onPickAttachments}
              title={t('aiAttachTitle')}
            >
              <img src={attachIcon} alt="" aria-hidden />
            </button>
          }
          textareaRef={inputRef}
          onChange={onPromptChange}
          onSend={send}
          onStop={onStop}
          onPasteFiles={onPasteFiles}
        />
      </div>
    </aside>
  )
}

const MAX_PREVIEW_CELL_ROWS = 50

function formatCell(
  cell: { readonly value: unknown; readonly formula?: string | undefined },
  t: TFunc,
): string {
  if (cell.formula) return cell.formula
  if (cell.value === null) return t('aiCellEmpty')
  return String(cell.value)
}

function Svg({ size, children }: { size: number; children: React.ReactNode }): React.JSX.Element {
  return (
    <svg
      width={size}
      height={size}
      viewBox="0 0 16 16"
      fill="none"
      stroke="currentColor"
      strokeWidth="1.2"
      strokeLinecap="round"
      aria-hidden
    >
      {children}
    </svg>
  )
}

function IconNewChat({ size }: { size: number }): React.JSX.Element {
  return (
    <Svg size={size}>
      <path
        d="M13.5 7.2v-3A1.7 1.7 0 0 0 11.8 2.5H4.2a1.7 1.7 0 0 0-1.7 1.7v6.1a1.7 1.7 0 0 0 1.7 1.7h1.1v2l2.6-2h1.3"
        strokeLinejoin="round"
      />
      <path d="M12.2 9.4v4M10.2 11.4h4" />
    </Svg>
  )
}

function IconCollapse({ size }: { size: number }): React.JSX.Element {
  // Mirrored glyph: the AI panel docks on the LEFT, so the divider and arrow point left
  return (
    <Svg size={size}>
      <rect x="1.5" y="2.5" width="13" height="11" rx="1" />
      <path d="M5.5 2.5v11" />
      <path d="M12.5 8H8.1M9.8 5.9 7.7 8l2.1 2.1" strokeWidth="1.3" strokeLinejoin="round" />
    </Svg>
  )
}

function IconPaperclip({ size }: { size: number }): React.JSX.Element {
  return (
    <Svg size={size}>
      <path
        d="M13 7.2 8.2 12a3.4 3.4 0 0 1-4.8-4.8l5-5a2.3 2.3 0 0 1 3.2 3.2l-5 5a1.1 1.1 0 0 1-1.6-1.6l4.6-4.6"
        strokeLinejoin="round"
      />
    </Svg>
  )
}

/** Tool row list (unified with docs/slides): dot + summary; rows with output
 * expand to details, with the arrow shown on hover */
/** Step-row status icons (timeline glyphs: 14px in a 20px slot, 1.6 stroke) */
function StepIcon({ status }: { status: 'running' | 'done' | 'error' }) {
  if (status === 'running') {
    return (
      <svg
        viewBox="0 0 24 24"
        width="14"
        height="14"
        fill="none"
        stroke="currentColor"
        strokeWidth="1.6"
        strokeLinecap="round"
        strokeLinejoin="round"
        aria-hidden
      >
        <path d="M6.5 3.5h11M6.5 20.5h11M8 3.5v3.2c0 2.6 4 4.2 4 5.3 0 1.1 4 2.7 4 5.3v3.2M16 3.5v3.2c0 2.6-4 4.2-4 5.3 0 1.1-4 2.7-4 5.3v3.2" />
      </svg>
    )
  }
  if (status === 'error') {
    return (
      <svg
        viewBox="0 0 24 24"
        width="14"
        height="14"
        fill="none"
        stroke="currentColor"
        strokeWidth="1.6"
        strokeLinecap="round"
        strokeLinejoin="round"
        aria-hidden
      >
        <circle cx="12" cy="12" r="9" />
        <path d="m9.2 9.2 5.6 5.6M14.8 9.2l-5.6 5.6" />
      </svg>
    )
  }
  return (
    <svg
      viewBox="0 0 24 24"
      width="14"
      height="14"
      fill="none"
      stroke="currentColor"
      strokeWidth="1.6"
      strokeLinecap="round"
      strokeLinejoin="round"
      aria-hidden
    >
      <circle cx="12" cy="12" r="9" />
      <path d="m8.5 12.4 2.4 2.4 4.6-5" />
    </svg>
  )
}

/** Tool activity group: a single quiet summary row
 *  that auto-opens while tools run, auto-collapses into "Worked · N steps" when they finish,
 *  and a manual toggle that always wins. Rows inside are step rows with 1px connectors. */
function ToolChipList({ tools }: { tools: readonly AiToolChip[] }): React.JSX.Element {
  const { t: tr } = useI18n()
  const [expanded, setExpanded] = useState<Set<number>>(new Set())
  const [userOpen, setUserOpen] = useState<boolean | null>(null)

  const toggle = (j: number) => {
    setExpanded((prev) => {
      const next = new Set(prev)
      if (next.has(j)) next.delete(j)
      else next.add(j)
      return next
    })
  }

  const anyRunning = tools.some((tool) => tool.running)
  const open = userOpen ?? anyRunning
  const label = anyRunning ? tr('aiGroupWorking') : tr('aiWorkedSteps', { n: tools.length })

  return (
    <div className="ai-work-group">
      <button
        type="button"
        className={`ai-work-group-summary${anyRunning ? ' running' : ''}`}
        aria-expanded={open}
        onClick={() => setUserOpen(!open)}
      >
        {anyRunning && !open && <span className="ai-tool-chip-spinner" aria-hidden />}
        <span className="ai-work-group-label">{label}</span>
        <span className={`ai-tool-chip-caret${open ? ' open' : ''}`} aria-hidden>
          ›
        </span>
      </button>
      <div className={`ai-work-group-body${open ? ' open' : ''}`}>
        <div className="ai-work-group-body-inner">
          {tools.map((tool, j) => {
            const hasOutput = !tool.running && !!tool.output
            const isOpen = expanded.has(j)
            const stepStatus = tool.running ? 'running' : tool.isError ? 'error' : 'done'
            return (
              <div key={j} className="ai-step-row">
                <span className={`ai-step-icon ${stepStatus}`} aria-hidden>
                  <StepIcon status={stepStatus} />
                </span>
                <div className="ai-step-content">
                  {hasOutput ? (
                    <button
                      type="button"
                      className="ai-step-title clickable"
                      title={tool.name}
                      aria-expanded={isOpen}
                      onClick={() => toggle(j)}
                    >
                      {tool.summary}
                    </button>
                  ) : (
                    <span className="ai-step-title" title={tool.name}>
                      {tool.summary}
                    </span>
                  )}
                  {hasOutput && isOpen && (
                    <div className="ai-step-detail">
                      <div className="ai-tool-output">
                        <div className="ai-tool-output-pre">{tool.output}</div>
                      </div>
                    </div>
                  )}
                </div>
              </div>
            )
          })}
        </div>
      </div>
    </div>
  )
}
