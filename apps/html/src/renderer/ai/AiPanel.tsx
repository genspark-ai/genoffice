import { useEffect, useRef, useState } from 'react'
import type { PointerEvent as ReactPointerEvent, ReactElement, ReactNode } from 'react'
import { AgentLoop, composeSkills } from '@genoffice/agent-core'
import type { AgentImage } from '@genoffice/agent-core'
import type { AiSettings } from '@genoffice/ai-provider'
import { ATTACHMENT_IMAGE_EXTS } from '../../shared/ipc'
import type { AttachmentAddResult, AttachmentMeta } from '../../shared/ipc'
import {
  AiComposer,
  AiScopeQuote,
  AiTypingIndicator,
  Markdown,
  type AiScopeQuoteData,
} from '@genoffice/ui'
import { aiLangDirective, t as tGlobal, useI18n } from '../i18n/locale'
import sendEnterOn from '../assets/send-enter-on.png'
import sendEnterOff from '../assets/send-enter-off.png'
import sendStop from '../assets/send-stop.png'
import attachIcon from '../assets/attach-icon.png'
import fileDocumentIcon from '../assets/file-document.png'
import fileExcelIcon from '../assets/file-excel.png'
import fileGeneralIcon from '../assets/file-general.png'
import fileImageIcon from '../assets/file-image.png'
import filePdfIcon from '../assets/file-pdf.png'
import filePptIcon from '../assets/file-ppt.png'
import fileWordIcon from '../assets/file-word.png'
import { createDocumentSkill } from './html-skill'
import { createFilesSkill } from './files-skill'
import { createIntentSkill, type PageIntent } from './intent'
import {
  buildPageWriterRequest,
  streamPage,
  type PageWriteResult,
  type PageWriteSpec,
} from './page-writer'
import type { BriefDecision, ClarifyQuestion, HtmlDocAccess } from './tools'
import type { Brief } from '../document/brief'
import { isDocEmpty } from '../document/blank'
import { ClarifyCard } from '../components/ClarifyCard'
import { BriefCard } from '../components/BriefCard'
import { createSearchSkill } from './search-skill'
import { createElectronTransport } from './transport'
import { DOC_NAV_SCHEME, parseDocNavHref } from './doc-nav'
import { EditQueueCard } from './EditQueueCard'
import {
  buildQueueInstruction,
  buildQueueSummary,
  liveItems,
  resolveQueue,
  type EditQueueItem,
} from './edit-queue'

/** [chip label, composer prefill]: blank page → design something new; page with content → rework it */
const GENERATE_STARTERS = [
  ['aiStarterLanding', 'aiStarterLandingPrompt'],
  ['aiStarterReport', 'aiStarterReportPrompt'],
  ['aiStarterPoster', 'aiStarterPosterPrompt'],
  ['aiStarterWorkspace', 'aiStarterWorkspacePrompt'],
] as const
const WRITE_STARTERS = [
  ['aiStarterArticle', 'aiStarterArticlePrompt'],
  ['aiStarterAnnouncement', 'aiStarterAnnouncementPrompt'],
  ['aiStarterGuide', 'aiStarterGuidePrompt'],
] as const
const EDIT_STARTERS = [
  ['aiStarterExtractBrief', 'aiStarterExtractBriefPrompt'],
  ['aiStarterRecolor', 'aiStarterRecolorPrompt'],
  ['aiStarterUnify', 'aiStarterUnifyPrompt'],
] as const

/** clipboard bitmap MIME → attachment extension (matches ATTACHMENT_IMAGE_EXTS) */
const PASTE_MIME_EXT: Record<string, string> = {
  'image/png': 'png',
  'image/jpeg': 'jpg',
  'image/gif': 'gif',
  'image/webp': 'webp',
}
const ATTACHMENT_CARD_ICONS: Record<string, string> = Object.fromEntries(
  (
    [
      [fileWordIcon, ['doc', 'docx']],
      [fileExcelIcon, ['xls', 'xlsx', 'xlsm', 'csv', 'tsv']],
      [filePptIcon, ['ppt', 'pptx']],
      [filePdfIcon, ['pdf']],
      [fileImageIcon, ['png', 'jpg', 'jpeg', 'gif', 'webp']],
    ] as [string, string[]][]
  ).flatMap(([icon, exts]) => exts.map((ext) => [ext, icon])),
)
const MAX_IMAGES_PER_MESSAGE = 20

function formatAttachmentSize(bytes: number): string {
  return bytes >= 1024 * 1024
    ? `${(bytes / (1024 * 1024)).toFixed(2)} MB`
    : `${(bytes / 1024).toFixed(2)} KB`
}

/** composer chips (removable) and the read-only echo on a sent message share one renderer */
function AttachmentList({
  atts,
  previews,
  onRemove,
  removeLabel,
}: {
  atts: AttachmentMeta[]
  previews: Record<string, string>
  onRemove?: (path: string) => void
  removeLabel?: string
}) {
  const remove = (path: string) =>
    onRemove && (
      <button
        type="button"
        className="ai-attachment-thumb-remove"
        onClick={() => onRemove(path)}
        data-tip={removeLabel}
        aria-label={removeLabel}
      >
        <svg width="16" height="16" viewBox="0 0 32 32" aria-hidden>
          <path
            d="M24 9.4L22.6 8L16 14.6L9.4 8L8 9.4l6.6 6.6L8 22.6L9.4 24l6.6-6.6l6.6 6.6l1.4-1.4l-6.6-6.6L24 9.4z"
            fill="currentColor"
          />
        </svg>
      </button>
    )
  return (
    <>
      {atts.map((a) =>
        ATTACHMENT_IMAGE_EXTS.has(a.ext) ? (
          <span key={a.path} className="ai-attachment-thumb" title={a.name}>
            {previews[a.path] ? (
              <img src={previews[a.path]} alt={a.name} />
            ) : (
              <span className="ai-attachment-thumb-pending" aria-hidden>
                <img src={fileImageIcon} alt="" />
              </span>
            )}
            {remove(a.path)}
          </span>
        ) : (
          <span key={a.path} className="ai-attachment-card" title={a.path}>
            <span className="ai-attachment-card-icon">
              <img
                src={ATTACHMENT_CARD_ICONS[a.ext] ?? fileDocumentIcon ?? fileGeneralIcon}
                alt=""
                aria-hidden
              />
            </span>
            <span className="ai-attachment-card-meta">
              <span className="ai-attachment-card-name">{a.name}</span>
              <span className="ai-attachment-card-size">{formatAttachmentSize(a.sizeBytes)}</span>
            </span>
            {remove(a.path)}
          </span>
        ),
      )}
    </>
  )
}

const PANEL_WIDTH_KEY = 'html-ai-panel-width'
const PANEL_WIDTH_DEFAULT = 360
const PANEL_WIDTH_MIN = 280
const MAX_SNAPSHOTS = 20
const TOOL_OUTPUT_MAX_CHARS = 2000

function clampPanelWidth(w: number): number {
  // The viewport can be transiently tiny (a WebContentsView is 0×0 until the
  // shell lays it out), so never let the ceiling drop below the minimum
  const max = Math.max(PANEL_WIDTH_MIN, Math.min(720, Math.round(window.innerWidth * 0.6)))
  return Math.min(Math.max(w, PANEL_WIDTH_MIN), max)
}

function loadPanelWidth(): number {
  const saved = Number(localStorage.getItem(PANEL_WIDTH_KEY))
  // static bounds only — clamping against the window here would bake a
  // transiently small viewport into the restored preference
  return Number.isFinite(saved) && saved > 0
    ? Math.min(Math.max(saved, PANEL_WIDTH_MIN), 720)
    : PANEL_WIDTH_DEFAULT
}

interface ToolActivity {
  name: string
  summary: string
  /** still executing: rendered as a spinner chip, replaced in place when the tool finishes */
  running?: boolean
  isError?: boolean
  output?: string
}

interface ChatEntry {
  role: 'user' | 'assistant'
  text: string
  streaming?: boolean
  isError?: boolean
  /** the run failed and this user message was rolled back out of the model context */
  undelivered?: boolean
  /** instruction actually sent when it differs from the bubble text (sid-pinned Ask AI / queue batch); retries resend this */
  instruction?: string
  /** the bubble was a queue batch: retry must hide the live selection again */
  queueRun?: boolean
  /** attachments consumed from the composer by this message (echoed read-only; retries resend them) */
  attachments?: AttachmentMeta[]
  tools?: ToolActivity[]
  /** the element this user message targeted, frozen at send */
  scope?: AiScopeQuoteData
}

/** the whole source text: one string is the exact rollback unit */
export interface DocSnapshot {
  text: string
}

interface Snapshot {
  label: string
  time: string
  doc: DocSnapshot
}

/** Preset instruction (ribbon / Ask AI "send now"); a new nonce triggers one auto-send */
export interface AiPreset {
  text: string
  nonce: number
  /** what the chat bubble shows when the instruction itself carries protocol text */
  displayText?: string
  scope?: AiScopeQuoteData
}

/** Prefill the composer without sending (Ask AI about the selected element) */
export interface AiDraft {
  text: string
  nonce: number
}

export interface HtmlAiDeps {
  /** live document access for the skill's tools */
  access: HtmlDocAccess
  getSnapshot(): DocSnapshot
  restoreSnapshot(snapshot: DocSnapshot): void
  /** the user sent a request (the text as shown in the bubble); names an untitled document */
  onPrompt(text: string): void
  /** fired when a run with at least one mutation finishes (auto-save hook) */
  onRunDone(mutated: boolean): void
  clearHighlights(): void
  /** [label](htmlnav://sid/N) citation clicked */
  navigateTo(sid: number): void
  /** the user confirmed a brief on the card; the app pins it into the next generated document */
  onBriefConfirmed(brief: Brief): void
  /** page the AI is still writing (null when done): mirrored over the preview */
  previewDraft(html: string | null): void
}

/** how often the streaming draft is pushed into the preview mirror */
const DRAFT_PREVIEW_MS = 400

export function AiPanel({
  deps,
  filePath,
  preset,
  draft,
  editQueue,
  onQueueEditInstruction,
  onQueueRemove,
  onQueueClear,
  onQueueFocus,
  onQueueConsume,
  onCollapse,
}: {
  deps: HtmlAiDeps
  filePath: string | null
  preset?: AiPreset | null
  draft?: AiDraft | null
  editQueue: EditQueueItem[]
  onQueueEditInstruction: (qid: string, instruction: string) => void
  onQueueRemove: (qid: string) => void
  onQueueClear: () => void
  onQueueFocus: (qid: string) => void
  onQueueConsume: (qids: string[]) => void
  onCollapse: () => void
}): ReactElement {
  const { lang, t } = useI18n()
  const [chat, setChat] = useState<ChatEntry[]>([])
  const [prompt, setPrompt] = useState('')
  const [busy, setBusy] = useState(false)
  const [copiedIdx, setCopiedIdx] = useState<number | null>(null)
  const [snapshots, setSnapshots] = useState<Snapshot[]>([])
  /** questionnaire / brief cards docked in the composer slot while a tool waits for the user */
  const [activeClarify, setActiveClarify] = useState<ClarifyQuestion[] | null>(null)
  const clarifyResolverRef = useRef<((r: { answers: string; cancelled?: boolean }) => void) | null>(
    null,
  )
  const [activeBrief, setActiveBrief] = useState<Brief | null>(null)
  const briefResolverRef = useRef<((d: BriefDecision) => void) | null>(null)
  /** the page writer stopped early: keep-or-discard card for what arrived */
  const [activePartial, setActivePartial] = useState<{ lines: number } | null>(null)
  const partialResolverRef = useRef<((keep: boolean) => void) | null>(null)
  /** answered receipts rendered after the user message they belong to (view-only, not chat data) */
  const [receipts, setReceipts] = useState<
    Array<{ afterIdx: number; qa?: Array<{ q: string; a: string }>; brief?: Brief }>
  >([])
  const chatRef = useRef<HTMLDivElement>(null)
  const inputRef = useRef<HTMLTextAreaElement>(null)
  const stickToBottomRef = useRef(true)
  const [intent, setIntent] = useState<PageIntent>('design')
  const intentRef = useRef(intent)
  intentRef.current = intent
  const [attachments, setAttachments] = useState<AttachmentMeta[]>([])
  const attachmentsRef = useRef(attachments)
  attachmentsRef.current = attachments
  const sendSeqRef = useRef(0)
  const [attachNotice, setAttachNotice] = useState<string | null>(null)
  const [dragOver, setDragOver] = useState(false)
  /** data-URL previews for image attachments, keyed by path */
  const [attachmentPreviews, setAttachmentPreviews] = useState<Record<string, string>>({})
  const previewRequestedRef = useRef(new Set<string>())
  /** attachments consumed by earlier sends: the files skill keeps reading them in follow-up turns */
  const sentAttachmentsRef = useRef<AttachmentMeta[]>([])
  useEffect(() => {
    const wanted = [...attachments, ...chat.flatMap((e) => e.attachments ?? [])]
    const alive = new Set(wanted.map((a) => a.path))
    setAttachmentPreviews((prev) => {
      const stale = Object.keys(prev).filter((path) => !alive.has(path))
      if (stale.length === 0) return prev
      const next = { ...prev }
      for (const path of stale) delete next[path]
      return next
    })
    for (const path of previewRequestedRef.current) {
      if (!alive.has(path)) previewRequestedRef.current.delete(path)
    }
    for (const a of wanted) {
      if (!ATTACHMENT_IMAGE_EXTS.has(a.ext) || previewRequestedRef.current.has(a.path)) continue
      previewRequestedRef.current.add(a.path)
      void window.htmlApi
        .readAttachmentImage(a.path)
        .then((r) => {
          if (!previewRequestedRef.current.has(a.path)) return
          if (r.ok && r.base64 && r.mime) {
            setAttachmentPreviews((prev) => ({
              ...prev,
              [a.path]: `data:${r.mime};base64,${r.base64}`,
            }))
          }
        })
        .catch(() => previewRequestedRef.current.delete(a.path))
    }
  }, [attachments, chat])
  const availableAttachments = (): AttachmentMeta[] => {
    const seen = new Set<string>()
    return [...sentAttachmentsRef.current, ...attachmentsRef.current].filter((a) => {
      if (seen.has(a.path)) return false
      seen.add(a.path)
      return true
    })
  }
  const showAttachNotice = (lines: string[]) => {
    if (lines.length === 0) return
    setAttachNotice(lines.join('; '))
    window.setTimeout(() => setAttachNotice(null), 5000)
  }
  const mergeAttachments = (result: AttachmentAddResult | null) => {
    if (!result) return
    if (result.accepted.length > 0) {
      setAttachments((prev) => {
        const seen = new Set(prev.map((a) => a.path))
        return [...prev, ...result.accepted.filter((a) => !seen.has(a.path))]
      })
    }
    showAttachNotice(result.rejected)
  }
  const pickAttachments = async () => mergeAttachments(await window.htmlApi.pickAttachments())
  const onDrop = async (e: React.DragEvent) => {
    e.preventDefault()
    e.stopPropagation()
    setDragOver(false)
    const paths = Array.from(e.dataTransfer.files)
      .map((f) => window.htmlApi.getPathForFile(f))
      .filter(Boolean)
    if (paths.length > 0) mergeAttachments(await window.htmlApi.addAttachmentPaths(paths))
  }
  /** pasted files with a local path go through the regular route; pure bitmaps (screenshots) hit a temp file first */
  const onPasteFiles = async (files: File[]) => {
    const paths: string[] = []
    for (const f of files) {
      const path = window.htmlApi.getPathForFile(f)
      if (path) {
        paths.push(path)
        continue
      }
      const ext = PASTE_MIME_EXT[f.type] ?? f.name.split('.').pop()?.toLowerCase() ?? 'bin'
      mergeAttachments(await window.htmlApi.addPastedImage(await f.arrayBuffer(), ext))
    }
    if (paths.length > 0) mergeAttachments(await window.htmlApi.addAttachmentPaths(paths))
  }
  const removeAttachment = (path: string) =>
    setAttachments((prev) => prev.filter((a) => a.path !== path))
  /** image attachments ride along as multimodal input (≤5MB each, capped per message) */
  const collectImages = async (atts: AttachmentMeta[]): Promise<AgentImage[]> => {
    const imageAtts = atts.filter((a) => ATTACHMENT_IMAGE_EXTS.has(a.ext))
    const images: AgentImage[] = []
    const failures: string[] = []
    for (const att of imageAtts.slice(0, MAX_IMAGES_PER_MESSAGE)) {
      const result = await window.htmlApi.readAttachmentImage(att.path)
      if (result.ok && result.base64 && result.mime) {
        images.push({ base64: result.base64, mime: result.mime })
      } else {
        failures.push(result.error ?? t('aiImageReadFail', { name: att.name }))
      }
    }
    if (imageAtts.length > MAX_IMAGES_PER_MESSAGE) {
      failures.push(t('aiTooManyImages', { max: MAX_IMAGES_PER_MESSAGE }))
    }
    showAttachNotice(failures)
    return images
  }
  // preferred = the user's chosen width (the only value persisted); panelWidth =
  // what fits the current window. Deriving the display width from the preference
  // means a transiently small window never permanently shrinks the panel.
  const preferredWidthRef = useRef(loadPanelWidth())
  const [panelWidth, setPanelWidth] = useState(() => clampPanelWidth(preferredWidthRef.current))
  const [resizing, setResizing] = useState(false)
  const asideRef = useRef<HTMLElement>(null)
  const mountedRef = useRef(true)

  useEffect(() => {
    const dock = asideRef.current?.closest('.ai-dock') as HTMLElement | null
    dock?.style.setProperty('--ai-panel-width', `${panelWidth}px`)
  }, [panelWidth])

  const settingsRef = useRef<AiSettings | null>(null)
  const langRef = useRef(lang)
  langRef.current = lang
  const depsRef = useRef(deps)
  depsRef.current = deps
  const filePathRef = useRef(filePath)
  /** instruction of the in-flight run, labels its rollback snapshot */
  const runInstructionRef = useRef('')
  /** a queue batch names its own targets: the live selection is hidden from the model while it runs */
  const queueRunRef = useRef(false)
  /** whether the last started run was a queue batch (survives onDone/onError for the header retry) */
  const lastQueueRunRef = useRef(false)
  /** what the user saw for that instruction (queue submissions show a summary) */
  const runDisplayRef = useRef('')
  /** the scope quote of the last send, carried over by the header retry */
  const lastScopeRef = useRef<AiScopeQuoteData | undefined>(undefined)
  const runMutatedRef = useRef(false)
  /** tool activity of the whole run, for transcript persistence */
  const runToolsRef = useRef<ToolActivity[]>([])
  const chatIdsRef = useRef<{ projectId: string; chatId: string } | null>(null)
  /** messages sent before resolveChat returned, flushed once the chat id is known */
  const pendingPersistRef = useRef<
    Array<{
      role: 'user' | 'assistant'
      text: string
      tools?: ToolActivity[]
      attachments?: AttachmentMeta[]
      scope?: AiScopeQuoteData
    }>
  >([])

  const patchLast = (patch: Partial<ChatEntry> | ((last: ChatEntry) => Partial<ChatEntry>)) => {
    setChat((prev) => {
      const next = [...prev]
      const last = next[next.length - 1]
      if (!last || last.role !== 'assistant') return prev
      next[next.length - 1] = { ...last, ...(typeof patch === 'function' ? patch(last) : patch) }
      return next
    })
  }

  const persistMessage = (
    role: 'user' | 'assistant',
    text: string,
    tools?: ToolActivity[],
    attachments?: AttachmentMeta[],
    scope?: AiScopeQuoteData,
  ) => {
    const ids = chatIdsRef.current
    if (!window.projectApi) return
    if (!ids) {
      pendingPersistRef.current.push({ role, text, tools, attachments, scope })
      return
    }
    void window.projectApi
      .appendChat({
        projectId: ids.projectId,
        chatId: ids.chatId,
        role,
        text,
        ...(tools && tools.length > 0 ? { tools } : {}),
        ...(attachments && attachments.length > 0 ? { attachments } : {}),
        ...(scope ? { scope } : {}),
      })
      .catch(() => {
        /* persistence failures are silent */
      })
  }

  // The loop is built once; every mutable value goes through a ref getter
  const transportRef = useRef<ReturnType<typeof createElectronTransport> | null>(null)
  if (!transportRef.current) {
    transportRef.current = createElectronTransport(() => settingsRef.current!)
  }

  /**
   * Whole-page generation: one tool-less request whose reply is the HTML, streamed
   * into the preview mirror as it arrives. A stream that stops early leaves the user
   * a keep-or-discard choice; a stream that produced nothing is retried once.
   */
  const runPageWriter = async (
    spec: PageWriteSpec,
    signal?: AbortSignal,
  ): Promise<PageWriteResult> => {
    const { system, user } = buildPageWriterRequest(spec, aiLangDirective(langRef.current))
    let draft = ''
    let timer: number | null = null
    let closed = false
    const flush = () => {
      timer = null
      if (closed) return
      depsRef.current.previewDraft(draft)
      const lines = draft.split('\n').length
      patchLast((last) => ({
        tools: last.tools?.map((tl) =>
          tl.running ? { ...tl, summary: tGlobal('aiWritingPage', { lines }) } : tl,
        ),
      }))
    }
    const attempt = () =>
      streamPage({
        transport: transportRef.current!,
        system,
        user,
        signal,
        onProgress: (html) => {
          if (closed) return
          draft = html
          if (timer === null) timer = window.setTimeout(flush, DRAFT_PREVIEW_MS)
        },
      })
    let outcome = await attempt()
    if (outcome.status === 'empty' && !signal?.aborted) outcome = await attempt()
    // nothing may touch the mirror after this point, or the overlay would outlive the landed page
    closed = true
    if (timer !== null) window.clearTimeout(timer)
    timer = null
    const done = () => depsRef.current.previewDraft(null)
    if (outcome.status === 'complete') {
      done()
      return { ok: true, html: outcome.html }
    }
    if (outcome.status === 'empty') {
      done()
      return { ok: false, error: outcome.error }
    }
    // the draft stays visible while the user decides what to do with it
    depsRef.current.previewDraft(outcome.html)
    const keep = await new Promise<boolean>((resolve) => {
      partialResolverRef.current = resolve
      setActivePartial({ lines: outcome.html.split('\n').length })
    })
    done()
    return keep
      ? { ok: true, html: outcome.html, truncated: true }
      : {
          ok: false,
          error: `${outcome.reason}${outcome.error ? `: ${outcome.error}` : ''}; the user discarded the partial page`,
        }
  }
  const runPageWriterRef = useRef(runPageWriter)
  runPageWriterRef.current = runPageWriter

  const loopRef = useRef<AgentLoop<DocSnapshot> | null>(null)
  if (!loopRef.current) {
    loopRef.current = new AgentLoop<DocSnapshot>({
      transport: transportRef.current,
      skill: composeSkills('html+search', '', [
        createDocumentSkill({
          getText: () => depsRef.current.access.getText(),
          getVersion: () => depsRef.current.access.getVersion(),
          getMap: () => depsRef.current.access.getMap(),
          getLastManualVersion: () => depsRef.current.access.getLastManualVersion(),
          getFilePath: () => depsRef.current.access.getFilePath(),
          getSelectedSid: () =>
            queueRunRef.current ? null : depsRef.current.access.getSelectedSid(),
          applyOps: (ops, label) => depsRef.current.access.applyOps(ops, label),
          replaceAll: (html, label) => depsRef.current.access.replaceAll(html, label),
          askClarification: (questions) =>
            new Promise((resolve) => {
              clarifyResolverRef.current = resolve
              setActiveClarify(questions)
            }),
          confirmBrief: (brief) =>
            new Promise((resolve) => {
              briefResolverRef.current = resolve
              setActiveBrief(brief)
            }),
          writePage: (spec, signal) => runPageWriterRef.current(spec, signal),
          getInstruction: () => runInstructionRef.current,
        }),
        createSearchSkill(),
        createFilesSkill(availableAttachments),
        createIntentSkill(
          () => intentRef.current,
          () => isDocEmpty(depsRef.current.access.getText()),
        ),
      ]),
      captureSnapshot: () => depsRef.current.getSnapshot(),
      systemSuffix: () => aiLangDirective(langRef.current),
      events: {
        onText: (text) => patchLast({ text }),
        onToolStart: (call) => {
          // Live "running" chip: replaced in place by onToolExecuted
          patchLast((last) => ({
            tools: [
              ...(last.tools ?? []),
              { name: call.name, summary: call.name.replace(/[_-]+/g, ' '), running: true },
            ],
          }))
        },
        onToolExecuted: ({ call, execution, snapshotBefore }) => {
          if (execution.mutated) runMutatedRef.current = true
          if (snapshotBefore !== undefined) {
            const label = runInstructionRef.current.slice(0, 40)
            const time = new Date().toLocaleTimeString([], {
              hour: '2-digit',
              minute: '2-digit',
            })
            setSnapshots((prev) =>
              [...prev, { label, time, doc: snapshotBefore }].slice(-MAX_SNAPSHOTS),
            )
          }
          const activity: ToolActivity = {
            name: call.name,
            summary: execution.summary,
            isError: execution.isError,
            output: execution.output?.slice(0, TOOL_OUTPUT_MAX_CHARS),
          }
          runToolsRef.current.push(activity)
          patchLast((last) => {
            // Swap out the running placeholder pushed by onToolStart (parse-fail calls have none)
            const tools = [...(last.tools ?? [])]
            if (tools.at(-1)?.running) tools.pop()
            return { tools: [...tools, activity] }
          })
        },
        onTurnEnd: () => {
          patchLast({ streaming: false })
          setChat((prev) => [...prev, { role: 'assistant', text: '', streaming: true }])
        },
        onDone: ({ text, cancelled, turnLimit, truncated }) => {
          const base = turnLimit
            ? [text, tGlobal('aiTurnLimit')].filter(Boolean).join('\n\n')
            : text || (cancelled ? tGlobal('aiStopped') : '')
          // A reasoning model can spend the entire output budget on thinking and close the
          // turn with finish_reason=length and no prose at all — the bare "(no reply)" read
          // as the assistant ignoring the user. Name the truncation, as docs already does.
          const final = truncated
            ? [base, tGlobal('aiTruncatedNote')].filter(Boolean).join('\n\n')
            : base
          patchLast((last) => ({
            streaming: false,
            text: final || (last.tools?.length ? last.text : tGlobal('aiNoReply')),
            // A stop mid-tool can leave a running placeholder behind — drop it
            tools: last.tools?.filter((tl) => !tl.running),
          }))
          persistMessage('assistant', final, runToolsRef.current)
          depsRef.current.clearHighlights()
          depsRef.current.onRunDone(runMutatedRef.current)
          queueRunRef.current = false
          setBusy(false)
        },
        onError: (error) => {
          // once a tool ran the message was delivered; a later turn failing is not a send failure
          const undelivered = runToolsRef.current.length === 0
          setChat((prev) => {
            const next = [...prev]
            for (let i = undelivered ? next.length - 1 : -1; i >= 0; i--) {
              const entry = next[i]!
              if (entry.role === 'user') {
                next[i] = { ...entry, undelivered: true }
                break
              }
            }
            const last = next.at(-1)
            if (last?.role === 'assistant') {
              next[next.length - 1] = {
                ...last,
                streaming: false,
                text: error,
                isError: true,
                tools: last.tools?.filter((tl) => !tl.running),
              }
            }
            return next
          })
          queueRunRef.current = false
          setBusy(false)
        },
      },
    })
  }

  useEffect(() => {
    mountedRef.current = true
    return () => {
      mountedRef.current = false
      clarifyResolverRef.current?.({ answers: '', cancelled: true })
      briefResolverRef.current?.({ kind: 'cancelled' })
      partialResolverRef.current?.(false)
      loopRef.current?.cancel()
      depsRef.current.clearHighlights()
    }
  }, [])

  // ── chat-history persistence: bind to the file, restore prior transcript ──
  useEffect(() => {
    const api = window.projectApi
    if (!api) return
    const tempChatId = `unsaved-${Date.now()}`
    void api
      .resolveChat({ filePath: filePathRef.current ?? null, tempChatId })
      .then((ids) => {
        chatIdsRef.current = ids
        for (const msg of pendingPersistRef.current.splice(0)) {
          persistMessage(msg.role, msg.text, msg.tools, msg.attachments, msg.scope)
        }
        return api.loadChat({ projectId: ids.projectId, chatId: ids.chatId, limit: 200 })
      })
      .then((msgs) => {
        if (msgs.length === 0) return
        // the user may have sent a message while history was loading — never
        // replace a live transcript (and don't clobber the loop context)
        let applied = false
        // stored metadata only: the chips render name/size, a still-readable image gets its thumbnail
        const restoredAtts = (m: (typeof msgs)[number]): AttachmentMeta[] | undefined =>
          m.attachments
            ?.filter((a) => a.path)
            .map((a) => ({
              name: a.name,
              path: a.path ?? '',
              ext: a.ext ?? '',
              sizeBytes: a.sizeBytes ?? 0,
            }))
        setChat((prev) => {
          if (prev.length > 0) return prev
          applied = true
          return msgs.map((m) => ({
            role: m.role,
            text: m.text,
            tools: m.tools?.map((tool) => ({
              name: tool.name,
              summary: tool.summary,
              isError: tool.isError,
              output: tool.output ? tool.output.slice(0, TOOL_OUTPUT_MAX_CHARS) : undefined,
            })),
            attachments: restoredAtts(m),
            ...(m.scope ? { scope: m.scope } : {}),
          }))
        })
        if (applied && !loopRef.current?.busy) {
          loopRef.current?.restore(msgs.map((m) => ({ role: m.role, text: m.text })))
          // follow-up turns can keep reading the files those messages attached
          const seen = new Set(sentAttachmentsRef.current.map((a) => a.path))
          for (const a of msgs.flatMap((m) => restoredAtts(m) ?? [])) {
            if (seen.has(a.path)) continue
            seen.add(a.path)
            sentAttachmentsRef.current.push(a)
          }
        }
      })
      .catch(() => {
        /* history load failures are silent */
      })
  }, [])

  /** after an untitled document's first save, bind the unsaved-* history to the real path */
  useEffect(() => {
    filePathRef.current = filePath
    const ids = chatIdsRef.current
    if (!window.projectApi || !ids || !filePath || !ids.chatId.startsWith('unsaved-')) return
    void window.projectApi
      .rebindChat({ projectId: ids.projectId, tempChatId: ids.chatId, newFilePath: filePath })
      .then((r) => {
        if (r?.chatId) chatIdsRef.current = r
      })
      .catch(() => {
        /* silent */
      })
  }, [filePath])

  useEffect(() => {
    if (stickToBottomRef.current) {
      chatRef.current?.scrollTo({ top: chatRef.current.scrollHeight })
    }
  }, [chat, busy])

  const onChatScroll = (): void => {
    const el = chatRef.current
    if (!el) return
    stickToBottomRef.current = el.scrollHeight - el.scrollTop - el.clientHeight < 48
  }

  /** false when nothing was started (empty text or a run already active) */
  const send = (
    text: string,
    displayText?: string,
    queueRun = false,
    attachmentsOverride?: AttachmentMeta[],
    scope?: AiScopeQuoteData,
  ): boolean => {
    const instruction = text.trim()
    const loop = loopRef.current
    if (!instruction || !loop || loop.busy) return false
    // the message consumes the composer attachments: echoed on the bubble, images multimodal, files via the skill
    const sentAtts = attachmentsOverride ?? attachmentsRef.current
    if (sentAtts.length > 0) {
      const seen = new Set(sentAttachmentsRef.current.map((a) => a.path))
      sentAttachmentsRef.current = [
        ...sentAttachmentsRef.current,
        ...sentAtts.filter((a) => !seen.has(a.path)),
      ]
      if (!attachmentsOverride) setAttachments([])
    }
    stickToBottomRef.current = true
    queueRunRef.current = queueRun
    lastQueueRunRef.current = queueRun
    runInstructionRef.current = instruction
    runDisplayRef.current = displayText ?? instruction
    lastScopeRef.current = scope
    runMutatedRef.current = false
    depsRef.current.onPrompt(displayText ?? instruction)
    runToolsRef.current = []
    setChat((prev) => [
      ...prev,
      {
        role: 'user',
        text: displayText ?? instruction,
        instruction: displayText ? instruction : undefined,
        queueRun: queueRun || undefined,
        attachments: sentAtts.length > 0 ? sentAtts : undefined,
        scope,
      },
      { role: 'assistant', text: '', streaming: true },
    ])
    setPrompt('')
    setBusy(true)
    // persist what the user saw — a restored transcript must not surface the
    // internal batch protocol text behind a queue submission
    persistMessage('user', displayText ?? instruction, undefined, sentAtts, scope)
    // Stop / New chat during the pre-run reads (settings, image bytes) bump the sequence; a stale send never starts the loop
    const seq = ++sendSeqRef.current
    void (async () => {
      try {
        settingsRef.current = await window.htmlApi.getAiSettings()
        const images = sentAtts.length > 0 ? await collectImages(sentAtts) : []
        if (!mountedRef.current || seq !== sendSeqRef.current) return
        await loop.run(instruction, images.length > 0 ? images : undefined)
      } catch (err) {
        // a send cancelled during its pre-run reads must not paint its error onto a later turn's bubble
        if (!mountedRef.current || seq !== sendSeqRef.current) return
        patchLast({
          streaming: false,
          text: err instanceof Error ? err.message : String(err),
          isError: true,
        })
        queueRunRef.current = false
        setBusy(false)
      }
    })()
    return true
  }

  /** one run for the whole queue; every item is consumed up front, failures go through retry */
  const sendQueue = (): void => {
    const access = depsRef.current.access
    const entries = liveItems(resolveQueue(access.getText(), access.getMap(), editQueue))
    if (entries.length === 0) {
      onQueueClear()
      return
    }
    const started = send(
      buildQueueInstruction(entries),
      buildQueueSummary(t('aiQueueSubmitted', { count: entries.length }), entries),
      true,
    )
    // consume only once the run is under way; a failed run keeps its retry via the bubble's instruction
    if (started) onQueueConsume(editQueue.map((item) => item.qid))
  }

  /** finish pending cards as skipped so no promise is left dangling */
  const dismissCards = (): void => {
    clarifyResolverRef.current?.({ answers: '', cancelled: true })
    clarifyResolverRef.current = null
    setActiveClarify(null)
    briefResolverRef.current?.({ kind: 'cancelled' })
    briefResolverRef.current = null
    setActiveBrief(null)
    decidePartial(false)
  }

  const decidePartial = (keep: boolean): void => {
    partialResolverRef.current?.(keep)
    partialResolverRef.current = null
    setActivePartial(null)
  }

  /** pre-run send in flight (reads before loop.run): cancelling it drops the empty reply and marks the message undelivered */
  const cancelPendingSend = (): void => {
    sendSeqRef.current++
    setChat((prev) => {
      const last = prev.at(-1)
      if (!last || last.role !== 'assistant' || !last.streaming || last.text) return prev
      const rest = prev.slice(0, -1)
      const user = rest.at(-1)
      if (user?.role === 'user') rest[rest.length - 1] = { ...user, undelivered: true }
      return rest
    })
    queueRunRef.current = false
    setBusy(false)
  }

  const stop = (): void => {
    dismissCards()
    const loop = loopRef.current
    if (loop?.busy) loop.cancel()
    else cancelPendingSend()
  }

  const retry = (): void => {
    send(
      runInstructionRef.current,
      runDisplayRef.current,
      lastQueueRunRef.current,
      undefined,
      lastScopeRef.current,
    )
  }

  /** [label](htmlnav://sid/N) links in replies select that element */
  const docNav = {
    scheme: DOC_NAV_SCHEME,
    onNavigate: (href: string) => {
      const sid = parseDocNavHref(href)
      if (sid !== null) depsRef.current.navigateTo(sid)
    },
  }

  const draftNonceRef = useRef(0)
  useEffect(() => {
    if (!draft || draft.nonce === draftNonceRef.current) return
    draftNonceRef.current = draft.nonce
    setPrompt(draft.text)
    inputRef.current?.focus()
  }, [draft])

  // ribbon presets auto-send; while a run is active they land in the composer instead
  const presetNonceRef = useRef(0)
  useEffect(() => {
    if (!preset || preset.nonce === presetNonceRef.current) return
    presetNonceRef.current = preset.nonce
    // while a run is active the request lands in the composer as the user phrased it, never as protocol text
    if (loopRef.current?.busy) setPrompt(preset.displayText ?? preset.text)
    else send(preset.text, preset.displayText, false, undefined, preset.scope)
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [preset])

  const copyMessage = (text: string, idx: number): void => {
    void navigator.clipboard.writeText(text)
    setCopiedIdx(idx)
    window.setTimeout(() => setCopiedIdx((cur) => (cur === idx ? null : cur)), 1200)
  }

  const rollback = (snapshot: Snapshot): void => {
    if (busy) return
    depsRef.current.restoreSnapshot(snapshot.doc)
    setSnapshots((prev) => prev.filter((s) => s !== snapshot))
  }

  // Re-derive the display width on window resize (max is 60% of the window);
  // growing the window back restores the preferred width
  useEffect(() => {
    const onResize = (): void => setPanelWidth(clampPanelWidth(preferredWidthRef.current))
    window.addEventListener('resize', onResize)
    return () => window.removeEventListener('resize', onResize)
  }, [])

  const resizeCleanupRef = useRef<(() => void) | null>(null)
  useEffect(() => () => resizeCleanupRef.current?.(), [])

  /** Drag the right edge to resize: the panel is flush with the window's left edge, so width = clientX */
  const startResize = (e: ReactPointerEvent<HTMLDivElement>): void => {
    e.preventDefault()
    const resizer = e.currentTarget
    setResizing(true)
    document.body.style.cursor = 'col-resize'
    document.body.style.userSelect = 'none'
    const onMove = (ev: PointerEvent): void => {
      const w = clampPanelWidth(ev.clientX)
      preferredWidthRef.current = w
      setPanelWidth(w)
    }
    let done = false
    const cleanup = (): void => {
      if (done) return
      done = true
      resizeCleanupRef.current = null
      window.removeEventListener('pointermove', onMove)
      window.removeEventListener('pointerup', cleanup)
      window.removeEventListener('pointercancel', cleanup)
      resizer.removeEventListener('lostpointercapture', cleanup)
      document.body.style.cursor = ''
      document.body.style.userSelect = ''
      setResizing(false)
      localStorage.setItem(PANEL_WIDTH_KEY, String(Math.round(preferredWidthRef.current)))
    }
    resizeCleanupRef.current = cleanup
    window.addEventListener('pointermove', onMove)
    window.addEventListener('pointerup', cleanup)
    window.addEventListener('pointercancel', cleanup)
    resizer.addEventListener('lostpointercapture', cleanup)
    resizer.setPointerCapture(e.pointerId)
  }

  const docEmpty = chat.length === 0 && isDocEmpty(deps.access.getText())

  return (
    <aside
      ref={asideRef}
      className={`copilot${resizing ? ' ai-panel-resizing' : ''}${dragOver ? ' ai-panel-dragover' : ''}`}
      style={{ width: '100%' }}
      dir={lang === 'ar' || lang === 'he' ? 'rtl' : undefined}
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
      onDrop={(e) => void onDrop(e)}
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
          {chat.length > 0 && (
            <button
              className="ai-header-btn"
              onClick={() => {
                stop()
                loopRef.current?.reset()
                setBusy(false)
                setChat([])
                setReceipts([])
                setAttachments([])
                setAttachNotice(null)
                sentAttachmentsRef.current = []
              }}
              data-tip={t('aiNewChat')}
              aria-label={t('aiNewChat')}
            >
              <IconNewChat />
            </button>
          )}
          <button
            className="ai-header-btn"
            onClick={onCollapse}
            data-tip={t('aiCollapsePanel')}
            aria-label={t('aiCollapsePanel')}
          >
            <IconCollapse />
          </button>
        </div>
      </header>

      <div className="ai-chat" ref={chatRef} onScroll={onChatScroll}>
        {chat.length === 0 && (
          <div className="ai-chat-empty">
            <div className="ai-chat-empty-title">
              {t(
                docEmpty
                  ? intent === 'write'
                    ? 'aiEmptyWriteTitle'
                    : 'aiEmptyTitle'
                  : 'aiEmptyDocTitle',
              )}
            </div>
            <div className="ai-chat-empty-body">
              {t(
                docEmpty
                  ? intent === 'write'
                    ? 'aiEmptyWriteBody'
                    : 'aiEmptyBody'
                  : 'aiEmptyDocBody',
              )}
            </div>
            <div className="ai-starter-list">
              {(docEmpty
                ? intent === 'write'
                  ? WRITE_STARTERS
                  : GENERATE_STARTERS
                : EDIT_STARTERS
              ).map(([label, promptKey]) => (
                <button
                  key={label}
                  className="ai-starter"
                  onClick={() => {
                    setPrompt(t(promptKey))
                    inputRef.current?.focus()
                  }}
                >
                  {t(label)}
                </button>
              ))}
            </div>
          </div>
        )}
        {chat.map((entry, i) => {
          if (entry.role === 'user') {
            const own = receipts.filter((r) => r.afterIdx === i)
            return (
              <div key={i} className="ai-msg ai-msg-user">
                {entry.scope && <AiScopeQuote scope={entry.scope} />}
                {entry.attachments && entry.attachments.length > 0 && (
                  <div className="ai-msg-attachments">
                    <AttachmentList atts={entry.attachments} previews={attachmentPreviews} />
                  </div>
                )}
                <span dir="auto">{entry.text}</span>
                {own.map((r, k) =>
                  r.qa ? (
                    <div key={`ca${k}`} className="ai-clarify-answered">
                      {r.qa.map((pair, m) => (
                        <div key={m} className="ai-clarify-answered-row">
                          <div className="ai-clarify-answered-q">{pair.q}</div>
                          <div className="ai-clarify-answered-a">{pair.a}</div>
                        </div>
                      ))}
                    </div>
                  ) : r.brief ? (
                    <div key={`br${k}`} className="brief-receipt">
                      <b>{t('briefTitle')}</b> ·{' '}
                      {t('briefReceipt', { hook: r.brief.core_hook, n: r.brief.sections.length })}
                    </div>
                  ) : null,
                )}
                {entry.undelivered && (
                  <div className="ai-msg-undelivered">
                    {t('aiUndelivered')}
                    {!busy && (
                      <button
                        className="ai-retry-btn"
                        onClick={() =>
                          send(
                            entry.instruction ?? entry.text,
                            entry.text,
                            entry.queueRun,
                            entry.attachments ?? [],
                            entry.scope,
                          )
                        }
                      >
                        {t('aiRetry')}
                      </button>
                    )}
                  </div>
                )}
              </div>
            )
          }
          const hasTools = (entry.tools?.length ?? 0) > 0
          if (!entry.text && !entry.streaming && !hasTools) return null
          const isLast = i === chat.length - 1
          // Action row appears once per completed reply: on the turn's final segment only
          // (mid-turn segments have a following assistant entry; the live turn ends when !busy)
          const nextEntry = chat[i + 1]
          const turnEnded = nextEntry ? nextEntry.role === 'user' : !busy
          const showToolbar = !entry.streaming && turnEnded && !!entry.text && !entry.isError
          return (
            <div
              key={i}
              className={`ai-msg ai-msg-assistant${entry.isError ? ' ai-msg-error' : ''}${entry.streaming ? ' ai-msg-streaming' : ''}`}
            >
              {!entry.text && entry.streaming ? (
                <span className="ai-typing-row">
                  <AiTypingIndicator label={hasTools ? t('aiWorking') : t('aiThinking')} />
                </span>
              ) : (
                entry.text && (
                  <div dir="auto">
                    <Markdown text={entry.text} nav={docNav} />
                  </div>
                )
              )}
              {hasTools && <ToolChipList tools={entry.tools!} />}
              {showToolbar && (
                <div className="ai-msg-toolbar">
                  <button
                    className="ai-msg-tool-btn"
                    onClick={() => copyMessage(entry.text, i)}
                    aria-label={t('aiCopyReplyTitle')}
                    data-tip={t('aiCopyReplyTitle')}
                  >
                    {copiedIdx === i ? (
                      <svg
                        width="14"
                        height="14"
                        viewBox="0 0 24 24"
                        fill="none"
                        stroke="currentColor"
                        strokeWidth="1.8"
                        strokeLinecap="round"
                        strokeLinejoin="round"
                      >
                        <polyline points="20 6 9 17 4 12" />
                      </svg>
                    ) : (
                      <svg width="20" height="20" viewBox="0 0 20 20" fill="none">
                        <path
                          d="M14.6113 5.34253C16.0608 5.3428 17.2363 6.518 17.2363 7.96753V15.5066C17.2361 16.956 16.0607 18.1313 14.6113 18.1316H7.07227C5.62267 18.1316 4.44751 16.9561 4.44727 15.5066V7.96753C4.44732 6.51783 5.62255 5.34253 7.07227 5.34253H14.6113ZM7.07227 6.59253C6.31291 6.59253 5.69732 7.20819 5.69727 7.96753V15.5066C5.69751 16.2658 6.31302 16.8816 7.07227 16.8816H14.6113C15.3703 16.8813 15.9861 16.2656 15.9863 15.5066V7.96753C15.9863 7.20835 15.3705 6.5928 14.6113 6.59253H7.07227ZM10.0176 2.8689C10.3626 2.86905 10.6426 3.14882 10.6426 3.4939C10.6425 3.83888 10.3626 4.11874 10.0176 4.1189H4.59961C3.84022 4.1189 3.22461 4.73451 3.22461 5.4939V11.324C3.22433 11.6689 2.94461 11.949 2.59961 11.949C2.25461 11.949 1.97489 11.6689 1.97461 11.324V5.4939C1.97461 4.04415 3.14987 2.8689 4.59961 2.8689H10.0176Z"
                          fill="currentColor"
                        />
                      </svg>
                    )}
                  </button>
                  {isLast && !busy && runInstructionRef.current && (
                    <button
                      className="ai-msg-tool-btn"
                      onClick={retry}
                      aria-label={t('aiRegenerateTitle')}
                      data-tip={t('aiRegenerateTitle')}
                    >
                      <svg
                        width="20"
                        height="20"
                        viewBox="0 0 16 16"
                        fill="none"
                        stroke="currentColor"
                        strokeWidth="1"
                        strokeLinecap="round"
                        strokeLinejoin="round"
                        aria-hidden
                      >
                        <path d="M 12.68 6.65 a 4.86 4.86 0 0 0 -9 -1.08 M 3.32 9.35 a 4.86 4.86 0 0 0 9 1.08" />
                        <path d="M 12.95 3.05 v 2.7 h -2.7 M 3.05 12.95 v -2.7 h 2.7" />
                      </svg>
                    </button>
                  )}
                </div>
              )}
            </div>
          )
        })}
      </div>

      {snapshots.length > 0 && (
        <div className="ai-versions">
          <div className="ai-versions-title">
            <IconClock />
            {t('aiSnapshotsTitle')}
          </div>
          {snapshots.map((s, i) => (
            <div key={i} className="ai-version-row">
              <span className="ai-version-label" data-tip={s.label}>
                <span className="ai-version-time">{s.time}</span>
                {s.label}
              </span>
              <button className="ai-version-rollback" disabled={busy} onClick={() => rollback(s)}>
                {t('aiRollback')}
              </button>
            </div>
          ))}
        </div>
      )}

      {activeClarify && (
        <div className="ai-composer">
          <ClarifyCard
            questions={activeClarify}
            onSubmit={(answers, qa) => {
              const idx = chat.map((c) => c.role).lastIndexOf('user')
              clarifyResolverRef.current?.({ answers })
              clarifyResolverRef.current = null
              setActiveClarify(null)
              setReceipts((prev) => [...prev, { afterIdx: idx, qa }])
            }}
            onSkip={() => {
              clarifyResolverRef.current?.({ answers: '', cancelled: true })
              clarifyResolverRef.current = null
              setActiveClarify(null)
            }}
          />
        </div>
      )}
      {activeBrief && !activeClarify && (
        <div className="ai-composer">
          <BriefCard
            key={activeBrief.core_hook + activeBrief.sections.length}
            brief={activeBrief}
            onDecide={(decision) => {
              const idx = chat.map((c) => c.role).lastIndexOf('user')
              if (decision.kind === 'confirmed') {
                depsRef.current.onBriefConfirmed(decision.brief)
                setReceipts((prev) => [...prev, { afterIdx: idx, brief: decision.brief }])
              }
              briefResolverRef.current?.(decision)
              briefResolverRef.current = null
              setActiveBrief(null)
            }}
          />
        </div>
      )}
      {activePartial && !activeClarify && !activeBrief && (
        <div className="ai-composer">
          <div className="brief-card ai-partial-card" role="group" aria-label={t('aiPartialTitle')}>
            <div className="ai-partial-title">{t('aiPartialTitle')}</div>
            <div className="ai-partial-body">
              {t('aiPartialBody', { lines: activePartial.lines })}
            </div>
            <div className="brief-actions">
              <button type="button" className="brief-btn" onClick={() => decidePartial(false)}>
                {t('aiPartialDiscard')}
              </button>
              <button
                type="button"
                className="brief-btn primary"
                onClick={() => decidePartial(true)}
              >
                {t('aiPartialAdopt')}
              </button>
            </div>
          </div>
        </div>
      )}
      <div
        className="ai-composer"
        style={activeClarify || activeBrief || activePartial ? { display: 'none' } : undefined}
      >
        {chat.length === 0 && docEmpty && (
          <div className="ai-intent-bar">
            <div className="ai-intent-label">{t('aiIntentLabel')}</div>
            <div className="ai-intent-cards" role="radiogroup" aria-label={t('aiIntentLabel')}>
              {(
                [
                  ['design', 'aiIntentDesign', 'aiIntentDesignDesc'],
                  ['write', 'aiIntentWrite', 'aiIntentWriteDesc'],
                ] as const
              ).map(([kind, title, desc]) => (
                <button
                  key={kind}
                  type="button"
                  role="radio"
                  aria-checked={intent === kind}
                  className={`ai-intent-card${intent === kind ? ' selected' : ''}`}
                  data-tip={t(desc)}
                  onClick={() => {
                    setIntent(kind)
                    inputRef.current?.focus()
                  }}
                >
                  <span className="ai-intent-card-icon" aria-hidden>
                    {kind === 'design' ? (
                      <svg width="16" height="16" viewBox="0 0 20 20" fill="none">
                        <rect x="2.5" y="2.5" width="15" height="15" rx="2.5" />
                        <path d="M2.5 7.5h15M7.5 7.5v10" />
                      </svg>
                    ) : (
                      <svg width="16" height="16" viewBox="0 0 20 20" fill="none">
                        <path d="M4 15.5h12M4 4.5h12M4 8.2h12M4 11.8h8" />
                      </svg>
                    )}
                  </span>
                  <span className="ai-intent-card-title">{t(title)}</span>
                  {intent === kind && (
                    <span className="ai-intent-card-check" aria-hidden>
                      <svg width="10" height="10" viewBox="0 0 10 10" fill="none">
                        <path d="M2 5.2l2.2 2.2L8 3" />
                      </svg>
                    </span>
                  )}
                </button>
              ))}
            </div>
          </div>
        )}
        {editQueue.length > 0 && (
          <EditQueueCard
            items={editQueue}
            text={deps.access.getText()}
            map={deps.access.getMap()}
            busy={busy}
            onEditInstruction={onQueueEditInstruction}
            onRemove={onQueueRemove}
            onDiscardAll={onQueueClear}
            onFocus={onQueueFocus}
            onSend={sendQueue}
          />
        )}
        {attachments.length > 0 && (
          <div className="ai-attachments">
            <AttachmentList
              atts={attachments}
              previews={attachmentPreviews}
              onRemove={removeAttachment}
              removeLabel={t('aiRemoveAttachmentTitle')}
            />
          </div>
        )}
        {attachNotice && <div className="ai-attach-notice">{attachNotice}</div>}
        <AiComposer
          value={prompt}
          busy={busy}
          placeholder={t('aiComposerPlaceholder')}
          hintIdle={t('aiHintIdle')}
          hintBusy={t('aiHintBusy')}
          sendLabel={t('aiSend')}
          stopLabel={t('aiStop')}
          iconOnly
          sendIconEnabled={<img src={sendEnterOn} alt="" aria-hidden />}
          sendIconDisabled={<img src={sendEnterOff} alt="" aria-hidden />}
          stopIcon={<img src={sendStop} alt="" aria-hidden />}
          textareaRef={inputRef}
          onChange={setPrompt}
          onSend={() => send(prompt)}
          onStop={stop}
          onPasteFiles={(files) => void onPasteFiles(files)}
          footerStart={
            <button
              type="button"
              className="ai-attach-btn"
              onClick={() => void pickAttachments()}
              data-tip={t('aiAttachTitle')}
              aria-label={t('aiAttachTitle')}
            >
              <img src={attachIcon} alt="" aria-hidden />
            </button>
          }
        />
      </div>
    </aside>
  )
}

/** Step-row status icons (timeline glyphs, unified with the other apps) */
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

/** Tool activity group (docs parity): auto-opens while tools run, auto-collapses into
 *  "Worked · N steps" when they finish; a manual toggle always wins */
function ToolChipList({ tools }: { tools: ToolActivity[] }) {
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
                      data-tip={tool.name}
                      aria-expanded={isOpen}
                      onClick={() => toggle(j)}
                    >
                      {tool.summary}
                    </button>
                  ) : (
                    <span className="ai-step-title" data-tip={tool.name}>
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

function Svg({ children }: { children: ReactNode }): ReactElement {
  return (
    <svg
      width="15"
      height="15"
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

function IconNewChat(): ReactElement {
  return (
    <Svg>
      <path
        d="M13.5 7.2v-3A1.7 1.7 0 0 0 11.8 2.5H4.2a1.7 1.7 0 0 0-1.7 1.7v6.1a1.7 1.7 0 0 0 1.7 1.7h1.1v2l2.6-2h1.3"
        strokeLinejoin="round"
      />
      <path d="M12.2 9.4v4M10.2 11.4h4" />
    </Svg>
  )
}

function IconCollapse(): ReactElement {
  return (
    <svg
      width={15}
      height={15}
      viewBox="0 0 16 16"
      fill="none"
      stroke="currentColor"
      strokeWidth="1.2"
      strokeLinecap="round"
      aria-hidden
    >
      <rect x="1.5" y="2.5" width="13" height="11" rx="1" />
      <path d="M5.5 2.5v11" />
      <path d="M12.5 8H8.1M9.8 5.9 7.7 8l2.1 2.1" strokeWidth="1.3" strokeLinejoin="round" />
    </svg>
  )
}

function IconClock(): ReactElement {
  return (
    <svg
      width="12"
      height="12"
      viewBox="0 0 16 16"
      fill="none"
      stroke="currentColor"
      strokeWidth="1.3"
      strokeLinecap="round"
      aria-hidden
    >
      <circle cx="8" cy="8" r="6" />
      <path d="M8 4.8V8l2.2 1.6" />
    </svg>
  )
}

/** Genspark brand mark, inline for crisp device-resolution rendering */
export function GensparkMark({ size = 18 }: { size?: number }): React.JSX.Element {
  return (
    <svg
      width={size}
      height={size}
      viewBox="0 0 130 130.025"
      fill="none"
      xmlns="http://www.w3.org/2000/svg"
      aria-hidden
    >
      <path
        d="M105.115 0H24.6428C11.0443 0 0 11.0686 0 24.6915V105.334C0 118.981 11.0199 130.025 24.6428 130.025H105.115C118.714 130.025 129.758 118.957 129.758 105.334V24.6915C129.758 11.0443 118.714 0 105.115 0ZM71.5201 35.2735C85.5078 33.1571 86.7729 31.9164 88.865 17.88C88.938 17.4421 89.3028 17.1259 89.7407 17.1259C90.1786 17.1259 90.5435 17.4421 90.6164 17.88C92.7328 31.8921 93.9735 33.1571 107.961 35.2735C108.399 35.3465 108.715 35.7114 108.715 36.1493C108.715 36.5871 108.399 36.952 107.961 37.025C93.9249 39.1414 92.7085 40.4064 90.5677 54.6131C90.5191 54.9537 90.2516 55.197 89.911 55.197C89.5704 55.197 89.3028 54.9537 89.2542 54.6131C87.1134 40.4064 85.5565 39.1658 71.4958 37.025C71.0579 36.952 70.7417 36.5871 70.7417 36.1493C70.7417 35.7114 71.0579 35.3465 71.4958 35.2735H71.5201ZM101.758 78.5261C101.758 78.8181 101.563 79.037 101.271 79.0856C92.3193 80.4236 91.5652 81.2264 90.2029 90.2759C90.1786 90.4948 89.9839 90.6408 89.7893 90.6408C89.5703 90.6408 89.4001 90.4948 89.3758 90.2759C88.0135 81.2507 87.0161 80.4479 78.0883 79.0856C77.7964 79.037 77.6017 78.7937 77.6017 78.5261C77.6017 78.2342 77.7964 78.0153 78.0883 77.9666C86.9918 76.6287 87.7703 75.8259 89.1326 66.898C89.1812 66.6061 89.4244 66.4115 89.692 66.4115C89.9839 66.4115 90.2028 66.6061 90.2515 66.898C91.5894 75.8259 92.3923 76.6043 101.296 77.9666C101.588 78.0153 101.782 78.2585 101.782 78.5261H101.758ZM16.5178 54.8077C16.5178 54.1023 17.0286 53.4941 17.7341 53.3968C40.1388 50.0154 42.1093 47.9963 45.4907 25.5672C45.588 24.8861 46.1961 24.3509 46.9016 24.3509C47.6071 24.3509 48.191 24.8617 48.3126 25.5672C51.694 47.9963 53.6887 50.0154 76.0691 53.3968C76.7503 53.4941 77.2855 54.1023 77.2855 54.8077C77.2855 55.5132 76.7746 56.1214 76.0691 56.2187C53.5914 59.6244 51.6696 61.6192 48.2639 84.3645C48.1909 84.8754 47.7287 85.2889 47.2179 85.2889C46.707 85.2889 46.2448 84.8997 46.1718 84.3645C42.7418 61.6435 40.2604 59.6244 17.7584 56.2187C17.0772 56.1214 16.542 55.5132 16.542 54.8077H16.5178ZM112.097 109.591C112.097 111.416 110.613 112.9 108.813 112.9H21.2614C19.4369 112.9 17.9774 111.416 17.9774 109.591V102.658C17.9774 100.834 19.4612 99.3497 21.2614 99.3497H108.813C110.637 99.3497 112.097 100.834 112.097 102.658V109.591Z"
        fill="currentColor"
      />
    </svg>
  )
}
