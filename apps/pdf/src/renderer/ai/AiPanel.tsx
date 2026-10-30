import {
  AI_PANEL_WIDTH_DEFAULT,
  AI_PANEL_WIDTH_MAX,
  aiPanelAutoWidth,
  aiPanelClampWidth,
  aiPanelWidthAtPointer,
  AiPanelSideButton,
  AiModelPicker,
  type AiModelPickerBridge,
  useAiPanelPrefs,
} from '@genoffice/ui'
import { useEffect, useRef, useState } from 'react'
import type {
  KeyboardEvent as ReactKeyboardEvent,
  PointerEvent as ReactPointerEvent,
  ReactElement,
} from 'react'
import {
  AgentLoop,
  skillContextVars,
  withAttachedSkills,
  type SkillVars,
} from '@genoffice/agent-core'
import {
  activeProvider,
  endpointSupportsVision,
  imageGenerationAvailable,
  type AiSettings,
} from '@genoffice/ai-provider/browser'
import {
  AiComposer,
  AiQueueStrip,
  AI_QUEUE_LABELS,
  AI_CONFIRM_LABELS,
  AiConfirmCard,
  AiScopeQuote,
  AiSkillChips,
  AiSkillsMenu,
  AiTypingIndicator,
  AiUsageChip,
  addTokenUsage,
  copyMarkdownRich,
  useAiSlash,
  useChatRunQueue,
  useToolConfirm,
  type AiScopeQuoteData,
  type AiTokenUsageView,
} from '@genoffice/ui'
import { aiLangDirective, t as tGlobal, useI18n } from '../i18n/locale'
import { createFileNav, Markdown } from '@genoffice/ui'
import sendEnterOn from '../assets/send-enter-on.png'
import sendEnterOff from '../assets/send-enter-off.png'
import sendStop from '../assets/send-stop.png'
import { createPdfSkill } from './pdf-skill'
import {
  collectPageImages,
  EMPTY_PAGE_IMAGE_PLAN,
  pageImageAnnotation,
  type PageImagePlan,
} from './page-images'
import { createElectronTransport } from './transport'
import { PDF_NAV_SCHEME, parsePdfNavHref } from './pdf-nav'
import type { FileOpConfirm, PdfAiDeps, PdfAppDeps } from './tools'

// Word-parity count (same as docs/markdown): Asian chars one by one + non-Asian words
const ASIAN_RE =
  /[ᄀ-ᇿ⺀-⿟、-〿぀-ヿ㄀-ㄯ㄰-㆏㇀-ㇿ㐀-䶿一-鿿가-힯豈-﫿！-｠￠-￦]|[\uD840-\uD87F][\uDC00-\uDFFF]/g
const NON_ASIAN_WORD_RE = /[A-Za-z0-9À-ɏ]+(?:['-][A-Za-z0-9À-ɏ]+)*/g

function countWords(text: string): number {
  return (text.match(ASIAN_RE) ?? []).length + (text.match(NON_ASIAN_WORD_RE) ?? []).length
}

const PANEL_WIDTH_KEY = 'pdf-ai-panel-width'

/** the user's dragged width (the only persisted value); null = never dragged → auto width */
function loadPanelWidth(): number | null {
  const saved = Number(localStorage.getItem(PANEL_WIDTH_KEY))
  // static bounds only — clamping against the window here would bake a
  // transiently small viewport into the restored preference
  return Number.isFinite(saved) && saved > 0
    ? Math.min(Math.max(saved, AI_PANEL_WIDTH_DEFAULT), AI_PANEL_WIDTH_MAX)
    : null
}

/** width actually displayed: the auto rule until the user drags, the clamped
 *  preference afterwards (the panel grows/shrinks with the window either way) */
function fitPanelWidth(preferred: number | null): number {
  return preferred === null
    ? aiPanelAutoWidth(window.innerWidth)
    : aiPanelClampWidth(preferred, window.innerWidth)
}

interface ToolActivity {
  name: string
  summary: string
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
  tools?: ToolActivity[]
  /** the passage this user message targeted, frozen at send */
  scope?: AiScopeQuoteData
  /** skills the user attached to this message from the composer menu */
  skills?: string[]
  /** tokens the provider billed for this reply, summed across its turns (absent when unreported) */
  usage?: AiTokenUsageView
  /** model reasoning streamed for this assistant message (kept after the run ends) */
  reasoning?: string
}

/** longest selection excerpt echoed on a user bubble */
const SCOPE_TEXT_MAX = 200

/** longest selection text offered to a skill body's ${SELECTION}; the core
 *  leaves the cap to the caller, and 500 chars is plenty for a model hint */
const SKILL_SELECTION_MAX = 500

type Phase = 'thinking' | 'replying' | 'working'

interface PendingConfirm {
  req: FileOpConfirm
  settle: (ok: boolean) => void
}

const MODEL_BRIDGE: AiModelPickerBridge = {
  getSettings: () => window.pdfApi.getAiSettings(),
  setSettings: (settings) => window.pdfApi.setAiSettings(settings),
  onSettingsChanged: (handler) => window.pdfApi.onAiSettingsChanged(handler),
  gskLoggedIn: () => window.pdfApi.gskStatus().then((s) => !!s?.loggedIn),
  openModelSettings: () => window.pdfApi.openAiModelSettings().catch(() => {}),
}

export function AiPanel({
  api,
  filePath,
  onCollapse,
  preset,
  onRunDone,
  onClearSelection,
}: {
  api: PdfAppDeps
  /** Absolute path of the open PDF (chat history is keyed to it) */
  filePath?: string
  onCollapse: () => void
  /** Ribbon AI buttons push a one-shot prompt; a new nonce triggers an auto-run */
  preset?: { text: string; nonce: number } | null
  /** Fired when a run that mutated the document finishes (drives the untitled-blank auto-save) */
  onRunDone?: () => void
  /** The × on the scope chip: drop the cached selection so runs target the whole document */
  onClearSelection?: () => void
}): ReactElement {
  const { lang, t } = useI18n()
  // AI panel display toggles (shell Settings, live-updated). Render paths read the
  // hook values; the once-built loop callbacks read showStreaming through a ref.
  const { showStreaming, showThinking, showToolCalls } = useAiPanelPrefs()
  const showStreamingRef = useRef(showStreaming)
  showStreamingRef.current = showStreaming
  const [chat, setChat] = useState<ChatEntry[]>([])
  const [prompt, setPrompt] = useState('')
  /** the composer textarea; the slash palette anchors above it */
  const inputRef = useRef<HTMLTextAreaElement | null>(null)
  const [attachedSkills, setAttachedSkills] = useState<string[]>([])
  const attachedSkillsRef = useRef(attachedSkills)
  attachedSkillsRef.current = attachedSkills
  const [busy, setBusy] = useState(false)
  /** messages queued while a reply runs; the pump re-runs them through send on settle */
  const msgQueue = useChatRunQueue<void>({
    busy,
    submit: (text) => {
      const loop = loopRef.current
      if (!loop || loop.busy) return false
      send(text, null)
      return true
    },
  })
  const [phase, setPhase] = useState<Phase>('thinking')
  /** the scope chip's expandable preview of the selected text */
  const [scopePreviewOpen, setScopePreviewOpen] = useState(false)
  const chatRef = useRef<HTMLDivElement>(null)
  const stickToBottomRef = useRef(true)
  /** which assistant message shows the copied checkmark (null = none) */
  const [copiedIdx, setCopiedIdx] = useState<number | null>(null)

  // ── Chat-history persistence (r142): same shared store Docs/Sheets use ──
  const chatIdsRef = useRef<{ projectId: string; chatId: string } | null>(null)
  /** current turn's streamed text; completed turns collect into runTextsRef */
  const segTextRef = useRef('')
  /** current turn's streamed reasoning (rendered under the showThinking pref) */
  const segReasoningRef = useRef('')
  /** whole-run accumulation: one stored assistant message per run (consecutive
      assistant rows would break restore() on strict-alternation providers) */
  const runTextsRef = useRef<string[]>([])
  const runToolsRef = useRef<ToolActivity[]>([])
  /** tokens billed so far in the in-flight run (summed across turns) */
  const runUsageRef = useRef<AiTokenUsageView | undefined>(undefined)
  /** tokens billed across the panel's session (the composer chip) */
  const [sessionUsage, setSessionUsage] = useState<AiTokenUsageView | undefined>(undefined)
  const chatStore = () =>
    (
      window as Window & {
        projectApi?: {
          resolveChat(args: {
            filePath: string | null
            tempChatId?: string
          }): Promise<{ projectId: string; chatId: string }>
          appendChat(args: {
            projectId: string
            chatId: string
            role: 'user' | 'assistant'
            text: string
            tools?: Array<{ name: string; summary: string; isError?: boolean; output?: string }>
            scope?: AiScopeQuoteData
          }): Promise<void>
          loadChat(args: { projectId: string; chatId: string; limit?: number }): Promise<
            Array<{
              role: 'user' | 'assistant'
              text: string
              tools?: Array<{ name: string; summary: string; isError?: boolean; output?: string }>
              scope?: AiScopeQuoteData
            }>
          >
          rebindChat(args: {
            projectId: string
            tempChatId: string
            newFilePath: string
          }): Promise<{ projectId: string; chatId: string } | null>
        }
      }
    ).projectApi
  const persistMessage = (
    role: 'user' | 'assistant',
    text: string,
    tools?: ToolActivity[],
    scope?: AiScopeQuoteData,
  ): void => {
    const ids = chatIdsRef.current
    const store = chatStore()
    if (!ids || !store || (!text && !tools?.length)) return
    void store
      .appendChat({
        projectId: ids.projectId,
        chatId: ids.chatId,
        role,
        text,
        ...(tools && tools.length > 0
          ? {
              tools: tools.map((tool) => ({
                name: tool.name,
                summary: tool.summary,
                isError: tool.isError,
                output: tool.output,
              })),
            }
          : {}),
        ...(scope ? { scope } : {}),
      })
      .catch(() => {
        /* silent */
      })
  }
  /** persist the whole run as ONE assistant message (docs parity: restore()
      feeds these back verbatim, and providers require user/assistant
      alternation; cancelled runs persist nothing — the unanswered user
      message is filtered out by restore()) */
  const persistRun = (): void => {
    const texts = [...runTextsRef.current, segTextRef.current].filter(Boolean)
    const tools = runToolsRef.current
    segTextRef.current = ''
    runTextsRef.current = []
    runToolsRef.current = []
    if (texts.length > 0 || tools.length > 0) {
      persistMessage('assistant', texts.join('\n\n'), tools)
    }
  }
  useEffect(() => {
    const store = chatStore()
    if (!store) return
    const tempChatId = `unsaved-${Date.now()}`
    void store
      .resolveChat({ filePath: filePath || null, tempChatId })
      .then((ids) => {
        chatIdsRef.current = ids
        return store.loadChat({ projectId: ids.projectId, chatId: ids.chatId, limit: 200 })
      })
      .then((msgs) => {
        if (msgs.length === 0) return
        setChat((prev) => [
          ...msgs.map((m) => ({
            role: m.role,
            text: m.text,
            tools: m.tools?.map((tool) => ({
              name: tool.name,
              summary: tool.summary,
              isError: tool.isError,
              output: tool.output,
            })),
            ...(m.scope ? { scope: m.scope } : {}),
          })),
          ...prev,
        ])
        // follow-ups after reopening continue the previous conversation
        loopRef.current?.restore(msgs.map((m) => ({ role: m.role, text: m.text })))
      })
      .catch(() => {
        /* history load failures are silent */
      })
    // eslint-disable-next-line react-hooks/exhaustive-deps -- mount-only, like Docs
  }, [])
  /** blank/generated PDFs get a path on first save: bind the unsaved-* history to it */
  useEffect(() => {
    const ids = chatIdsRef.current
    const store = chatStore()
    if (!store || !ids || !filePath || !ids.chatId.startsWith('unsaved-')) return
    void store
      .rebindChat({ projectId: ids.projectId, tempChatId: ids.chatId, newFilePath: filePath })
      .then((rebound) => {
        if (rebound?.chatId) chatIdsRef.current = rebound
      })
      .catch(() => {
        /* silent */
      })
  }, [filePath])
  // preferred = the user's dragged width (the only value persisted; null = auto);
  // panelWidth = what fits the current window. Deriving the display width from the
  // preference means a transiently small window never permanently shrinks the panel.
  const preferredWidthRef = useRef<number | null>(loadPanelWidth())
  const [panelWidth, setPanelWidth] = useState(() => fitPanelWidth(preferredWidthRef.current))
  const [resizing, setResizing] = useState(false)
  const asideRef = useRef<HTMLElement>(null)

  // The .ai-dock wrapper owns the animated width (docs-style 180ms slide);
  // it tracks the resizable panel width through this variable
  useEffect(() => {
    const dock = asideRef.current?.closest('.ai-dock') as HTMLElement | null
    dock?.style.setProperty('--ai-panel-width', `${panelWidth}px`)
  }, [panelWidth])
  const settingsRef = useRef<AiSettings | null>(null)

  /** gsk login state for the cloud-tools gate (refreshed on mount and window focus) */
  const gskLoggedInRef = useRef(false)
  useEffect(() => {
    let alive = true
    const refresh = () => {
      void window.pdfApi
        ?.gskStatus()
        .then((s) => {
          if (alive) gskLoggedInRef.current = !!s?.loggedIn
        })
        .catch(() => {})
    }
    refresh()
    window.addEventListener('focus', refresh)
    return () => {
      alive = false
      window.removeEventListener('focus', refresh)
    }
  }, [])

  // No usable chat provider and no Genspark login: the composer gates sending
  // until setup, mirroring the docs panel's empty-state CTA rule. Refreshed
  // with the model chip (mount, window focus, ai-settings broadcasts); the
  // Promise.resolve wrapping keeps a throwing bridge harmless in tests.
  const [aiSetupNeeded, setAiSetupNeeded] = useState(false)
  useEffect(() => {
    const refresh = () => {
      void Promise.all([
        Promise.resolve()
          .then(() => MODEL_BRIDGE.getSettings())
          .catch(() => null),
        Promise.resolve()
          .then(() => MODEL_BRIDGE.gskLoggedIn?.() ?? false)
          .catch(() => false),
      ]).then(([settings, loggedIn]) => {
        setAiSetupNeeded(
          settings != null &&
            !loggedIn &&
            activeProvider(settings) === 'genspark' &&
            !(settings.customEndpoints ?? []).some((e) => e.baseUrl?.trim() && e.model?.trim()),
        )
      })
    }
    refresh()
    window.addEventListener('focus', refresh)
    let off: (() => void) | undefined
    try {
      off = MODEL_BRIDGE.onSettingsChanged?.(refresh)
    } catch {
      off = undefined
    }
    return () => {
      window.removeEventListener('focus', refresh)
      off?.()
    }
  }, [])
  const langRef = useRef(lang)
  langRef.current = lang
  const apiRef = useRef(api)
  apiRef.current = api
  const onRunDoneRef = useRef(onRunDone)
  onRunDoneRef.current = onRunDone
  /** Any tool in the current run reported mutated: true */
  const runMutatedRef = useRef(false)
  /** Annotation lines for the images attached to the in-flight run (vision injection);
      null while the run carries none. buildContext reads it through the skill deps. */
  const visionAttachmentRef = useRef<string | null>(null)

  /** Confirmation card for irreversible file operations; one at a time, a second request is refused */
  const [fileOpConfirm, setFileOpConfirm] = useState<FileOpConfirm | null>(null)
  const confirmRef = useRef<PendingConfirm | null>(null)
  /** the approval gate for the document-editing tools (risk:'high' in tools.ts) */
  const confirmTool = useToolConfirm()
  const requestFileOpConfirm = (req: FileOpConfirm, signal?: AbortSignal): Promise<boolean> =>
    new Promise((resolve) => {
      if (confirmRef.current || signal?.aborted) {
        resolve(false)
        return
      }
      const pending: PendingConfirm = {
        req,
        settle: (ok) => {
          if (confirmRef.current !== pending) return
          confirmRef.current = null
          setFileOpConfirm(null)
          resolve(ok)
        },
      }
      confirmRef.current = pending
      setFileOpConfirm(req)
      signal?.addEventListener('abort', () => pending.settle(false), { once: true })
    })
  // another document or the panel going away answers a pending card with "no"
  useEffect(() => () => confirmRef.current?.settle(false), [filePath])

  const patchLast = (patch: Partial<ChatEntry> | ((last: ChatEntry) => Partial<ChatEntry>)) => {
    setChat((prev) => {
      const next = [...prev]
      const last = next[next.length - 1]
      if (!last || last.role !== 'assistant') return prev
      next[next.length - 1] = { ...last, ...(typeof patch === 'function' ? patch(last) : patch) }
      return next
    })
  }

  // The loop is built once; every mutable value goes through a ref getter
  const loopRef = useRef<AgentLoop | null>(null)
  if (!loopRef.current) {
    const deps: PdfAiDeps = {
      doc: () => apiRef.current.doc(),
      fileName: () => apiRef.current.fileName(),
      pageCount: () => apiRef.current.pageCount(),
      currentPage: () => apiRef.current.currentPage(),
      readOnly: () => apiRef.current.readOnly(),
      ocrText: (idx) => apiRef.current.ocrText(idx),
      selection: () => apiRef.current.selection(),
      pendingSummary: () => apiRef.current.pendingSummary(),
      outline: () => apiRef.current.outline(),
      searchIndex: () => apiRef.current.searchIndex(),
      isDeleted: (i) => apiRef.current.isDeleted(i),
      gotoPage: (p) => apiRef.current.gotoPage(p),
      addMarkup: (type, idx, rects, color) => apiRef.current.addMarkup(type, idx, rects, color),
      annotationSummary: () => apiRef.current.annotationSummary(),
      pageImageAttachment: () => visionAttachmentRef.current,
      createDocument: (request) => apiRef.current.createDocument(request),
      confirmFileOp: requestFileOpConfirm,
      insertBlankPage: (afterVis) => apiRef.current.insertBlankPage(afterVis),
      setPageSize: (w, h) => apiRef.current.setPageSize(w, h),
      cropPages: (vis, rect) => apiRef.current.cropPages(vis, rect),
      replacePages: (vis) => apiRef.current.replacePages(vis),
      extractPages: (vis) => apiRef.current.extractPages(vis),
      splitPdf: (n) => apiRef.current.splitPdf(n),
      splitPages: (n) => apiRef.current.splitPages(n),
      mergePages: (n, direction, separator) => apiRef.current.mergePages(n, direction, separator),
      stamps: () => apiRef.current.stamps(),
      setStamps: (cfg) => apiRef.current.setStamps(cfg),
      annotationsOn: (idx) => apiRef.current.annotationsOn(idx),
      addNote: (idx, at, contents, color) => apiRef.current.addNote(idx, at, contents, color),
      findNoteRoot: (idx, key) => apiRef.current.findNoteRoot(idx, key),
      replyToThread: (idx, root, contents) => apiRef.current.replyToThread(idx, root, contents),
      editNote: (idx, item, contents) => apiRef.current.editNote(idx, item, contents),
      deleteMarkups: (idx, keys) => apiRef.current.deleteMarkups(idx, keys),
      deleteNoteThread: (idx, root) => apiRef.current.deleteNoteThread(idx, root),
      editText: (input) => apiRef.current.editText(input),
      moveTextBlock: (idx, block, d) => apiRef.current.moveTextBlock(idx, block, d),
      insertText: (input) => apiRef.current.insertText(input),
      addFormMark: (idx, kind, rect) => apiRef.current.addFormMark(idx, kind, rect),
      formWidgets: (idx) => apiRef.current.formWidgets(idx),
      pageBoxes: (idx) => apiRef.current.pageBoxes(idx),
      textInserts: () => apiRef.current.textInserts(),
      updateTextInsert: (id, edit) => apiRef.current.updateTextInsert(id, edit),
      moveTextInsert: (id, origin) => apiRef.current.moveTextInsert(id, origin),
      deleteTextInsert: (id) => apiRef.current.deleteTextInsert(id),
      editFonts: () => apiRef.current.editFonts(),
      formEdits: () => apiRef.current.formEdits(),
      applyOps: (ops, opts) => apiRef.current.applyOps(ops, opts),
      metadata: () => apiRef.current.metadata(),
      pageOrder: () => apiRef.current.pageOrder(),
      pageGeom: (idx) => apiRef.current.pageGeom(idx),
      listImages: () => apiRef.current.listImages(),
      isImageClaimed: (ref) => apiRef.current.isImageClaimed(ref),
      insertImage: (idx, png, rect, layer) => apiRef.current.insertImage(idx, png, rect, layer),
      transformImage: (ref, rect, layer, quarterTurns) =>
        apiRef.current.transformImage(ref, rect, layer, quarterTurns),
      replaceImage: (ref, png) => apiRef.current.replaceImage(ref, png),
      bakeImage: (ref, op, signal) => apiRef.current.bakeImage(ref, op, signal),
      deleteImage: (ref) => apiRef.current.deleteImage(ref),
      searchImages: (query, max) => apiRef.current.searchImages(query, max),
      generateImage: (op) => apiRef.current.generateImage(op),
      imageGenAvailable: () =>
        imageGenerationAvailable(settingsRef.current, gskLoggedInRef.current),
      fetchImage: (url) => apiRef.current.fetchImage(url),
    }
    loopRef.current = new AgentLoop({
      transport: createElectronTransport(
        () => settingsRef.current!,
        (text) => {
          // reasoning rides on the message (the showThinking pref gates rendering)
          segReasoningRef.current += text
          patchLast({ reasoning: segReasoningRef.current })
        },
      ),
      skill: createPdfSkill(deps),
      // the document-editing tools (risk:'high', see the policy note in
      // tools.ts) stop at the inline allow/deny card; a denial feeds the model
      // a do-not-retry result so it can answer without editing. The
      // file-level operations keep their own richer card (confirmFileOp).
      confirmTool: (call) => confirmTool.confirm(call),
      systemSuffix: () => aiLangDirective(langRef.current),
      events: {
        onText: (text) => {
          setPhase('replying')
          segTextRef.current = text
          // showStreaming off: the reply stays pending (typing indicator) and is
          // revealed whole when its turn completes
          if (showStreamingRef.current) patchLast({ text })
        },
        onToolExecuted: ({ call, execution }) => {
          setPhase('working')
          if (execution.mutated) runMutatedRef.current = true
          runToolsRef.current.push({
            name: call.name,
            summary: execution.summary,
            isError: execution.isError,
            output: execution.output?.slice(0, 2000),
          })
          patchLast((last) => ({
            tools: [
              ...(last.tools ?? []),
              {
                name: call.name,
                summary: execution.summary,
                isError: execution.isError,
                output: execution.output?.slice(0, 2000),
              },
            ],
          }))
        },
        onTurnEnd: () => {
          setPhase('thinking')
          runTextsRef.current.push(segTextRef.current)
          // showStreaming off: the held turn text is revealed now that its turn completed
          if (!showStreamingRef.current) patchLast({ text: segTextRef.current })
          segTextRef.current = ''
          segReasoningRef.current = ''
          patchLast({ streaming: false })
          setChat((prev) => [...prev, { role: 'assistant', text: '', streaming: true }])
        },
        onUsage: (usage) => {
          runUsageRef.current = addTokenUsage(runUsageRef.current, usage)
          setSessionUsage((prev) => addTokenUsage(prev, usage))
          patchLast({ usage: runUsageRef.current })
        },
        onDone: ({ text, cancelled, turnLimit, truncated }) => {
          const base = turnLimit
            ? [text, tGlobal('aiTurnLimit')].filter(Boolean).join('\n\n')
            : text || (cancelled ? tGlobal('aiStopped') : '')
          // finish_reason=length with no prose (a reasoning model that spent the whole
          // output budget thinking) must say so instead of showing the bare "(no reply)",
          // which reads like the assistant ignored the user — same handling as docs
          const final = truncated
            ? [base, tGlobal('aiTruncatedNote')].filter(Boolean).join('\n\n')
            : base
          if (cancelled) {
            segTextRef.current = ''
            segReasoningRef.current = ''
            runTextsRef.current = []
            runToolsRef.current = []
          } else {
            segTextRef.current = final || segTextRef.current
            persistRun()
          }
          patchLast((last) => ({
            streaming: false,
            text: final || (last.tools?.length ? last.text : tGlobal('aiNoReply')),
          }))
          setBusy(false)
          if (runMutatedRef.current) {
            runMutatedRef.current = false
            onRunDoneRef.current?.()
          }
        },
        onError: (error) => {
          setChat((prev) => {
            const next = [...prev]
            // the loop rolled this run's user message out of the model context — surface that
            for (let i = next.length - 1; i >= 0; i--) {
              const entry = next[i]!
              if (entry.role === 'user') {
                next[i] = { ...entry, undelivered: true }
                break
              }
            }
            const last = next.at(-1)
            if (last?.role === 'assistant') {
              next[next.length - 1] = { ...last, streaming: false, text: error, isError: true }
            }
            return next
          })
          setBusy(false)
        },
      },
    })
  }

  useEffect(() => {
    if (stickToBottomRef.current) {
      chatRef.current?.scrollTo({ top: chatRef.current.scrollHeight })
    }
  }, [chat, busy, fileOpConfirm])

  const onChatScroll = (): void => {
    const el = chatRef.current
    if (!el) return
    stickToBottomRef.current = el.scrollHeight - el.scrollTop - el.clientHeight < 48
  }
  /** Send-time vision collection over the open file's content-stream images.
      Failures degrade to a no-image run: the model still gets its text context. */
  const collectVisionPlan = (): Promise<PageImagePlan> => {
    const path = filePath
    if (!path) return Promise.resolve(EMPTY_PAGE_IMAGE_PLAN)
    return collectPageImages({
      listImages: () => window.pdfApi.listPageImages(path),
      imageBytes: (ref) => window.pdfApi.pageImageBytes({ path, ...ref }),
      isDeleted: (i) => apiRef.current.isDeleted(i),
    })
  }

  /** retryScope: null = a retry that had no scope; undefined = capture the live selection */
  const send = (
    text: string,
    retryScope?: AiScopeQuoteData | null,
    /** only the composer's own send takes the attached skills along */
    attachSkills = false,
  ): void => {
    const instruction = text.trim()
    const loop = loopRef.current
    if (!instruction || !loop || loop.busy) return
    // a new run owns the stream: the previous reply's confirm card is history
    confirmTool.clear()
    stickToBottomRef.current = true
    const scope = retryScope !== undefined ? (retryScope ?? undefined) : selectionScopeQuote()
    persistMessage('user', instruction, undefined, scope)
    segTextRef.current = ''
    segReasoningRef.current = ''
    runTextsRef.current = []
    runToolsRef.current = []
    runUsageRef.current = undefined
    // skills ride with this one message only: the chips clear like a draft does
    const sentSkills = attachSkills ? attachedSkillsRef.current : []
    // context vars freeze at send like the scope quote does — a selection made
    // while the skill bodies load must not leak into the run that already left
    const sentVars = sentSkills.length > 0 ? skillVars() : undefined
    if (sentSkills.length > 0) setAttachedSkills([])
    setChat((prev) => [
      ...prev,
      {
        role: 'user',
        text: instruction,
        ...(sentSkills.length > 0 ? { skills: sentSkills } : {}),
        ...(scope ? { scope } : {}),
      },
      { role: 'assistant', text: '', streaming: true },
    ])
    setPrompt('')
    setBusy(true)
    setPhase('thinking')
    runMutatedRef.current = false
    // Vision injection: cleared per run; set below only when the endpoint supports
    // vision AND the collection found images — otherwise the context stays text-only
    visionAttachmentRef.current = null
    void (async () => {
      try {
        settingsRef.current = await window.pdfApi.getAiSettings()
        const settings = settingsRef.current
        // Vision gate: text-only endpoints never collect (zero behavior change)
        const plan = settings && endpointSupportsVision(settings) ? await collectVisionPlan() : null
        visionAttachmentRef.current = plan ? pageImageAnnotation(plan) : null
        await loop.run(await withSkillBodies(instruction, sentSkills, sentVars), plan?.images)
      } catch (err) {
        patchLast({
          streaming: false,
          text: err instanceof Error ? err.message : String(err),
          isError: true,
        })
        setBusy(false)
      }
    })()
  }

  const stop = (): void => {
    msgQueue.hold()
    // a confirmation left open when Stop is pressed counts as denied
    confirmTool.settleAll(false)
    loopRef.current?.cancel()
  }

  /** Selection text for a skill body's ${SELECTION}: the mouseup-cached
   *  selection, whitespace-collapsed and capped where the caller must. */
  const skillSelection = (): string | undefined => {
    const sel = apiRef.current.selection()?.text.replace(/\s+/g, ' ').trim()
    if (!sel) return undefined
    return sel.length > SKILL_SELECTION_MAX ? `${sel.slice(0, SKILL_SELECTION_MAX)}…` : sel
  }

  /** The standard context variables offered to attached skill bodies, sourced
   *  at send time from the same surface the tools read: the open file's name
   *  (omitted when no document is open, never an empty substitution), the
   *  cached selection, and the UI language. The product keeps the name "PDF"
   *  in every language, so it needs no translation. */
  const skillVars = (): SkillVars =>
    skillContextVars({
      lang: langRef.current,
      app: 'PDF',
      fileName: apiRef.current.fileName() || undefined,
      selection: skillSelection(),
    })

  /** the attached skills' full text in front of the request; a skill that vanished is skipped */
  const withSkillBodies = async (
    instruction: string,
    names: readonly string[],
    vars?: SkillVars,
  ) => {
    if (names.length === 0) return instruction
    const bodies = await Promise.all(
      names.map(async (name) => {
        try {
          return { name, body: await window.pdfApi.skillText(name) }
        } catch {
          return null
        }
      }),
    )
    return withAttachedSkills(
      instruction,
      bodies.filter((b): b is { name: string; body: string } => b !== null),
      vars,
    )
  }

  const listSkillsForMenu = async () =>
    (await window.pdfApi.skillsList()).map(({ name, description }) => ({ name, description }))
  const toggleSkill = (name: string) =>
    setAttachedSkills((prev) =>
      prev.includes(name) ? prev.filter((n) => n !== name) : [...prev, name],
    )

  /** the palette's pick: strike the `/query` token, keep the caret where it
   *  stood, and attach the skill (a pick must never detach an attached one) */
  const attachSlashSkill = (name: string, range: { start: number; end: number }): void => {
    setPrompt((prev) => prev.slice(0, range.start) + prev.slice(range.end))
    setAttachedSkills((prev) => (prev.includes(name) ? prev : [...prev, name]))
    requestAnimationFrame(() => {
      inputRef.current?.focus()
      inputRef.current?.setSelectionRange(range.start, range.start)
    })
  }
  const slash = useAiSlash({
    lang,
    listSkills: listSkillsForMenu,
    inputRef,
    value: prompt,
    onAttachSkill: attachSlashSkill,
  })

  /** Enter while a reply runs: the draft joins the queue above the composer */
  const enqueueDraft = () => {
    const text = prompt.trim()
    if (!text) return
    msgQueue.enqueue(text, undefined)
    setPrompt('')
  }

  // One-click AI actions from the ribbon / Ask popover; while a run is active the
  // preset lands in the composer instead of being dropped silently (markdown parity)
  useEffect(() => {
    if (!preset) return
    if (loopRef.current?.busy) setPrompt(preset.text)
    else send(preset.text)
    // eslint-disable-next-line react-hooks/exhaustive-deps -- run once per nonce
  }, [preset?.nonce])

  // Re-derive the display width on window resize: the auto rule tracks the
  // window until the user drags, a dragged preference just re-clamps into it
  useEffect(() => {
    const onResize = (): void => setPanelWidth(fitPanelWidth(preferredWidthRef.current))
    window.addEventListener('resize', onResize)
    return () => window.removeEventListener('resize', onResize)
  }, [])

  const resizeCleanupRef = useRef<(() => void) | null>(null)
  useEffect(() => () => resizeCleanupRef.current?.(), [])

  /** Drag the inner panel edge to resize from the selected window side. */
  const startResize = (e: ReactPointerEvent<HTMLDivElement>): void => {
    e.preventDefault()
    const resizer = e.currentTarget
    setResizing(true)
    document.body.style.cursor = 'col-resize'
    document.body.style.userSelect = 'none'
    const onMove = (ev: PointerEvent): void => {
      const w = aiPanelClampWidth(aiPanelWidthAtPointer(ev.clientX), window.innerWidth)
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
      const preferred = preferredWidthRef.current
      if (preferred !== null) localStorage.setItem(PANEL_WIDTH_KEY, String(Math.round(preferred)))
    }
    resizeCleanupRef.current = cleanup
    window.addEventListener('pointermove', onMove)
    window.addEventListener('pointerup', cleanup)
    window.addEventListener('pointercancel', cleanup)
    // lostpointercapture also fires if the resizer is unmounted mid-drag (panel collapse)
    resizer.addEventListener('lostpointercapture', cleanup)
    resizer.setPointerCapture(e.pointerId)
  }

  const typingLabel =
    phase === 'replying' ? t('aiReplying') : phase === 'working' ? t('aiWorking') : t('aiThinking')

  // scope chip data, read per render (App re-renders on every selection change)
  const scopeSel = api.selection()
  const hasScopeSelection = !!scopeSel && scopeSel.text.trim().length > 0

  const selectionScopeQuote = (): AiScopeQuoteData | undefined => {
    const sel = api.selection()
    const text = sel?.text.replace(/\s+/g, ' ').trim() ?? ''
    if (!sel || !text) return undefined
    return {
      label: t('aiScopeSelection', {
        page: sel.lastPage > sel.page ? `${sel.page}-${sel.lastPage}` : sel.page,
        words: countWords(text),
      }),
      text: text.length > SCOPE_TEXT_MAX ? `${text.slice(0, SCOPE_TEXT_MAX)}…` : text,
    }
  }

  // the selection can vanish without the × (click-away, another file): close the preview too
  useEffect(() => {
    if (!hasScopeSelection) setScopePreviewOpen(false)
  }, [hasScopeSelection])

  /** [p.N](pdfnav://page/N) links in replies scroll the reading view to that page */
  const pdfNav = {
    scheme: PDF_NAV_SCHEME,
    onNavigate: (href: string) => {
      const page = parsePdfNavHref(href)
      if (page !== null) apiRef.current.gotoPage(page)
    },
  }
  // [name](filenav:///abs/path) citations open the cited file in its own app
  const fileNav = createFileNav((path) => {
    void window.pdfApi.openSourcePath(path)
  })

  const copyMessage = (text: string, idx: number) => {
    // rich clipboard first (text/html + text/plain), plain-text fallback inside
    void copyMarkdownRich(text).finally(() => {
      setCopiedIdx(idx)
      window.setTimeout(() => setCopiedIdx((cur) => (cur === idx ? null : cur)), 1200)
    })
  }

  return (
    <aside
      ref={asideRef}
      className={`copilot${resizing ? ' ai-panel-resizing' : ''}`}
      style={{ width: '100%' }}
      dir={lang === 'ar' || lang === 'he' ? 'rtl' : undefined}
    >
      <div
        className="ai-panel-resizer"
        onPointerDown={startResize}
        role="separator"
        aria-orientation="vertical"
        aria-label={t('ribbonAiAssistant')}
      />
      <header className="ai-panel-header">
        <span className="ai-panel-title">
          <GensparkMark size={22} />
          Genspark
        </span>
        <div className="ai-panel-header-actions">
          <AiPanelSideButton
            lang={lang}
            onMove={(side) => window.pdfApi.setAiPanelPrefs({ side })}
          />
          {chat.length > 0 && (
            <button
              className="ai-header-btn"
              onClick={() => {
                stop()
                loopRef.current?.reset()
                msgQueue.clear()
                setBusy(false)
                setChat([])
                setAttachedSkills([])
              }}
              data-tip={t('aiNewChat')}
              aria-label={t('aiNewChat')}
            >
              <IconNewChat />
            </button>
          )}
          <button
            className="ai-header-btn ai-panel-collapse"
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
            <div className="ai-chat-empty-title">{t('aiEmptyTitle')}</div>
            <div className="ai-chat-empty-body">{t('aiEmptyBody')}</div>
            <div className="ai-quick-actions">
              <button
                className="ai-quick-btn"
                onClick={() =>
                  send(t(hasScopeSelection ? 'aiQuickSummarySelPrompt' : 'aiQuickSummaryPrompt'))
                }
              >
                {t('aiQuickSummary')}
              </button>
              <button
                className="ai-quick-btn"
                onClick={() =>
                  send(
                    t(hasScopeSelection ? 'aiQuickKeyPointsSelPrompt' : 'aiQuickKeyPointsPrompt'),
                  )
                }
              >
                {t('aiQuickKeyPoints')}
              </button>
            </div>
          </div>
        )}
        {chat.map((entry, i) => {
          if (entry.role === 'user') {
            return (
              <div key={i} className="ai-msg ai-msg-user">
                {entry.scope && <AiScopeQuote scope={entry.scope} />}
                {entry.skills && <AiSkillChips lang={lang} names={entry.skills} />}
                <span dir="auto">{entry.text}</span>
                {entry.undelivered && (
                  <div className="ai-msg-undelivered">
                    {t('aiUndelivered')}
                    {!busy && (
                      <button
                        className="ai-retry-btn"
                        onClick={() => send(entry.text, entry.scope ?? null)}
                      >
                        {t('aiRetry')}
                      </button>
                    )}
                  </div>
                )}
              </div>
            )
          }
          const hasTools = showToolCalls && (entry.tools?.length ?? 0) > 0
          const showReasoning = showThinking && !!entry.reasoning
          if (!entry.text && !hasTools && !showReasoning) return null
          return (
            <div
              key={i}
              className={`ai-msg ai-msg-assistant${entry.isError ? ' ai-msg-error' : ''}`}
            >
              {showReasoning && (
                <ReasoningBlock text={entry.reasoning!} streaming={entry.streaming} />
              )}
              {hasTools && <ToolChipList tools={entry.tools!} />}
              {entry.usage && (
                <div className="ai-message-usage">
                  <AiUsageChip usage={entry.usage} lang={lang} />
                </div>
              )}
              {entry.text && (
                <div dir="auto">
                  <Markdown
                    text={entry.text}
                    navs={[pdfNav, fileNav]}
                    codeCopyLabel={t('aiCopyCode')}
                  />
                </div>
              )}
              {/* Reply actions: copy with a 1.2s checkmark flip (docs parity), shown
                  once the reply finished streaming and holds text to copy */}
              {!entry.streaming && entry.text && (
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
                </div>
              )}
            </div>
          )
        })}
        {confirmTool.card && (
          <div className="ai-msg ai-msg-assistant">
            <AiConfirmCard
              labels={AI_CONFIRM_LABELS[lang]}
              card={confirmTool.card}
              onAllow={confirmTool.allow}
              onDeny={confirmTool.deny}
            />
          </div>
        )}
        {fileOpConfirm && (
          <div className="ai-confirm-card" role="group" aria-label={t('aiFileOpConfirmTitle')}>
            <div className="ai-confirm-title">{t('aiFileOpConfirmTitle')}</div>
            <div className="ai-confirm-summary">{fileOpConfirm.summary}</div>
            {fileOpConfirm.detail && (
              <div className="ai-confirm-detail">{fileOpConfirm.detail}</div>
            )}
            <div className="ai-confirm-warning">{t('aiFileOpConfirmWarning')}</div>
            <div className="ai-confirm-actions">
              <button
                type="button"
                className="pdf-modal-btn"
                onClick={() => confirmRef.current?.settle(false)}
              >
                {t('cancel')}
              </button>
              <button
                type="button"
                className="pdf-modal-btn primary"
                onClick={() => confirmRef.current?.settle(true)}
              >
                {t('aiFileOpConfirm')}
              </button>
            </div>
          </div>
        )}
        {/* In-progress state: a standalone three-dot row at the end of the stream, kept until done */}
        {busy && <AiTypingIndicator label={typingLabel} />}
      </div>

      <div
        className="ai-composer"
        onKeyDownCapture={(e: ReactKeyboardEvent) => {
          // the shared AiComposer owns the textarea's keydown: the palette
          // runs first here in capture, and a key it consumed is stopped
          // before the composer's Enter-send/Escape-stop can see it
          if (slash.onKeyDown(e as ReactKeyboardEvent<HTMLTextAreaElement>)) e.stopPropagation()
        }}
      >
        {slash.element}
        <AiComposer
          queueStrip={
            <AiQueueStrip
              items={msgQueue.queued}
              labels={AI_QUEUE_LABELS[lang]}
              onUpdate={msgQueue.update}
              onRemove={msgQueue.remove}
              onClear={msgQueue.clear}
              paused={msgQueue.paused}
              onTogglePause={msgQueue.togglePaused}
              onMove={msgQueue.move}
            />
          }
          onQueue={enqueueDraft}
          queuePlaceholder={AI_QUEUE_LABELS[lang].queuePlaceholder}
          value={prompt}
          textareaRef={inputRef}
          busy={busy}
          header={
            (hasScopeSelection || attachedSkills.length > 0) && (
              <>
                <AiSkillChips lang={lang} names={attachedSkills} onRemove={toggleSkill} />
                {hasScopeSelection && (
                  <div className="ai-scope-row">
                    <span className="ai-scope-hint">
                      <button
                        className="ai-scope-label"
                        onClick={() => setScopePreviewOpen((v) => !v)}
                        aria-expanded={scopePreviewOpen}
                        data-tip={t('aiScopeSelectionTip')}
                      >
                        {t('aiScopeSelection', {
                          page:
                            scopeSel!.lastPage > scopeSel!.page
                              ? `${scopeSel!.page}-${scopeSel!.lastPage}`
                              : scopeSel!.page,
                          words: countWords(scopeSel!.text),
                        })}
                      </button>
                      <button
                        className="ai-scope-clear"
                        onClick={() => {
                          setScopePreviewOpen(false)
                          onClearSelection?.()
                        }}
                        data-tip={t('aiScopeClearTitle')}
                        aria-label={t('aiScopeClearTitle')}
                      >
                        <svg width="10" height="10" viewBox="0 0 16 16" aria-hidden>
                          <path
                            d="M4 4l8 8M12 4l-8 8"
                            fill="none"
                            stroke="currentColor"
                            strokeWidth="1.8"
                            strokeLinecap="round"
                          />
                        </svg>
                      </button>
                    </span>
                    {scopePreviewOpen && (
                      <div className="ai-scope-preview">
                        {scopeSel!.text.length > 400
                          ? `${scopeSel!.text.slice(0, 400)}…`
                          : scopeSel!.text}
                      </div>
                    )}
                  </div>
                )}
              </>
            )
          }
          placeholder={t('aiComposerPlaceholder')}
          hintIdle={t('aiHintIdle')}
          hintBusy={t('aiHintBusy')}
          setupRequired={aiSetupNeeded}
          setupHint={t('aiSetupNeeded')}
          sendLabel={t('aiSend')}
          stopLabel={t('aiStop')}
          iconOnly
          footerStart={
            <>
              <AiModelPicker bridge={MODEL_BRIDGE} lang={lang} />
              {sessionUsage && <AiUsageChip usage={sessionUsage} session lang={lang} />}
              <AiSkillsMenu
                lang={lang}
                listSkills={listSkillsForMenu}
                selected={attachedSkills}
                onToggle={toggleSkill}
                onManage={() => void window.pdfApi.openSkillsSettings?.().catch(() => {})}
              />
            </>
          }
          sendIconEnabled={<img src={sendEnterOn} alt="" aria-hidden />}
          sendIconDisabled={<img src={sendEnterOff} alt="" aria-hidden />}
          stopIcon={<img src={sendStop} alt="" aria-hidden />}
          onChange={setPrompt}
          onSend={() => send(prompt, undefined, true)}
          onStop={stop}
        />
      </div>
    </aside>
  )
}

/** Tool row list (unified with docs/slides/sheets): dot + summary, expandable details when there's output */
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

  const open = userOpen ?? false
  const label = tr('aiWorkedSteps', { n: tools.length })

  return (
    <div className="ai-work-group">
      <button
        type="button"
        className={`ai-work-group-summary`}
        aria-expanded={open}
        onClick={() => setUserOpen(!open)}
      >
        <span className="ai-work-group-label">{label}</span>
        <span className={`ai-tool-chip-caret${open ? ' open' : ''}`} aria-hidden>
          ›
        </span>
      </button>
      <div className={`ai-work-group-body${open ? ' open' : ''}`}>
        <div className="ai-work-group-body-inner">
          {tools.map((tool, j) => {
            const hasOutput = !!tool.output
            const isOpen = expanded.has(j)
            const stepStatus = tool.isError ? 'error' : 'done'
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

/** Reasoning block: the model's thinking attached to an assistant message.
 *  Live-expanded while the reply streams (the manual toggle always wins), a
 *  quiet collapsible after — same affordance language as the tool group. */
function ReasoningBlock({ text, streaming }: { text: string; streaming?: boolean }) {
  const { t } = useI18n()
  const [userOpen, setUserOpen] = useState<boolean | null>(null)

  const open = userOpen ?? !!streaming

  return (
    <div className="ai-reasoning">
      <button
        type="button"
        className="ai-reasoning-summary"
        aria-expanded={open}
        onClick={() => setUserOpen(!open)}
      >
        <span className="ai-reasoning-label">{t('aiReasoning')}</span>
        <span className={`ai-tool-chip-caret${open ? ' open' : ''}`} aria-hidden>
          ›
        </span>
      </button>
      <div className={`ai-reasoning-body${open ? ' open' : ''}`}>
        <div className="ai-reasoning-body-inner">
          <div className="ai-reasoning-pre" dir="auto">
            {text}
          </div>
        </div>
      </div>
    </div>
  )
}

function Svg({ children }: { children: React.ReactNode }): ReactElement {
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

/* Same glyph as the sheets IconCollapse (16×16 viewBox, 1.2/1.3 stroke), rendered at 15px */
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
      {/* Mirrored: the AI panel docks on the LEFT, so the divider and arrow point left */}
      <rect x="1.5" y="2.5" width="13" height="11" rx="1" />
      <path d="M5.5 2.5v11" />
      <path d="M12.5 8H8.1M9.8 5.9 7.7 8l2.1 2.1" strokeWidth="1.3" strokeLinejoin="round" />
    </svg>
  )
}

/** Genspark brand mark (rounded-square sparkle badge), inline so it renders
 * crisply at device resolution instead of going through <img> rasterization */
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
