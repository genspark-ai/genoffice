import { useEffect, useMemo, useRef, useState } from 'react'
import {
  AgentLoop,
  ATTACHMENT_IMAGE_EXTS,
  composeSkills,
  createAttachmentsSkill,
  createCliSkill,
  createDocumentSkill,
  createIpcTransport,
  createKnowledgeBaseSkill,
  RUN_GENOFFICE_TOOL,
  withAttachedSkills,
  type AgentImage,
  type AttachmentInfo,
  type CliContext,
  type CreateDocumentFormat,
  type KbFileInfo,
} from '@genoffice/agent-core'
import {
  AiComposer,
  AiModelPicker,
  type AiModelPickerBridge,
  AiSkillChips,
  AiSkillsMenu,
  AiTypingIndicator,
  AiUsageChip,
  addTokenUsage,
  createFileNav,
  Markdown,
  type AiTokenUsageView,
} from '@genoffice/ui'
import type { AiSettings } from '@genoffice/ai-provider'
import { activeProvider } from '@genoffice/ai-provider/browser'
import type { ChatMessage, ChatMeta, ProjectApi } from '@genoffice/project-store'
import { useI18n } from './locale'
import type { AttachmentAddResult } from '../../shared/home-api'
import { AskAttachments } from './AskAttachments'
import { artifactFromCall, retryInstruction, settleArtifact, type Artifact } from './ask-artifacts'
import { appendToolNote, toolNoteDetail, toolNoteKind, type ToolNote } from './ask-tool-notes'
import sendEnterOn from './assets/send-enter-on.png'
import sendEnterOff from './assets/send-enter-off.png'
import sendStop from './assets/send-stop.png'

/**
 * The home screen's knowledge-base ask panel: chat with the reader's starred
 * corpus without opening a document. First app-side consumer of the shell's
 * kb:* gateway — the loop, streaming and rendering reuse what every editor's
 * AI panel uses (AgentLoop + createIpcTransport + Markdown with a filenav
 * nav), so answers cite the folder and file they came from and clicking a
 * citation opens that file in the app that owns it.
 *
 * The conversation persists through project-store under a fixed chat id, the
 * one place in the app where a conversation belongs to no file: with no
 * `filePath` the store takes the caller's id verbatim, so the panel passes a
 * constant and finds its own transcript again on the next launch. An empty
 * corpus gets an onboarding hero instead of a bare hint: the only useful move
 * there is going to Recents and starring files, so the panel says so and
 * offers the jump as a button.
 *
 * That constant is also the limit of what exists so far: one transcript, no
 * list, no name, and nothing to start a second conversation. A reader with two
 * questions on two days gets one long thread — a list, per-conversation ids
 * and a name are the next layer, not this one.
 */

/** The one transcript the panel keeps. Stable across launches by design. */
const CHAT_ID = 'kb-ask'

const TOOL_NOTE_LABEL = {
  search: 'askToolSearch',
  read: 'askToolRead',
  create: 'askToolCreate',
  cli: 'askToolCli',
} as const

interface AskMessage {
  role: 'user' | 'assistant'
  text: string
  toolNotes: ToolNote[]
  /** the files this answer set out to build, with their progress (task card) */
  artifacts: Artifact[]
  streaming?: boolean
  isError?: boolean
  /** tokens the provider billed for this answer, summed across its turns (absent when unreported) */
  usage?: AiTokenUsageView
  /** skills the user attached to this question from the composer menu */
  skills?: string[]
  /** files the user attached to this question */
  attachments?: AttachmentInfo[]
  /** the "save as" buttons' outcome for this answer */
  saved?: { path: string; name: string } | { error: string } | 'saving'
}

/**
 * The panel's conversations, newest first, labelled by their own first question.
 *
 * A stored conversation has no title — the store writes `<chat-id>.jsonl` and
 * nothing else — so the first thing said in it is the only honest label. A
 * blank one falls back to the id, which at least is unique and copyable.
 *
 * The list is sorted here rather than trusting the store to: `listChats`
 * returns directory order, and a reader opening the panel expects the thread
 * they were just in to be on top.
 */
export interface ListedThread {
  readonly chatId: string
  readonly label: string
  readonly updatedAt: string
  readonly approxCount: number
}

export function listThreads(
  metas: readonly ChatMeta[],
  firstQuestionOf: (chatId: string) => string | undefined,
): ListedThread[] {
  return metas
    .map((meta) => ({
      chatId: meta.chatId,
      label: firstQuestionOf(meta.chatId)?.replace(/\s+/g, ' ').trim().slice(0, 60) || meta.chatId,
      updatedAt: meta.updatedAt,
      approxCount: meta.approxCount,
    }))
    .sort((a, b) => b.updatedAt.localeCompare(a.updatedAt))
}

/**
 * This panel's threads, and nothing else. Unsaved editor documents write
 * their AI history into the same default project under `unsaved-` ids, and a
 * list that showed those would offer deleting a document's transcript from a
 * panel that has nothing to do with it.
 */
export function panelThreads(metas: readonly ChatMeta[]): ChatMeta[] {
  return metas.filter((meta) => meta.chatId === CHAT_ID || meta.chatId.startsWith(`${CHAT_ID}-`))
}

/** The first question in a stored conversation, which is how it is labelled. */
export function firstQuestionOf(stored: readonly ChatMessage[]): string | undefined {
  return stored.find((message) => message.role === 'user')?.text
}

/**
 * Claims the panel's transcript: resolves its ids, reads back what is already
 * stored, then writes whatever turns were asked before the ids landed — the
 * stored thread is older than anything asked meanwhile, so it is shown and
 * written first. `queued` is drained; a turn queued after the drain is the
 * caller's to flush once the ids are in hand.
 *
 * Split out of the effect so the ordering is testable without a DOM — the
 * panel is the only conversation in the app that belongs to no file, so
 * nothing else exercises the file-less `resolveChat` path end to end.
 */
export async function claimThread(
  api: ProjectApi,
  chatId: string,
  queued: Array<{ role: 'user' | 'assistant'; text: string }>,
  show: (messages: AskMessage[]) => void,
  stillCurrent: () => boolean,
): Promise<{ projectId: string; chatId: string }> {
  const ids = await api.resolveChat({ filePath: null, tempChatId: chatId })
  const stored = await api.loadChat({ ...ids, limit: 200 })
  if (!stillCurrent()) return ids
  if (stored.length > 0) show(restoreAskMessages(stored))
  for (const entry of queued.splice(0)) {
    await api.appendChat({ ...ids, ...entry })
  }
  return ids
}

/**
 * The stored turns as the panel renders them. Only role and text come back:
 * a turn read from disk has this session's badges, token chip and "save as"
 * outcome long gone, and rendering them as absent is truer than guessing.
 */
export function restoreAskMessages(stored: ChatMessage[]): AskMessage[] {
  return stored
    .filter((message) => message.role === 'user' || message.role === 'assistant')
    .map((message) => ({ role: message.role, text: message.text, toolNotes: [], artifacts: [] }))
}

export function KbAskPanel({
  onOpenPath,
  onGoToRecents,
  onOpenModelSettings,
}: {
  onOpenPath: (path: string) => void
  onGoToRecents: () => void
  /** Settings › AI Model, the model chip's last row */
  onOpenModelSettings: () => void
}): React.JSX.Element {
  const { t, lang } = useI18n()
  const [messages, setMessages] = useState<AskMessage[]>([])
  const [draft, setDraft] = useState('')
  const [attachedSkills, setAttachedSkills] = useState<string[]>([])
  const [attachments, setAttachments] = useState<AttachmentInfo[]>([])
  /** rejection / image-read messages, cleared after a few seconds */
  const [attachNotice, setAttachNotice] = useState<string | null>(null)
  const [dragOver, setDragOver] = useState(false)
  const [busy, setBusy] = useState(false)
  const [corpusCount, setCorpusCount] = useState<number | null>(null)
  const [gskEmail, setGskEmail] = useState<string | null>(null)
  const [provider, setProvider] = useState<string | null>(null)
  const corpusRef = useRef<KbFileInfo[]>([])
  const settingsRef = useRef<AiSettings | null>(null)
  const loopRef = useRef<AgentLoop<unknown> | null>(null)
  const scrollRef = useRef<HTMLDivElement>(null)
  /** bumped by Stop: an ask still refreshing the corpus must not start its run */
  const askSeqRef = useRef(0)
  /** tokens billed so far in the in-flight answer (summed across turns) */
  const runUsageRef = useRef<AiTokenUsageView | undefined>(undefined)
  /** tokens billed across the panel's session (the composer chip) */
  const [sessionUsage, setSessionUsage] = useState<AiTokenUsageView | undefined>(undefined)
  /** the rendered thread, readable from callbacks that close over an older one */
  const messagesRef = useRef<AskMessage[]>([])
  messagesRef.current = messages
  /** where this panel's transcript lives, once project-store has answered */
  const [activeChat, setActiveChat] = useState<string>(CHAT_ID)
  const [threads, setThreads] = useState<ListedThread[]>([])
  const chatIdsRef = useRef<{ projectId: string; chatId: string } | null>(null)
  /** the thread id the CLI tools key their work directory on */
  const activeChatRef = useRef<string>(CHAT_ID)
  activeChatRef.current = activeChat
  const cliContextRef = useRef<CliContext | null>(null)
  const attachmentsRef = useRef<AttachmentInfo[]>([])
  attachmentsRef.current = attachments
  /** attachments of earlier questions in this thread: a follow-up may still read them */
  const sentAttachmentsRef = useRef<AttachmentInfo[]>([])
  /** composer attachments plus everything already sent this thread (deduped by path) */
  const availableAttachments = (): AttachmentInfo[] => {
    const seen = new Set<string>()
    return [...sentAttachmentsRef.current, ...attachmentsRef.current].filter((a) =>
      seen.has(a.path) ? false : (seen.add(a.path), true),
    )
  }
  /** turns asked before the ids came back, written once they do */
  const pendingRef = useRef<Array<{ role: 'user' | 'assistant'; text: string }>>([])

  const persist = (entry: { role: 'user' | 'assistant'; text: string }): void => {
    const ids = chatIdsRef.current
    if (!ids) {
      pendingRef.current.push(entry)
      return
    }
    void window.projectApi?.appendChat({ ...ids, ...entry }).catch(() => undefined)
  }

  useEffect(() => {
    // Claim whichever thread is active — the fixed home thread on mount, the
    // chosen id after a chip click or a new thread. The ids resolve in the
    // background, so a turn asked before they land is queued by persist() and
    // flushed here rather than lost.
    const api = window.projectApi
    if (!api) return
    let current = true
    // a question asked while the thread was loading stays on screen, below it
    const show = (restored: AskMessage[]): void => {
      setMessages((prev) => [...restored, ...prev])
      messagesRef.current = [...restored, ...messagesRef.current]
      loopRef.current?.restore(restored.map((m) => ({ role: m.role, text: m.text })))
    }
    cliContextRef.current = null
    sentAttachmentsRef.current = []
    void window.aiOffice
      .kbCliContext?.(activeChat)
      .then((ctx) => {
        if (current) cliContextRef.current = ctx
      })
      .catch(() => undefined)
    void claimThread(api, activeChat, pendingRef.current, show, () => current)
      .then((ids) => {
        if (!current) return
        chatIdsRef.current = ids
        for (const entry of pendingRef.current.splice(0)) persist(entry)
      })
      .catch(() => undefined)
    void refreshThreads(api)
    return () => {
      current = false
    }
  }, [activeChat])

  const refreshThreads = async (api: ProjectApi): Promise<void> => {
    try {
      const ids = await api.resolveChat({ filePath: null, tempChatId: CHAT_ID })
      // Only this panel's own threads belong in the list: the default project
      // also keeps every editor document's AI history, and listing those
      // would offer deleting them from here.
      const metas = panelThreads(await api.listChats({ projectId: ids.projectId }))
      // Labels ride on the listing itself, so a refresh is one call however
      // many threads are stored.
      setThreads(
        listThreads(metas, (chatId) => metas.find((meta) => meta.chatId === chatId)?.firstQuestion),
      )
    } catch {
      setThreads([])
    }
  }

  // The model's history and the stored ids belong to the thread being left;
  // a question typed before the next claim resolves queues in pendingRef.
  const leaveThread = (): void => {
    loopRef.current?.reset()
    chatIdsRef.current = null
    pendingRef.current = []
    messagesRef.current = []
    setMessages([])
    setAttachedSkills([])
    // staged files belong to the thread they were dropped into
    setAttachments([])
    sentAttachmentsRef.current = []
  }

  /** Switching is a state change: the effect re-reads whichever thread is now active. */
  const openThread = (chatId: string): void => {
    leaveThread()
    setActiveChat(chatId)
  }

  /** A fresh id, so a new thread does not append to the one being read. */
  const startThread = (): void => {
    leaveThread()
    setActiveChat(`kb-ask-${Date.now()}`)
  }

  const removeThread = async (chatId: string, label: string): Promise<void> => {
    // A transcript gone is gone; the ×  on a chip should not be a hair-trigger.
    const thread = threads.find((entry) => entry.chatId === chatId)
    if (!window.confirm(t('askThreadDeleteConfirm', { label: thread?.label ?? label }))) {
      return
    }
    const api = window.projectApi
    if (!api) return
    try {
      const ids = await api.resolveChat({ filePath: null, tempChatId: CHAT_ID })
      await api.deleteChat({ projectId: ids.projectId, chatId })
      if (chatId === activeChat) startThread()
      await refreshThreads(api)
    } catch {
      // A delete that did not happen leaves the list as it was, which is the
      // honest outcome; the reader sees the thread still listed.
    }
  }

  const refreshSettings = (): void => {
    void window.aiOffice
      .aiGetSettings()
      .then((settings) => {
        settingsRef.current = settings
        setProvider(activeProvider(settings))
      })
      .catch(() => undefined)
  }

  useEffect(() => window.aiOffice.onAiSettingsChanged?.(refreshSettings), [])

  const modelBridge = useMemo<AiModelPickerBridge>(
    () => ({
      getSettings: () => window.aiOffice.aiGetSettings(),
      setSettings: (settings) => window.aiOffice.setAiSettings(settings),
      onSettingsChanged: (handler) => window.aiOffice.onAiSettingsChanged(handler),
      gskLoggedIn: () => window.aiOffice.aiGskStatus().then((s) => !!s.loggedIn),
      openModelSettings: onOpenModelSettings,
    }),
    [onOpenModelSettings],
  )

  useEffect(() => {
    // one corpus snapshot per mount; sends refresh it again (see ask()).
    // A bridge without the kb channel (defensive: the shell always exposes
    // it) leaves the corpus empty, which keeps the panel on its guidance.
    if (typeof window.aiOffice?.kbList !== 'function') {
      setCorpusCount(0)
      return
    }
    void window.aiOffice
      .kbList()
      .then((files) => {
        corpusRef.current = files
        setCorpusCount(files.length)
      })
      .catch(() => setCorpusCount(0))
    refreshSettings()
    void window.aiOffice
      .aiGskStatus()
      .then((status) => setGskEmail(status.loggedIn ? (status.email ?? '') : null))
      .catch(() => undefined)
    const cliSkill = createCliSkill({
      context: () => cliContextRef.current,
      run: (request) =>
        window.aiOffice.kbRunCli(activeChatRef.current, {
          ...request,
          sources: availableAttachments().map((a) => a.path),
        }),
    })
    const filesSkill = createAttachmentsSkill({
      list: availableAttachments,
      read: async (path, offset, maxChars) => {
        const result = await window.aiOffice.readAttachment(path, offset, maxChars)
        if (!result.ok) throw new Error(result.error ?? 'read failed')
        return {
          text: result.text ?? '',
          totalChars: result.totalChars ?? 0,
          offset: result.offset ?? 0,
        }
      },
      pathsInContext: true,
    })
    const docSkill = createDocumentSkill({
      create: (request) => window.aiOffice.kbCreateDocument(request),
      alsoSaves: [RUN_GENOFFICE_TOOL],
    })
    loopRef.current = new AgentLoop({
      transport: createIpcTransport<AiSettings>({
        onStream: (listener) => window.aiOffice.onAiStream(listener),
        start: (request) => window.aiOffice.aiStream(request),
        cancel: (requestId) => void window.aiOffice.aiStreamCancel(requestId),
        getSettings: () => settingsRef.current as AiSettings,
        unknownErrorText: () => t('askUnknownError'),
      }),
      skill: composeSkills('home-ask', '', [
        createKnowledgeBaseSkill({
          listFiles: () => corpusRef.current,
          search: (query, limit) => window.aiOffice.kbSearch({ q: query, limit }),
          read: (path, offset) => window.aiOffice.kbRead(path, offset),
          alsoCitable: () => [
            ...docSkill.createdPaths(),
            ...cliSkill.createdPaths(),
            ...availableAttachments().map((a) => a.path),
          ],
        }),
        filesSkill,
        docSkill,
        cliSkill,
      ]),
      events: {
        onText: (text) => patchLast({ text }),
        onToolExecuted: ({ call, execution }) =>
          patchLast((last) => ({ artifacts: settleArtifact(last.artifacts, call, execution) })),
        onToolStart: (call) =>
          patchLast((last) => ({
            artifacts: (() => {
              const started = artifactFromCall(call, cliContextRef.current?.outDir)
              return started ? [...last.artifacts, started] : last.artifacts
            })(),
            toolNotes: appendToolNote(last.toolNotes, {
              kind: toolNoteKind(call),
              label: t(TOOL_NOTE_LABEL[toolNoteKind(call)]),
              detail: toolNoteDetail(
                call,
                (index) => corpusRef.current[index]?.name,
                (index) => availableAttachments()[index]?.name,
              ),
              count: 1,
            }),
          })),
        onDone: () => finishLast(),
        onUsage: (usage) => {
          runUsageRef.current = addTokenUsage(runUsageRef.current, usage)
          setSessionUsage((prev) => addTokenUsage(prev, usage))
          patchLast({ usage: runUsageRef.current })
        },
        onError: (error) =>
          patchLast((last) => ({
            text: last.text || error,
            streaming: false,
            isError: !last.text,
          })),
      },
    })
    return () => loopRef.current?.cancel()
    // eslint-disable-next-line react-hooks/exhaustive-deps -- mount-only; the loop reads live refs
  }, [])

  useEffect(() => {
    scrollRef.current?.scrollTo({ top: scrollRef.current.scrollHeight })
  }, [messages])

  // Stream patches land on the ref synchronously as well as on state: onText
  // and onDone can arrive in one tick, and finishLast must see the final text
  // before React has rendered it.
  const patchLast = (
    patch: Partial<AskMessage> | ((last: AskMessage) => Partial<AskMessage>),
  ): void => {
    const apply = (prev: AskMessage[]): AskMessage[] => {
      if (prev.length === 0 || prev[prev.length - 1]!.role !== 'assistant') return prev
      const last = prev[prev.length - 1]!
      const delta = typeof patch === 'function' ? patch(last) : patch
      return [...prev.slice(0, -1), { ...last, ...delta }]
    }
    messagesRef.current = apply(messagesRef.current)
    setMessages(apply)
  }

  const finishLast = (): void => {
    // The answer is whole for the first time here; streaming it as it grows
    // would store every intermediate state. The `streaming` guard is what
    // keeps a second finish (Stop, then the loop's own onDone) from storing
    // the same turn twice.
    const last = messagesRef.current[messagesRef.current.length - 1]
    if (last?.role === 'assistant' && last.streaming && last.text) {
      persist({ role: 'assistant', text: last.text })
    }
    patchLast({ streaming: false })
    setBusy(false)
  }

  const ask = (): Promise<void> => {
    const instruction = draft.trim()
    if (!instruction || busy) return Promise.resolve()
    setDraft('')
    // skills and files ride with this one question only: the chips clear like the draft does
    const skills = attachedSkills
    setAttachedSkills([])
    const files = attachments
    setAttachments([])
    return askWith(instruction, skills, files)
  }

  const notice = (text: string): void => {
    setAttachNotice(text)
    window.setTimeout(() => setAttachNotice((cur) => (cur === text ? null : cur)), 5000)
  }

  const mergeAttachments = (result: AttachmentAddResult | null): void => {
    if (!result) return
    if (result.rejected.length > 0) notice(result.rejected.join('; '))
    if (result.accepted.length === 0) return
    setAttachments((prev) => {
      const seen = new Set(prev.map((a) => a.path))
      return [...prev, ...result.accepted.filter((a) => !seen.has(a.path))]
    })
  }

  const addDroppedFiles = async (files: File[]): Promise<void> => {
    const paths: string[] = []
    for (const f of files) {
      const p = window.aiOffice.pathForFile(f)
      if (p) {
        paths.push(p)
        continue
      }
      // a pasted screenshot has no path: the main process saves it to a temp file first
      const ext = f.type.split('/')[1]?.replace('jpeg', 'jpg') ?? ''
      mergeAttachments(await window.aiOffice.addPastedImage(await f.arrayBuffer(), ext))
    }
    if (paths.length > 0) mergeAttachments(await window.aiOffice.addAttachmentPaths(paths))
  }

  const removeAttachment = (path: string): void =>
    setAttachments((prev) => prev.filter((a) => a.path !== path))

  /** image attachments go multimodal with the question; a failed read is reported, not fatal */
  const collectImages = async (atts: readonly AttachmentInfo[]): Promise<AgentImage[]> => {
    const images: AgentImage[] = []
    const failures: string[] = []
    for (const att of atts.filter((a) => ATTACHMENT_IMAGE_EXTS.has(a.ext))) {
      try {
        const result = await window.aiOffice.readAttachmentImage(att.path)
        if (result.ok && result.base64 && result.mime) {
          images.push({ base64: result.base64, mime: result.mime })
          continue
        }
        failures.push(result.error ?? t('askImageReadFail', { name: att.name }))
      } catch {
        failures.push(t('askImageReadFail', { name: att.name }))
      }
    }
    if (failures.length > 0) notice(failures.join('; '))
    return images
  }

  /** the attached skills' full text in front of the question; a skill that vanished is skipped */
  const withSkillBodies = async (instruction: string, names: readonly string[]) => {
    if (names.length === 0) return instruction
    const bodies = await Promise.all(
      names.map(async (name) => {
        try {
          return { name, body: await window.aiOffice.skillText(name) }
        } catch {
          return null
        }
      }),
    )
    return withAttachedSkills(
      instruction,
      bodies.filter((b): b is { name: string; body: string } => b !== null),
    )
  }

  const listSkillsForMenu = async () =>
    (await window.aiOffice.skillsList()).map(({ name, description }) => ({ name, description }))
  const toggleSkill = (name: string) =>
    setAttachedSkills((prev) =>
      prev.includes(name) ? prev.filter((n) => n !== name) : [...prev, name],
    )

  // Retry and other programmatic turns: the composer draft is not theirs to clear
  const askWith = async (
    instruction: string,
    skills: readonly string[] = [],
    files: readonly AttachmentInfo[] = [],
  ): Promise<void> => {
    if (!instruction || busy) return
    setBusy(true)
    refreshSettings()
    if (files.length > 0) {
      const seen = new Set(sentAttachmentsRef.current.map((a) => a.path))
      sentAttachmentsRef.current = [
        ...sentAttachmentsRef.current,
        ...files.filter((a) => !seen.has(a.path)),
      ]
    }
    const seq = ++askSeqRef.current
    runUsageRef.current = undefined
    // freshly starred files are answerable this very turn
    try {
      const files = await window.aiOffice.kbList()
      corpusRef.current = files
      setCorpusCount(files.length)
    } catch {
      // keep the mount snapshot
    }
    if (askSeqRef.current !== seq) return
    // a question about an attached file needs no starred corpus
    if (corpusRef.current.length === 0 && availableAttachments().length === 0) {
      setMessages((prev) => [
        ...prev,
        { role: 'user', text: instruction, toolNotes: [], artifacts: [] },
        { role: 'assistant', text: t('askEmptyCorpus'), toolNotes: [], artifacts: [] },
      ])
      persist({ role: 'user', text: instruction })
      persist({ role: 'assistant', text: t('askEmptyCorpus') })
      setBusy(false)
      return
    }
    setMessages((prev) => [
      ...prev,
      {
        role: 'user',
        text: instruction,
        toolNotes: [],
        artifacts: [],
        ...(skills.length > 0 ? { skills: [...skills] } : {}),
        ...(files.length > 0 ? { attachments: [...files] } : {}),
      },
      { role: 'assistant', text: '', toolNotes: [], artifacts: [], streaming: true },
    ])
    persist({ role: 'user', text: instruction })
    try {
      const [text, images] = await Promise.all([
        withSkillBodies(instruction, skills),
        collectImages(files),
      ])
      if (askSeqRef.current !== seq) return
      await loopRef.current!.run(text, images)
    } catch {
      finishLast()
    }
  }

  const fileNav = createFileNav((path) => onOpenPath(path))

  // the deterministic half of "keep this answer": the question becomes the
  // file name, the answer's markdown the body, same write path as the tool
  const saveAnswer = async (index: number, format: CreateDocumentFormat): Promise<void> => {
    const answer = messages[index]
    if (!answer || answer.saved === 'saving') return
    const question = [...messages.slice(0, index)].reverse().find((m) => m.role === 'user')
    const title = (question?.text ?? '').replace(/\s+/g, ' ').trim().slice(0, 60)
    const patchAt = (saved: AskMessage['saved']): void =>
      setMessages((prev) => prev.map((m, i) => (i === index ? { ...m, saved } : m)))
    patchAt('saving')
    try {
      const { path } = await window.aiOffice.kbCreateDocument({
        title,
        format,
        content: answer.text,
      })
      patchAt({ path, name: path.split(/[\\/]/).pop() ?? path })
    } catch (error) {
      patchAt({ error: error instanceof Error ? error.message : String(error) })
    }
  }

  const hasFiles = (e: React.DragEvent): boolean => e.dataTransfer.types.includes('Files')
  // files dropped on the panel attach to the question; stopPropagation keeps
  // them from Home's drop-to-open bridge and its full-window overlay
  const dropProps = {
    onDragEnter: (e: React.DragEvent) => {
      if (!hasFiles(e)) return
      e.preventDefault()
      e.stopPropagation()
      setDragOver(true)
    },
    onDragOver: (e: React.DragEvent) => {
      if (!hasFiles(e)) return
      e.preventDefault()
      e.stopPropagation()
      e.dataTransfer.dropEffect = 'copy'
    },
    onDragLeave: (e: React.DragEvent) => {
      if (!hasFiles(e)) return
      e.stopPropagation()
      if (e.currentTarget.contains(e.relatedTarget as Node | null)) return
      setDragOver(false)
    },
    onDrop: (e: React.DragEvent) => {
      if (!hasFiles(e)) return
      e.preventDefault()
      e.stopPropagation()
      setDragOver(false)
      void addDroppedFiles(Array.from(e.dataTransfer.files))
    },
  }

  return (
    <div className={`kb-ask${dragOver ? ' kb-ask-dragover' : ''}`} {...dropProps}>
      {dragOver && (
        <div className="kb-ask-drop" aria-hidden="true">
          {t('askDropFiles')}
        </div>
      )}
      {corpusCount !== null && corpusCount > 0 && (
        <div className="kb-ask-headbar">
          <b>{t('askTitle')}</b>
          <span>{t('askCorpusCount', { n: String(corpusCount) })}</span>
        </div>
      )}
      {threads.length > 0 && (
        <div className="kb-ask-threads" role="tablist" aria-label={t('askThreadsTitle')}>
          {threads.map((thread) => (
            <div
              key={thread.chatId}
              className={`kb-ask-thread-chip${thread.chatId === activeChat ? ' active' : ''}`}
            >
              <button
                type="button"
                role="tab"
                aria-selected={thread.chatId === activeChat}
                title={thread.label}
                onClick={() => openThread(thread.chatId)}
              >
                {thread.label}
              </button>
              <button
                type="button"
                className="kb-ask-thread-del"
                aria-label={t('askThreadDelete', { label: thread.label })}
                onClick={() => void removeThread(thread.chatId, thread.label)}
              >
                ×
              </button>
            </div>
          ))}
          <button
            type="button"
            className="kb-ask-thread-new"
            aria-label={t('askThreadNew')}
            title={t('askThreadNew')}
            onClick={startThread}
          >
            +
          </button>
        </div>
      )}
      <div className="kb-ask-thread" ref={scrollRef}>
        {messages.length === 0 && corpusCount === 0 && (
          <div className="kb-ask-hero">
            <div className="kb-ask-hero-icon" aria-hidden="true">
              <IconBookOpen />
            </div>
            <h3>{t('askTitle')}</h3>
            <p>{t('askHeroTagline')}</p>
            <div className="kb-ask-steps">
              <div className="kb-ask-step">
                <span className="kb-ask-step-num">1</span>
                <div>
                  <b>{t('askStep1Title')}</b>
                  <span>{t('askStep1Desc')}</span>
                </div>
              </div>
              <div className="kb-ask-step">
                <span className="kb-ask-step-num">2</span>
                <div>
                  <b>{t('askStep2Title')}</b>
                  <span>{t('askStep2Desc')}</span>
                </div>
              </div>
              <div className="kb-ask-step">
                <span className="kb-ask-step-num">3</span>
                <div>
                  <b>{t('askStep3Title')}</b>
                  <span>{t('askStep3Desc')}</span>
                </div>
              </div>
            </div>
            <button className="kb-ask-cta" type="button" onClick={onGoToRecents}>
              <IconStar />
              {t('askCtaRecents')}
            </button>
          </div>
        )}
        {messages.length === 0 && corpusCount !== null && corpusCount > 0 && (
          <p className="kb-ask-hint">{t('askIntro')}</p>
        )}
        {messages.map((m, i) => (
          <div key={i} className={`kb-ask-msg kb-ask-${m.role}${m.isError ? ' kb-ask-error' : ''}`}>
            {m.role === 'user' && m.attachments && <AskAttachments items={m.attachments} />}
            {m.role === 'user' && m.skills && <AiSkillChips lang={lang} names={m.skills} />}
            {m.role === 'assistant' && m.toolNotes.length > 0 && (
              <div className="kb-ask-tools">
                {m.toolNotes.map((note, j) => (
                  <span key={j} className="kb-ask-tool-note">
                    {note.kind === 'search' ? (
                      <IconSearch />
                    ) : note.kind === 'create' || note.kind === 'cli' ? (
                      <IconFilePlus />
                    ) : (
                      <IconFileText />
                    )}
                    {note.label}
                    {note.detail && <span className="kb-ask-tool-detail">{note.detail}</span>}
                    {note.count > 1 && <span className="kb-ask-tool-count">×{note.count}</span>}
                    {m.streaming && j === m.toolNotes.length - 1 ? '…' : ''}
                  </span>
                ))}
              </div>
            )}
            {m.role === 'assistant' && m.artifacts.length > 0 && (
              <div className="kb-ask-card" role="list">
                {m.artifacts.map((a) => (
                  <div
                    key={a.callId}
                    className={`kb-ask-card-row kb-ask-card-${a.status}`}
                    role="listitem"
                  >
                    {a.status === 'done' && a.path ? (
                      <button
                        type="button"
                        className="kb-ask-card-name"
                        onClick={() => onOpenPath(a.path!)}
                      >
                        <IconFileText />
                        {a.name}
                      </button>
                    ) : (
                      <span className="kb-ask-card-name">
                        <IconFileText />
                        {a.name}
                      </span>
                    )}
                    {a.status === 'building' && (
                      <span className="kb-ask-card-status">{t('askArtifactBuilding')}…</span>
                    )}
                    {a.status === 'failed' && (
                      <>
                        <span className="kb-ask-card-status kb-ask-card-error" title={a.error}>
                          {t('askArtifactFailed')}
                        </span>
                        <button
                          type="button"
                          className="kb-ask-action"
                          disabled={busy}
                          onClick={() => void askWith(retryInstruction(a))}
                        >
                          {t('askArtifactRetry')}
                        </button>
                      </>
                    )}
                  </div>
                ))}
              </div>
            )}
            {m.text ? (
              m.role === 'assistant' ? (
                <Markdown text={m.text} navs={[fileNav]} />
              ) : (
                m.text
              )
            ) : m.streaming ? (
              <AiTypingIndicator label={t('askHintBusy')} />
            ) : null}
            {m.role === 'assistant' && m.usage && (
              <div className="ai-message-usage">
                <AiUsageChip usage={m.usage} lang={lang} />
              </div>
            )}
            {m.role === 'assistant' && m.text && !m.streaming && !m.isError && (
              <div className="kb-ask-actions">
                {m.saved && typeof m.saved === 'object' && 'path' in m.saved ? (
                  <button
                    type="button"
                    className="kb-ask-action kb-ask-action-saved"
                    onClick={() => onOpenPath((m.saved as { path: string }).path)}
                  >
                    <IconFileText />
                    {t('askSavedAs', { name: m.saved.name })}
                  </button>
                ) : (
                  <>
                    <button
                      type="button"
                      className="kb-ask-action"
                      disabled={m.saved === 'saving'}
                      onClick={() => void saveAnswer(i, 'md')}
                    >
                      <IconFilePlus />
                      {t('askSaveMd')}
                    </button>
                    <button
                      type="button"
                      className="kb-ask-action"
                      disabled={m.saved === 'saving'}
                      onClick={() => void saveAnswer(i, 'docx')}
                    >
                      <IconFilePlus />
                      {t('askSaveDocx')}
                    </button>
                    {m.saved && typeof m.saved === 'object' && 'error' in m.saved && (
                      <span className="kb-ask-action-error">
                        {t('askSaveFailed', { error: m.saved.error })}
                      </span>
                    )}
                  </>
                )}
              </div>
            )}
          </div>
        ))}
      </div>
      {messages.length > 0 && (
        <p className="kb-ask-nosave">
          <IconInfo />
          {t('askNoSave')}
        </p>
      )}
      {gskEmail === null && provider === 'genspark' && (
        <p className="kb-ask-login">{t('askNeedsLogin')}</p>
      )}
      {attachNotice && (
        <p className="kb-ask-notice" role="status">
          {attachNotice}
        </p>
      )}
      <AiComposer
        value={draft}
        busy={busy}
        placeholder={t('askPlaceholder')}
        hintIdle={t('askHintIdle')}
        hintBusy={t('askHintBusy')}
        sendLabel={t('askSend')}
        stopLabel={t('askStop')}
        ariaLabel={t('askPlaceholder')}
        iconOnly
        sendIconEnabled={<img src={sendEnterOn} alt="" aria-hidden />}
        sendIconDisabled={<img src={sendEnterOff} alt="" aria-hidden />}
        stopIcon={<img src={sendStop} alt="" aria-hidden />}
        header={
          <>
            <AskAttachments
              items={attachments}
              onRemove={removeAttachment}
              removeLabel={t('askRemoveAttachment')}
            />
            <AiSkillChips lang={lang} names={attachedSkills} onRemove={toggleSkill} />
          </>
        }
        footerStart={
          <>
            <AiModelPicker bridge={modelBridge} lang={lang} />
            {sessionUsage && <AiUsageChip usage={sessionUsage} session lang={lang} />}
            <AiSkillsMenu
              lang={lang}
              listSkills={listSkillsForMenu}
              selected={attachedSkills}
              onToggle={toggleSkill}
              onAddFile={() =>
                void window.aiOffice
                  .pickAttachments()
                  .then(mergeAttachments)
                  .catch(() => {})
              }
              onManage={() => void window.aiOffice.openSkillsSettings?.().catch(() => {})}
            />
          </>
        }
        onChange={setDraft}
        onPasteFiles={(files) => void addDroppedFiles(files)}
        onSend={() => void ask()}
        onStop={() => {
          askSeqRef.current++
          loopRef.current?.cancel()
          finishLast()
        }}
      />
    </div>
  )
}

/** open book for the empty-corpus hero (inline, like every Home icon) */
function IconBookOpen(): React.JSX.Element {
  return (
    <svg
      width="26"
      height="26"
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth="1.5"
      strokeLinecap="round"
      strokeLinejoin="round"
      aria-hidden="true"
    >
      <path d="M12 6.8C10.2 5.4 7.9 4.7 5.5 4.7c-.4 0-.8 0-1.2.1v12.9c.4-.1.8-.1 1.2-.1 2.4 0 4.7.7 6.5 2.1 1.8-1.4 4.1-2.1 6.5-2.1.4 0 .8 0 1.2.1V4.8c-.4-.1-.8-.1-1.2-.1-2.4 0-4.7.7-6.5 2.1z" />
      <path d="M12 6.8v12.9" />
    </svg>
  )
}

function IconStar(): React.JSX.Element {
  return (
    <svg
      width="15"
      height="15"
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth="1.8"
      strokeLinejoin="round"
      aria-hidden="true"
    >
      <path d="M12 3.6l2.6 5.2 5.8.9-4.2 4.1 1 5.8-5.2-2.7-5.2 2.7 1-5.8L3.6 9.7l5.8-.9z" />
    </svg>
  )
}

function IconInfo(): React.JSX.Element {
  return (
    <svg
      width="13"
      height="13"
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth="1.8"
      strokeLinecap="round"
      aria-hidden="true"
    >
      <circle cx="12" cy="12" r="9" />
      <path d="M12 8.2v.01M12 11.5V16" />
    </svg>
  )
}

function IconSearch(): React.JSX.Element {
  return (
    <svg
      width="11"
      height="11"
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth="2.2"
      strokeLinecap="round"
      aria-hidden="true"
    >
      <circle cx="11" cy="11" r="7" />
      <path d="m20 20-3.8-3.8" />
    </svg>
  )
}

function IconFilePlus(): React.JSX.Element {
  return (
    <svg
      width="11"
      height="11"
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth="2.2"
      strokeLinecap="round"
      strokeLinejoin="round"
      aria-hidden="true"
    >
      <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" />
      <polyline points="14 2 14 8 20 8" />
      <line x1="12" y1="18" x2="12" y2="12" />
      <line x1="9" y1="15" x2="15" y2="15" />
    </svg>
  )
}

function IconFileText(): React.JSX.Element {
  return (
    <svg
      width="11"
      height="11"
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth="2.2"
      strokeLinejoin="round"
      aria-hidden="true"
    >
      <path d="M14 3H7a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8z" />
      <path d="M14 3v5h5M9 13h6M9 17h6" />
    </svg>
  )
}
