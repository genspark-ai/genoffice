import { useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react'
import {
  Dropdown,
  FindPanel,
  type FindFocusRequest,
  type AiScopeQuoteData,
  useAutoSavePref,
  type FindPanelStrings,
  type FindTarget,
} from '@genoffice/ui'
import { useI18n } from './i18n/locale'
import { parseDocText, serializeDocText, type Envelope } from './document/envelope'
import { SourceEditor, type CursorInfo, type SourceEditorHandle } from './source/SourceEditor'
import { PreviewFrame, type PreviewFrameHandle } from './preview/PreviewFrame'
import { instrumentForPreview } from './preview/instrument'
import type { ComputedSnapshot, ElementRect, FromInspector } from './preview/inspector-protocol'
import inspectorSource from './preview/inspector.js?raw'
import { AiPanel, GensparkMark, type AiPreset, type HtmlAiDeps } from './ai/AiPanel'
import { AiAskPopover, type AnchorRect, type AskMode } from './components/AiAskPopover'
import {
  EDIT_QUEUE_MAX,
  buildSelectionInstruction,
  excerptOf,
  resolveQueueItem,
  type EditQueueItem,
} from './ai/edit-queue'
import { Breadcrumb, type NodeState } from './components/Breadcrumb'
import {
  Ribbon,
  type PresentKind,
  VIEW_MODES,
  type CanvasMode,
  type ViewMode,
} from './components/Ribbon'
import { CropDialog, CutoutDialog, type ImageDialogLabels } from '@genoffice/ui'
import { FloatToolbar } from './components/FloatToolbar'
import { StylePanel } from './components/StylePanel'
import { floatPosition, parseDeclarations } from './document/float-position'
import {
  buildParseMap,
  childrenOf,
  elementCovering,
  sourceText,
  type ParseMap,
} from './document/parse-map'
import { compileOps, type HtmlOp, type OpError } from './document/ops'
import { injectBrief, parseBrief, type Brief } from './document/brief'
import { applyPatches } from './document/patch'
import { deriveAutoFileName, deriveNameFromPrompt } from './document/auto-name'
import type { ExportFormat, SaveMode } from '../shared/ipc'

type LoadStatus = 'loading' | 'ready' | 'error'
type SaveState = 'idle' | 'saving' | 'saved' | 'failed'

interface TextSel {
  sid: number
  textNodeIndex: number
  start: number
  end: number
}

const MIN_ZOOM = 50
const MAX_ZOOM = 200
const ZOOM_STEP = 10
const clampZoom = (z: number) => Math.min(MAX_ZOOM, Math.max(MIN_ZOOM, z))
const inTextField = (target: EventTarget | null) =>
  target instanceof HTMLElement &&
  !!target.closest('.cm-editor, input, textarea, [contenteditable=""], [contenteditable="true"]')
const PREVIEW_DEBOUNCE_MS = 150
const DEFAULT_ENVELOPE: Envelope = { bom: false, eol: '\n', trailingNewline: true }
/** live style pokes are written to the source as one set_style op this long after the last change */
const STYLE_COMMIT_MS = 600
type Device = 'desktop' | 'tablet' | 'mobile'
const DEVICES: Device[] = ['desktop', 'tablet', 'mobile']
/** document scaffolding: selectable in the source, never a target for element ops */
const STRUCTURAL = new Set(['html', 'head', 'body'])
const IS_MAC = /Mac/i.test(navigator.platform)

function readViewMode(): ViewMode {
  const stored = localStorage.getItem('htmlapp.viewMode')
  return VIEW_MODES.includes(stored as ViewMode) ? (stored as ViewMode) : 'preview'
}

function isDarkTheme(): boolean {
  const explicit = document.documentElement.getAttribute('data-theme')
  if (explicit === 'dark') return true
  if (explicit === 'light') return false
  return window.matchMedia('(prefers-color-scheme: dark)').matches
}

/** picture bytes as a data URL: inline data, a remote URL through the main-process fetcher, or a file next to the document */
async function loadImageDataUrl(src: string): Promise<string | null> {
  if (/^data:image\//i.test(src)) return src
  if (/^https?:/i.test(src)) {
    const r = await window.htmlApi.fetchImage(src)
    return r ? `data:${r.mime};base64,${r.base64}` : null
  }
  if (/^[a-z][a-z0-9+.-]*:/i.test(src)) return null
  const r = await window.htmlApi.readImage(src)
  return r ? `data:${r.mime};base64,${r.base64}` : null
}

export default function App() {
  const { t } = useI18n()
  const [status, setStatus] = useState<LoadStatus>('loading')
  const [path, setPath] = useState<string | null>(null)
  const [text, setText] = useState('')
  const [savedText, setSavedText] = useState('')
  const [saveState, setSaveState] = useState<SaveState>('idle')
  const [view, setView] = useState<ViewMode>(readViewMode)
  const [canvasMode, setCanvasModeState] = useState<CanvasMode>('edit')
  /** Present → Fullscreen: the presenting view also takes the whole screen */
  const [presentFull, setPresentFull] = useState(false)
  const [zoom, setZoom] = useState(100)
  const [autoSave, setAutoSave] = useAutoSavePref('htmlapp.autoSave', window.htmlApi)
  const [findTarget, setFindTarget] = useState<FindTarget | null>(null)
  const [findFocus, setFindFocus] = useState<FindFocusRequest>({ field: 'find', nonce: 0 })
  const [cursor, setCursor] = useState<CursorInfo>({ line: 1, col: 1, pos: 0 })
  const [previewUrl, setPreviewUrl] = useState<string | null>(null)
  const [previewNonce, setPreviewNonce] = useState(0)
  const [draftHtml, setDraftHtml] = useState<string | null>(null)
  const [historyState, setHistoryState] = useState({ undo: false, redo: false })
  const [aiOpen, setAiOpen] = useState(() => localStorage.getItem('htmlapp.showAi') !== '0')
  const [aiPreset, setAiPreset] = useState<AiPreset | null>(null)
  const [editQueue, setEditQueue] = useState<EditQueueItem[]>([])
  const [askMode, setAskMode] = useState<AskMode | null>(null)
  const [selectedSid, setSelectedSid] = useState<number | null>(null)
  const [selectedState, setSelectedState] = useState<NodeState>('static')
  const [textSel, setTextSel] = useState<TextSel | null>(null)
  const wrapSelectionRef = useRef<(tag: string) => void>(() => {})
  // selected element as the preview sees it: drives the floating toolbar and the style panel
  const [selRect, setSelRect] = useState<ElementRect | null>(null)
  const [selComputed, setSelComputed] = useState<ComputedSnapshot | null>(null)
  const [selText, setSelText] = useState<{ run: string | null; index: number }>({
    run: null,
    index: -1,
  })
  const [pendingCount, setPendingCount] = useState(0)
  const [barSize, setBarSize] = useState({ w: 440, h: 32 })
  const [stageTick, setStageTick] = useState(0)
  const [device, setDevice] = useState<Device>(() => {
    const stored = localStorage.getItem('htmlapp.device')
    return DEVICES.includes(stored as Device) ? (stored as Device) : 'desktop'
  })
  const [panelOpen, setPanelOpen] = useState(
    () => localStorage.getItem('htmlapp.stylePanel') !== '0',
  )
  const [panelDismissedSid, setPanelDismissedSid] = useState<number | null>(null)
  const [notice, setNotice] = useState<string | null>(null)
  const [pictureDialog, setPictureDialog] = useState<{
    kind: 'crop' | 'cutout'
    sid: number
    /** authored src at open time: the result only lands on an <img> that still carries it */
    src: string
    image: string
  } | null>(null)

  const editorRef = useRef<SourceEditorHandle>(null)
  const previewRef = useRef<PreviewFrameHandle>(null)
  const envelopeRef = useRef<Envelope>(DEFAULT_ENVELOPE)
  const textRef = useRef('')
  const savedTextRef = useRef('')
  const savingRef = useRef(false)
  const statusRef = useRef<LoadStatus>('loading')
  const pushedTextRef = useRef<string | null>(null)
  /** parse-map version of the copy currently served to the preview; messages from older copies are ignored */
  const pushedVersionRef = useRef(-1)
  /** bumped on every text change; AI staleness checks compare against lastManualVersionRef */
  const versionRef = useRef(0)
  const lastManualVersionRef = useRef(0)
  const mapRef = useRef<ParseMap | null>(null)
  const mapSourceRef = useRef('')
  const pathRef = useRef<string | null>(null)
  const selectedSidRef = useRef<number | null>(null)
  const editQueueRef = useRef<EditQueueItem[]>([])
  const canvasModeRef = useRef<CanvasMode>('edit')
  const queueSeqRef = useRef(0)
  /** brief confirmed on the AI card; pinned into the head of the next generated document */
  const briefRef = useRef<Brief | null>(null)
  /** name taken from the first AI request of an untitled document: tab title now, file name at the first save */
  const provisionalNameRef = useRef<string | null>(null)
  const pendingStylesRef = useRef<Record<string, string | null>>({})
  const styleTimerRef = useRef<number | null>(null)
  const flushingStylesRef = useRef(false)
  const stageRef = useRef<HTMLDivElement>(null)
  const barRef = useRef<HTMLDivElement>(null)
  pathRef.current = path
  textRef.current = text
  editQueueRef.current = editQueue
  canvasModeRef.current = canvasMode
  savedTextRef.current = savedText
  statusRef.current = status
  selectedSidRef.current = selectedSid
  const dirty = text !== savedText

  const getMap = useCallback((): ParseMap => {
    const cached = mapRef.current
    // the initial load swaps the text without bumping the version, so the source itself is part of the key
    if (cached && cached.version === versionRef.current && mapSourceRef.current === textRef.current)
      return cached
    const next = buildParseMap(textRef.current, versionRef.current, cached)
    mapRef.current = next
    mapSourceRef.current = textRef.current
    return next
  }, [])

  useEffect(() => {
    let cancelled = false
    void (async () => {
      try {
        const [pending, info] = await Promise.all([
          window.htmlApi.consumePending(),
          window.htmlApi.getPreviewInfo(),
        ])
        const raw = pending ? await window.htmlApi.readFile(pending) : ''
        if (cancelled) return
        const doc = parseDocText(raw)
        envelopeRef.current = doc.envelope
        textRef.current = doc.text
        // the frame must never race the first push: serve the buffer before the URL is known to React
        const map0 = getMap()
        window.htmlApi.updatePreview(instrumentForPreview(doc.text, map0, inspectorSource))
        pushedTextRef.current = doc.text
        pushedVersionRef.current = map0.version
        setPath(pending)
        setText(doc.text)
        setSavedText(doc.text)
        setPreviewUrl(info.url)
        setStatus('ready')
      } catch (err) {
        console.error('[html] open failed:', err)
        if (!cancelled) setStatus('error')
      }
    })()
    return () => {
      cancelled = true
    }
  }, [getMap])

  // mirror dirtiness to the main process (close prompt) — untitled blank docs never count
  useEffect(() => {
    if (status !== 'ready') return
    window.htmlApi.setDirty(dirty || pendingCount > 0)
  }, [dirty, pendingCount, status])

  const pushPreview = useCallback(
    (nextText: string) => {
      if (pushedTextRef.current === nextText) return
      const map = getMap()
      window.htmlApi.updatePreview(instrumentForPreview(nextText, map, inspectorSource))
      pushedTextRef.current = nextText
      pushedVersionRef.current = map.version
      setPreviewNonce((n) => n + 1)
    },
    [getMap],
  )
  // push the instrumented buffer to html-preview:// and reload the frame, debounced per keystroke
  useEffect(() => {
    if (status !== 'ready' || pushedTextRef.current === text) return
    const id = window.setTimeout(() => pushPreview(text), PREVIEW_DEBOUNCE_MS)
    return () => window.clearTimeout(id)
  }, [text, status, pushPreview])

  useEffect(() => {
    localStorage.setItem('htmlapp.viewMode', view)
  }, [view])
  const frameMode = (mode: CanvasMode) => (mode === 'edit' ? 'inspect' : 'browse')
  const prevCanvasModeRef = useRef(canvasMode)
  useEffect(() => {
    // a presentation lets links navigate the frame; coming back reloads the instrumented document
    if (prevCanvasModeRef.current === 'present' && canvasMode !== 'present')
      setPreviewNonce((n) => n + 1)
    prevCanvasModeRef.current = canvasMode
    previewRef.current?.post({ type: 'gx:setMode', mode: frameMode(canvasMode) })
  }, [canvasMode])

  useEffect(() => {
    localStorage.setItem('htmlapp.showAi', aiOpen ? '1' : '0')
  }, [aiOpen])

  useEffect(() => {
    localStorage.setItem('htmlapp.device', device)
  }, [device])

  useEffect(() => {
    localStorage.setItem('htmlapp.stylePanel', panelOpen ? '1' : '0')
  }, [panelOpen])

  // the device host is centred in the stage: any stage resize (split view, AI dock, device) moves it
  useEffect(() => {
    const stage = stageRef.current
    if (!stage || status !== 'ready') return
    const ro = new ResizeObserver(() => setStageTick((n) => n + 1))
    ro.observe(stage)
    return () => ro.disconnect()
  }, [status])

  useLayoutEffect(() => {
    const el = barRef.current
    if (!el) return
    const r = el.getBoundingClientRect()
    if (r.width && (r.width !== barSize.w || r.height !== barSize.h))
      setBarSize({ w: r.width, h: r.height })
  }, [selRect, selText.run, barSize.w, barSize.h])

  useEffect(() => {
    if (!notice) return
    const id = window.setTimeout(() => setNotice(null), 4000)
    return () => window.clearTimeout(id)
  }, [notice])

  const refreshHistory = useCallback(() => {
    const editor = editorRef.current
    if (editor) setHistoryState({ undo: editor.canUndo(), redo: editor.canRedo() })
  }, [])

  /** every text change goes through here so the version counter and the map cache stay coherent */
  const commitText = useCallback(
    (next: string, manual: boolean) => {
      textRef.current = next
      versionRef.current += 1
      if (manual) lastManualVersionRef.current = versionRef.current
      // text-range coordinates never survive a document change
      setTextSel(null)
      // any other document change reloads the preview and drops the live pokes; forget them too rather than
      // committing them later against sids the rebuilt parse map may have reassigned
      if (!flushingStylesRef.current) {
        if (styleTimerRef.current !== null) window.clearTimeout(styleTimerRef.current)
        styleTimerRef.current = null
        pendingStylesRef.current = {}
        setPendingCount(0)
      }
      setText(next)
      refreshHistory()
    },
    [refreshHistory],
  )

  const onEditorChange = useCallback(
    (next: string) => {
      editorRef.current?.clearHighlights()
      commitText(next, true)
    },
    [commitText],
  )

  /** compile + apply one batch; manual batches are not highlighted and count as user edits for AI staleness */
  const applyOps = useCallback(
    (
      ops: HtmlOp[],
      manual: boolean,
    ): { ok: true; ranges: Array<[number, number]> } | { ok: false; errors: OpError[] } => {
      const base = textRef.current
      const compiled = compileOps(base, getMap(), ops)
      if (compiled.errors.length > 0) return { ok: false, errors: compiled.errors }
      const editor = editorRef.current
      const ranges = editor
        ? editor.applyPatches(compiled.patches, !manual)
        : compiled.patches.map((p) => [p.from, p.from + p.text.length] as [number, number])
      commitText(applyPatches(base, compiled.patches), manual)
      return { ok: true, ranges }
    },
    [commitText, getMap],
  )

  const replaceAll = useCallback(
    (html: string, highlight: boolean) => {
      // a generated document carries the confirmed brief so later turns (and re-opens) stay anchored to it
      const pinned =
        highlight && briefRef.current && !parseBrief(html)
          ? injectBrief(html, briefRef.current)
          : html
      editorRef.current?.replaceDoc(pinned, highlight)
      commitText(pinned, false)
    },
    [commitText],
  )

  const flushStylesRef = useRef<(() => void) | null>(null)
  const flushDraftsRef = useRef<(() => void) | null>(null)
  /** land live style pokes and open panel drafts in the source before anything reads, saves or edits it */
  const flushPending = useCallback(() => {
    flushStylesRef.current?.()
    flushDraftsRef.current?.()
  }, [])

  /** apply a toolbar/inspector batch; the selection follows the edited element (or clears when it is gone) */
  const runManual = useCallback(
    (ops: HtmlOp[], follow: 'reselect' | 'clear' | 'keep' = 'reselect') => {
      // live style pokes and open panel drafts are written first, while their sid still names the element they were made on
      flushPending()
      const before = selectedSidRef.current
      const r = applyOps(ops, true)
      if (!r.ok) {
        setNotice(r.errors[0]?.message ?? 'edit rejected')
        return false
      }
      if (follow === 'clear') {
        selectSidRef.current?.(null, {})
        return true
      }
      if (follow === 'reselect' && before !== null) {
        const map = getMap()
        // sids are matched by path, so after a move / rename / replace the old sid may now name the
        // sibling that took the old position; re-derive the target from the largest applied range
        const relocates = ops.some(
          (o) => o.op === 'move' || o.op === 'set_tag' || o.op === 'replace_element',
        )
        const widest = r.ranges.reduce<[number, number] | null>(
          (best, cur) => (!best || cur[1] - cur[0] > best[1] - best[0] ? cur : best),
          null,
        )
        // moved chunks carry their line's indentation and newline; trim so the covering
        // element is the moved element itself, not its parent
        const trimmed = (() => {
          if (!widest) return null
          const src = textRef.current
          let [a, b] = widest
          while (a < b && /\s/.test(src[a]!)) a++
          while (b > a && /\s/.test(src[b - 1]!)) b--
          return [a, b] as [number, number]
        })()
        const covering = trimmed ? elementCovering(map, trimmed[0], trimmed[1]) : null
        if (relocates && trimmed) {
          selectSidRef.current?.(covering && !STRUCTURAL.has(covering.tag) ? covering.sid : null, {
            reveal: true,
          })
        } else if (!map.bySid.has(before)) {
          selectSidRef.current?.(covering && !STRUCTURAL.has(covering.tag) ? covering.sid : null, {
            reveal: true,
          })
        }
      }
      return true
    },
    [applyOps, getMap, flushPending],
  )

  // ── selection model: one current element shared by the preview, the source pane, the toolbar and the AI ──

  const selectSidRef = useRef<
    | ((
        sid: number | null,
        opts: { reveal?: boolean; toPreview?: boolean; state?: NodeState },
      ) => void)
    | null
  >(null)

  const selectSid = useCallback(
    (
      sid: number | null,
      opts: { reveal?: boolean; toPreview?: boolean; state?: NodeState } = {},
    ) => {
      if (sid !== selectedSidRef.current) flushPending()
      setSelectedSid(sid)
      setSelectedState(opts.state ?? 'static')
      setTextSel(null)
      // geometry, styles and text run all belong to the previous element; the frame re-reports them via gx:rect
      setSelRect(null)
      setSelComputed(null)
      setSelText({ run: null, index: -1 })
      if (opts.toPreview !== false) previewRef.current?.post({ type: 'gx:select', sid })
      if (sid !== null && opts.reveal) {
        const e = getMap().bySid.get(sid)
        if (e) editorRef.current?.revealRange(e.range[0], e.range[1], false)
      }
    },
    [getMap, flushPending],
  )
  selectSidRef.current = selectSid

  /** numbered pins for queued edits; stale items (element gone) get no pin */
  const postMarks = useCallback(() => {
    const items = canvasModeRef.current === 'present' ? [] : editQueueRef.current
    // no getMap() before the document is loaded: it would cache an empty map under version 0
    const map = items.length ? getMap() : null
    const src = textRef.current
    // one pin per element; several queued edits on the same element share it and list their ordinals
    const bySid = new Map<number, string[]>()
    items.forEach((item, i) => {
      const target = map && resolveQueueItem(src, map, item).target
      if (target) bySid.set(target.sid, [...(bySid.get(target.sid) ?? []), String(i + 1)])
    })
    const marks = [...bySid].map(([sid, ordinals]) => ({ sid, label: ordinals.join('·') }))
    previewRef.current?.post({ type: 'gx:mark', marks })
  }, [getMap])
  useEffect(postMarks, [editQueue, canvasMode, postMarks])

  const onInspectorMessage = useCallback(
    (msg: FromInspector) => {
      // a click that lands while the frame is reloading carries sids from the previous copy
      if (msg.version !== pushedVersionRef.current) return
      switch (msg.type) {
        case 'gx:ready': {
          previewRef.current?.post({ type: 'gx:theme', dark: isDarkTheme() })
          previewRef.current?.post({ type: 'gx:setMode', mode: frameMode(canvasModeRef.current) })
          const sid = selectedSidRef.current
          if (sid !== null) previewRef.current?.post({ type: 'gx:select', sid })
          postMarks()
          return
        }
        case 'gx:markClick': {
          const map = getMap()
          // a shared pin opens the most recent edit on that element; the others stay reachable from the queue card
          const item = editQueueRef.current
            .filter((q) => resolveQueueItem(textRef.current, map, q).target?.sid === msg.sid)
            .at(-1)
          if (!item) return
          selectSid(item.sid, { reveal: true })
          setAskMode({ kind: 'edit', qid: item.qid })
          return
        }
        case 'gx:select': {
          if (!msg.element || msg.dynamic || msg.element.sid === null) {
            if (msg.element) {
              flushPending()
              setSelectedSid(null)
              setSelectedState('dynamic')
              setNotice(t('nodeDynamic'))
            } else selectSid(null, { toPreview: false })
            return
          }
          const map = getMap()
          const e = map.bySid.get(msg.element.sid)
          if (!e) {
            selectSid(null, { toPreview: false })
            return
          }
          const src = textRef.current
          const tag = src.slice(e.startTag[0], e.startTag[1])
          const cls = (/\sclass\s*=\s*["']([^"']*)["']/i.exec(tag)?.[1] ?? '')
            .trim()
            .split(/\s+/)
            .filter(Boolean)
            .sort()
            .join(' ')
          const liveCls = msg.element.className.trim().split(/\s+/).filter(Boolean).sort().join(' ')
          const srcText = sourceText(src, e)
          const liveText = msg.element.text
          const textMatches =
            srcText.length > 400 ? srcText.startsWith(liveText.slice(0, 380)) : srcText === liveText
          const dirty =
            cls !== liveCls ||
            childrenOf(map, e.sid).length !== msg.element.childElementCount ||
            !textMatches
          selectSid(e.sid, { reveal: true, toPreview: false, state: dirty ? 'dirty' : 'static' })
          setSelRect(msg.element.rect)
          setSelComputed(msg.element.computed)
          setSelText({ run: msg.element.textRun, index: msg.element.textRunIndex })
          return
        }
        case 'gx:hover':
          return
        case 'gx:rect':
          if (msg.sid !== selectedSidRef.current) return
          setSelRect(msg.rect)
          setSelComputed(msg.computed)
          setSelText({ run: msg.textRun, index: msg.textRunIndex })
          return
        case 'gx:textSelect':
          setTextSel({
            sid: msg.sid,
            textNodeIndex: msg.textNodeIndex,
            start: msg.start,
            end: msg.end,
          })
          return
        case 'gx:textEditCommit':
          runManual([
            { op: 'set_text_node', sid: msg.sid, index: msg.textNodeIndex, text: msg.newText },
          ])
          return
        case 'gx:textEditCancel':
          return
        case 'gx:zoom':
          if (canvasModeRef.current !== 'present') setZoom((z) => clampZoom(z - msg.delta * 0.6))
          return
        case 'gx:keyCommand': {
          if (canvasModeRef.current === 'present') {
            if (msg.command === 'escape') setCanvasModeState('edit')
            return
          }
          if (msg.command === 'undo' || msg.command === 'redo') {
            editorRef.current?.[msg.command]()
            return
          }
          if (msg.command === 'zoomIn' || msg.command === 'zoomOut') {
            const dir = msg.command === 'zoomIn' ? 1 : -1
            setZoom((z) => clampZoom(Math.round(z) + dir * ZOOM_STEP))
            return
          }
          if (msg.command === 'zoomReset') {
            setZoom(100)
            return
          }
          const sid = selectedSidRef.current
          if (sid === null) return
          const map = getMap()
          const e = map.bySid.get(sid)
          if (!e) return
          if (msg.command === 'delete') runManual([{ op: 'remove', sid }], 'clear')
          else if (msg.command === 'escape') selectSid(null)
          else if (msg.command === 'askAi') setAskMode({ kind: 'new' })
          else if (msg.command === 'bold') wrapSelectionRef.current('strong')
          else if (msg.command === 'italic') wrapSelectionRef.current('em')
          else if (msg.command === 'parent' && e.parentSid !== null) {
            const parent = map.bySid.get(e.parentSid)
            if (parent && parent.tag !== 'body' && parent.tag !== 'html')
              selectSid(parent.sid, { reveal: true })
          } else if (msg.command === 'child') {
            const first = childrenOf(map, sid)[0]
            if (first) selectSid(first.sid, { reveal: true })
          } else if (msg.command === 'next' || msg.command === 'prev') {
            const siblings = e.parentSid !== null ? childrenOf(map, e.parentSid) : []
            const i = siblings.findIndex((s) => s.sid === sid)
            const target = siblings[msg.command === 'next' ? i + 1 : i - 1]
            if (target) selectSid(target.sid, { reveal: true })
          }
          return
        }
        case 'gx:navigateBlocked':
          window.open(msg.href)
          setNotice(t('openExternal'))
          return
        default:
          return
      }
    },
    [getMap, postMarks, runManual, selectSid, t, flushPending],
  )

  // the source cursor picks the covering element (no reveal: the user is already there)
  const onCursor = useCallback(
    (c: CursorInfo) => {
      setCursor(c)
      // [pos, pos+1): a cursor sitting on a boundary belongs to the element that starts there
      const e = elementCovering(getMap(), c.pos, Math.min(c.pos + 1, textRef.current.length))
      const sid = e && !STRUCTURAL.has(e.tag) ? e.sid : null
      if (sid !== selectedSidRef.current) selectSid(sid, { toPreview: true })
    },
    [getMap, selectSid],
  )

  const selectedEntry = selectedSid !== null ? getMap().bySid.get(selectedSid) : undefined

  const duplicateSelected = () => {
    if (!selectedEntry) return
    const html = textRef.current.slice(selectedEntry.range[0], selectedEntry.range[1])
    runManual([{ op: 'insert_html', sid: selectedEntry.sid, position: 'after', html: `\n${html}` }])
  }
  const moveSelected = (dir: -1 | 1) => {
    if (!selectedEntry || selectedEntry.parentSid === null) return
    const siblings = childrenOf(getMap(), selectedEntry.parentSid)
    const i = siblings.findIndex((s) => s.sid === selectedEntry.sid)
    const ref = siblings[i + dir]
    if (!ref) return
    runManual([
      {
        op: 'move',
        sid: selectedEntry.sid,
        position: dir < 0 ? 'before' : 'after',
        ref_sid: ref.sid,
      },
    ])
  }
  const wrapSelection = (tag: string, attrs?: Record<string, string>) => {
    if (!textSel) return
    if (runManual([{ op: 'wrap_text', ...textSel, index: textSel.textNodeIndex, tag, attrs }]))
      setTextSel(null)
  }
  wrapSelectionRef.current = wrapSelection
  const askAi = () => {
    if (selectedEntry) setAskMode({ kind: 'new' })
  }
  const setAttr = (name: string, value: string | null) => {
    // may run from the panel's unmount after the element was deleted
    if (selectedEntry && getMap().bySid.has(selectedEntry.sid))
      runManual([{ op: 'set_attr', sid: selectedEntry.sid, name, value }])
  }
  /** picks a file, copies it into the document's assets/ and points src at it; needs a saved document */
  const replaceImage = async () => {
    if (!selectedEntry) return
    if (!pathRef.current) {
      setNotice(t('imageNeedsSave'))
      return
    }
    const sid = selectedEntry.sid
    const rel = await window.htmlApi.pickImage()
    if (rel && selectedSidRef.current === sid)
      runManual([{ op: 'set_attr', sid, name: 'src', value: rel }])
  }
  /** crop / remove background edit the pixels: read the picture, open the dialog, write a new asset */
  const openPictureDialog = async (kind: 'crop' | 'cutout') => {
    const src = selComputed?.image?.src
    if (!selectedEntry || !src) return
    if (!pathRef.current) {
      setNotice(t('imageNeedsSave'))
      return
    }
    const sid = selectedEntry.sid
    const image = await loadImageDataUrl(src).catch(() => null)
    if (selectedSidRef.current !== sid) return
    if (!image) {
      setNotice(t('imageLoadFail'))
      return
    }
    setPictureDialog({ kind, sid, src, image })
  }
  const applyPictureBytes = async (sid: number, src: string, png: string) => {
    setPictureDialog(null)
    const rel = await window.htmlApi.saveImage({ base64: png, ext: 'png' })
    // sids are matched by path, so after an edit during the write the sid may name another
    // element: the target must still be the <img> with the src the dialog was opened on
    const entry = rel ? getMap().bySid.get(sid) : undefined
    // compare through the same attribute decoding the frame used (entities such as &amp; in the query)
    const authoredSrc =
      entry?.tag === 'img'
        ? new DOMParser()
            .parseFromString(
              textRef.current.slice(entry.startTag[0], entry.startTag[1]),
              'text/html',
            )
            .querySelector('img')
            ?.getAttribute('src')
        : null
    const stillThere = authoredSrc === src
    if (!rel || !stillThere) {
      // an unreferenced pending asset is discarded by the save-time reconciliation
      setNotice(t('imageProcessFail'))
      return
    }
    runManual([{ op: 'set_attr', sid, name: 'src', value: rel }])
  }
  const imageDialogLabels: ImageDialogLabels = {
    cancel: t('cancel'),
    apply: t('imageApply'),
    applying: t('imageProcessing'),
    loading: t('imageLoading'),
    loadFailed: t('imageLoadFail'),
    processFailed: t('imageProcessFail'),
    cutoutTitle: t('imageRemoveBg'),
    tolerance: t('imageCutoutTolerance'),
    cutoutHint: (pct) => t('imageCutoutHint', { pct }),
    cropTitle: t('imageCrop'),
    cropHint: t('imageCropHint'),
  }

  // ── element-scoped AI edits: queue them on the selected element or run one right away ──
  const askTarget = useMemo(() => {
    if (!selectedEntry || STRUCTURAL.has(selectedEntry.tag)) return null
    return {
      sid: selectedEntry.sid,
      tag: selectedEntry.tag,
      excerpt: excerptOf(text, selectedEntry),
      start: selectedEntry.range[0],
    }
  }, [selectedEntry, text])
  const queueAdd = (instruction: string) => {
    setAskMode(null)
    if (!askTarget || editQueue.length >= EDIT_QUEUE_MAX) return
    const qid = `q${++queueSeqRef.current}`
    setEditQueue((prev) => [
      ...prev,
      { qid, sid: askTarget.sid, tag: askTarget.tag, capturedText: askTarget.excerpt, instruction },
    ])
    setAiOpen(true)
  }
  const queueUpdate = (qid: string, instruction: string) => {
    setAskMode(null)
    setEditQueue((prev) => prev.map((q) => (q.qid === qid ? { ...q, instruction } : q)))
  }
  const queueRemove = (qid: string) => {
    setAskMode(null)
    setEditQueue((prev) => prev.filter((q) => q.qid !== qid))
  }
  const queueConsume = (qids: string[]) =>
    setEditQueue((prev) => prev.filter((q) => !qids.includes(q.qid)))
  const queueFocus = (qid: string) => {
    const item = editQueue.find((q) => q.qid === qid)
    const target = item && resolveQueueItem(text, getMap(), item).target
    if (target) selectSid(target.sid, { reveal: true })
  }
  const askSendNow = (instruction: string) => {
    setAskMode(null)
    if (!askTarget) return
    flushPending()
    setAiOpen(true)
    const scope: AiScopeQuoteData = {
      label: t('aiScopeElement', { tag: askTarget.tag }),
      ...(askTarget.excerpt.trim() ? { text: askTarget.excerpt.trim() } : {}),
    }
    setAiPreset({
      text: buildSelectionInstruction(askTarget, instruction),
      displayText: instruction,
      nonce: Date.now(),
      scope,
    })
  }
  /** viewport rect of the selected element, clipped to the stage; identity changes whenever the geometry does */
  const getAskAnchorRect = useCallback((): AnchorRect | null => {
    // stageTick: the stage moved or resized (AI dock, split view, device recentre) without the frame-local rect changing
    void stageTick
    const stage = stageRef.current?.getBoundingClientRect()
    const host = stageRef.current?.querySelector('.preview-host')?.getBoundingClientRect()
    if (!selRect || !stage || !host) return null
    const z = zoom / 100
    const left = Math.max(stage.left, host.left + selRect.x * z)
    const top = Math.max(stage.top, host.top + selRect.y * z)
    const right = Math.min(stage.right, host.left + (selRect.x + selRect.width) * z)
    const bottom = Math.min(stage.bottom, host.top + (selRect.y + selRect.height) * z)
    if (right <= left || bottom <= top) return null
    return { left, top, right, bottom, viewTop: stage.top, viewBottom: stage.bottom }
  }, [selRect, zoom, stageTick])

  // ── live style editing: poke the preview DOM now, write one set_style op to the source shortly after ──
  const flushStyles = useCallback(() => {
    if (styleTimerRef.current !== null) window.clearTimeout(styleTimerRef.current)
    styleTimerRef.current = null
    const styles = pendingStylesRef.current
    pendingStylesRef.current = {}
    setPendingCount(0)
    const sid = selectedSidRef.current
    if (sid === null || Object.keys(styles).length === 0) return
    flushingStylesRef.current = true
    try {
      runManual([{ op: 'set_style', sid, styles }])
    } finally {
      flushingStylesRef.current = false
    }
  }, [runManual])
  flushStylesRef.current = flushStyles

  const previewStyle = (styles: Record<string, string | null>) => {
    const sid = selectedSidRef.current
    if (sid === null) return
    previewRef.current?.post({ type: 'gx:previewStyle', sid, styles })
    pendingStylesRef.current = { ...pendingStylesRef.current, ...styles }
    setPendingCount(Object.keys(pendingStylesRef.current).length)
    if (styleTimerRef.current !== null) window.clearTimeout(styleTimerRef.current)
    styleTimerRef.current = window.setTimeout(flushStyles, STYLE_COMMIT_MS)
  }

  /** drop uncommitted previews: reload the frame from the last pushed source */
  const revertStyles = () => {
    if (styleTimerRef.current !== null) window.clearTimeout(styleTimerRef.current)
    styleTimerRef.current = null
    pendingStylesRef.current = {}
    setPendingCount(0)
    setPreviewNonce((n) => n + 1)
  }

  const setCanvasMode = useCallback(
    (mode: CanvasMode) => {
      setCanvasModeState((prev) => {
        if (prev === mode) return prev
        if (mode === 'present') selectSid(null)
        return mode
      })
      flushPending()
      setAskMode(null)
    },
    [flushPending, selectSid],
  )
  useEffect(() => {
    if (canvasMode !== 'present') return
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') setCanvasMode('edit')
    }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [canvasMode, setCanvasMode])

  const startPresent = useCallback(
    (kind: PresentKind) => {
      if (kind === 'newTab') {
        // the new tab reads html-preview:// as it stands: land pending pokes and the debounced push first
        flushPending()
        pushPreview(textRef.current)
        void window.htmlApi.presentInNewTab(path?.split(/[\\/]/).pop() ?? '')
        return
      }
      setPresentFull(kind === 'fullscreen')
      setCanvasMode('present')
      // HTML fullscreen needs the click's user activation; macOS snaps via simpleFullScreen instead
      if (kind === 'fullscreen' && !IS_MAC)
        void document.documentElement.requestFullscreen?.().catch(() => {})
    },
    [path, setCanvasMode, flushPending, pushPreview],
  )

  // a link navigation inside the presented frame keeps focus in a document without our
  // inspector; hand it back to the app so Esc still ends the presentation
  const onPreviewLoad = useCallback(() => {
    if (canvasModeRef.current !== 'present') return
    const active = document.activeElement
    if (active instanceof HTMLIFrameElement) active.blur()
  }, [])

  // Same shape as the slides show: the main process snaps the window (tab-strip bleed,
  // macOS simpleFullScreen without the Space animation), HTML fullscreen only elsewhere.
  // Entry is read from fullscreenchange, not the request promise, and only a loss that
  // sticks (not a strict-mode remount blip) ends the presentation.
  useEffect(() => {
    if (canvasMode !== 'present' || !presentFull) return
    let entered = !!document.fullscreenElement
    let exitTimer = 0
    void window.htmlApi.setPresentFullScreen(true).catch(() => {})
    const onFsChange = () => {
      if (document.fullscreenElement) {
        entered = true
        window.clearTimeout(exitTimer)
        return
      }
      if (!entered) return
      window.clearTimeout(exitTimer)
      exitTimer = window.setTimeout(() => {
        if (!document.fullscreenElement) setCanvasMode('edit')
      }, 150)
    }
    document.addEventListener('fullscreenchange', onFsChange)
    return () => {
      window.clearTimeout(exitTimer)
      document.removeEventListener('fullscreenchange', onFsChange)
      if (document.fullscreenElement) void document.exitFullscreen().catch(() => {})
      void window.htmlApi.setPresentFullScreen(false).catch(() => {})
    }
  }, [canvasMode, presentFull, setCanvasMode])

  const setRunText = (text: string) => {
    if (!selectedEntry || selText.index < 0 || !getMap().bySid.has(selectedEntry.sid)) return
    runManual([{ op: 'set_text_node', sid: selectedEntry.sid, index: selText.index, text }])
  }

  const doSave = useCallback(
    async (mode: SaveMode, suggestedName?: string): Promise<boolean> => {
      if (statusRef.current !== 'ready') return false
      // uncommitted live style pokes belong to the document being saved
      flushPending()
      // a menu save or ⌘S during an autosave queues behind it instead of being dropped;
      // the main-process waiter is resolved by the save() call that follows
      while (savingRef.current) await new Promise((r) => setTimeout(r, 50))
      savingRef.current = true
      setSaveState('saving')
      try {
        const textAtSave = textRef.current
        const serialized = serializeDocText({ text: textAtSave, envelope: envelopeRef.current })
        const result = await window.htmlApi.save({
          text: serialized,
          imageSources: [],
          mode,
          suggestedName,
          defaultName: provisionalNameRef.current ?? undefined,
        })
        if (result.ok && 'path' in result) {
          setPath((previous) => {
            // a new path changes the preview's <base>; relative assets only resolve after a reload
            if (previous !== result.path) setPreviewNonce((n) => n + 1)
            return result.path
          })
          setSavedText(textAtSave)
          // edits that landed during the write keep the document dirty
          setSaveState(textRef.current === textAtSave ? 'saved' : 'idle')
          if (textRef.current !== textAtSave) window.htmlApi.setDirty(true)
          return true
        }
        setSaveState(result.ok ? 'idle' : 'failed')
        return false
      } catch (err) {
        console.error('[html] save failed:', err)
        setSaveState('failed')
        return false
      } finally {
        savingRef.current = false
      }
    },
    [flushPending],
  )

  const zoomIn = useCallback(() => setZoom((z) => clampZoom(Math.round(z) + ZOOM_STEP)), [])
  const zoomOut = useCallback(() => setZoom((z) => clampZoom(Math.round(z) - ZOOM_STEP)), [])
  const cycleView = useCallback(
    () => setView((v) => VIEW_MODES[(VIEW_MODES.indexOf(v) + 1) % VIEW_MODES.length]!),
    [],
  )

  const openFind = useCallback((replace: boolean) => {
    if (statusRef.current !== 'ready' || canvasModeRef.current === 'present') return
    const target = editorRef.current?.findTarget()
    if (!target) return
    setFindTarget(target)
    setFindFocus((f) => ({ field: replace ? 'replace' : 'find', nonce: f.nonce + 1 }))
    // hits live in the source pane, so it must be on screen
    setView((v) => (v === 'preview' ? 'split' : v))
  }, [])
  const closeFind = useCallback(() => setFindTarget(null), [])
  useEffect(() => {
    if (canvasMode === 'present') closeFind()
  }, [canvasMode, closeFind])

  const exportingRef = useRef(false)
  const runExport = useCallback(async (format: ExportFormat) => {
    if (statusRef.current !== 'ready' || exportingRef.current) return
    exportingRef.current = true
    try {
      flushPending()
      const html = textRef.current
      const suggestedName =
        pathRef.current?.replace(/^.*[/\\]/, '').replace(/\.html?$/i, '') ||
        deriveAutoFileName(html) ||
        ''
      const result =
        format === 'pdf'
          ? await window.htmlApi.exportPdf({ html, suggestedName })
          : await window.htmlApi.exportDocx({ html, suggestedName })
      if (!result.ok) {
        console.error('[html] export failed:', result.error)
        setNotice(t('exportFailed'))
      }
    } finally {
      exportingRef.current = false
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps -- t is not referentially stable
  }, [])

  useEffect(() => {
    const offSave = window.htmlApi.onSaveRequest((mode) => {
      if (statusRef.current !== 'ready') {
        window.htmlApi.sendSaveRequestAck(false)
        return
      }
      void doSave(mode)
    })
    const offClose = window.htmlApi.onCloseSaveRequest(() => {
      void (async () => {
        while (savingRef.current) await new Promise((r) => setTimeout(r, 50))
        flushPending()
        if (textRef.current === savedTextRef.current) {
          window.htmlApi.sendCloseSaveResult(true)
          return
        }
        window.htmlApi.sendCloseSaveResult(await doSave('save'))
      })()
    })
    const offRenamed = window.htmlApi.onFileRenamed((next) => setPath(next))
    const offExport = window.htmlApi.onExportRequest((format) => void runExport(format))
    const offTheme = window.htmlApi.onThemeChanged(() => {
      // let main.tsx flip data-theme first
      window.setTimeout(
        () => previewRef.current?.post({ type: 'gx:theme', dark: isDarkTheme() }),
        0,
      )
    })
    const onKeyDown = (event: KeyboardEvent) => {
      if (!(event.metaKey || event.ctrlKey) || event.altKey) return
      const key = event.key.toLowerCase()
      if (key === 's') {
        event.preventDefault()
        void doSave(event.shiftKey ? 'saveAs' : 'save')
      } else if (key === '\\') {
        event.preventDefault()
        cycleView()
      } else if (key === 'f' && !event.shiftKey) {
        event.preventDefault()
        openFind(false)
      } else if (key === 'h' && !event.shiftKey) {
        // Word's replace shortcut; macOS Cmd+H is the system hide role and never reaches here
        event.preventDefault()
        openFind(true)
      } else if (key === '=' || key === '+') {
        event.preventDefault()
        zoomIn()
      } else if (key === '-' || key === '_') {
        event.preventDefault()
        zoomOut()
      } else if (key === '0') {
        event.preventDefault()
        setZoom(100)
      } else if (key === 'b' || key === 'i') {
        if (inTextField(event.target)) return
        event.preventDefault()
        wrapSelectionRef.current(key === 'b' ? 'strong' : 'em')
      } else if (key === 'z' || (key === 'y' && !event.shiftKey)) {
        // CodeMirror and form fields keep their own history
        if (inTextField(event.target)) return
        event.preventDefault()
        if (key === 'z' && !event.shiftKey) editorRef.current?.undo()
        else editorRef.current?.redo()
      }
    }
    window.addEventListener('keydown', onKeyDown, true)
    return () => {
      offSave()
      offClose()
      offRenamed()
      offExport()
      offTheme()
      window.removeEventListener('keydown', onKeyDown, true)
    }
  }, [doSave, runExport, cycleView, zoomIn, zoomOut, flushPending, openFind])

  // pinch / ctrl+wheel over the stage chrome around the frame; wheel inside the frame arrives as gx:zoom
  useEffect(() => {
    const onWheel = (event: WheelEvent) => {
      if (!event.ctrlKey && !event.metaKey) return
      if (canvasModeRef.current === 'present') return
      if (!(event.target as HTMLElement | null)?.closest?.('.preview-stage')) return
      event.preventDefault()
      setZoom((z) => clampZoom(z - event.deltaY * 0.6))
    }
    window.addEventListener('wheel', onWheel, { passive: false })
    return () => window.removeEventListener('wheel', onWheel)
  }, [])

  // autosave: every 30s and on window blur; untitled documents wait for an explicit first save
  useEffect(() => {
    if (!autoSave || !path) return
    const tick = () => {
      // a native Save As dialog blurs the window: autosave must not queue a write behind
      // a user-driven save, only explicit saves do
      if (savingRef.current) return
      flushPending()
      if (textRef.current === savedTextRef.current) return
      void doSave('save')
    }
    const id = window.setInterval(tick, 30_000)
    window.addEventListener('blur', tick)
    return () => {
      window.clearInterval(id)
      window.removeEventListener('blur', tick)
    }
  }, [autoSave, path, doSave, flushPending])

  const aiDeps: HtmlAiDeps = {
    access: {
      getText: () => textRef.current,
      getVersion: () => versionRef.current,
      getMap,
      getLastManualVersion: () => lastManualVersionRef.current,
      getFilePath: () => pathRef.current,
      getSelectedSid: () => selectedSidRef.current,
      applyOps: (ops) => {
        flushPending()
        return applyOps(ops, false)
      },
      replaceAll: (html) => replaceAll(html, true),
    },
    getSnapshot: () => ({ text: textRef.current }),
    restoreSnapshot: (snapshot) => replaceAll(snapshot.text, false),
    onPrompt: (text) => {
      if (pathRef.current || provisionalNameRef.current) return
      const name = deriveNameFromPrompt(text)
      if (!name) return
      provisionalNameRef.current = name
      window.htmlApi.setProvisionalTitle(name)
    },
    onRunDone: (mutated) => {
      // AI wrote into a never-saved document: save it silently under the name of the request that started it
      if (!mutated || pathRef.current) return
      const name = provisionalNameRef.current || deriveAutoFileName(textRef.current)
      if (name) void doSave('save', name)
    },
    clearHighlights: () => editorRef.current?.clearHighlights(),
    onBriefConfirmed: (brief) => {
      briefRef.current = brief
    },
    previewDraft: setDraftHtml,
    navigateTo: (sid) => {
      const e = getMap().bySid.get(sid)
      if (!e) return
      selectSid(sid, { reveal: true })
    },
  }

  const statusText = useMemo(() => {
    if (saveState === 'saving') return t('saving')
    if (saveState === 'failed') return t('saveFailedStatus')
    if (dirty) return t('unsaved')
    if (saveState === 'saved') return t('savedOk')
    return ''
  }, [saveState, dirty, t])

  if (status === 'error') return <div className="center-note">{t('loadFailed')}</div>
  if (status === 'loading') return <div className="center-note">{t('loading')}</div>

  const hasElement = !!selectedEntry && !STRUCTURAL.has(selectedEntry.tag)
  const panelShown = panelOpen && selectedEntry?.sid !== panelDismissedSid

  /** stage geometry for the floating toolbar: frame coordinates scale with the CSS zoom and shift by the centred device host */
  const floatLayout = () => {
    void stageTick
    const stage = stageRef.current?.getBoundingClientRect()
    const host = stageRef.current?.querySelector('.preview-host')?.getBoundingClientRect()
    return {
      zoom,
      offsetX: stage && host ? host.left - stage.left : 0,
      offsetY: stage && host ? host.top - stage.top : 0,
      // the style panel floats over the right edge of the stage
      stageWidth: Math.max(0, (stage?.width ?? 0) - (panelShown && hasElement ? 280 : 0)),
      barWidth: barSize.w,
      barHeight: barSize.h,
    }
  }

  const findStrings: FindPanelStrings = {
    findPlaceholder: t('findPlaceholder'),
    replacePlaceholder: t('replacePlaceholder'),
    matchCase: t('matchCase'),
    wholeWord: t('wholeWord'),
    noResults: t('noResults'),
    prevMatch: t('prevMatch'),
    nextMatch: t('nextMatch'),
    closeEsc: t('closeEsc'),
    replace: t('replace'),
    replaceAll: t('replaceAll'),
  }

  return (
    <div
      className={`app${canvasMode === 'present' ? ` present${presentFull ? ' present-full' : ''}` : ''}`}
    >
      <Ribbon
        disabled={status !== 'ready'}
        dirty={dirty}
        onSave={() => void doSave('save')}
        onFind={() => openFind(false)}
        canUndo={historyState.undo}
        canRedo={historyState.redo}
        onUndo={() => {
          editorRef.current?.undo()
          editorRef.current?.focus()
        }}
        onRedo={() => {
          editorRef.current?.redo()
          editorRef.current?.focus()
        }}
        autoSave={autoSave}
        onToggleAutoSave={setAutoSave}
        view={view}
        onView={setView}
        aiOpen={aiOpen}
        onToggleAi={() => setAiOpen((v) => !v)}
        onAiPreset={(text) => {
          flushPending()
          setAiOpen(true)
          setAiPreset({ text, nonce: Date.now() })
        }}
        canvasMode={canvasMode}
        onPresent={startPresent}
      />

      <div className="app-main">
        <div className={`ai-dock${aiOpen ? '' : ' collapsed'}`}>
          {!aiOpen && (
            <button
              className="ai-rail"
              data-tip={t('aiOpenAssistant')}
              aria-label={t('aiOpenAssistant')}
              onClick={() => setAiOpen(true)}
            >
              <GensparkMark size={18} />
            </button>
          )}
          {/* stays mounted while collapsed: an in-flight run, its snapshots and the loop context survive */}
          <AiPanel
            deps={aiDeps}
            filePath={path}
            preset={aiPreset}
            editQueue={editQueue}
            onQueueEditInstruction={queueUpdate}
            onQueueRemove={queueRemove}
            onQueueClear={() => setEditQueue([])}
            onQueueFocus={queueFocus}
            onQueueConsume={queueConsume}
            onCollapse={() => setAiOpen(false)}
          />
        </div>
        <div className="app-content">
          {findTarget && (
            <FindPanel
              target={findTarget}
              strings={findStrings}
              onClose={closeFind}
              focusRequest={findFocus}
            />
          )}
          {view !== 'preview' && (
            <Breadcrumb
              text={text}
              map={getMap()}
              sid={selectedSid}
              state={selectedState}
              onSelect={(sid) => selectSid(sid, { reveal: true })}
            />
          )}
          <div className={`workspace view-${canvasMode === 'present' ? 'preview' : view}`}>
            <div className="pane pane-preview">
              <div
                className="preview-stage"
                ref={stageRef}
                data-device={device}
                onPointerDown={(e) => {
                  if (e.target === e.currentTarget) selectSid(null)
                }}
              >
                <PreviewFrame
                  ref={previewRef}
                  url={previewUrl}
                  nonce={previewNonce}
                  zoom={zoom}
                  onMessage={onInspectorMessage}
                  onLoad={onPreviewLoad}
                  draft={draftHtml}
                />
                {canvasMode === 'present' && (
                  <button
                    type="button"
                    className="present-exit"
                    onClick={() => setCanvasMode('edit')}
                  >
                    {t('presentExit')}
                  </button>
                )}
                {canvasMode === 'edit' && hasElement && selectedEntry && selRect && selComputed && (
                  <FloatToolbar
                    barRef={barRef}
                    {...floatPosition(selRect, floatLayout())}
                    computed={selComputed}
                    pending={pendingStylesRef.current}
                    canEditText={selText.run !== null}
                    onStyle={previewStyle}
                    onEditText={() =>
                      previewRef.current?.post({ type: 'gx:beginTextEdit', sid: selectedEntry.sid })
                    }
                    onReplaceImage={replaceImage}
                    onCropImage={() => void openPictureDialog('crop')}
                    onCutoutImage={() => void openPictureDialog('cutout')}
                    canAskAi={askTarget !== null}
                    onAskAi={askAi}
                    tag={selectedEntry.tag}
                    onMove={moveSelected}
                    onDuplicate={duplicateSelected}
                    onDelete={() => runManual([{ op: 'remove', sid: selectedEntry.sid }], 'clear')}
                    panelOpen={panelShown}
                    onTogglePanel={() => {
                      if (panelShown) setPanelOpen(false)
                      else {
                        setPanelDismissedSid(null)
                        setPanelOpen(true)
                      }
                    }}
                  />
                )}
                {canvasMode === 'edit' &&
                  hasElement &&
                  selectedEntry &&
                  selComputed &&
                  panelShown && (
                    <StylePanel
                      key={selectedEntry.sid}
                      tag={selectedEntry.tag}
                      computed={selComputed}
                      textRun={selText.run}
                      onStyle={previewStyle}
                      onText={setRunText}
                      onAttr={setAttr}
                      draftRef={flushDraftsRef}
                      onCustomCss={(css) => {
                        const decls = parseDeclarations(css)
                        if (Object.keys(decls).length) previewStyle(decls)
                      }}
                      pending={pendingCount > 0}
                      onRevert={revertStyles}
                      onClose={() => setPanelDismissedSid(selectedEntry.sid)}
                    />
                  )}
              </div>
            </div>
            <div className="pane pane-source">
              <SourceEditor
                ref={editorRef}
                className="source-editor"
                initialText={text}
                onChange={onEditorChange}
                onCursor={onCursor}
              />
            </div>
          </div>

          <footer className="status-bar">
            <div className="status-left">
              {notice ? (
                <span className="status-item status-notice">{notice}</span>
              ) : (
                !path && <span className="status-item status-hint">{t('previewNeedsSave')}</span>
              )}
            </div>
            <div className="status-right">
              {statusText && (
                <span className={`status-save status-${saveState}`}>{statusText}</span>
              )}
              {view !== 'preview' && (
                <span className="status-item">
                  {t('cursorPos', { line: cursor.line, col: cursor.col })}
                </span>
              )}
              <span className="status-item">{t('charCount', { n: text.length })}</span>
              <Dropdown
                className="device-dd"
                value={device}
                ariaLabel={t('device')}
                tip={t('device')}
                options={DEVICES.map((d) => ({
                  value: d,
                  label: t(
                    d === 'desktop' ? 'devDesktop' : d === 'tablet' ? 'devTablet' : 'devPhone',
                  ),
                }))}
                onPick={setDevice}
              />
              <button
                type="button"
                className="zoom-btn"
                aria-label={t('zoomOut')}
                onClick={zoomOut}
                disabled={zoom <= MIN_ZOOM}
              >
                −
              </button>
              <input
                className="zoom-slider"
                type="range"
                min={MIN_ZOOM}
                max={MAX_ZOOM}
                step={ZOOM_STEP}
                value={Math.round(zoom)}
                aria-label={t('zoom')}
                onChange={(event) => setZoom(Number(event.target.value))}
              />
              <button
                type="button"
                className="zoom-btn"
                aria-label={t('zoomIn')}
                onClick={zoomIn}
                disabled={zoom >= MAX_ZOOM}
              >
                +
              </button>
              <span className="zoom-value">{Math.round(zoom)}%</span>
            </div>
          </footer>
        </div>
      </div>
      {askTarget && askMode && canvasMode !== 'present' && (
        <AiAskPopover
          key={askMode.kind === 'edit' ? askMode.qid : 'new'}
          target={askTarget}
          mode={askMode}
          initialText={
            askMode.kind === 'edit'
              ? editQueue.find((q) => q.qid === askMode.qid)?.instruction
              : undefined
          }
          getAnchorRect={getAskAnchorRect}
          onSubmit={(instruction) =>
            askMode.kind === 'edit' ? queueUpdate(askMode.qid, instruction) : queueAdd(instruction)
          }
          onCancel={() => setAskMode(null)}
          onSendNow={askSendNow}
          onRemove={() => askMode.kind === 'edit' && queueRemove(askMode.qid)}
          queueFull={editQueue.length >= EDIT_QUEUE_MAX}
        />
      )}
      {pictureDialog?.kind === 'cutout' && (
        <CutoutDialog
          labels={imageDialogLabels}
          image={pictureDialog.image}
          onApply={(png) => void applyPictureBytes(pictureDialog.sid, pictureDialog.src, png)}
          onCancel={() => setPictureDialog(null)}
        />
      )}
      {pictureDialog?.kind === 'crop' && (
        <CropDialog
          labels={imageDialogLabels}
          image={pictureDialog.image}
          onApply={(png) => void applyPictureBytes(pictureDialog.sid, pictureDialog.src, png)}
          onCancel={() => setPictureDialog(null)}
        />
      )}
    </div>
  )
}
