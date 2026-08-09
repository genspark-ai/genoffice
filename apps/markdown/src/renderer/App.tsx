import { useCallback, useEffect, useMemo, useRef, useState } from 'react'
import { EditorContent, useEditor } from '@tiptap/react'
import type { Editor } from '@tiptap/core'
import { useI18n } from './i18n/locale'
import {
  buildFrontmatterRaw,
  frontmatterInner,
  parseDocText,
  serializeDocText,
  stripLegacyFencedDivs,
  type DocEnvelope,
} from './markdown/docText'
import { buildExtensions } from './editor/extensions'
import { buildSlashItems } from './editor/slashCommand'
import type { SlashController, SlashMenuState } from './editor/slashCommand'
import { setImageBaseDir } from './editor/localImage'
import { Ribbon } from './components/Ribbon'
import { SlashMenu, type SlashMenuHandle } from './components/SlashMenu'
import { TableMenu } from './components/TableMenu'
import { FrontmatterPanel } from './components/FrontmatterPanel'
import { AiPanel, GensparkMark, type AiPreset, type MarkdownAiDeps } from './ai/AiPanel'
import { DOCX_MAX_IMAGE_PX, exportDocxBytes } from './export/docxExport'
import { buildPrintHtml } from './export/printHtml'
import { resolveImageSrc } from './editor/localImage'
import type { ExportFormat, SaveMode } from '../shared/ipc'

type LoadStatus = 'loading' | 'ready' | 'error'
type SaveState = 'idle' | 'saving' | 'saved' | 'failed'

const EMPTY_ENVELOPE: DocEnvelope = {
  frontmatter: '',
  body: '',
  eol: '\n',
  trailingNewline: true,
  bom: false,
}

function dirOf(path: string): string {
  const i = Math.max(path.lastIndexOf('/'), path.lastIndexOf('\\'))
  return i > 0 ? path.slice(0, i) : path
}

function bytesToBase64(bytes: Uint8Array): string {
  let binary = ''
  const CHUNK = 0x8000
  for (let i = 0; i < bytes.length; i += CHUNK) {
    binary += String.fromCharCode(...bytes.subarray(i, i + CHUNK))
  }
  return btoa(binary)
}

/** Measure a document image via the DOM (the editor already displays it) */
function measureImage(displaySrc: string): Promise<{ width: number; height: number } | null> {
  return new Promise((resolvePromise) => {
    const img = new Image()
    img.onload = () => resolvePromise({ width: img.naturalWidth, height: img.naturalHeight })
    img.onerror = () => resolvePromise(null)
    img.src = displaySrc
  })
}

/** File name for an AI-generated untitled document: first heading, else first words */
export function deriveAutoFileName(editor: Editor): string {
  const doc = editor.state.doc
  for (let i = 0; i < doc.childCount; i++) {
    const node = doc.child(i)
    const text = node.textContent.replace(/\s+/g, ' ').trim()
    if (!text) continue
    if (node.type.name === 'heading') return text.slice(0, 60)
    return text.split(' ').slice(0, 8).join(' ').slice(0, 60)
  }
  return ''
}

export default function App() {
  const { t } = useI18n()
  const [status, setStatus] = useState<LoadStatus>('loading')
  const [filePath, setFilePath] = useState<string | null>(null)
  const [dirty, setDirty] = useState(false)
  const [saveState, setSaveState] = useState<SaveState>('idle')
  const [slashState, setSlashState] = useState<SlashMenuState | null>(null)
  const [fmOpen, setFmOpen] = useState(false)
  const [fmText, setFmText] = useState('')
  const [aiOpen, setAiOpen] = useState(true)
  const [aiPreset, setAiPreset] = useState<AiPreset | null>(null)
  const [autoSave, setAutoSave] = useState(() => localStorage.getItem('mdapp.autoSave') === '1')

  const statusRef = useRef<LoadStatus>('loading')
  const dirtyRef = useRef(false)
  const savingRef = useRef(false)
  const envelopeRef = useRef<DocEnvelope>(EMPTY_ENVELOPE)
  const editorRef = useRef<Editor | null>(null)
  const filePathRef = useRef<string | null>(null)
  const slashMenuRef = useRef<SlashMenuHandle>(null)
  const scrollRef = useRef<HTMLDivElement>(null)

  const markDirty = useCallback(() => {
    if (statusRef.current !== 'ready' || dirtyRef.current) return
    dirtyRef.current = true
    setDirty(true)
    setSaveState('idle')
    window.markdownApi.setDirty(true)
  }, [])

  const insertImage = useCallback(() => {
    void (async () => {
      const relPath = await window.markdownApi.pickImage()
      const current = editorRef.current
      if (relPath && current) current.chain().focus().setImage({ src: relPath }).run()
    })()
  }, [])

  const extensions = useMemo(() => {
    const controller: SlashController = {
      onOpen: setSlashState,
      onUpdate: setSlashState,
      onKeyDown: (event) => slashMenuRef.current?.handleKey(event) ?? false,
      onClose: () => setSlashState(null),
    }
    return buildExtensions({
      slashController: controller,
      slashItems: () =>
        buildSlashItems({ insertImage: filePathRef.current ? insertImage : undefined }),
    })
  }, [insertImage])

  const editor = useEditor({
    extensions,
    content: '',
    autofocus: true,
    editorProps: { attributes: { class: 'doc-editor' } },
    // uiOnly transactions (toggle fold state) never reach the file — not dirty
    onUpdate: ({ transaction }) => {
      if (!transaction.getMeta('uiOnly')) markDirty()
    },
  })
  editorRef.current = editor
  filePathRef.current = filePath

  useEffect(() => {
    setImageBaseDir(filePath ? dirOf(filePath) : null)
  }, [filePath])

  useEffect(() => {
    if (!editor) return
    let cancelled = false
    void (async () => {
      try {
        const path = await window.markdownApi.consumePending()
        if (cancelled) return
        if (path) {
          const raw = await window.markdownApi.readFile(path)
          if (cancelled) return
          const envelope = parseDocText(raw)
          envelopeRef.current = envelope
          setImageBaseDir(dirOf(path))
          // the initial load must not be undoable — Cmd+Z right after opening
          // would otherwise blank the document (and Cmd+S overwrite the file)
          editor
            .chain()
            .setMeta('addToHistory', false)
            .setContent(stripLegacyFencedDivs(envelope.body), { contentType: 'markdown' })
            .run()
          setFilePath(path)
          const inner = frontmatterInner(envelope.frontmatter)
          setFmText(inner)
          if (inner) setFmOpen(true)
        } else {
          envelopeRef.current = { ...EMPTY_ENVELOPE }
        }
        statusRef.current = 'ready'
        setStatus('ready')
      } catch (err) {
        console.error('[markdown] load failed:', err)
        if (!cancelled) {
          statusRef.current = 'error'
          setStatus('error')
        }
      }
    })()
    return () => {
      cancelled = true
    }
  }, [editor])

  const onFrontmatterChange = useCallback(
    (inner: string) => {
      setFmText(inner)
      envelopeRef.current.frontmatter = buildFrontmatterRaw(inner)
      markDirty()
    },
    [markDirty],
  )

  /** Serialize and write to disk; false when canceled/failed (caller keeps the tab open) */
  const doSave = useCallback(async (mode: SaveMode, suggestedName?: string): Promise<boolean> => {
    const current = editorRef.current
    if (!current || statusRef.current !== 'ready' || savingRef.current) return false
    savingRef.current = true
    setSaveState('saving')
    try {
      // edits landing while the write is in flight (AI streaming, fast typing)
      // must keep the document dirty — compare doc identity after the await
      const docAtSave = current.state.doc
      const fmAtSave = envelopeRef.current.frontmatter
      const body = current.getMarkdown()
      const text = serializeDocText(envelopeRef.current, body)
      const result = await window.markdownApi.save({ text, mode, suggestedName })
      if (result.ok && 'path' in result) {
        setFilePath(result.path)
        const unchanged =
          editorRef.current?.state.doc === docAtSave && envelopeRef.current.frontmatter === fmAtSave
        if (unchanged) {
          dirtyRef.current = false
          setDirty(false)
          window.markdownApi.setDirty(false)
          setSaveState('saved')
        } else {
          // the main process cleared its dirty flag on write — re-assert it
          dirtyRef.current = true
          setDirty(true)
          window.markdownApi.setDirty(true)
          setSaveState('idle')
        }
        return true
      }
      setSaveState(result.ok ? 'idle' : 'failed')
      return false
    } catch (err) {
      console.error('[markdown] save failed:', err)
      setSaveState('failed')
      return false
    } finally {
      savingRef.current = false
    }
  }, [])

  const runExport = useCallback(async (format: ExportFormat) => {
    const current = editorRef.current
    if (!current || statusRef.current !== 'ready') return
    const suggestedName =
      (filePathRef.current
        ? filePathRef.current.replace(/^.*[/\\]/, '').replace(/\.(md|markdown)$/i, '')
        : deriveAutoFileName(current)) || 'Untitled'
    try {
      if (format === 'pdf') {
        const html = buildPrintHtml(current.view.dom, suggestedName)
        const result = await window.markdownApi.exportPdf({ html, suggestedName })
        if (!result.ok) console.error('[markdown] pdf export failed:', result.error)
        return
      }
      const loadImage = async (src: string) => {
        const data = await window.markdownApi.readImage(src)
        if (!data) return null
        const dims = await measureImage(resolveImageSrc(src))
        let width = dims?.width || 400
        let height = dims?.height || 300
        if (width > DOCX_MAX_IMAGE_PX) {
          height = Math.round((height * DOCX_MAX_IMAGE_PX) / width)
          width = DOCX_MAX_IMAGE_PX
        }
        return { base64: data.base64, mime: data.mime, widthPx: width, heightPx: height }
      }
      const bytes = await exportDocxBytes(current.getJSON(), loadImage)
      const result = await window.markdownApi.exportDocx({
        base64: bytesToBase64(bytes),
        suggestedName,
        mode: format === 'docs' ? 'openInDocs' : 'dialog',
      })
      if (!result.ok) console.error('[markdown] docx export failed:', result.error)
    } catch (err) {
      console.error('[markdown] export failed:', err)
    }
  }, [])

  useEffect(() => {
    const offExport = window.markdownApi.onExportRequest((format) => void runExport(format))
    return offExport
  }, [runExport])

  useEffect(() => {
    const offSave = window.markdownApi.onSaveRequest(
      (mode) => void doSave(mode).then((ok) => window.markdownApi.sendSaveRequestAck(ok)),
    )
    const offClose = window.markdownApi.onCloseSaveRequest(() => {
      void doSave('save').then((ok) => window.markdownApi.sendCloseSaveResult(ok))
    })
    const offRenamed = window.markdownApi.onFileRenamed((newPath) => setFilePath(newPath))
    const onKeyDown = (event: KeyboardEvent) => {
      if ((event.metaKey || event.ctrlKey) && !event.altKey && event.key.toLowerCase() === 's') {
        event.preventDefault()
        void doSave(event.shiftKey ? 'saveAs' : 'save')
      }
    }
    window.addEventListener('keydown', onKeyDown, true)
    return () => {
      offSave()
      offClose()
      offRenamed()
      window.removeEventListener('keydown', onKeyDown, true)
    }
  }, [doSave])

  useEffect(() => {
    localStorage.setItem('mdapp.autoSave', autoSave ? '1' : '0')
  }, [autoSave])

  // autosave: every 30s and on window blur, silently persist pending changes
  // (same policy as the docs app; untitled documents are skipped — the first
  // save must go through the explicit save path that names the file)
  useEffect(() => {
    if (!autoSave || !filePath) return
    const tick = () => {
      if (!dirtyRef.current) return
      if (editorRef.current?.view.composing) return // don't interrupt IME input
      void doSave('save')
    }
    const id = window.setInterval(tick, 30_000)
    window.addEventListener('blur', tick)
    return () => {
      window.clearInterval(id)
      window.removeEventListener('blur', tick)
    }
  }, [autoSave, filePath, doSave])

  const aiDeps: MarkdownAiDeps = {
    getEditor: () => editorRef.current,
    getSnapshot: () => editorRef.current?.getMarkdown() ?? '',
    restoreSnapshot: (markdown) => {
      const current = editorRef.current
      if (!current) return
      current.commands.setContent(markdown, { contentType: 'markdown' })
      markDirty()
    },
    onRunDone: (mutated) => {
      // AI wrote into a never-saved document → name it from the content and save silently
      if (!mutated || filePathRef.current || !editorRef.current) return
      const name = deriveAutoFileName(editorRef.current)
      if (name) void doSave('save', name)
    },
  }

  const fileName = filePath ? filePath.replace(/^.*[/\\]/, '') : null
  const statusText =
    saveState === 'saving'
      ? t('saving')
      : saveState === 'failed'
        ? t('saveFailed')
        : dirty
          ? t('unsaved')
          : saveState === 'saved'
            ? t('savedOk')
            : ''

  if (status === 'error') {
    return (
      <div className="app">
        <div className="center-note">{t('loadError')}</div>
      </div>
    )
  }

  return (
    <div className="app">
      <Ribbon
        editor={editor}
        disabled={status !== 'ready'}
        dirty={dirty}
        onSave={() => void doSave('save')}
        autoSave={autoSave}
        onToggleAutoSave={setAutoSave}
        imageEnabled={Boolean(filePath)}
        onInsertImage={insertImage}
        frontmatterOpen={fmOpen}
        onToggleFrontmatter={() => setFmOpen((v) => !v)}
        aiOpen={aiOpen}
        onToggleAi={() => setAiOpen((v) => !v)}
        onAiPreset={(text) => {
          setAiOpen(true)
          setAiPreset((prev) => ({ text, nonce: (prev?.nonce ?? 0) + 1 }))
        }}
      />
      {status === 'loading' && <div className="center-note">{t('loading')}</div>}
      <div className="app-main" style={status === 'ready' ? undefined : { display: 'none' }}>
        <div className={`ai-dock${aiOpen ? '' : ' collapsed'}`}>
          {!aiOpen && (
            <button
              className="ai-rail"
              title={t('aiOpenAssistant')}
              onClick={() => setAiOpen(true)}
            >
              <GensparkMark size={22} />
            </button>
          )}
          {/* mounted only after the file is loaded so chat history resolves against the real path */}
          {status === 'ready' && (
            <AiPanel
              deps={aiDeps}
              filePath={filePath}
              preset={aiPreset}
              onCollapse={() => setAiOpen(false)}
            />
          )}
        </div>
        <div className="editor-scroll" ref={scrollRef}>
          <div className="doc-page">
            {fmOpen && <FrontmatterPanel value={fmText} onChange={onFrontmatterChange} />}
            <EditorContent editor={editor} />
          </div>
        </div>
      </div>
      <SlashMenu ref={slashMenuRef} state={slashState} onDismiss={() => setSlashState(null)} />
      <TableMenu editor={editor} scrollRef={scrollRef} />
      <div className="status-bar">
        {fileName && <span className="status-file">{fileName}</span>}
        {statusText && <span className={`status-save status-${saveState}`}>{statusText}</span>}
      </div>
    </div>
  )
}
