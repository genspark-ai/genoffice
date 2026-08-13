import { useCallback, useEffect, useRef, useState } from 'react'
import { useEditor, EditorContent } from '@tiptap/react'
import type { AnyExtension } from '@tiptap/core'
import { markdownExtensions } from './editor/extensions'
import { markdownToProseMirror, prosemirrorToMarkdown } from './markdown-io'
import { Ribbon } from './components/Ribbon'
import { LocaleProvider, useI18n } from './i18n/locale'
import type { MarkdownMenuCommand, OpenFileResult } from '../shared/ipc'

function EditorApp() {
  const { t } = useI18n()
  const [filePath, setFilePath] = useState<string | null>(null)
  const [fileName, setFileName] = useState(t('untitled'))
  const [dirty, setDirty] = useState(false)
  const [zoom, setZoom] = useState(100)
  const [showAi, setShowAi] = useState(false)
  const [ribbonTab, setRibbonTab] = useState<'home' | 'insert' | 'view'>('home')
  const [wordCount, setWordCount] = useState({ words: 0, chars: 0 })
  const frontmatterRef = useRef<Record<string, unknown> | null>(null)
  const dirtyRef = useRef(false)
  const wordCountRef = useRef(wordCount)
  const wordCountTimerRef = useRef<number | undefined>(undefined)

  useEffect(() => {
    dirtyRef.current = dirty
  }, [dirty])

  useEffect(() => {
    wordCountRef.current = wordCount
  }, [wordCount])

  const editor = useEditor({
    // @ts-expect-error TipTap duplicate module issue in npm workspace monorepo
    extensions: markdownExtensions as AnyExtension[],
    content: { type: 'doc', content: [{ type: 'paragraph' }] },
    editorProps: {
      attributes: { class: 'tiptap', spellcheck: 'true' },
    },
    onUpdate: () => {
      setDirty(true)
      debouncedWordCount()
    },
  })

  // ── Word count (debounced) ──────────────────────────────────────────────

  const debouncedWordCount = useCallback(() => {
    clearTimeout(wordCountTimerRef.current)
    wordCountTimerRef.current = window.setTimeout(() => {
      if (!editor) return
      const text = editor.getText()
      const words = text.trim() ? text.trim().split(/\s+/).length : 0
      const chars = text.length
      setWordCount({ words, chars })
    }, 300)
  }, [editor])

  // ── Theme ───────────────────────────────────────────────────────────────

  const [themeExplicit, setThemeExplicit] = useState<'light' | 'dark' | null>(null)

  // Read initial theme from shell
  useEffect(() => {
    void window.markdownApi.getTheme().then((theme: string) => {
      if (theme === 'dark' || theme === 'light') {
        setThemeExplicit(theme as 'light' | 'dark')
      }
    })
  }, [])

  // Listen for theme changes from shell
  useEffect(() => {
    return window.markdownApi.onThemeChanged((theme: string) => {
      if (theme === 'dark' || theme === 'light') {
        setThemeExplicit(theme as 'light' | 'dark')
      } else {
        setThemeExplicit(null)
      }
    })
  }, [])

  useEffect(() => {
    if (themeExplicit) {
      document.documentElement.setAttribute('data-theme', themeExplicit)
    } else {
      document.documentElement.removeAttribute('data-theme')
    }
  }, [themeExplicit])

  // ── File I/O ────────────────────────────────────────────────────────────

  const loadFile = useCallback(
    async (result: OpenFileResult) => {
      if (!editor) return
      const text = result.data
      // @ts-expect-error TipTap duplicate module issue in npm workspace monorepo
      const json = markdownToProseMirror(text, markdownExtensions)
      const fmNode = json.content?.find((n) => n.type === 'frontmatter')
      frontmatterRef.current = (fmNode?.attrs?.data as Record<string, unknown>) ?? null
      editor.commands.setContent(json)
      setFilePath(result.path)
      setFileName(result.name)
      setDirty(false)
      debouncedWordCount()
    },
    [editor, debouncedWordCount],
  )

  const saveFile = useCallback(
    async (saveAs = false): Promise<boolean> => {
      if (!editor) return false
      const json = editor.getJSON()
      if (frontmatterRef.current && Object.keys(frontmatterRef.current).length > 0) {
        const fmNode = json.content?.find((n) => n.type === 'frontmatter')
        if (fmNode) {
          fmNode.attrs = { ...fmNode.attrs, data: frontmatterRef.current }
        } else {
          json.content = [
            { type: 'frontmatter', attrs: { data: frontmatterRef.current } },
            ...(json.content ?? []),
          ]
        }
      }
      const markdown = prosemirrorToMarkdown(json)
      const result = saveAs
        ? await window.markdownApi.saveAs(markdown)
        : await window.markdownApi.save(markdown, filePath)
      if (result.ok) {
        setDirty(false)
        if (result.path) {
          setFilePath(result.path)
          const parts = result.path.split('/')
          setFileName(parts[parts.length - 1])
        }
        return true
      }
      return false
    },
    [editor, filePath],
  )

  // ── Boot: consume pending open ──────────────────────────────────────────

  useEffect(() => {
    void (async () => {
      const pending = await window.markdownApi.consumePendingOpen()
      if (pending) await loadFile(pending)
    })()
  }, [loadFile])

  // ── Close guard (uses ref to avoid re-registration on every dirty change)

  const saveFileRef = useRef(saveFile)
  useEffect(() => {
    saveFileRef.current = saveFile
  }, [saveFile])

  useEffect(() => {
    const unsubCheck = window.markdownApi.onCloseCheck(() => {
      window.markdownApi.reportCloseCheckResult({ dirty: dirtyRef.current })
    })
    const unsubSave = window.markdownApi.onCloseSaveRequest(async () => {
      const ok = await saveFileRef.current()
      window.markdownApi.reportCloseSaveResult(ok)
    })
    return () => {
      unsubCheck()
      unsubSave()
    }
  }, [])

  // ── Dirty tracking ──────────────────────────────────────────────────────

  useEffect(() => {
    window.markdownApi.onDirtyChanged(dirty)
  }, [dirty])

  // ── Teardown ────────────────────────────────────────────────────────────

  useEffect(() => {
    return window.markdownApi.onTeardown(() => {
      editor?.destroy()
    })
  }, [editor])

  // ── Menu commands ───────────────────────────────────────────────────────

  useEffect(() => {
    return window.markdownApi.onMenuCommand((command: MarkdownMenuCommand) => {
      if (!editor) return
      switch (command) {
        // File
        case 'new':
          editor.commands.clearContent()
          setFilePath(null)
          setFileName(t('untitled'))
          setDirty(false)
          frontmatterRef.current = null
          break
        case 'open':
          void (async () => {
            const result = await window.markdownApi.open()
            if (result) await loadFile(result)
          })()
          break
        case 'open-path':
          break
        case 'save':
          void saveFile(false)
          break
        case 'save-as':
          void saveFile(true)
          break
        // Edit
        case 'undo':
          // @ts-expect-error TipTap duplicate module issue in npm workspace monorepo
          editor.commands.undo()
          break
        case 'redo':
          // @ts-expect-error TipTap duplicate module issue in npm workspace monorepo
          editor.commands.redo()
          break
        case 'find':
          break
        case 'word-count': {
          const wc = wordCountRef.current
          alert(`${wc.words} words, ${wc.chars} characters`)
          break
        }
        // View
        case 'zoom-in':
          setZoom((z) => Math.min(200, z + 10))
          break
        case 'zoom-out':
          setZoom((z) => Math.max(50, z - 10))
          break
        case 'zoom-100':
          setZoom(100)
          break
        case 'toggle-ai':
          setShowAi((v) => !v)
          break
        case 'toggle-dark':
          setThemeExplicit((v) => (v === 'dark' ? 'light' : 'dark'))
          break
        // Format
        case 'bold':
          editor.chain().focus().toggleMark('bold').run()
          break
        case 'italic':
          editor.chain().focus().toggleMark('italic').run()
          break
        case 'strike':
          editor.chain().focus().toggleMark('strike').run()
          break
        case 'code':
          editor.chain().focus().toggleMark('code').run()
          break
        case 'underline':
          editor.chain().focus().toggleMark('underline').run()
          break
        case 'heading-1':
          editor.chain().focus().toggleNode('heading', 'paragraph', { level: 1 }).run()
          break
        case 'heading-2':
          editor.chain().focus().toggleNode('heading', 'paragraph', { level: 2 }).run()
          break
        case 'heading-3':
          editor.chain().focus().toggleNode('heading', 'paragraph', { level: 3 }).run()
          break
        case 'bullet-list':
          editor.chain().focus().toggleList('bulletList', 'listItem').run()
          break
        case 'ordered-list':
          editor.chain().focus().toggleList('orderedList', 'listItem').run()
          break
        case 'task-list':
          editor.chain().focus().toggleList('taskList', 'taskItem').run()
          break
        case 'blockquote':
          editor.chain().focus().toggleNode('blockquote', 'paragraph').run()
          break
        case 'code-block':
          editor.chain().focus().toggleNode('codeBlock', 'paragraph').run()
          break
        case 'horizontal-rule':
          // @ts-expect-error TipTap duplicate module issue in npm workspace monorepo
          editor.commands.setHorizontalRule()
          break
        case 'insert-table':
          // @ts-expect-error TipTap duplicate module issue in npm workspace monorepo
          editor.commands.insertTable({ rows: 3, cols: 3, withHeaderRow: true })
          break
        case 'insert-link': {
          const url = window.prompt('Enter URL:')
          // @ts-expect-error TipTap duplicate module issue in npm workspace monorepo
          if (url) editor.commands.setLink({ href: url })
          break
        }
        case 'insert-image': {
          const url = window.prompt('Enter image URL:')
          // @ts-expect-error TipTap duplicate module issue in npm workspace monorepo
          if (url) editor.commands.setImage({ src: url })
          break
        }
        case 'align-left':
          // @ts-expect-error TipTap duplicate module issue in npm workspace monorepo
          editor.commands.setTextAlign('left')
          break
        case 'align-center':
          // @ts-expect-error TipTap duplicate module issue in npm workspace monorepo
          editor.commands.setTextAlign('center')
          break
        case 'align-right':
          // @ts-expect-error TipTap duplicate module issue in npm workspace monorepo
          editor.commands.setTextAlign('right')
          break
        case 'align-justify':
          // @ts-expect-error TipTap duplicate module issue in npm workspace monorepo
          editor.commands.setTextAlign('justify')
          break
        case 'print':
          window.markdownApi.print()
          break
        case 'export-pdf':
          window.markdownApi.exportPdf()
          break
      }
    })
  }, [editor, loadFile, saveFile, t])

  // ── Keyboard shortcuts ──────────────────────────────────────────────────

  useEffect(() => {
    const handler = (e: KeyboardEvent) => {
      if ((e.metaKey || e.ctrlKey) && e.key === 's') {
        e.preventDefault()
        void saveFile(e.shiftKey)
      }
    }
    window.addEventListener('keydown', handler)
    return () => window.removeEventListener('keydown', handler)
  }, [saveFile])

  // ── Title ───────────────────────────────────────────────────────────────

  useEffect(() => {
    document.title = `${dirty ? '* ' : ''}${fileName} — GenOffice Markdown`
  }, [fileName, dirty])

  if (!editor) return null

  return (
    <div className="app">
      <Ribbon
        editor={editor}
        activeTab={ribbonTab}
        onTabChange={setRibbonTab}
        zoom={zoom}
        onToggleDark={() => setThemeExplicit((v) => (v === 'dark' ? 'light' : 'dark'))}
      />
      <div className="app-main">
        <div className="editor-container">
          <div className="editor-scroll">
            <div
              className="editor-page"
              style={{ transform: `scale(${zoom / 100})`, transformOrigin: 'top center' }}
            >
              <EditorContent editor={editor} />
            </div>
          </div>
        </div>
        {showAi && (
          <div className="ai-panel">
            <div className="ai-panel-header">
              <span>AI Assistant</span>
              <button className="ribbon-btn" onClick={() => setShowAi(false)}>
                ×
              </button>
            </div>
            <div className="ai-panel-messages">
              <p style={{ color: 'var(--text-secondary)', fontSize: 13, padding: 12 }}>
                AI features coming soon.
              </p>
            </div>
          </div>
        )}
      </div>
      <div className="status-bar">
        <span>{wordCount.words} words</span>
        <span>{wordCount.chars} characters</span>
        {filePath && <span style={{ marginLeft: 'auto' }}>{filePath}</span>}
      </div>
    </div>
  )
}

export function App() {
  return (
    <LocaleProvider>
      <EditorApp />
    </LocaleProvider>
  )
}
