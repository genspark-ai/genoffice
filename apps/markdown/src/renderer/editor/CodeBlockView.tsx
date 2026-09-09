import { useEffect, useRef, useState } from 'react'
import { NodeViewContent, NodeViewWrapper } from '@tiptap/react'
import type { NodeViewProps } from '@tiptap/react'
import { Dropdown } from '@genoffice/ui'
import { t } from '../i18n/locale'
import { MERMAID_LANGUAGE, renderMermaid } from './mermaid'
import type { MermaidResult } from './mermaid'

const LANGUAGES = [
  'plaintext',
  'bash',
  'c',
  'cpp',
  'csharp',
  'css',
  'diff',
  'dockerfile',
  'go',
  'graphql',
  'html',
  'java',
  'javascript',
  'json',
  'kotlin',
  'lua',
  'markdown',
  MERMAID_LANGUAGE,
  'objectivec',
  'php',
  'python',
  'r',
  'ruby',
  'rust',
  'scala',
  'scss',
  'sql',
  'swift',
  'typescript',
  'xml',
  'yaml',
]

const RERENDER_DEBOUNCE_MS = 300

export function CodeBlockView({ node, updateAttributes, editor, getPos }: NodeViewProps) {
  const [copied, setCopied] = useState(false)
  const copyTimerRef = useRef<number | null>(null)
  const mountedRef = useRef(true)
  const language = String(node.attrs.language ?? '') || 'plaintext'
  const isMermaid = language === MERMAID_LANGUAGE
  const source = node.textContent

  const [diagram, setDiagram] = useState<MermaidResult | null>(null)
  const hasDiagramRef = useRef(false)
  hasDiagramRef.current = diagram !== null
  const [caretInside, setCaretInside] = useState(false)

  useEffect(() => {
    mountedRef.current = true
    return () => {
      mountedRef.current = false
      if (copyTimerRef.current !== null) window.clearTimeout(copyTimerRef.current)
    }
  }, [])

  useEffect(() => {
    if (!isMermaid) return
    const update = () => {
      const pos = getPos()
      if (pos === undefined) return
      // size from the live document: the `node` prop lags one render behind a paste/IME commit
      const current = editor.state.doc.nodeAt(pos)
      if (!current) return
      const { from, to } = editor.state.selection
      setCaretInside(from >= pos && to <= pos + current.nodeSize)
    }
    update()
    editor.on('selectionUpdate', update)
    return () => {
      editor.off('selectionUpdate', update)
    }
  }, [editor, getPos, isMermaid])

  useEffect(() => {
    if (!isMermaid || !source.trim()) {
      setDiagram(null)
      return
    }
    let cancelled = false
    // first paint right away, then debounce while the user types
    const delay = hasDiagramRef.current ? RERENDER_DEBOUNCE_MS : 0
    const timer = window.setTimeout(() => {
      void renderMermaid(source).then((result) => {
        if (!cancelled) setDiagram(result)
      })
    }, delay)
    return () => {
      cancelled = true
      window.clearTimeout(timer)
    }
  }, [isMermaid, source])

  const copy = () => {
    void navigator.clipboard
      .writeText(node.textContent)
      .then(() => {
        if (!mountedRef.current) return
        if (copyTimerRef.current !== null) window.clearTimeout(copyTimerRef.current)
        setCopied(true)
        copyTimerRef.current = window.setTimeout(() => {
          copyTimerRef.current = null
          setCopied(false)
        }, 1500)
      })
      .catch(() => {})
  }

  const editSource = () => {
    const pos = getPos()
    if (pos === undefined || !editor.isEditable) return
    editor
      .chain()
      .focus()
      .setTextSelection(pos + node.nodeSize - 1)
      .run()
  }

  const rendered = isMermaid && diagram?.ok === true
  const showSource = !rendered || caretInside
  const error = isMermaid && diagram && !diagram.ok && source.trim() ? diagram.error : null

  const className = [
    'md-codeblock',
    isMermaid && 'md-mermaid',
    !showSource && 'md-mermaid-collapsed',
  ]
    .filter(Boolean)
    .join(' ')

  return (
    <NodeViewWrapper className={className} data-mermaid={rendered ? 'rendered' : undefined}>
      <div className="md-codeblock-bar" contentEditable={false}>
        <Dropdown
          className="md-codeblock-lang"
          value={LANGUAGES.includes(language) ? language : 'plaintext'}
          disabled={!editor.isEditable}
          options={LANGUAGES.map((lang) => ({ value: lang, label: lang }))}
          onPick={(lang) => updateAttributes({ language: lang === 'plaintext' ? null : lang })}
        />
        <button type="button" className="md-codeblock-copy" onClick={copy}>
          {copied ? t('codeCopied') : t('codeCopy')}
        </button>
      </div>
      <pre>
        <NodeViewContent<'code'> as="code" />
      </pre>
      {error && (
        <div className="md-mermaid-error" contentEditable={false}>
          {t('mermaidError')}: {error}
        </div>
      )}
      {isMermaid && diagram?.ok && (
        <div
          className="md-mermaid-preview"
          contentEditable={false}
          onClick={editSource}
          dangerouslySetInnerHTML={{ __html: diagram.svg }}
        />
      )}
    </NodeViewWrapper>
  )
}
