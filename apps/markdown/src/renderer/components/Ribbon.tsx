import { useEffect, useRef, useState } from 'react'
import type { ReactNode } from 'react'
import type { Editor } from '@tiptap/core'
import { useEditorState } from '@tiptap/react'
import { useI18n } from '../i18n/locale'
import type { StringKey } from '../i18n/locale'
import { GensparkMark } from '../ai/AiPanel'
import { liftFromList } from '../editor/slashCommand'
import {
  IconBullets,
  IconHighlight,
  IconHr,
  IconIndentDec,
  IconIndentInc,
  IconInlineCode,
  IconLink,
  IconNumbered,
  IconPicture,
  IconProperties,
  IconRedo,
  IconTable,
  IconTaskList,
  IconUndo,
} from './icons'

interface Props {
  editor: Editor | null
  disabled: boolean
  imageEnabled: boolean
  onInsertImage: () => void
  frontmatterOpen: boolean
  onToggleFrontmatter: () => void
  aiOpen: boolean
  onToggleAi: () => void
  onAiPreset: (instruction: string) => void
}

type BlockStyle =
  | 'paragraph'
  | 'h1'
  | 'h2'
  | 'h3'
  | 'h4'
  | 'h5'
  | 'h6'
  | 'quote'
  | 'codeBlock'
  | 'callout'
  | 'toggle'

const STYLE_LABEL: Record<BlockStyle, StringKey> = {
  paragraph: 'styleParagraph',
  h1: 'styleH1',
  h2: 'styleH2',
  h3: 'styleH3',
  h4: 'styleH4',
  h5: 'styleH5',
  h6: 'styleH6',
  quote: 'styleQuote',
  codeBlock: 'styleCodeBlock',
  callout: 'insertCallout',
  toggle: 'insertToggle',
}

function applyBlockStyle(editor: Editor, style: BlockStyle): void {
  // block-type conversions are illegal inside list items — leave the list first
  liftFromList(editor)
  const chain = editor.chain().focus()
  // unwrap wrappers the target style does not keep
  if (style !== 'callout' && editor.isActive('callout')) chain.unsetCallout()
  if (style !== 'toggle' && editor.isActive('toggle')) chain.unsetToggle()
  switch (style) {
    case 'paragraph':
      chain.setParagraph().run()
      break
    case 'quote':
      chain.setParagraph().setBlockquote().run()
      break
    case 'codeBlock':
      chain.setCodeBlock().run()
      break
    case 'callout':
      chain.setCallout().run()
      break
    case 'toggle':
      chain.setToggle().run()
      break
    default:
      chain.setHeading({ level: Number(style.slice(1)) as 1 | 2 | 3 | 4 | 5 | 6 }).run()
  }
}

/** doc + sparkle / pen + sparkle / lines + sparkle, same glyphs as the docs ribbon */
function AiFeatureIcon({ kind }: { kind: 'summarize' | 'polish' | 'tidy' }) {
  return (
    <svg
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth="1.5"
      strokeLinecap="round"
      strokeLinejoin="round"
      aria-hidden
    >
      {kind === 'summarize' && (
        <>
          <path d="M13.875 21H6.5C5.39543 21 4.5 20.1046 4.5 19V5C4.5 3.89543 5.39543 3 6.5 3H17.5C18.6046 3 19.5 3.89543 19.5 5V13" />
          <path d="M8 7H16" />
          <path d="M8 10.2H14" />
          <path d="M8 13.4H12" />
        </>
      )}
      {kind === 'polish' && (
        <>
          <path d="M5.00012 20.7481L8.80319 20.7482L21.7482 7.80317L17.945 4L5 16.945L5.00012 20.7481Z" />
          <path d="M15.1406 6.80469L18.9438 10.6079" />
          <path d="M8 3L8.22106 3.59745C8.51094 4.38087 8.65589 4.77259 8.94166 5.05833C9.22743 5.34409 9.61914 5.48903 10.4026 5.77893L11 6L10.4026 6.22107C9.61914 6.51097 9.22743 6.65592 8.94166 6.94167C8.65589 7.22741 8.51094 7.61913 8.22106 8.40255L8 9L7.77894 8.40255C7.48906 7.61913 7.34411 7.22741 7.05834 6.94167C6.77257 6.65592 6.38086 6.51097 5.59743 6.22107L5 6L5.59743 5.77893C6.38086 5.48903 6.77257 5.34409 7.05834 5.05833C7.34411 4.77259 7.48906 4.38087 7.77894 3.59745L8 3Z" />
        </>
      )}
      {kind === 'tidy' && (
        <>
          <path d="M4 5H20" />
          <path d="M4 9H16" />
          <path d="M4 13H11" />
          <path d="M4 17H10" />
        </>
      )}
      {kind !== 'polish' && (
        <path d="M17 14L17.2579 14.697C17.5961 15.611 17.7652 16.068 18.0986 16.4014C18.432 16.7348 18.889 16.9039 19.803 17.2421L20.5 17.5L19.803 17.7579C18.889 18.0961 18.432 18.2652 18.0986 18.5986C17.7652 18.932 17.5961 19.389 17.2579 20.303L17 21L16.7421 20.303C16.4039 19.389 16.2348 18.932 15.9014 18.5986C15.568 18.2652 15.111 18.0961 14.197 17.7579L13.5 17.5L14.197 17.2421C15.111 16.9039 15.568 16.7348 15.9014 16.4014C16.2348 16.068 16.4039 15.611 16.7421 14.697L17 14Z" />
      )}
    </svg>
  )
}

function IconBtn({
  title,
  active,
  disabled,
  onClick,
  children,
}: {
  title: string
  active?: boolean
  disabled?: boolean
  onClick: () => void
  children: ReactNode
}) {
  return (
    <button
      type="button"
      className={`rb-btn${active ? ' active' : ''}`}
      title={title}
      disabled={disabled}
      onMouseDown={(e) => e.preventDefault()}
      onClick={onClick}
    >
      {children}
    </button>
  )
}

export function Ribbon({
  editor,
  disabled,
  imageEnabled,
  onInsertImage,
  frontmatterOpen,
  onToggleFrontmatter,
  aiOpen,
  onToggleAi,
  onAiPreset,
}: Props) {
  const { t } = useI18n()
  const [linkOpen, setLinkOpen] = useState(false)
  const [linkUrl, setLinkUrl] = useState('')
  const linkInputRef = useRef<HTMLInputElement>(null)

  const state = useEditorState({
    editor,
    selector: ({ editor: e }) => {
      if (!e) return null
      const style: BlockStyle = e.isActive('callout')
        ? 'callout'
        : e.isActive('toggle')
          ? 'toggle'
          : e.isActive('codeBlock')
            ? 'codeBlock'
            : e.isActive('blockquote')
              ? 'quote'
              : e.isActive('heading')
                ? (`h${e.getAttributes('heading').level}` as BlockStyle)
                : 'paragraph'
      return {
        style,
        empty: e.isEmpty,
        bold: e.isActive('bold'),
        italic: e.isActive('italic'),
        strike: e.isActive('strike'),
        underline: e.isActive('underline'),
        highlight: e.isActive('highlight'),
        code: e.isActive('code'),
        link: e.isActive('link'),
        bullet: e.isActive('bulletList'),
        ordered: e.isActive('orderedList'),
        task: e.isActive('taskList'),
        inList: e.isActive('listItem') || e.isActive('taskItem'),
        canUndo: e.can().undo(),
        canRedo: e.can().redo(),
      }
    },
  })

  useEffect(() => {
    if (linkOpen) linkInputRef.current?.focus()
  }, [linkOpen])

  const off = disabled || !editor || !state

  const openLink = () => {
    if (!editor) return
    setLinkUrl(String(editor.getAttributes('link').href ?? ''))
    setLinkOpen((v) => !v)
  }

  const applyLink = () => {
    if (!editor) return
    const url = linkUrl.trim()
    const chain = editor.chain().focus().extendMarkRange('link')
    if (url) chain.setLink({ href: url }).run()
    else chain.unsetLink().run()
    setLinkOpen(false)
  }

  const listItemName = state?.task ? 'taskItem' : 'listItem'
  const ICON = 16

  const aiPresets = [
    { kind: 'summarize', btn: 'aiSummarizeBtn', prompt: 'aiSummarizePrompt' },
    { kind: 'polish', btn: 'aiPolishBtn', prompt: 'aiPolishPrompt' },
    { kind: 'tidy', btn: 'aiTidyBtn', prompt: 'aiTidyPrompt' },
  ] as const

  return (
    <div className="ribbon">
      <div className="ribbon-group">
        <div className="ribbon-group-items">
          <button
            type="button"
            className={`rb-big ai-entry${aiOpen ? ' active' : ''}`}
            title={t('aiOpenAssistant')}
            disabled={disabled}
            onMouseDown={(e) => e.preventDefault()}
            onClick={onToggleAi}
          >
            <span className="rb-big-icon">
              <GensparkMark size={26} />
            </span>
            <span>Genspark AI</span>
          </button>
          {aiPresets.map(({ kind, btn, prompt }) => (
            <button
              key={kind}
              type="button"
              className="rb-big ai-entry"
              title={t(prompt)}
              disabled={off || state?.empty}
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => onAiPreset(t(prompt))}
            >
              <span className="rb-big-icon">
                <span className="ai-feature-icon" aria-hidden="true">
                  <AiFeatureIcon kind={kind} />
                </span>
              </span>
              <span>{t(btn)}</span>
            </button>
          ))}
        </div>
      </div>

      <div className="rb-sep" />

      <div className="rb-group">
        <IconBtn
          title={t('undo')}
          disabled={off || !state?.canUndo}
          onClick={() => editor?.chain().focus().undo().run()}
        >
          <IconUndo size={ICON} />
        </IconBtn>
        <IconBtn
          title={t('redo')}
          disabled={off || !state?.canRedo}
          onClick={() => editor?.chain().focus().redo().run()}
        >
          <IconRedo size={ICON} />
        </IconBtn>
      </div>

      <div className="rb-sep" />

      <select
        className="rb-style"
        value={state?.style ?? 'paragraph'}
        disabled={off}
        onChange={(e) => editor && applyBlockStyle(editor, e.target.value as BlockStyle)}
      >
        {(Object.keys(STYLE_LABEL) as BlockStyle[]).map((s) => (
          <option key={s} value={s}>
            {t(STYLE_LABEL[s])}
          </option>
        ))}
      </select>

      <div className="rb-sep" />

      <div className="rb-group">
        <IconBtn
          title={t('bold')}
          active={state?.bold}
          disabled={off}
          onClick={() => editor?.chain().focus().toggleBold().run()}
        >
          <span className="rb-glyph" style={{ fontWeight: 800 }}>
            B
          </span>
        </IconBtn>
        <IconBtn
          title={t('italic')}
          active={state?.italic}
          disabled={off}
          onClick={() => editor?.chain().focus().toggleItalic().run()}
        >
          <span className="rb-glyph" style={{ fontStyle: 'italic' }}>
            I
          </span>
        </IconBtn>
        <IconBtn
          title={t('underline')}
          active={state?.underline}
          disabled={off}
          onClick={() => editor?.chain().focus().toggleUnderline().run()}
        >
          <span className="rb-glyph" style={{ textDecoration: 'underline' }}>
            U
          </span>
        </IconBtn>
        <IconBtn
          title={t('strike')}
          active={state?.strike}
          disabled={off}
          onClick={() => editor?.chain().focus().toggleStrike().run()}
        >
          <span className="rb-glyph" style={{ textDecoration: 'line-through' }}>
            S
          </span>
        </IconBtn>
        <IconBtn
          title={t('highlight')}
          active={state?.highlight}
          disabled={off}
          onClick={() => editor?.chain().focus().toggleHighlight().run()}
        >
          <IconHighlight size={ICON} />
        </IconBtn>
        <IconBtn
          title={t('inlineCode')}
          active={state?.code}
          disabled={off}
          onClick={() => editor?.chain().focus().toggleCode().run()}
        >
          <IconInlineCode size={ICON} />
        </IconBtn>
        <span className="rb-link-anchor">
          <IconBtn title={t('link')} active={state?.link} disabled={off} onClick={openLink}>
            <IconLink size={ICON} />
          </IconBtn>
          {linkOpen && (
            <span className="rb-link-pop" onMouseDown={(e) => e.stopPropagation()}>
              <input
                ref={linkInputRef}
                value={linkUrl}
                placeholder={t('linkPlaceholder')}
                onChange={(e) => setLinkUrl(e.target.value)}
                onKeyDown={(e) => {
                  if (e.key === 'Enter') applyLink()
                  if (e.key === 'Escape') setLinkOpen(false)
                }}
              />
              <button type="button" onClick={applyLink}>
                {t('linkApply')}
              </button>
            </span>
          )}
        </span>
      </div>

      <div className="rb-sep" />

      <div className="rb-group">
        <IconBtn
          title={t('bulletList')}
          active={state?.bullet}
          disabled={off}
          onClick={() => editor?.chain().focus().toggleBulletList().run()}
        >
          <IconBullets size={ICON} />
        </IconBtn>
        <IconBtn
          title={t('orderedList')}
          active={state?.ordered}
          disabled={off}
          onClick={() => editor?.chain().focus().toggleOrderedList().run()}
        >
          <IconNumbered size={ICON} />
        </IconBtn>
        <IconBtn
          title={t('taskList')}
          active={state?.task}
          disabled={off}
          onClick={() => editor?.chain().focus().toggleTaskList().run()}
        >
          <IconTaskList size={ICON} />
        </IconBtn>
        <IconBtn
          title={t('outdent')}
          disabled={off || !state?.inList}
          onClick={() => editor?.chain().focus().liftListItem(listItemName).run()}
        >
          <IconIndentDec size={ICON} />
        </IconBtn>
        <IconBtn
          title={t('indent')}
          disabled={off || !state?.inList}
          onClick={() => editor?.chain().focus().sinkListItem(listItemName).run()}
        >
          <IconIndentInc size={ICON} />
        </IconBtn>
      </div>

      <div className="rb-sep" />

      <div className="rb-group">
        <IconBtn
          title={t('insertTable')}
          disabled={off}
          onClick={() =>
            editor?.chain().focus().insertTable({ rows: 3, cols: 3, withHeaderRow: true }).run()
          }
        >
          <IconTable size={ICON} />
        </IconBtn>
        <IconBtn title={t('insertImage')} disabled={off || !imageEnabled} onClick={onInsertImage}>
          <IconPicture size={ICON} />
        </IconBtn>
        <IconBtn
          title={t('insertHr')}
          disabled={off}
          onClick={() => editor?.chain().focus().setHorizontalRule().run()}
        >
          <IconHr size={ICON} />
        </IconBtn>
      </div>

      <div className="rb-spacer" />

      <IconBtn
        title={t('fmProperties')}
        active={frontmatterOpen}
        disabled={disabled}
        onClick={onToggleFrontmatter}
      >
        <IconProperties size={ICON} />
      </IconBtn>
    </div>
  )
}
