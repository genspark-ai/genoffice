import { useEffect, useRef, useState } from 'react'
import aofDocsLogo from '../assets/aof-docs-logo.png'
import type { RibbonProps } from './Ribbon'
import { GDocsMenuBar, compact, menuDivider, mi, type MenuSpec } from './GDocsMenuBar'
import { GDocsToolbar } from './GDocsToolbar'
import { IconComments } from './icons'
import { showToast } from './toast-bus'

/** Extra commands the Ribbon never needed (it had no File/Edit/... menu bar of
 *  its own — those lived only in the native OS menu). GDocsHeader needs them
 *  to build real File/Insert/Tools menus instead of dead buttons. */
export interface GDocsHeaderExtraProps {
  onNewFile: () => void
  onFind: () => void
  onFontDialog: () => void
  onWordCount: () => void
  onExportPdf: () => void
  onLink: () => void
  onEquation: () => void
  onInsertTable: () => void
  onInsertImage: () => void
  onInsertPageBreak: () => void
  /** set once the current document has been imported from / sent to Google Docs */
  googleFileId: string | null
  onImportFromGoogle: () => void
  onSendToGoogle: () => void
  /** true while a send/update is in flight — disables the send controls so a
   *  double-click can't fire two create-doc requests before googleFileId commits */
  sendingToGoogle: boolean
  onOpenGoogleShare: () => void
  /** File ▸ Google Drive folder… — set the default destination folder */
  onOpenGoogleFolderSettings: () => void
  /** File ▸ Make a copy… — duplicate the linked Google Doc; only enabled once googleFileId is set */
  onMakeGoogleCopy: () => void
  /** File ▸ Move to Drive folder… / header move icon — move the linked Google Doc; only enabled once googleFileId is set */
  onMoveGoogleFile: () => void
  /** true once the linked Google Doc is app-owned and can be renamed via Drive */
  googleWritable: boolean
  /** the title to display — App owns it (user override ?? file name) */
  displayTitle: string
  /** editable title committed (non-empty, changed). When the doc is linked +
   *  writable this also renames the Drive file; the title is always kept as
   *  the default name for the next Send/Copy regardless of link state. */
  onTitleCommit: (title: string) => void
}

type Props = RibbonProps & GDocsHeaderExtraProps

/** Send/update-to-Google-Docs glyph (upload cloud); kept local, not in
 *  icons.tsx, since it's specific to this one header button. */
function IconGoogleDrive() {
  return (
    <svg
      width={18}
      height={18}
      viewBox="0 0 16 16"
      fill="none"
      stroke="currentColor"
      strokeWidth={1.4}
      strokeLinecap="round"
      strokeLinejoin="round"
      aria-hidden
    >
      <path d="M4.5 11.5h7a2.25 2.25 0 0 0 .4-4.47 3 3 0 0 0-5.66-1.4A2.5 2.5 0 0 0 3 8.25a2.25 2.25 0 0 0 1.5 3.25z" />
      <path d="M8 6.5v4.2" />
      <path d="M6.3 9l1.7-1.7L9.7 9" />
    </svg>
  )
}

/** globe glyph shown on the left of the Share pill label, matching current
 *  Google Docs (a lock glyph shows instead once the doc is link-restricted —
 *  we always know the doc is at least owner-only, so globe is the steady
 *  state here; lock is exposed for parity via IconLock if ever needed). */
/** folder-with-arrow glyph for the "Move to Drive folder" header button. */
function IconMoveFolder() {
  return (
    <svg
      width={18}
      height={18}
      viewBox="0 0 16 16"
      fill="none"
      stroke="currentColor"
      strokeWidth={1.4}
      strokeLinecap="round"
      strokeLinejoin="round"
      aria-hidden
    >
      <path d="M1.5 4.25A1.25 1.25 0 0 1 2.75 3H6l1.25 1.5h6a1.25 1.25 0 0 1 1.25 1.25v5A1.25 1.25 0 0 1 13.25 12h-10.5A1.25 1.25 0 0 1 1.5 10.75z" />
      <path d="M8 6.75v3.1" />
      <path d="M6.4 8.35 8 6.75l1.6 1.6" />
    </svg>
  )
}

/** two-page copy glyph for the header "Make a copy" button. */
function IconCopyDoc() {
  return (
    <svg
      width={18}
      height={18}
      viewBox="0 0 16 16"
      fill="none"
      stroke="currentColor"
      strokeWidth={1.4}
      strokeLinecap="round"
      strokeLinejoin="round"
      aria-hidden
    >
      <rect x="5.5" y="5.5" width="8" height="9" rx="1.5" />
      <path d="M10.5 3h-6A1.5 1.5 0 0 0 3 4.5v7" />
    </svg>
  )
}

function IconGlobeSmall() {
  return (
    <svg width={16} height={16} viewBox="0 0 24 24" fill="none" aria-hidden>
      <circle cx="12" cy="12" r="9" stroke="currentColor" strokeWidth="1.6" />
      <path
        d="M3 12h18M12 3a14 14 0 0 1 0 18M12 3a14 14 0 0 0 0 18"
        stroke="currentColor"
        strokeWidth="1.6"
      />
    </svg>
  )
}

function IconCaretSmall() {
  return (
    <svg width={12} height={12} viewBox="0 0 24 24" fill="none" aria-hidden>
      <path
        d="M6 9l6 6 6-6"
        stroke="currentColor"
        strokeWidth="2"
        strokeLinecap="round"
        strokeLinejoin="round"
      />
    </svg>
  )
}

function isMac() {
  return navigator.platform.toLowerCase().includes('mac')
}
const MOD = isMac() ? '⌘' : 'Ctrl+'

export function GDocsHeader(props: Props) {
  const {
    editor,
    formatState,
    hasDoc,
    blocks,
    allocateNumId,
    createListDef,
    docDefaults,
    themeFonts,
    zoom,
    onZoom,
    onZoomFit,
    filePath,
    commentCount,
    canComment,
    onShowComments,
    onNewComment,
    showMarks,
    onShowMarks,
    showRuler,
    onShowRuler,
    darkCanvas,
    onDarkCanvas,
    viewMode,
    onViewMode,
    trackChanges,
    onTrackChanges,
    isProtected,
    onToggleProtection,
    onOpen,
    onSave,
    onSaveAs,
    onInsertNote,
    onCompare,
    onPagePreview,
    onNewFile,
    onFind,
    onFontDialog,
    onWordCount,
    onExportPdf,
    onLink,
    onEquation,
    onInsertTable,
    onInsertImage,
    onInsertPageBreak,
    googleFileId,
    googleWritable,
    onImportFromGoogle,
    onSendToGoogle,
    sendingToGoogle,
    onOpenGoogleShare,
    onOpenGoogleFolderSettings,
    onMakeGoogleCopy,
    onMoveGoogleFile,
    displayTitle,
    onTitleCommit,
  } = props

  // The displayed title is CONTROLLED by App (displayTitle = user override ??
  // file name) — this component holds only a draft while the field is being
  // edited. Local always-on title state kept reverting: any filePath change
  // (autosave assigning a name, save-as) re-synced over the user's text.
  const [draftTitle, setDraftTitle] = useState<string | null>(null)
  const editingTitle = draftTitle !== null
  const title = draftTitle ?? displayTitle
  const setTitle = setDraftTitle
  const setEditingTitle = (on: boolean) => setDraftTitle(on ? displayTitle : null)
  const titleInputRef = useRef<HTMLInputElement>(null)

  useEffect(() => {
    if (editingTitle) titleInputRef.current?.select()
  }, [editingTitle])

  // NOTE: genoffice has no local-disk rename command (only Save As, which
  // prompts a native file dialog) — this field never renames the file on
  // disk. When the doc is linked + writable to Google, committing it DOES
  // rename the Drive file. Either way it becomes the default name for the
  // next Send to Google Docs / Make a copy.
  const commitTitle = () => {
    const trimmed = (draftTitle ?? '').trim()
    setDraftTitle(null)
    if (trimmed === '' || trimmed === displayTitle) return // empty or unchanged
    onTitleCommit(trimmed)
    if (googleFileId && googleWritable) {
      void window.desktop.googleRenameFile(googleFileId, trimmed).then((result) => {
        if (!result.ok) {
          showToast(`Rename in Google Drive failed: ${result.error}`, 'error')
        }
      })
    }
  }

  const menus: MenuSpec[] = [
    {
      key: 'file',
      label: 'File',
      items: compact([
        mi('New document', onNewFile),
        mi('Open…', onOpen, { shortcut: `${MOD}O` }),
        menuDivider(),
        mi('Save', onSave, { shortcut: `${MOD}S` }),
        mi('Save as…', onSaveAs, { shortcut: `${MOD}⇧S` }),
        menuDivider(),
        mi('Import from Google Docs…', onImportFromGoogle),
        mi('Send to Google Docs', hasDoc && !sendingToGoogle ? onSendToGoogle : undefined),
        mi('Make a copy…', googleFileId ? onMakeGoogleCopy : undefined),
        mi('Move to Drive folder…', googleFileId ? onMoveGoogleFile : undefined),
        mi('Google Drive folder…', onOpenGoogleFolderSettings),
        menuDivider(),
        mi('Download as PDF', onExportPdf),
        menuDivider(),
        mi('Print', () => window.print(), { shortcut: `${MOD}P` }),
      ]),
    },
    {
      key: 'edit',
      label: 'Edit',
      items: compact([
        mi('Undo', () => editor.chain().focus().undo().run(), { shortcut: `${MOD}Z` }),
        mi('Redo', () => editor.chain().focus().redo().run(), { shortcut: `${MOD}⇧Z` }),
        menuDivider(),
        mi('Find…', onFind, { shortcut: `${MOD}F` }),
      ]),
    },
    {
      key: 'view',
      label: 'View',
      items: compact([
        mi('Show ruler', () => onShowRuler(!showRuler), { checked: showRuler }),
        mi('Show formatting marks', () => onShowMarks(!showMarks), { checked: showMarks }),
        mi('Dark background', () => onDarkCanvas(!darkCanvas), { checked: darkCanvas }),
        menuDivider(),
        mi('Print layout', () => onViewMode('print'), { checked: viewMode === 'print' }),
        mi('Web layout', () => onViewMode('web'), { checked: viewMode === 'web' }),
        mi('Outline', () => onViewMode('outline'), { checked: viewMode === 'outline' }),
        menuDivider(),
        mi('Page preview…', onPagePreview),
      ]),
    },
    {
      key: 'insert',
      label: 'Insert',
      items: compact([
        mi('Image…', onInsertImage),
        mi('Table (3×3)', onInsertTable),
        mi('Link', onLink, { shortcut: `${MOD}K` }),
        mi('Comment', () => {
          onShowComments()
          onNewComment()
        }),
        menuDivider(),
        mi('Page break', onInsertPageBreak, { shortcut: `${MOD}⏎` }),
        mi('Equation', onEquation),
        mi('Footnote', onInsertNote ? () => onInsertNote('footnote') : undefined),
        mi('Endnote', onInsertNote ? () => onInsertNote('endnote') : undefined),
      ]),
    },
    {
      key: 'format',
      label: 'Format',
      items: compact([
        mi('Text', undefined, {
          submenu: compact([
            mi('Bold', () => editor.chain().focus().toggleMark('bold').run(), {
              shortcut: `${MOD}B`,
              checked: formatState.bold,
            }),
            mi('Italic', () => editor.chain().focus().toggleMark('italic').run(), {
              shortcut: `${MOD}I`,
              checked: formatState.italic,
            }),
            mi('Underline', () => editor.chain().focus().toggleMark('underline').run(), {
              shortcut: `${MOD}U`,
              checked: formatState.underline,
            }),
            mi('Strikethrough', () => editor.chain().focus().toggleMark('strike').run(), {
              checked: formatState.strike,
            }),
          ]),
        }),
        mi('Align & indent', undefined, {
          submenu: compact([
            mi('Left', () => setAlign(editor, 'left')),
            mi('Center', () => setAlign(editor, 'center')),
            mi('Right', () => setAlign(editor, 'right')),
            mi('Justify', () => setAlign(editor, 'justify')),
          ]),
        }),
        mi('Bulleted list', () => toggleList(editor, blocks, 'bullet', allocateNumId), {
          checked: formatState.listBullet,
        }),
        mi('Numbered list', () => toggleList(editor, blocks, 'ordered', allocateNumId), {
          checked: formatState.listOrdered,
        }),
        menuDivider(),
        mi('Font…', onFontDialog),
        mi('Clear formatting', () => editor.chain().focus().unsetAllMarks().run(), {
          shortcut: `${MOD}\\`,
        }),
      ]),
    },
    {
      key: 'tools',
      label: 'Tools',
      items: compact([
        mi('Word count…', onWordCount, { shortcut: `${MOD}⇧C` }),
        mi('Track changes', () => onTrackChanges(!trackChanges), { checked: trackChanges }),
        mi('Protect document', onToggleProtection, { checked: isProtected }),
        menuDivider(),
        mi('Compare documents…', onCompare),
      ]),
    },
  ]

  // createListDef currently unused directly in the header (list creation from
  // scratch reuses the ribbon's numbering flow); kept in props for parity.
  void createListDef

  return (
    <div className="gdh-root">
      <div className="gdh-titlebar" />
      <div className="gdh-toprow">
        <img className="gdh-logo" src={aofDocsLogo} alt="Artisans Docs" height={66} />
        <div className="gdh-title-menu">
          <div className="gdh-title-row">
            {editingTitle ? (
              <input
                ref={titleInputRef}
                className="gdh-title-input"
                value={title}
                onChange={(e) => setTitle(e.target.value)}
                onBlur={commitTitle}
                onKeyDown={(e) => {
                  if (e.key === 'Enter') (e.target as HTMLInputElement).blur()
                  if (e.key === 'Escape') setDraftTitle(null) // cancel edit, keep current title
                }}
              />
            ) : (
              <button
                type="button"
                className="gdh-title-display"
                title={
                  googleFileId && googleWritable
                    ? 'Rename — renames the Google Doc and sets the default name for future sends'
                    : 'Rename — sets the default name for the next Send to Google Docs (use File ▸ Save as… to change the saved file name)'
                }
                onClick={() => setEditingTitle(true)}
              >
                {title}
              </button>
            )}
          </div>
          <GDocsMenuBar menus={menus} />
        </div>
        <div className="gdh-toprow-right">
          <button
            type="button"
            className="gdh-icon-btn"
            title={`Comments${commentCount ? ` (${commentCount})` : ''}`}
            data-tip={`Comments${commentCount ? ` (${commentCount})` : ''}`}
            onClick={onShowComments}
          >
            <IconComments />
            {commentCount > 0 && <span className="gdh-badge">{commentCount}</span>}
          </button>
          <button
            type="button"
            className="gdh-icon-btn gdh-google-send-btn"
            disabled={!hasDoc || sendingToGoogle}
            title={googleFileId ? 'Update Google Doc' : 'Send to Google Docs'}
            data-tip={googleFileId ? 'Update Google Doc' : 'Send to Google Docs'}
            onClick={onSendToGoogle}
          >
            <IconGoogleDrive />
          </button>
          {googleFileId && (
            <button
              type="button"
              className="gdh-icon-btn"
              title="Move to Drive folder"
              data-tip="Move to Drive folder"
              onClick={onMoveGoogleFile}
            >
              <IconMoveFolder />
            </button>
          )}
          {googleFileId && (
            <button
              type="button"
              className="gdh-icon-btn"
              title="Make a copy"
              data-tip="Make a copy"
              onClick={onMakeGoogleCopy}
            >
              <IconCopyDoc />
            </button>
          )}
          <div className="gdh-share-split">
            <button
              type="button"
              className="gdh-share-btn"
              disabled={!hasDoc}
              title="Share"
              onClick={onOpenGoogleShare}
            >
              <IconGlobeSmall />
              Share
            </button>
            <button
              type="button"
              className="gdh-share-caret"
              disabled={!hasDoc}
              title="Share options"
              onClick={onOpenGoogleShare}
            >
              <IconCaretSmall />
            </button>
          </div>
        </div>
      </div>

      <GDocsToolbar
        editor={editor}
        formatState={formatState}
        hasDoc={hasDoc}
        blocks={blocks}
        allocateNumId={allocateNumId}
        createListDef={createListDef}
        docDefaults={docDefaults}
        themeFonts={themeFonts}
        zoom={zoom}
        onZoom={onZoom}
        onZoomFit={onZoomFit}
        onLink={onLink}
        onInsertImage={onInsertImage}
        onNewComment={() => {
          onShowComments()
          onNewComment()
        }}
        canComment={canComment}
      />
    </div>
  )
}

function baseTitle(filePath: string | null): string {
  if (!filePath) return 'Untitled document'
  const name = filePath.split(/[\\/]/).pop() ?? filePath
  return name.replace(/\.docx$/i, '')
}

function setAlign(editor: RibbonProps['editor'], align: 'left' | 'center' | 'right' | 'justify') {
  editor
    .chain()
    .focus()
    .updateAttributes('docParagraph', { align })
    .updateAttributes('docHeading', { align })
    .updateAttributes('docListItem', { align })
    .run()
}

function toggleList(
  editor: RibbonProps['editor'],
  blocks: RibbonProps['blocks'],
  kind: 'bullet' | 'ordered',
  allocateNumId: RibbonProps['allocateNumId'],
) {
  if (editor.isActive('docListItem', { kind })) {
    editor.chain().focus().setNode('docParagraph').run()
    return
  }
  const existing = blocks.find((b) => b.type === 'listItem' && b.list?.kind === kind)
  const numId = existing?.list?.numId ?? allocateNumId?.(kind) ?? null
  editor.chain().focus().setNode('docListItem', { kind, numId, ilvl: 0 }).run()
}
