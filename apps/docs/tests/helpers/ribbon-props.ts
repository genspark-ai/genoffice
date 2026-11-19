import type { Editor } from '@tiptap/core'
import type { ActionCommandId } from '../../src/shared/commands'
import { runMenuCommand, type CommandHost } from '../../src/renderer/commands/menu-commands'
import { applyGalleryStyle } from '../../src/renderer/components/ribbon-tabs'
import { FALLBACK_STYLES } from '../../src/renderer/style-gallery'
import type { RibbonFormatState } from '../../src/renderer/components/ribbon-format-state'
import { MERGE_RANGE_ALL } from '../../src/renderer/mail-merge/range'

const noop = () => {}

/** runs ribbon commands on the editor; beyond applying a built-in style, what a command asks of the app does nothing */
export function commandRunner(editor: Editor, hasDoc = true) {
  const applyStyle: CommandHost['applyStyle'] = (styleId) => {
    const info = FALLBACK_STYLES.get(styleId)
    if (info) applyGalleryStyle(editor, null, info, undefined, null)
  }
  const base = { editor, hasDoc, tableSectionWidthPx: 600, applyStyle }
  const host = new Proxy(base as CommandHost, {
    get: (target, key) => (key in target ? target[key as keyof CommandHost] : noop),
  })
  return (id: ActionCommandId, payload?: string) => runMenuCommand(host, id, payload)
}

/** Every Ribbon prop, so a test can mount the real component and drive one control. */
export function ribbonProps(editor: Editor, formatState: RibbonFormatState) {
  return {
    editor,
    formatState,
    hasDoc: true,
    onOpen: noop,
    onSave: noop,
    onSaveAs: noop,
    onNameWithAi: noop,
    showAi: false,
    onCommand: commandRunner(editor),
    editTarget: {
      hasSelection: !editor.state.selection.empty,
      editable: editor.isEditable,
      field: false,
    },
    section: null,
    onSection: noop,
    activeSection: null,
    onInsertSectionBreak: noop,
    onPaperSizeAll: noop,
    onPageBorderAll: noop,
    mirrorMargins: false,
    onMirrorMargins: noop,
    pageColor: null,
    onPageColor: noop,
    watermark: null,
    watermarkPicture: null,
    watermarkDocImage: null,
    themeFonts: null,
    onThemeFonts: noop,
    themeColors: null,
    onThemeColors: noop,
    inkTool: 'select' as const,
    onInkTool: noop,
    inkPresets: [{ id: 'p1', kind: 'pen' as const, color: 'C00000', width: 2 }],
    inkActiveId: 'p1',
    onInkPresets: noop,
    onInkActivate: noop,
    inkCount: 0,
    onInkClearAll: noop,
    onInsertNote: noop,
    sources: [],
    onAddSource: noop,
    zoom: 100,
    onZoom: noop,
    onZoomFit: noop,
    onZoomDialog: noop,
    onTrackChangesOptions: noop,
    darkPage: false,
    onDarkPage: noop,
    header: null,
    onHeader: noop,
    onPageNumFormat: noop,
    footer: null,
    onFooter: noop,
    hfEditing: null,
    onHfAction: noop,
    onHfEdit: noop,
    showMarks: false,
    showRuler: false,
    onShowRuler: noop,
    showNav: false,
    onShowNav: noop,
    commentCount: 0,
    openCommentCount: 0,
    resolvedCommentCount: 0,
    onShowComments: noop,
    showComments: false,
    onToggleComments: noop,
    commentsMode: 'contextual' as const,
    onCommentsMode: noop,
    onNewComment: noop,
    commentAtCaret: false,
    onDeleteComment: noop,
    onDeleteAllComments: noop,
    onGotoComment: noop,
    trackChanges: false,
    onTrackChanges: noop,
    spellcheck: true,
    onSpellcheck: noop,
    revisionDisplay: 'all' as const,
    onRevisionDisplay: noop,
    balloonMode: 'comments' as const,
    onBalloonMode: noop,
    revisionCount: 0,
    onAcceptRevision: noop,
    onRejectRevision: noop,
    onGotoRevision: noop,
    revisionFilter: { insDel: true, format: true, hiddenAuthors: [] },
    onRevisionFilter: noop,
    revisionAuthors: [],
    showReviewPane: false,
    onToggleReviewPane: noop,
    isProtected: false,
    commentsAllowed: false,
    trackChangesForced: false,
    protectActive: false,
    onProtectDoc: noop,
    onCompare: noop,
    mergeSource: null,
    mergeRecordIndex: 0,
    onMergeRecordIndex: noop,
    mergeRange: MERGE_RANGE_ALL,
    onMergeRange: noop,
    filePath: null,
    viewMode: 'print' as const,
    onViewMode: noop,
    readMode: false,
    onReadMode: noop,
    showGrid: false,
    onShowGrid: noop,
    splitView: false,
    onSplitView: noop,
    onPagePreview: noop,
  }
}
