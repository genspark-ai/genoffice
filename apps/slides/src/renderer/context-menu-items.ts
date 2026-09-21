/**
 * Context-menu item builder extracted from App.tsx. Builds the
 * thumbnail/section/canvas/element menus from the latest App state (ActionCtx);
 * item callbacks dispatch into the extracted action modules.
 */
import { canEditPoints } from './edit-points'
import type { PictureRenderNode } from '@genoffice/pptx-render'
import type { ActionCtx } from './action-context'
import type { CtxItem } from './components/ContextMenu'
import { isEditableText } from './konva-adapter'
import { restoreEditSelection } from './TextEditOverlay'
import * as clipboardActions from './clipboard-actions'
import * as slideActions from './slide-actions'
import * as showActions from './show-actions'
import * as arrangeActions from './arrange-actions'
import * as insertActions from './insert-actions'
import * as pictureEditActions from './picture-edit-actions'
import * as styleActions from './style-actions'
import * as tableActions from './table-actions'
import { TABLE_SHADING_COLORS } from './components/table-shading-colors'
import { saveDefaultShapeStyle, shapeStyleOf } from './default-shape'
import { layoutLabel } from './layout-names'
import { t } from './i18n/locale'
import { groupSections, setAllCollapsed } from './section-groups'

export function buildCtxItems(ctx: ActionCtx): Array<CtxItem | null> {
  const { ctxMenu, slides, sections, selectedIds, slide, current } = ctx
  if (!ctxMenu) return []
  if (ctxMenu.kind === 'thumb') {
    // PowerPoint's slide-thumbnail menu order; a right-click inside the multi-selection acts on
    // all of it and paste lands after the last selected slide
    const i = ctxMenu.index
    const sel = ctx.selectedSlides.includes(i) ? ctx.selectedSlides : [i]
    const last = sel[sel.length - 1]!
    return [
      {
        label: t('appCtxCutSlide'),
        disabled: slides.length <= sel.length,
        onClick: () => void slideActions.cutSlides(ctx, sel),
      },
      { label: t('appCtxCopySlide'), onClick: () => void clipboardActions.copySlides(ctx, sel) },
      {
        label: t('appCtxPasteSlide'),
        disabled: !ctx.canPasteSlide,
        onClick: () => void clipboardActions.pasteSlideAfter(ctx, last),
      },
      {
        label: t('appCtxPasteSlideKeepSource'),
        disabled: !ctx.canPasteSlide,
        onClick: () => void clipboardActions.pasteSlideAfter(ctx, last, 'source'),
      },
      {
        label: t('appCtxPasteSlideAsPicture'),
        disabled: !ctx.canPasteSlide,
        onClick: () => void clipboardActions.pasteSlideAfter(ctx, last, 'picture'),
      },
      null,
      { label: t('appCtxNewSlide'), onClick: () => void slideActions.addSlideAt(ctx, last + 1) },
      {
        label: t('appCtxDuplicateSlide'),
        onClick: () => void slideActions.duplicateSlides(ctx, sel),
      },
      {
        label: t('appCtxDeleteSlide'),
        danger: true,
        disabled: slides.length <= sel.length,
        onClick: () => void slideActions.deleteSlides(ctx, sel),
      },
      null,
      { label: t('appCtxAddSectionBefore'), onClick: () => void slideActions.addSectionAt(ctx, i) },
      null,
      ...layoutItems(ctx, i),
      null,
      {
        label: t('appCtxChangeBgImage'),
        onClick: () => {
          ctx.setCurrent(i)
          void styleActions.onBackground(ctx, {
            kind: 'image',
            mode: 'stretch',
            pick: true,
            slideIndex: i,
          })
        },
      },
      {
        label: t('appCtxFormatBackground'),
        onClick: () => {
          ctx.setCurrent(i)
          ctx.openBgFormat()
        },
      },
      null,
      {
        label: slides[i]?.hidden ? t('appCtxUnhideSlide') : t('appCtxHideSlide'),
        onClick: () => void showActions.setSlidesHidden(ctx, sel, !slides[i]?.hidden),
      },
      null,
      {
        label: t('ribbonNewComment'),
        onClick: () => {
          ctx.setCurrent(i)
          ctx.newComment()
        },
      },
    ]
  }
  if (ctxMenu.kind === 'gap') {
    // PowerPoint's blank-area menu: new slide / paste land at the insertion point,
    // a section can only start on a slide (not below the last one)
    const pos = ctxMenu.pos
    return [
      { label: t('appCtxCutSlide'), disabled: true },
      { label: t('appCtxCopySlide'), disabled: true },
      {
        label: t('appCtxPasteSlide'),
        disabled: !ctx.canPasteSlide,
        onClick: () => void clipboardActions.pasteSlideAfter(ctx, pos - 1),
      },
      {
        label: t('appCtxPasteSlideKeepSource'),
        disabled: !ctx.canPasteSlide,
        onClick: () => void clipboardActions.pasteSlideAfter(ctx, pos - 1, 'source'),
      },
      {
        label: t('appCtxPasteSlideAsPicture'),
        disabled: !ctx.canPasteSlide,
        onClick: () => void clipboardActions.pasteSlideAfter(ctx, pos - 1, 'picture'),
      },
      null,
      { label: t('appCtxNewSlide'), onClick: () => void slideActions.addSlideAt(ctx, pos) },
      {
        label: t('ribbonAddSection'),
        disabled: pos >= slides.length,
        onClick: () => void slideActions.addSectionAt(ctx, pos),
      },
      null,
      { label: t('ribbonTabSlideShow'), onClick: () => showActions.startSlideShow(ctx, true) },
    ]
  }
  if (ctxMenu.kind === 'section') {
    // PowerPoint's section-header menu; the unsectioned lead group (sid null) gets it too,
    // minus Rename since the file never stores that header
    const sid = ctxMenu.sectionId
    const si = sections.findIndex((s) => s.id === sid)
    const group = groupSections(sections, slides.length)?.find((g) => g.id === sid)
    const ownSlides = group ? group.end - group.start : 0
    return [
      {
        label: t('appCtxRenameSection'),
        disabled: sid == null,
        onClick: () => {
          if (sid != null) ctx.setRenamingSec({ id: sid, value: group?.name ?? '' })
        },
      },
      {
        label: t('appCtxRemoveSection'),
        danger: true,
        onClick: () => void slideActions.removeSectionAt(ctx, sid),
      },
      {
        label: t('appCtxRemoveAllSections'),
        danger: true,
        onClick: () => void slideActions.removeAllSections(ctx),
      },
      {
        label: t('appCtxRemoveSectionSlides'),
        danger: true,
        disabled: !group || ownSlides >= slides.length,
        onClick: () => void slideActions.removeSectionWithSlides(ctx, sid),
      },
      null,
      {
        label: t('appCtxMoveSectionUp'),
        disabled: si <= 0,
        onClick: () => {
          if (sid != null) void slideActions.moveSectionDir(ctx, sid, 'up')
        },
      },
      {
        label: t('appCtxMoveSectionDown'),
        disabled: si < 0 || si >= sections.length - 1,
        onClick: () => {
          if (sid != null) void slideActions.moveSectionDir(ctx, sid, 'down')
        },
      },
      null,
      {
        label: t('appCtxCollapseAll'),
        onClick: () => ctx.setCollapsedSecs(setAllCollapsed(sections, true)),
      },
      {
        label: t('appCtxExpandAll'),
        onClick: () => ctx.setCollapsedSecs(setAllCollapsed(sections, false)),
      },
    ]
  }
  if (ctxMenu.kind === 'canvas') {
    // PowerPoint's blank-canvas menu order; cut/copy are placeholders (nothing is selected)
    return [
      { label: t('appCtxCut'), hint: '⌘X', disabled: true },
      { label: t('appCtxCopy'), hint: '⌘C', disabled: true },
      {
        label: t('appCtxPaste'),
        hint: '⌘V',
        disabled: !ctx.hasClipboard,
        onClick: () => void clipboardActions.pasteClipboard(ctx),
      },
      null,
      { label: t('ribbonRuler'), checked: ctx.showRuler, onClick: ctx.toggleRuler },
      { label: t('ribbonGridlines'), checked: ctx.showGrid, onClick: ctx.toggleGrid },
      { label: t('ribbonGuides'), checked: ctx.showGuides, onClick: ctx.toggleGuides },
      null,
      { label: t('appCtxNewSlide'), onClick: () => void slideActions.addSlide(ctx) },
      ...layoutItems(ctx, current),
      null,
      { label: t('appCtxFormatBackground'), onClick: () => ctx.openBgFormat() },
      { label: t('ribbonNewComment'), onClick: () => ctx.newComment() },
    ]
  }
  if (ctxMenu.kind === 'text') {
    // Every command runs against the editor selection saved when the menu opened, so the
    // edit session and its DOM range survive the menu taking the click
    const onSel = (fn: () => void) => () => {
      restoreEditSelection()
      fn()
    }
    const native = (op: 'cut' | 'copy' | 'paste') =>
      onSel(() => void window.slidesApi.nativeClipboard(op))
    const editedId = ctx.editing?.sourceId ?? ctx.editingCell?.sourceId
    return [
      { label: t('appCtxCut'), hint: '⌘X', disabled: ctxMenu.collapsed, onClick: native('cut') },
      { label: t('appCtxCopy'), hint: '⌘C', disabled: ctxMenu.collapsed, onClick: native('copy') },
      {
        label: t('appCtxPaste'),
        hint: '⌘V',
        disabled: !ctx.hasClipboard,
        onClick: native('paste'),
      },
      null,
      { label: t('ribbonBold'), hint: '⌘B', onClick: onSel(() => styleActions.onFormat('bold')) },
      {
        label: t('ribbonItalic'),
        hint: '⌘I',
        onClick: onSel(() => styleActions.onFormat('italic')),
      },
      {
        label: t('ribbonUnderline'),
        hint: '⌘U',
        onClick: onSel(() => styleActions.onFormat('underline')),
      },
      null,
      // Cell edits have no caret-paragraph bullet path (onParagraphFormat would restyle the whole table)
      ...(ctx.editingCell
        ? []
        : [
            {
              label: t('ribbonBullets'),
              onClick: onSel(() => styleActions.onParagraphFormat(ctx, { bullet: 'char' })),
            } as CtxItem,
            {
              label: t('ribbonNumbering'),
              onClick: onSel(() => styleActions.onParagraphFormat(ctx, { bullet: 'number' })),
            } as CtxItem,
            null,
          ]),
      { label: t('appCtxHyperlink'), onClick: onSel(() => void insertActions.openLinkDialog(ctx)) },
      {
        label: t('appCtxSelectAll'),
        hint: '⌘A',
        onClick: onSel(() => void document.execCommand('selectAll')),
      },
      null,
      {
        label: t('paneFormatTitleTyped', {
          type: nodeTypeName(editedId ? ctx.findNodeCtx(editedId)?.node : undefined),
        }),
        onClick: () => ctx.openFormat(),
      },
    ]
  }
  const node = slide?.nodes.find((n) => n.sourceId === ctxMenu.targetId)
  const single = selectedIds.length === 1
  const pictureIds = selectedIds.length ? selectedIds : [ctxMenu.targetId]
  const cell = ctxMenu.cell
  // Group: multi-select and all are text/shape/picture
  const canGroup =
    selectedIds.length >= 2 &&
    selectedIds.every((id) => {
      const n = slide?.nodes.find((nn) => nn.sourceId === id)
      return n && (n.type === 'text' || n.type === 'shape' || n.type === 'picture')
    })
  // Ungroup: single selection and it's a group
  const canUngroup = single && node?.type === 'group'
  const canRegroup =
    !!slide && !!arrangeActions.regroupCandidates(ctx.ungroupedSets, slide, selectedIds)
  const tableItems: Array<CtxItem | null> = cell
    ? [
        {
          label: t('appCtxInsertRowAbove'),
          onClick: () =>
            void tableActions.tableStructureOp(ctx, ctxMenu.targetId, 'insert-row', cell.row, true),
        },
        {
          label: t('appCtxInsertRowBelow'),
          onClick: () =>
            void tableActions.tableStructureOp(ctx, ctxMenu.targetId, 'insert-row', cell.row),
        },
        {
          label: t('appCtxInsertColLeft'),
          onClick: () =>
            void tableActions.tableStructureOp(ctx, ctxMenu.targetId, 'insert-col', cell.col, true),
        },
        {
          label: t('appCtxInsertColRight'),
          onClick: () =>
            void tableActions.tableStructureOp(ctx, ctxMenu.targetId, 'insert-col', cell.col),
        },
        {
          label: t('appCtxDeleteRow'),
          danger: true,
          onClick: () =>
            void tableActions.tableStructureOp(ctx, ctxMenu.targetId, 'delete-row', cell.row),
        },
        {
          label: t('appCtxDeleteCol'),
          danger: true,
          onClick: () =>
            void tableActions.tableStructureOp(ctx, ctxMenu.targetId, 'delete-col', cell.col),
        },
        null,
        {
          label: t('appCtxMergeRight'),
          onClick: () =>
            void tableActions.tableMergeOp(
              ctx,
              ctxMenu.targetId,
              'merge-right',
              cell.row,
              cell.col,
            ),
        },
        {
          label: t('appCtxMergeDown'),
          onClick: () =>
            void tableActions.tableMergeOp(ctx, ctxMenu.targetId, 'merge-down', cell.row, cell.col),
        },
        {
          label: t('appCtxCellAnchorTop'),
          onClick: () =>
            void window.slidesApi
              .setTableCellAnchor({
                slideIndex: current,
                sourceId: ctxMenu.targetId,
                row: cell.row,
                col: cell.col,
                anchor: 'top',
              })
              .then((r) => r && ctx.applySlide(current, r)),
        },
        {
          label: t('appCtxCellAnchorMiddle'),
          onClick: () =>
            void window.slidesApi
              .setTableCellAnchor({
                slideIndex: current,
                sourceId: ctxMenu.targetId,
                row: cell.row,
                col: cell.col,
                anchor: 'middle',
              })
              .then((r) => r && ctx.applySlide(current, r)),
        },
        {
          label: t('appCtxCellAnchorBottom'),
          onClick: () =>
            void window.slidesApi
              .setTableCellAnchor({
                slideIndex: current,
                sourceId: ctxMenu.targetId,
                row: cell.row,
                col: cell.col,
                anchor: 'bottom',
              })
              .then((r) => r && ctx.applySlide(current, r)),
        },
        {
          label: t('appCtxSplitCell'),
          disabled: (() => {
            const tbl = slide?.nodes.find((n) => n.sourceId === ctxMenu.targetId)
            const cr =
              tbl?.type === 'table'
                ? tbl.cells.find((c) => c.row === cell.row && c.col === cell.col)
                : undefined
            return !((cr?.gridSpan ?? 1) > 1 || (cr?.rowSpan ?? 1) > 1)
          })(),
          onClick: () =>
            void tableActions.tableMergeOp(ctx, ctxMenu.targetId, 'split', cell.row, cell.col),
        },
        null,
        {
          label: t('appCtxCellShading'),
          swatches: [...TABLE_SHADING_COLORS, 'none'],
          onSwatch: (c) =>
            void window.slidesApi
              .editTableStyle({
                slideIndex: current,
                sourceId: ctxMenu.targetId,
                shadingColor: c,
                cells: [{ row: cell.row, col: cell.col }],
              })
              .then((r) => {
                if (!r) return
                ctx.applySlide(current, r.slide)
                if (r.sourceId) ctx.setSelectedIds([r.sourceId])
              }),
        },
        null,
      ]
    : []
  const isLine = !!(node as { line?: unknown } | undefined)?.line
  const defaultStyle = single ? shapeStyleOf(node) : undefined
  const zOrder = (dir: 'front' | 'back' | 'forward' | 'backward') => ({
    label: t(
      dir === 'front'
        ? 'appCtxBringToFront'
        : dir === 'back'
          ? 'appCtxSendToBack'
          : dir === 'forward'
            ? 'appCtxBringForward'
            : 'appCtxSendBackward',
    ),
    onClick: () => void arrangeActions.reorderSelected(ctx, ctxMenu.targetId, dir),
  })
  // PowerPoint's shape/picture menu order; rotate/flip and align/distribute (ribbon-only
  // in PowerPoint) follow the Format entry, Delete stays last
  return [
    ...tableItems,
    { label: t('appCtxCut'), hint: '⌘X', onClick: () => void clipboardActions.cutSelected(ctx) },
    { label: t('appCtxCopy'), hint: '⌘C', onClick: () => void clipboardActions.copySelected(ctx) },
    {
      label: t('appCtxPaste'),
      hint: '⌘V',
      disabled: !ctx.hasClipboard,
      onClick: () => void clipboardActions.pasteClipboard(ctx),
    },
    null,
    ...(node && isEditableText(node)
      ? [{ label: t('appCtxEditText'), onClick: () => ctx.startEdit(ctxMenu.targetId) } as CtxItem]
      : []),
    // Connectors are endpoint-based (p:cxnSp): swapping their prstGeom would
    // orphan the connection metadata, so they keep their existing menu.
    ...(single && node?.type === 'shape' && !isLine
      ? [
          {
            label: t('appCtxChangeShape'),
            onClick: () => ctx.openChangeShape(ctxMenu.targetId, ctxMenu.x, ctxMenu.y),
          } as CtxItem,
        ]
      : []),
    ...(single && node && (node.type === 'shape' || node.type === 'text') && !isLine
      ? [
          {
            label: t('appCtxEditPoints'),
            disabled: !canEditPoints(node),
            onClick: () => ctx.setEditPointsTarget({ sourceId: ctxMenu.targetId, vertex: null }),
          } as CtxItem,
        ]
      : []),
    ...(node?.type !== 'table'
      ? [
          {
            label: t('appCtxGroup'),
            disabled: !canGroup && !canUngroup && !canRegroup,
            sub: [
              {
                label: t('appCtxGroup'),
                hint: '⌘G',
                disabled: !canGroup,
                onClick: () => void arrangeActions.groupSelected(ctx),
              },
              {
                label: t('appCtxUngroup'),
                hint: '⌘⇧G',
                disabled: !canUngroup,
                onClick: () => void arrangeActions.ungroupSelected(ctx),
              },
              {
                label: t('appCtxRegroup'),
                disabled: !canRegroup,
                onClick: () => void arrangeActions.regroupSelected(ctx),
              },
            ],
          } as CtxItem,
        ]
      : []),
    {
      label: t('appCtxBringToFront'),
      disabled: !single,
      sub: [zOrder('front'), zOrder('forward')],
    },
    {
      label: t('appCtxSendToBack'),
      disabled: !single,
      sub: [zOrder('back'), zOrder('backward')],
    },
    ...(node?.type !== 'table'
      ? [
          {
            label: t('appCtxSaveAsPicture'),
            disabled: !pictureEditActions.canSaveAsPicture(ctx, pictureIds),
            onClick: () => void pictureEditActions.saveSelectionAsPicture(ctx, pictureIds),
          } as CtxItem,
        ]
      : []),
    ...(single
      ? [
          {
            label: t('appCtxHyperlink'),
            onClick: () => void insertActions.openLinkDialog(ctx),
          } as CtxItem,
        ]
      : []),
    ...(node && node.type === 'picture'
      ? [
          {
            label: t('ribbonReplacePicture'),
            onClick: () => void pictureEditActions.replacePicture(ctx),
          } as CtxItem,
          {
            label: t('appCtxCropPicture'),
            onClick: () => pictureEditActions.startCrop(ctx),
          } as CtxItem,
          ...(!(node as PictureRenderNode).media
            ? [
                {
                  label: t('appCtxRemoveBackground'),
                  onClick: () => pictureEditActions.startCutout(ctx),
                } as CtxItem,
              ]
            : []),
        ]
      : []),
    ...(defaultStyle
      ? [
          {
            label: t('appCtxSetDefaultShape'),
            onClick: () => saveDefaultShapeStyle(defaultStyle),
          } as CtxItem,
        ]
      : []),
    ...(single
      ? [{ label: t('appCtxSizePosition'), onClick: () => ctx.openFormat('size') } as CtxItem]
      : []),
    {
      // findNodeCtx also resolves children of the group being edited (node is top-level only)
      label: t('paneFormatTitleTyped', {
        type: nodeTypeName(node ?? ctx.findNodeCtx(ctxMenu.targetId)?.node),
      }),
      onClick: () => ctx.openFormat(),
    },
    null,
    // Rotate/flip for shapes, pictures and groups; connectors are endpoint-based
    ...(node &&
    (node.type === 'shape' || node.type === 'picture' || node.type === 'group') &&
    !isLine
      ? [
          {
            label: t('appCtxRotateLeft90'),
            onClick: () => void arrangeActions.rotateSelected(ctx, -90),
          } as CtxItem,
          {
            label: t('appCtxRotateRight90'),
            onClick: () => void arrangeActions.rotateSelected(ctx, 90),
          } as CtxItem,
          null,
          {
            label: t('appCtxFlipH'),
            onClick: () => void arrangeActions.flipSelected(ctx, 'h'),
          } as CtxItem,
          {
            label: t('appCtxFlipV'),
            onClick: () => void arrangeActions.flipSelected(ctx, 'v'),
          } as CtxItem,
          null,
        ]
      : []),
    ...(selectedIds.length >= 1
      ? [
          {
            label: t('appCtxAlignLeft'),
            onClick: () => void arrangeActions.alignSelected(ctx, 'left'),
          } as CtxItem,
          {
            label: t('appCtxAlignCenterH'),
            onClick: () => void arrangeActions.alignSelected(ctx, 'center-h'),
          } as CtxItem,
          {
            label: t('appCtxAlignRight'),
            onClick: () => void arrangeActions.alignSelected(ctx, 'right'),
          } as CtxItem,
          {
            label: t('appCtxAlignTop'),
            onClick: () => void arrangeActions.alignSelected(ctx, 'top'),
          } as CtxItem,
          {
            label: t('appCtxAlignCenterV'),
            onClick: () => void arrangeActions.alignSelected(ctx, 'center-v'),
          } as CtxItem,
          {
            label: t('appCtxAlignBottom'),
            onClick: () => void arrangeActions.alignSelected(ctx, 'bottom'),
          } as CtxItem,
        ]
      : []),
    ...(selectedIds.length >= 3
      ? [
          {
            label: t('appCtxDistributeH'),
            onClick: () => void arrangeActions.alignSelected(ctx, 'distribute-h'),
          } as CtxItem,
          {
            label: t('appCtxDistributeV'),
            onClick: () => void arrangeActions.alignSelected(ctx, 'distribute-v'),
          } as CtxItem,
        ]
      : []),
    null,
    { label: t('ribbonNewComment'), onClick: () => ctx.newComment() },
    null,
    {
      label: t('appCtxDelete'),
      hint: '⌫',
      danger: true,
      onClick: () => void clipboardActions.deleteSelected(ctx),
    },
  ]
}

/** Layout ▸ (the deck's slideLayouts) + Reset Slide, applied to slide `index` */
function layoutItems(ctx: ActionCtx, index: number): Array<CtxItem | null> {
  const layouts = ctx.layouts ?? []
  return [
    {
      label: t('ribbonLayout'),
      disabled: layouts.length === 0,
      sub: layouts.map((lay) => ({
        label: layoutLabel(lay.name, t),
        onClick: () => void slideActions.setSlideLayoutAt(ctx, index, lay.path),
      })),
    },
    {
      label: t('appCtxResetSlide'),
      onClick: () => void slideActions.setSlideLayoutAt(ctx, index),
    },
  ]
}

/** Format-pane type name (same mapping as FormatPane's title) for the context-menu label */
function nodeTypeName(node: { type: string } | undefined): string {
  switch (node?.type) {
    case 'picture':
      return t('paneFormatPicture')
    case 'group':
      return t('paneFormatGroup')
    case 'text':
      return t('paneFormatTextBox')
    case 'table':
      return t('ribbonGroupTable')
    case 'chart':
      return t('ribbonChart')
    default:
      return t('paneFormatShape')
  }
}
