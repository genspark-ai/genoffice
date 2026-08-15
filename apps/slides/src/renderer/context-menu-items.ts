/**
 * Context-menu item builder extracted from App.tsx. Builds the
 * thumbnail/section/canvas/element menus from the latest App state (ActionCtx);
 * item callbacks dispatch into the extracted action modules.
 */
import type { PictureRenderNode } from '@genoffice/pptx-render'
import type { ActionCtx } from './action-context'
import type { CtxItem } from './components/ContextMenu'
import { isEditableText } from './konva-adapter'
import * as clipboardActions from './clipboard-actions'
import * as slideActions from './slide-actions'
import * as showActions from './show-actions'
import * as arrangeActions from './arrange-actions'
import * as insertActions from './insert-actions'
import * as pictureEditActions from './picture-edit-actions'
import * as tableActions from './table-actions'
import { TABLE_SHADING_COLORS } from './components/table-shading-colors'
import { t } from './i18n/locale'

export function buildCtxItems(ctx: ActionCtx): Array<CtxItem | null> {
  const { ctxMenu, slides, sections, selectedIds, slide, current } = ctx
  if (!ctxMenu) return []
  if (ctxMenu.kind === 'thumb') {
    const i = ctxMenu.index
    return [
      { label: t('appCtxNewSlide'), onClick: () => void slideActions.addSlide(ctx) },
      {
        label: t('appCtxDuplicateSlide'),
        onClick: () => void slideActions.duplicateSlideAt(ctx, i),
      },
      {
        label: slides[i]?.hidden ? t('appCtxUnhideSlide') : t('appCtxHideSlide'),
        onClick: () => void showActions.toggleHidden(ctx, i),
      },
      null,
      {
        label: t('appCtxCutSlide'),
        disabled: slides.length <= 1,
        onClick: () => void slideActions.cutSlideAt(ctx, i),
      },
      { label: t('appCtxCopySlide'), onClick: () => void clipboardActions.copySlideAt(ctx, i) },
      {
        label: t('appCtxPasteSlide'),
        disabled: !ctx.canPasteSlide,
        onClick: () => void clipboardActions.pasteSlideAfter(ctx, i),
      },
      {
        label: t('appCtxPasteSlideKeepSource'),
        disabled: !ctx.canPasteSlide,
        onClick: () => void clipboardActions.pasteSlideAfter(ctx, i, 'source'),
      },
      {
        label: t('appCtxPasteSlideAsPicture'),
        disabled: !ctx.canPasteSlide,
        onClick: () => void clipboardActions.pasteSlideAfter(ctx, i, 'picture'),
      },
      null,
      { label: t('appCtxAddSectionBefore'), onClick: () => void slideActions.addSectionAt(ctx, i) },
      null,
      {
        label: t('appCtxDeleteSlide'),
        danger: true,
        disabled: slides.length <= 1,
        onClick: () => void slideActions.deleteSlideAt(ctx, i),
      },
    ]
  }
  if (ctxMenu.kind === 'section') {
    const sid = ctxMenu.sectionId
    const si = sections.findIndex((s) => s.id === sid)
    const sec = sections[si]
    return [
      {
        label: t('appCtxRenameSection'),
        onClick: () => ctx.setRenamingSec({ id: sid, value: sec?.name ?? '' }),
      },
      {
        label: t('appCtxRemoveSection'),
        danger: true,
        onClick: () => void slideActions.removeSectionAt(ctx, sid),
      },
      null,
      {
        label: t('appCtxMoveSectionUp'),
        disabled: si <= 0,
        onClick: () => void slideActions.moveSectionDir(ctx, sid, 'up'),
      },
      {
        label: t('appCtxMoveSectionDown'),
        disabled: si < 0 || si >= sections.length - 1,
        onClick: () => void slideActions.moveSectionDir(ctx, sid, 'down'),
      },
    ]
  }
  if (ctxMenu.kind === 'canvas') {
    return [
      {
        label: t('appCtxPaste'),
        hint: '⌘V',
        disabled: !ctx.hasClipboard,
        onClick: () => void clipboardActions.pasteClipboard(ctx),
      },
      null,
      { label: t('appCtxNewSlide'), onClick: () => void slideActions.addSlide(ctx) },
    ]
  }
  const node = slide?.nodes.find((n) => n.sourceId === ctxMenu.targetId)
  const single = selectedIds.length === 1
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
    {
      label: t('appCtxBringToFront'),
      disabled: !single,
      onClick: () => void arrangeActions.reorderSelected(ctx, ctxMenu.targetId, 'front'),
    },
    {
      label: t('appCtxSendToBack'),
      disabled: !single,
      onClick: () => void arrangeActions.reorderSelected(ctx, ctxMenu.targetId, 'back'),
    },
    {
      label: t('appCtxBringForward'),
      disabled: !single,
      onClick: () => void arrangeActions.reorderSelected(ctx, ctxMenu.targetId, 'forward'),
    },
    {
      label: t('appCtxSendBackward'),
      disabled: !single,
      onClick: () => void arrangeActions.reorderSelected(ctx, ctxMenu.targetId, 'backward'),
    },
    null,
    // Rotate/flip group (PowerPoint parity) for shapes, pictures and groups; connectors
    // are endpoint-based and keep their existing menu
    ...(node &&
    (node.type === 'shape' || node.type === 'picture' || node.type === 'group') &&
    !(node as { line?: unknown }).line
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
    ...(node && isEditableText(node)
      ? [{ label: t('appCtxEditText'), onClick: () => ctx.startEdit(ctxMenu.targetId) } as CtxItem]
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
    ...(canGroup
      ? [
          {
            label: t('appCtxGroup'),
            hint: '⌘G',
            onClick: () => void arrangeActions.groupSelected(ctx),
          } as CtxItem,
        ]
      : []),
    ...(canUngroup
      ? [
          {
            label: t('appCtxUngroup'),
            hint: '⌘⇧G',
            onClick: () => void arrangeActions.ungroupSelected(ctx),
          } as CtxItem,
        ]
      : []),
    ...(selectedIds.length >= 1
      ? [
          null,
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
    {
      label: t('appCtxDelete'),
      hint: '⌫',
      danger: true,
      onClick: () => void clipboardActions.deleteSelected(ctx),
    },
  ]
}
