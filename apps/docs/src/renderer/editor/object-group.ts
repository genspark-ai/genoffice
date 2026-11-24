import type { Editor } from '@tiptap/core'
import { Fragment, type Node as PmNode } from '@tiptap/pm/model'
import type { EditorState } from '@tiptap/pm/state'
import type { EditorView } from '@tiptap/pm/view'
import {
  maxDrawingId,
  splitGroupParagraph,
  type Block,
  type GroupCorner,
  type TextboxDisplay,
  type UngroupedMember,
} from '@genoffice/docx-engine'
import { setObjectDragFollower, type ObjectDragFollow } from './extensions'
import {
  boundsOf,
  moveObjectsTo,
  objectBox,
  pageZoom,
  pinObjectOffsets,
  type Box,
} from './object-align'
import {
  groupMates,
  holdObjectSelection,
  isFileGroup,
  isFloatingObject,
  newGroupId,
  releaseHeldObjects,
  selectedObjectPositions,
  setCarriedSettler,
  setObjectSelection,
  type GroupBox,
  type GroupMemberSource,
} from './object-multi-select'
import { textboxBandBottom } from './obstacle-plan'
import { isStraightLineKind } from './shape-svg'

const EMU_PER_PX = 9525

/** a group as the file holds it: one drawing whose boxes are its children */
function isFileGroupNode(node: PmNode | null): boolean {
  const boxes = node?.attrs.textboxes as TextboxDisplay[] | null | undefined
  return node?.type.name === 'docProtected' && Array.isArray(boxes) && isFileGroup(boxes)
}

/**
 * What the save can merge into a group: a drawing that is one shape, text box
 * or picture, in a paragraph of the body itself.
 */
function canJoinGroup(state: EditorState, pos: number): boolean {
  const node = state.doc.nodeAt(pos)
  if (!node || isFileGroupNode(node)) return false
  const depth = state.doc.resolve(pos).depth
  return node.type.name === 'docInlineImage' ? depth === 1 : depth === 0
}

/** Group: several objects selected that are not one group already. */
export function canGroupObjects(state: EditorState): boolean {
  const selected = selectedObjectPositions(state)
  if (selected.length < 2 || !selected.every((pos) => canJoinGroup(state, pos))) return false
  const ids = new Set(selected.map((pos) => state.doc.nodeAt(pos)?.attrs.groupId ?? null))
  return ids.size > 1 || ids.has(null)
}

/** each member's box inside the bounds of them all, document px; null when one is not laid out */
function measuredGroupBoxes(view: EditorView, positions: number[]): GroupBox[] | null {
  const zoom = pageZoom()
  const boxes: Box[] = []
  for (const pos of positions) {
    const box = objectBox(view, pos)
    if (!box) return null
    boxes.push(box)
  }
  const bounds = boundsOf(boxes)
  const px = (v: number) => Math.round((v / zoom) * 100) / 100
  return boxes.map((b) => ({
    x: px(b.left - bounds.left),
    y: px(b.top - bounds.top),
    w: px(b.width),
    h: px(b.height),
  }))
}

/** Makes the selected objects one group. False when they cannot be. */
export function groupSelectedObjects(editor: Editor): boolean {
  const { view } = editor
  if (!canGroupObjects(view.state)) return false
  const positions = [...selectedObjectPositions(view.state)].sort((a, b) => a - b)
  // the save places the group from its first member's offsets
  const pinned = pinObjectOffsets(view, positions)
  const boxes = measuredGroupBoxes(view, positions)
  if (!boxes) return false
  const groupId = newGroupId()
  const tr = view.state.tr
  // one undo takes the group and the pinning back together
  if (pinned) tr.setMeta('appendedTransaction', pinned)
  positions.forEach((pos, i) => {
    const node = tr.doc.nodeAt(pos)!
    tr.setNodeMarkup(pos, undefined, { ...node.attrs, groupId, groupBox: boxes[i] })
  })
  view.dispatch(setObjectSelection(tr, selectedObjectPositions(view.state)))
  view.focus()
  return true
}

/** the file group's paragraph: its block of the file, or what an earlier Ungroup left */
function groupSource(node: PmNode, blocks: readonly Block[]): GroupMemberSource | null {
  const member = node.attrs.member as GroupMemberSource | null
  if (member) return member
  const block = blocks.find((b) => b.docxIndex != null && b.docxIndex === node.attrs.docxIndex)
  if (!block?.originalXml || !block.textboxes) return null
  return {
    xml: block.originalXml,
    boxes: block.textboxes,
    xEmu: block.imageOffsetXEmu ?? 0,
    yEmu: block.imageOffsetYEmu ?? 0,
  }
}

/** a box of the group as a box of the child it belongs to, one level of grouping gone */
function memberBox(box: TextboxDisplay, member: UngroupedMember): TextboxDisplay {
  const { grouped: _grouped, groupPath, ...rest } = box
  const path = (groupPath ?? []).slice(1)
  return {
    ...rest,
    ...(member.kind === 'group' && path.length > 0 ? { grouped: true, groupPath: path } : {}),
    // text boxes are counted within the paragraph, now the child's own
    ...(box.txbxIndex !== undefined ? { txbxIndex: box.txbxIndex - member.textBoxesBefore } : {}),
  }
}

/**
 * One node per child of a file group. The first is the group's paragraph; the
 * others are carried behind it, so the text keeps its lines. Null when the
 * group cannot be taken apart or a child has nothing to show.
 */
function fileGroupMembers(
  node: PmNode,
  source: GroupMemberSource,
  nextId: number,
  corner?: GroupCorner,
): PmNode[] | null {
  const split = splitGroupParagraph(source.xml, 0, nextId, corner)
  const shown = node.attrs.textboxes as TextboxDisplay[]
  if (!split || shown.length !== source.boxes.length) return null
  // text flowing above and below one band cannot be kept in place around separate drawings
  if (shown.some((box) => textboxBandBottom(box) > 0)) return null
  // how far the group was moved since it was read; a corner taken from the page is where it is now
  const movedX = split.fromCorner.h
    ? 0
    : Number(node.attrs.imageOffsetXEmu ?? source.xEmu) - source.xEmu
  const movedY = split.fromCorner.v
    ? 0
    : Number(node.attrs.imageOffsetYEmu ?? source.yEmu) - source.yEmu
  const out: PmNode[] = []
  for (const [index, member] of split.members.entries()) {
    const own = source.boxes.flatMap((box, i) => (box.groupPath?.[0] === index ? [i] : []))
    if (own.length === 0) return null
    const first = index === 0
    const origin: GroupMemberSource = {
      xml: split.paragraphs[index],
      boxes: own.map((i) => memberBox(source.boxes[i], member)),
      xEmu: member.xEmu,
      yEmu: member.yEmu,
    }
    const attrs = {
      // the paragraph itself (file index, text, line) stays with the first child
      ...(first
        ? node.attrs
        : {
            docxIndex: null,
            blockType: node.attrs.blockType,
            label: node.attrs.label,
            imageWrap: node.attrs.imageWrap,
            imageRelH: node.attrs.imageRelH,
            imageRelV: node.attrs.imageRelV,
            imageZOrder: node.attrs.imageZOrder,
            imageAnchorLocked: node.attrs.imageAnchorLocked,
            carriedDy: node.attrs.carriedDy,
          }),
      member: origin,
      carried: !first || !!node.attrs.carried,
      groupId: null,
      groupBox: null,
      textboxes: own.map((i) => memberBox(shown[i], member)),
      imageOffsetXEmu: member.xEmu + movedX,
      imageOffsetYEmu: member.yEmu + movedY,
      imagePosH: null,
      imagePosV: null,
    }
    out.push(node.type.create(attrs))
  }
  return out
}

/** stands for the corner while only asking whether a group placed by alignment can be taken apart */
const ANY_CORNER: GroupCorner = { xEmu: 0, yEmu: 0 }

/**
 * The children whose drawn box says where the group's corner is, with the
 * index of that box: one box of its own, and not a line, whose box is a band
 * around its stroke.
 */
function cornerMembers(
  members: readonly UngroupedMember[],
  boxes: readonly TextboxDisplay[],
): Array<{ member: UngroupedMember; at: number }> {
  return members.flatMap((member, index) => {
    const own = boxes.flatMap((box, i) => (box.groupPath?.[0] === index ? [i] : []))
    return member.kind === 'group' || own.length !== 1 || isStraightLineKind(boxes[own[0]].prst)
      ? []
      : [{ member, at: own[0] }]
  })
}

/**
 * Where the page drew the corner of the file group at `pos`, from its
 * paragraph's box: what a group placed by alignment has in place of offsets.
 * Read off one child's drawn box; a turned box is known by its centre only.
 */
function drawnGroupCorner(
  view: EditorView,
  pos: number,
  source: GroupMemberSource,
): GroupCorner | null {
  const probe = splitGroupParagraph(source.xml, 0, 1, ANY_CORNER)
  const dom = view.nodeDOM(pos)
  if (!probe || !(dom instanceof HTMLElement)) return null
  const drawn = Array.from(dom.querySelectorAll<HTMLElement>('.doc-textbox')).filter(
    (el) => !el.parentElement?.closest('.doc-textbox'),
  )
  if (drawn.length !== source.boxes.length) return null
  const frame = dom.getBoundingClientRect()
  const zoom = pageZoom()
  let turned: GroupCorner | null = null
  for (const { member, at } of cornerMembers(probe.members, source.boxes)) {
    const rect = drawn[at].getBoundingClientRect()
    if (!source.boxes[at].rotDeg) {
      return {
        xEmu: Math.round(((rect.left - frame.left) / zoom) * EMU_PER_PX - member.box.x),
        yEmu: Math.round(((rect.top - frame.top) / zoom) * EMU_PER_PX - member.box.y),
      }
    }
    const midX = ((rect.left + rect.right) / 2 - frame.left) / zoom
    const midY = ((rect.top + rect.bottom) / 2 - frame.top) / zoom
    turned ??= {
      xEmu: Math.round(midX * EMU_PER_PX - member.box.x - member.box.cx / 2),
      yEmu: Math.round(midY * EMU_PER_PX - member.box.y - member.box.cy / 2),
    }
  }
  return turned
}

/** the children of a group placed by alignment, as far as can be told before its corner is measured */
function alignedGroupMembers(node: PmNode, source: GroupMemberSource): PmNode[] | null {
  const probe = splitGroupParagraph(source.xml, 0, 1, ANY_CORNER)
  // with no child to read the corner off, Ungroup would be offered and do nothing
  if (!probe || cornerMembers(probe.members, source.boxes).length === 0) return null
  return fileGroupMembers(node, source, 1, ANY_CORNER)
}

let fileBlocks: readonly Block[] = []

/** the open file as read: where a file group's XML is found when no caller hands it over */
export function setGroupFileBlocks(blocks: readonly Block[]): void {
  fileBlocks = blocks
}

const splittable = new WeakMap<readonly Block[], WeakMap<PmNode, boolean>>()

/** a file group Ungroup can take apart: not one in line, or with text banded around it */
function canSplitFileGroup(node: PmNode, blocks: readonly Block[]): boolean {
  let known = splittable.get(blocks)
  if (!known) splittable.set(blocks, (known = new WeakMap()))
  let can = known.get(node)
  if (can === undefined) {
    const source = groupSource(node, blocks)
    can =
      !!source && (fileGroupMembers(node, source, 1) ?? alignedGroupMembers(node, source)) !== null
    known.set(node, can)
  }
  return can
}

/** Ungroup: a group among the selected objects that can be taken apart. */
export function canUngroupObjects(
  state: EditorState,
  blocks: readonly Block[] = fileBlocks,
): boolean {
  return selectedObjectPositions(state).some((pos) => {
    const node = state.doc.nodeAt(pos)
    return !!node?.attrs.groupId || (isFileGroupNode(node) && canSplitFileGroup(node!, blocks))
  })
}

function highestDrawingId(doc: PmNode, blocks: readonly Block[]): number {
  let max = 0
  for (const block of blocks)
    if (block.originalXml) max = Math.max(max, maxDrawingId(block.originalXml))
  doc.descendants((node) => {
    const xml = (node.attrs.member as GroupMemberSource | null)?.xml ?? node.attrs.genXml
    if (typeof xml === 'string') max = Math.max(max, maxDrawingId(xml))
  })
  return max
}

/** Takes the selected groups apart; a file group needs `blocks`, the file as read, for its XML. */
export function ungroupSelectedObjects(
  editor: Editor,
  blocks: readonly Block[] = fileBlocks,
): boolean {
  const { view } = editor
  const { state } = view
  if (!canUngroupObjects(state, blocks)) return false
  const selected = selectedObjectPositions(state)
  const tr = state.tr
  let nextId = highestDrawingId(state.doc, blocks) + 1
  const split = new Map<number, number>()
  const all = [...new Set([...selected, ...groupMates(state.doc, selected)])].sort((a, b) => b - a)
  // back to front: a split shifts only what comes after it
  for (const pos of all) {
    const node = state.doc.nodeAt(pos)!
    if (node.attrs.groupId) {
      tr.setNodeMarkup(pos, undefined, { ...node.attrs, groupId: null, groupBox: null })
      continue
    }
    if (!isFileGroupNode(node)) continue
    const source = groupSource(node, blocks)
    const corner = source && drawnGroupCorner(view, pos, source)
    const members =
      source &&
      (fileGroupMembers(node, source, nextId) ??
        (corner ? fileGroupMembers(node, source, nextId, corner) : null))
    if (!members) continue
    nextId += members.length
    split.set(pos, members.length)
    tr.replaceWith(pos, pos + node.nodeSize, Fragment.fromArray(members))
  }
  if (!tr.docChanged) return false
  const now: number[] = []
  for (const pos of [...all].reverse()) {
    let at = tr.mapping.map(pos, -1)
    for (let i = 0; i < (split.get(pos) ?? 1); i++) {
      const node = tr.doc.nodeAt(at)
      if (!node) break
      if (isFloatingObject(node)) now.push(at)
      at += node.nodeSize
    }
  }
  view.dispatch(now.length > 0 ? setObjectSelection(tr, now) : tr)
  settleCarried(view)
  view.focus()
  return true
}

/** how far below its paragraph a carried node stands, document px */
function carriedDepth(view: EditorView, pos: number, index: number): number | null {
  const { doc } = view.state
  let host = pos
  let at = index
  while (at > 0 && doc.child(at).attrs.carried) host -= doc.child(--at).nodeSize
  const own = view.nodeDOM(pos)
  const before = host < pos ? view.nodeDOM(host) : null
  if (!(own instanceof HTMLElement) || !(before instanceof HTMLElement)) return null
  const dy = (own.getBoundingClientRect().top - before.getBoundingClientRect().top) / pageZoom()
  return Math.round(dy * 100) / 100
}

/**
 * A carried child stands in the flow below the paragraph its offsets measure
 * from: its boxes are drawn that much higher, so it shows where the file has
 * it. Not a user edit: out of the undo history.
 */
export function settleCarried(view: EditorView): void {
  const tr = view.state.tr
  view.state.doc.forEach((node, pos, index) => {
    if (!node.attrs.carried) return
    const depth = carriedDepth(view, pos, index)
    const by = depth === null ? 0 : depth - Number(node.attrs.carriedDy ?? 0)
    if (Math.abs(by) < 0.01) return
    tr.setNodeMarkup(pos, undefined, {
      ...node.attrs,
      carriedDy: depth,
      // a box the canvas pins to its page is not drawn from the flow
      textboxes: (node.attrs.textboxes as TextboxDisplay[]).map((box) =>
        box.pageRelV || box.pagePinned
          ? box
          : { ...box, offsetYEmu: (box.offsetYEmu ?? 0) - Math.round(by * EMU_PER_PX) },
      ),
    })
  })
  if (!tr.docChanged) return
  const selected = selectedObjectPositions(view.state)
  view.dispatch(
    (selected.length > 0 ? setObjectSelection(tr, selected) : tr).setMeta('addToHistory', false),
  )
}

/**
 * Measures every group again, so the save writes the members where they show:
 * members anchored in different paragraphs drift apart when the text between
 * them changes. Not an edit the user made: it stays out of the undo history.
 */
export function refreshGroupBoxes(editor: Editor): void {
  const { view } = editor
  const groups = new Map<string, number[]>()
  view.state.doc.descendants((node, pos) => {
    const id = node.attrs.groupId as string | null | undefined
    if (!id) return
    const list = groups.get(id)
    if (list) list.push(pos)
    else groups.set(id, [pos])
  })
  settleCarried(view)
  if (groups.size === 0) return
  const tr = view.state.tr
  for (const positions of groups.values()) {
    const boxes = measuredGroupBoxes(view, positions)
    if (!boxes) continue
    positions.forEach((pos, i) => {
      const node = view.state.doc.nodeAt(pos)!
      const old = node.attrs.groupBox as GroupBox | null
      const box = boxes[i]
      if (old && old.x === box.x && old.y === box.y && old.w === box.w && old.h === box.h) return
      tr.setNodeMarkup(pos, undefined, { ...node.attrs, groupBox: box })
    })
  }
  if (!tr.docChanged) return
  const selected = selectedObjectPositions(view.state)
  view.dispatch(
    (selected.length > 0 ? setObjectSelection(tr, selected) : tr).setMeta('addToHistory', false),
  )
}

/** the elements that show the object at `pos`: what a drag slides along with the pointer */
function dragVisuals(view: EditorView, pos: number): HTMLElement[] {
  const node = view.state.doc.nodeAt(pos)
  const dom = view.nodeDOM(pos)
  if (!node || !(dom instanceof HTMLElement)) return []
  if (node.type.name === 'docInlineImage') {
    const img = dom instanceof HTMLImageElement ? dom : dom.querySelector('img')
    return img ? [img] : []
  }
  if (node.attrs.blockType === 'image') {
    const img = dom.querySelector<HTMLElement>('.doc-img-crop, .doc-protected-img')
    return img ? [img] : []
  }
  // a turned box sits inside the rotator that owns its slot
  const boxes = Array.from(dom.querySelectorAll<HTMLElement>('.doc-textbox'))
  return boxes.map((box) => box.closest<HTMLElement>('.box-rotator') ?? box)
}

/**
 * A drag moves every selected object, not only the one under the pointer:
 * the others slide with the pointer while it moves, and once the drop is in
 * the document they are moved by the same distance and all selected again.
 * Call when the press lands.
 */
export function beginObjectDrag(view: EditorView, pos: number): ObjectDragFollow {
  const before = objectBox(view, pos)
  const others = selectedObjectPositions(view.state).filter((p) => p !== pos)
  const starts = new Map(others.map((p) => [p, objectBox(view, p)]))
  // the drag handler slides the pressed object itself
  const riding = others
    .flatMap((p) => dragVisuals(view, p))
    .map((el) => ({ el, transform: el.style.transform }))
  holdObjectSelection(view)
  return {
    move(dx, dy) {
      for (const { el, transform } of riding)
        el.style.transform = `translate(${dx}px, ${dy}px)${transform ? ` ${transform}` : ''}`
    },
    drop(droppedAt) {
      for (const { el, transform } of riding) el.style.transform = transform
      const held = releaseHeldObjects(view)
      const after = objectBox(view, droppedAt)
      if (!before || !after || held.length !== others.length) return
      const dx = after.left - before.left
      const dy = after.top - before.top
      const targets = new Map<number, Box>()
      held.forEach((p, i) => {
        const start = starts.get(others[i])
        if (start) targets.set(p, { ...start, left: start.left + dx, top: start.top + dy })
      })
      view.dispatch(setObjectSelection(view.state.tr, [droppedAt, ...held]))
      // part of the drag: one undo takes the whole selection back
      if (Math.abs(dx) >= 0.5 || Math.abs(dy) >= 0.5) moveObjectsTo(view, targets, view.state.tr)
    },
  }
}

setObjectDragFollower(beginObjectDrag)
setCarriedSettler(settleCarried)
