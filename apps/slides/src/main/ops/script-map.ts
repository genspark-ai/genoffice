/**
 * Edit-script → op mapping: the whole script arrives as one IPC and compiles
 * to one atomic transaction. This is surface translation (px→EMU, group-local
 * scale, pt→EMU stroke widths, style patch → font/paragraph ops) — pure and
 * Node-testable; the executor owns validation/rollback/journal.
 */
import type { OpenedPptx } from '@genoffice/pptx-engine'
import type { ApplyEditScriptOp, ScriptStylePatch } from '../../shared/ipc'
import type { Op } from './registry'

const EMU_PER_PX_96 = 9525
const EMU_PER_PT = 12700

export function mapScriptOps(opened: OpenedPptx, req: ApplyEditScriptOp): Op[] {
  const baseWidthPx = opened.deck.size.cx / EMU_PER_PX_96
  const scale = req.fitWidthPx / baseWidthPx
  const toEmu = (px: number) => Math.round((px / scale) * EMU_PER_PX_96)
  const target = (el: string) => ({ slide: req.slideIndex, el })
  const ops: Op[] = []

  // Geometry first, top-level before group children (matching the legacy dispatch
  // order: children convert against the group's post-move state). Group-child boxes
  // travel as document-space absBox — the setTransform op converts them into the
  // child coordinate system at apply time, against the group's live state. The batch
  // path never redistributed table grids, so resizeTableGrid stays off here.
  const boxOps = [...req.boxes].sort((a, b) => (a.groupId ? 1 : 0) - (b.groupId ? 1 : 0))
  for (const b of boxOps) {
    const box = { x: toEmu(b.x), y: toEmu(b.y), cx: toEmu(b.w), cy: toEmu(b.h) }
    ops.push({
      op: 'setTransform',
      target: target(b.id),
      ...(b.groupId ? { group: b.groupId, absBox: box } : { box }),
      rotDeg: b.rotation,
    })
  }

  for (const e of req.edits) {
    const grp = e.groupId ? { group: e.groupId } : {}
    if (e.kind === 'text') {
      ops.push({ op: 'setText', target: target(e.id), paragraphs: e.paragraphs, ...grp })
    } else if (e.kind === 'style') {
      // Run-level fields ride setFont (the engine merges onto every run, keeping
      // unspecified fields); align is paragraph-level and rides setParagraphFormat.
      const { align, fontSize, ...rest } = e.style as ScriptStylePatch
      const font = { ...rest, ...(fontSize !== undefined ? { fontSizePt: fontSize } : {}) }
      if (Object.keys(font).length > 0) {
        ops.push({ op: 'setFont', target: target(e.id), font, ...grp })
      }
      if (align) {
        ops.push({ op: 'setParagraphFormat', target: target(e.id), format: { align }, ...grp })
      }
    } else if (e.kind === 'fill') {
      ops.push({ op: 'setFill', target: target(e.id), fill: e.fill, ...grp })
    } else {
      ops.push({
        op: 'setStroke',
        target: target(e.id),
        stroke: e.stroke
          ? { color: e.stroke.color, widthEmu: Math.round(e.stroke.widthPt * EMU_PER_PT) }
          : null,
        ...grp,
      })
    }
  }
  return ops
}
