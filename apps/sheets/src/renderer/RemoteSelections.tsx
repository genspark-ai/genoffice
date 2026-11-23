import { useEffect, useState, useSyncExternalStore } from 'react'
import { IRenderManagerService, SHEET_VIEWPORT_KEY } from '@univerjs/engine-render'
import { collabPresenceColor, currentCollabPresence, subscribeCollabPresence } from '@genoffice/ui'
import type { UniverRuntime } from './univer-state'
import {
  labelGoesInside,
  membersOnSheet,
  REMOTE_LABEL_HEIGHT,
  type RemoteSelectionMember,
} from './remote-selection-geometry'
import './remote-selections.css'

/**
 * The selections of the other members of the room, drawn over the grid.
 *
 * The grid is a canvas, so the remote selections are DOM boxes on top of it.
 * The coordinate mapping is the one the shape layer already uses: Univer's
 * cell rect is scene-space, so the viewport scroll is subtracted and the
 * result scaled by the zoom (see shape-draw.ts).
 *
 * Only the members on the sheet that is on screen are drawn — a selection on
 * another sheet has no cells to point at, and its skeleton is not even built.
 */

interface RemoteBox {
  id: number
  name: string
  x: number
  y: number
  width: number
  height: number
  /** true when the label would fall above the grid, so it goes inside the box */
  labelInside: boolean
}

/** the screen rect of a range, in viewport pixels */
function rangeRect(
  runtime: UniverRuntime | null,
  sheetId: string,
  range: string,
): Omit<RemoteBox, 'id' | 'name' | 'labelInside'> | null {
  if (!runtime) return null
  const workbook = runtime.univerAPI.getActiveWorkbook()
  const worksheet = workbook?.getSheetBySheetId(sheetId)
  if (!workbook || !worksheet) return null
  // the grid's render surface is the largest canvas in the container
  const hostEl = document.getElementById('univer-container')
  if (!hostEl) return null
  let surface: DOMRect | null = null
  for (const canvas of hostEl.querySelectorAll('canvas')) {
    const rect = canvas.getBoundingClientRect()
    if (!surface || rect.width * rect.height > surface.width * surface.height) surface = rect
  }
  if (!surface) return null
  const viewMain = runtime.univer
    .__getInjector()
    .get(IRenderManagerService)
    .getRenderById(workbook.getId())
    ?.scene?.getViewport(SHEET_VIEWPORT_KEY.VIEW_MAIN)
  const scrollX = viewMain?.viewportScrollX ?? 0
  const scrollY = viewMain?.viewportScrollY ?? 0
  const zoom = worksheet.getZoom() || 1
  try {
    const cellRect = worksheet.getRange(range).getCellRect()
    return {
      x: surface.x + (cellRect.x - scrollX) * zoom,
      y: surface.y + (cellRect.y - scrollY) * zoom,
      width: (cellRect.right - cellRect.x) * zoom,
      height: (cellRect.bottom - cellRect.y) * zoom,
    }
  } catch {
    return null
  }
}

export function RemoteSelections({
  runtime,
}: {
  /** read live: the runtime exists after the editor mounts */
  runtime: () => UniverRuntime | null
}) {
  const presence = useSyncExternalStore(subscribeCollabPresence, currentCollabPresence)
  const [boxes, setBoxes] = useState<RemoteBox[]>([])

  useEffect(() => {
    const members: RemoteSelectionMember[] = [...presence.values()]
    if (members.length === 0) {
      setBoxes([])
      return
    }
    // the boxes follow the local scroll and zoom, so they are placed again a
    // few times a second while somebody else is in the room
    const tick = (): void => {
      const current = runtime()
      const sheetId = current?.univerAPI.getActiveWorkbook()?.getActiveSheet()?.getSheetId()
      const surface = document.getElementById('univer-container')
      const surfaceTop = surface?.getBoundingClientRect().top ?? 0
      const next: RemoteBox[] = []
      for (const member of membersOnSheet(members, sheetId)) {
        const rect = rangeRect(current, member.sheetId!, member.range!)
        if (rect) {
          next.push({
            ...rect,
            id: member.id,
            name: member.name,
            labelInside: labelGoesInside(rect.y, surfaceTop),
          })
        }
      }
      setBoxes(next)
    }
    tick()
    const timer = setInterval(tick, 250)
    return () => clearInterval(timer)
  }, [presence, runtime])

  if (boxes.length === 0) return null
  return (
    <>
      {boxes.map((box) => {
        const color = collabPresenceColor(box.id)
        return (
          <div
            key={box.id}
            className="remote-selection"
            style={{
              left: box.x,
              top: box.y,
              width: box.width,
              height: box.height,
              borderColor: color,
            }}
          >
            <span
              className="remote-selection-label"
              style={{ background: color, top: box.labelInside ? 0 : -REMOTE_LABEL_HEIGHT }}
            >
              {box.name}
            </span>
          </div>
        )
      })}
    </>
  )
}
