/**
 * The decisions behind the remote selection overlay.
 *
 * They are kept apart from the component because they are the parts that can
 * be wrong without anybody noticing: a box on the wrong sheet, or a label that
 * disappears above the grid.
 */

/** what the overlay needs to know about a member of the room */
export interface RemoteSelectionMember {
  /** the member's awareness id */
  id: number
  /** the member's display name */
  name: string
  /** the sheet the member is on */
  sheetId?: string | undefined
  /** the range the member has selected, in A1 notation */
  range?: string | undefined
}

/**
 * The members whose selection is on the sheet that is on screen.
 *
 * A selection on another sheet has no cells to point at, so it is not drawn —
 * the member still shows up in the status bar with the sheet they are on.
 */
export function membersOnSheet(
  members: readonly RemoteSelectionMember[],
  activeSheetId: string | undefined,
): RemoteSelectionMember[] {
  if (!activeSheetId) return []
  return members.filter((member) => member.sheetId === activeSheetId && !!member.range)
}

/** the height of the name tag that rides on the top edge of a selection */
export const REMOTE_LABEL_HEIGHT = 16

/**
 * True when the name tag would fall above the grid, so it goes inside the box
 * instead. The first row of a sheet sits right under the grid's top edge, and a
 * tag above it would be drawn over the sheet tabs.
 */
export function labelGoesInside(
  boxTop: number,
  surfaceTop: number,
  labelHeight: number = REMOTE_LABEL_HEIGHT,
): boolean {
  return boxTop - surfaceTop < labelHeight
}
