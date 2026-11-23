import React from 'react'
import './collab-presence-chip.css'

/**
 * The presence chip of a shared document.
 *
 * One avatar per other member of the room, in the status bar next to the
 * share chip. The avatar's colour comes from the member's awareness id, so
 * both ends paint the same member the same colour without exchanging one.
 */

export interface CollabPresenceMember {
  /** the member's awareness id — the colour is derived from it */
  id: number
  /** the member's name */
  name: string
  /** where the member is right now (a sheet name, a slide name) */
  location?: string | undefined
}

/** a stable colour per member id */
export function collabPresenceColor(id: number): string {
  const hue = (id * 47) % 360
  return `hsl(${hue}, 68%, 46%)`
}

function initialOf(name: string): string {
  const trimmed = name.trim()
  if (!trimmed) return '?'
  // a name may start with a CJK character or a Latin one; the first code point
  // is what the reader recognises
  return Array.from(trimmed)[0]!.toUpperCase()
}

export function CollabPresenceChip({
  members,
}: {
  members: CollabPresenceMember[]
}): React.ReactElement | null {
  if (members.length === 0) return null
  const tip = members
    .map((member) => `${member.name}${member.location ? ` — ${member.location}` : ''}`)
    .join('\n')
  return (
    <span className="collab-presence-chip" data-tip={tip}>
      {members.map((member) => (
        <span
          key={member.id}
          className="collab-presence-avatar"
          style={{ background: collabPresenceColor(member.id) }}
        >
          {initialOf(member.name)}
        </span>
      ))}
      {members.length > 1 ? <span className="collab-presence-count">{members.length}</span> : null}
    </span>
  )
}
