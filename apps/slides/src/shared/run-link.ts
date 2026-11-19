/**
 * Run-hyperlink string encoding shared by renderer and main: external url, "slide:N"
 * (0-based) for in-doc jumps, "action:<name>" for named show actions — same encoding as
 * the engine's TextRun.hyperlink.
 */
import { NAMED_ACTIONS, type NamedAction } from '@genoffice/pptx-engine/named-action'
import type { LinkTargetOp } from './ipc'

export function encodeLinkTarget(target: LinkTargetOp): string {
  if (target.kind === 'slide') return `slide:${target.slideIndex}`
  if (target.kind === 'action') return `action:${target.action}`
  return target.url
}

export function decodeLinkTarget(s: string | null | undefined): LinkTargetOp | null {
  if (!s) return null
  const m = /^slide:(\d+)$/.exec(s)
  if (m) return { kind: 'slide', slideIndex: Number(m[1]) }
  const a = /^action:(\w+)$/.exec(s)
  if (a) {
    return (NAMED_ACTIONS as readonly string[]).includes(a[1]!)
      ? { kind: 'action', action: a[1] as NamedAction }
      : null
  }
  return { kind: 'url', url: s }
}

/** Schemes a deck link may carry into the package / export: the suite gate plus mailto. */
export const DECK_LINK_PROTOCOLS: readonly string[] = ['http:', 'https:', 'mailto:']
