import type { RenderNode } from '@genoffice/pptx-render'
import { contextualTabFor, type ContextTab } from './components/context-tabs'

export type DblClickAction =
  | { kind: 'editText' }
  | { kind: 'enterGroup' }
  | { kind: 'editCell' }
  | { kind: 'playMedia' }
  | { kind: 'openTab'; tab: ContextTab }

export interface DblClickContext {
  /** Text can be edited in place (connectors never; group children only when the overlay can follow) */
  editable: boolean
  insideGroup: boolean
  canEnterGroup: boolean
  canPlayMedia: boolean
  /** Table only: the pointer landed on a cell rather than the border */
  hitCell?: boolean
}

export function dblClickActionFor(node: RenderNode, ctx: DblClickContext): DblClickAction | null {
  if (ctx.editable) return { kind: 'editText' }
  if (!ctx.insideGroup) {
    if (node.type === 'group' && ctx.canEnterGroup) return { kind: 'enterGroup' }
    if (node.type === 'table' && ctx.hitCell) return { kind: 'editCell' }
    if (node.type === 'picture' && node.media && ctx.canPlayMedia) return { kind: 'playMedia' }
  }
  const tab = contextualTabFor(node)
  return tab ? { kind: 'openTab', tab } : null
}
