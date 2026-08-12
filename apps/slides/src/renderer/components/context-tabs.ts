import type { RenderNode } from '@genoffice/pptx-render'

export type ContextElementType = 'table' | 'chart' | 'picture' | 'shape' | 'textShape' | null
export type ContextTab = 'tableDesign' | 'chartDesign' | 'pictureFormat'

function nodeHasVisibleText(node: RenderNode): boolean {
  if (node.type === 'text' || node.type === 'shape') {
    return (node.text?.lines ?? []).some((line) =>
      line.runs.some((run) => run.text.trim().length > 0),
    )
  }
  return node.type === 'group' && node.children.some(nodeHasVisibleText)
}

export function contextElementTypeForNode(node: RenderNode): ContextElementType {
  if (node.type === 'table' || node.type === 'chart' || node.type === 'picture') return node.type
  if (node.type === 'shape') return nodeHasVisibleText(node) ? 'textShape' : 'shape'
  if (node.type === 'group') return nodeHasVisibleText(node) ? 'textShape' : 'shape'
  return null
}

export function contextTabForElement(type: ContextElementType): ContextTab | null {
  if (type === 'table') return 'tableDesign'
  if (type === 'chart') return 'chartDesign'
  if (type === 'picture' || type === 'shape' || type === 'textShape') return 'pictureFormat'
  return null
}

/**
 * Shapes share picture-format commands such as outline and arrange, but selecting a
 * text-bearing shape should not pull the user away from Home's text controls.
 */
export function autoContextTabForElement(type: ContextElementType): ContextTab | null {
  return type === 'textShape' ? null : contextTabForElement(type)
}
