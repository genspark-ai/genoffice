import type { RenderNode } from '@genoffice/pptx-render'

/** 'mixed' = multi-select spanning pictures/groups: only selection-wide commands (outline) apply */
export type ContextElementType =
  'table' | 'chart' | 'picture' | 'shape' | 'textShape' | 'mixed' | null
export type ContextTab = 'tableDesign' | 'chartDesign' | 'pictureFormat' | 'shapeFormat'

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
  // Mixed selections use picture-format: its outline command spans the whole
  // selection while picture-only tools (crop/transparency) stay disabled
  if (type === 'picture' || type === 'mixed') return 'pictureFormat'
  if (type === 'shape' || type === 'textShape') return 'shapeFormat'
  return null
}

/**
 * Shapes get their own format tab (styles/fill/outline), but selecting a text-bearing
 * shape should not pull the user away from Home's text controls.
 */
export function autoContextTabForElement(type: ContextElementType): ContextTab | null {
  return type === 'textShape' ? null : contextTabForElement(type)
}
