import type { Editor } from '@tiptap/core'
import type { DocsRangePreconditions } from '../../shared/edit-preconditions'
import type { DocsErrorCode } from '../../shared/operation-result'
import { projectedTargetHtml } from './document-query'

export function checkRangePreconditions(
  editor: Editor,
  input: { startBlockIndex: number; endBlockIndex: number; preconditions?: DocsRangePreconditions },
  documentId: string,
  revision: number,
): { error: string; errorCode: DocsErrorCode } | undefined {
  const conditions = input.preconditions
  if (!conditions) return
  const refuse = (error: string, errorCode: DocsErrorCode = 'invalid_input') => ({
    error,
    errorCode,
  })
  const start = input.startBlockIndex
  const end = input.endBlockIndex
  if (
    !Number.isSafeInteger(start) ||
    !Number.isSafeInteger(end) ||
    start < 0 ||
    end < start ||
    end >= editor.state.doc.childCount
  )
    return refuse('Invalid replacement range')
  const blocks: Array<{ from: number; to: number; type: string }> = []
  editor.state.doc.forEach((node, pos) =>
    blocks.push({ from: pos, to: pos + node.nodeSize, type: node.type.name }),
  )
  const selected = blocks.slice(start, end + 1)
  for (const target of [...(conditions.targets ?? []), ...(conditions.preserve ?? [])]) {
    if (target.documentId !== documentId || target.revision !== revision)
      return refuse('Stale replacement precondition; query_document again.', 'stale_target')
    if (target.kind && target.kind !== 'block')
      return refuse('Replacement preconditions require block targets')
    if (
      !blocks.some(
        (block) =>
          block.from === target.range.from &&
          block.to === target.range.to &&
          block.type === target.expectedType,
      )
    )
      return refuse('Replacement target type or extent precondition failed')
    if (
      target.expectedHtml !== undefined &&
      projectedTargetHtml(editor, target) !== target.expectedHtml
    )
      return refuse('Replacement target content precondition failed')
  }
  if (
    conditions.targets &&
    (conditions.targets.length !== selected.length ||
      selected.some(
        (block, i) =>
          block.from !== conditions.targets![i].range.from ||
          block.to !== conditions.targets![i].range.to,
      ))
  )
    return refuse('Replacement range does not match the exact ordered target sequence')
  if (
    conditions.allowedTypes &&
    selected.some((block) => !conditions.allowedTypes!.includes(block.type))
  )
    return refuse('Replacement range includes a block type excluded by allowedTypes')
  if (
    conditions.preserve?.some(
      (target) =>
        target.range.from < selected[selected.length - 1].to && target.range.to > selected[0].from,
    )
  )
    return refuse('Replacement range overlaps a block the caller requires preserved')
}
