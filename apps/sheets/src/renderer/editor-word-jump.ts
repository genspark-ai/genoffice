/**
 * Word-by-word caret movement in the cell editor and formula bar.
 *
 * Univer's doc editor registers only plain and Shift arrow shortcuts, so
 * Ctrl+←/→ (Option+←/→ on mac, Excel-for-mac convention) degraded to a
 * one-character move. These bindings are gated on `whenSheetEditorActivated`,
 * so the grid-level "jump to the data edge" Ctrl+arrows are untouched.
 *
 * Boundaries come from Intl.Segmenter, so accented words and CJK runs jump
 * as whole words rather than ASCII islands. Univer stream tokens (every
 * control character except TAB) are hard boundaries: the caret stops at a
 * paragraph end and crosses it on the next press, like Word.
 */
import {
  CommandType,
  FOCUSING_EDITOR_INPUT_FORMULA,
  ICommandService,
  IContextService,
  IUniverInstanceService,
  RANGE_DIRECTION,
} from '@univerjs/core'
import type { ICommand } from '@univerjs/core'
import { DocSelectionManagerService } from '@univerjs/docs'
import { IShortcutService, KeyCode, MetaKeys } from '@univerjs/ui'
import { whenSheetEditorActivated } from '@univerjs/sheets-ui'
import type { UniverRuntime } from './univer-state'

const MOVE_WORD_ID = 'genoffice.command.editor-move-word'
const SELECT_WORD_ID = 'genoffice.command.editor-select-word'

/** every control char is a stream token except TAB, which is just whitespace */
const isBoundary = (ch: string | undefined): boolean =>
  ch !== undefined && ch !== '\t' && ch.charCodeAt(0) < 32

const segmenter = new Intl.Segmenter(undefined, { granularity: 'word' })

interface SliceSegment {
  start: number
  end: number
  wordLike: boolean
}

/** word segments of the token-free slice around `offset` (absolute offsets) */
function sliceSegments(
  stream: string,
  offset: number,
): { start: number; segments: SliceSegment[] } {
  let start = offset
  while (start > 0 && !isBoundary(stream[start - 1])) start--
  let end = offset
  while (end < stream.length && !isBoundary(stream[end])) end++
  const segments: SliceSegment[] = []
  for (const s of segmenter.segment(stream.slice(start, end))) {
    segments.push({
      start: start + s.index,
      end: start + s.index + s.segment.length,
      wordLike: Boolean((s as { isWordLike?: boolean }).isWordLike),
    })
  }
  return { start, segments }
}

/**
 * The offset a word-jump lands on. Word semantics: → goes to the start of the
 * next word (end of text if none), ← to the start of the current word, or of
 * the previous one when already at a word start. A press at a slice edge
 * crosses the adjacent stream token.
 */
export function wordJumpOffset(stream: string, offset: number, dir: -1 | 1): number {
  // the stream's final "\r\n" terminator pair is never a caret home
  const limit = Math.max(0, stream.length - 2)
  if (dir === 1) {
    if (offset >= limit) return limit
    if (isBoundary(stream[offset])) {
      let next = offset + 1
      while (next < limit && isBoundary(stream[next])) next++
      return Math.min(next, limit)
    }
    const { segments } = sliceSegments(stream, offset)
    const inside = segments.find((s) => s.start <= offset && offset < s.end)
    for (const s of segments) {
      if (s.wordLike && s.start > offset && (!inside || s.start >= inside.end)) {
        return Math.min(s.start, limit)
      }
    }
    const end = segments.length > 0 ? segments[segments.length - 1]!.end : offset
    return Math.min(end, limit)
  }
  if (offset <= 0) return 0
  let at = offset
  if (isBoundary(stream[at - 1])) {
    at--
    while (at > 0 && isBoundary(stream[at - 1])) at--
    if (at === 0) return 0
  }
  const { start, segments } = sliceSegments(stream, at)
  let target = start
  for (const s of segments) {
    if (s.wordLike && s.start < at) target = s.start
    if (s.start >= at) break
  }
  return target
}

/**
 * Consecutive Shift-presses must extend from the same anchor. The active
 * range does not reliably carry which side the focus is on, so remember the
 * anchor across presses and drop it as soon as the selection is no longer
 * the one this command produced.
 */
let lastExtend: { anchor: number; start: number; end: number } | null = null

function registerCommands(commandService: ICommandService): void {
  const jump = (
    accessor: Parameters<NonNullable<ICommand['handler']>>[0],
    dir: -1 | 1,
    extend: boolean,
  ): boolean => {
    const docSelection = accessor.get(DocSelectionManagerService)
    const active = docSelection.getActiveTextRange()
    const doc = accessor.get(IUniverInstanceService).getCurrentUniverDocInstance()
    if (!active || !doc) return false
    const body = doc.getSelfOrHeaderFooterModel(active.segmentId).getBody()
    const stream = body?.dataStream
    if (!stream) return false
    const { startOffset, endOffset, collapsed, style } = active
    if (startOffset == null || endOffset == null) return false
    if (!extend) {
      lastExtend = null
      const from = dir === 1 ? endOffset : startOffset
      const target = wordJumpOffset(stream, from, dir)
      docSelection.replaceTextRanges(
        [{ startOffset: target, endOffset: target, ...(style ? { style } : {}) }],
        false,
      )
      return true
    }
    const remembered =
      lastExtend && lastExtend.start === startOffset && lastExtend.end === endOffset
        ? lastExtend.anchor
        : null
    const anchor =
      remembered ??
      (collapsed
        ? startOffset
        : active.direction === RANGE_DIRECTION.BACKWARD
          ? endOffset
          : startOffset)
    const focus =
      remembered != null || !collapsed
        ? anchor === startOffset
          ? endOffset
          : startOffset
        : startOffset
    const target = wordJumpOffset(stream, focus, dir)
    const next = { start: Math.min(anchor, target), end: Math.max(anchor, target) }
    lastExtend = { anchor, ...next }
    docSelection.replaceTextRanges(
      [{ startOffset: next.start, endOffset: next.end, ...(style ? { style } : {}) }],
      false,
    )
    return true
  }

  const moveWord: ICommand<{ dir: -1 | 1 }> = {
    id: MOVE_WORD_ID,
    type: CommandType.COMMAND,
    handler: (accessor, params) => (params ? jump(accessor, params.dir, false) : false),
  }
  const selectWord: ICommand<{ dir: -1 | 1 }> = {
    id: SELECT_WORD_ID,
    type: CommandType.COMMAND,
    handler: (accessor, params) => (params ? jump(accessor, params.dir, true) : false),
  }
  commandService.registerCommand(moveWord)
  commandService.registerCommand(selectWord)
}

/** Ctrl+←/→ (mac: Option) word moves in the cell editor / formula bar */
export function registerEditorWordJump(runtime: UniverRuntime): void {
  const injector = runtime.univer.__getInjector()
  registerCommands(injector.get(ICommandService))
  const shortcutService = injector.get(IShortcutService)
  const items = [
    { dir: 1 as const, key: KeyCode.ARROW_RIGHT, extend: false },
    { dir: -1 as const, key: KeyCode.ARROW_LEFT, extend: false },
    { dir: 1 as const, key: KeyCode.ARROW_RIGHT, extend: true },
    { dir: -1 as const, key: KeyCode.ARROW_LEFT, extend: true },
  ]
  for (const { dir, key, extend } of items) {
    const shift = extend ? MetaKeys.SHIFT : 0
    shortcutService.registerShortcut({
      id: extend ? SELECT_WORD_ID : MOVE_WORD_ID,
      binding: key | MetaKeys.CTRL_COMMAND | shift,
      mac: key | MetaKeys.ALT | shift,
      // Univer's formula-embedding editor intercepts every arrow binding at
      // priority 900 (it owns reference point-mode); word moves must outrank
      // it for plain text, and the precondition yields the formula case back
      // so Ctrl+arrow keeps its reference semantics inside a "=" edit.
      priority: 1000,
      preconditions: (contextService: IContextService) =>
        whenSheetEditorActivated(contextService) &&
        !contextService.getContextValue(FOCUSING_EDITOR_INPUT_FORMULA),
      staticParameters: { dir },
    })
  }
}
