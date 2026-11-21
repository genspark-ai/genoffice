import { save, type FileActionContext } from './file-actions'
import { reducePictures, type ReduceOptions, type ReduceOutcome } from './reduce-pictures'

export type ReduceResult =
  | { status: 'done' | 'none'; outcome: ReduceOutcome }
  /** the file was not written */
  | { status: 'failed' }
  /** the file was written, but the open document moved on meanwhile and still holds the old pictures */
  | { status: 'raced' }

/**
 * Reduce File Size is one step with the save: the pictures are resampled in
 * the serialized document, the file is written, and the editor reloads from
 * what was written, so the page and the file never hold different pictures.
 */
export async function reduceFileSize(
  ctx: FileActionContext,
  options: ReduceOptions,
): Promise<ReduceResult> {
  let outcome: ReduceOutcome | null = null
  const saved = await save(ctx, false, false, undefined, undefined, async (bytes) => {
    const reduced = await reducePictures(bytes, options)
    outcome = reduced.outcome
    return reduced.bytes
  })
  const result = outcome as ReduceOutcome | null
  if (!result) return { status: 'failed' }
  if (result.reduced === 0) return { status: 'none', outcome: result }
  if (!saved) return { status: 'failed' }
  if (ctx.saveIncompleteRef.current) return { status: 'raced' }
  return { status: 'done', outcome: result }
}
