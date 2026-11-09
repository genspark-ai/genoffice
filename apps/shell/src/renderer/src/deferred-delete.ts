export interface DeferredDeleteOptions {
  /** how long the Undo window stays open */
  delayMs: number
  /** runs once the window closes without an undo (the real delete) */
  commit: (paths: string[]) => void | Promise<void>
  /** fired whenever the set of hidden paths changes */
  onChange?: (pending: ReadonlySet<string>) => void
}

/**
 * Delete-with-undo: paths are hidden immediately and only handed to `commit`
 * after `delayMs`. They stay hidden until `commit` settles so rows never flash
 * back. Closing the app inside the window leaves the files untouched.
 */
export function createDeferredDelete(opts: DeferredDeleteOptions) {
  let nextId = 1
  const batches = new Map<
    number,
    { paths: string[]; timer: ReturnType<typeof setTimeout>; committing: boolean }
  >()

  const pending = (): ReadonlySet<string> => {
    const all = new Set<string>()
    for (const b of batches.values()) for (const p of b.paths) all.add(p)
    return all
  }
  const notify = (): void => opts.onChange?.(pending())

  const run = async (id: number): Promise<void> => {
    const batch = batches.get(id)
    if (!batch || batch.committing) return
    batch.committing = true
    try {
      await opts.commit(batch.paths)
    } finally {
      batches.delete(id)
      notify()
    }
  }

  return {
    schedule(paths: string[]): number {
      const id = nextId++
      const timer = setTimeout(() => void run(id), opts.delayMs)
      batches.set(id, { paths: [...paths], timer, committing: false })
      notify()
      return id
    },
    /** cancel a batch; returns its paths, or null when it already committed */
    undo(id: number): string[] | null {
      const batch = batches.get(id)
      if (!batch || batch.committing) return null
      clearTimeout(batch.timer)
      batches.delete(id)
      notify()
      return batch.paths
    },
    /** commit everything now (e.g. before leaving the view) */
    flush(): void {
      for (const [id, batch] of [...batches]) {
        clearTimeout(batch.timer)
        void run(id)
      }
    },
    pending,
  }
}
