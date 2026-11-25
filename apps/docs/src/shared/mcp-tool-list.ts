/** Check live authority while MCP clients may be holding a cached tools/list result. */
export function monitorMcpToolList(
  signal: AbortSignal,
  readNames: () => Promise<string[]>,
  notify: () => Promise<void>,
) {
  let last: string | undefined
  let timer: ReturnType<typeof setTimeout> | undefined
  let listing = 0
  const signature = (names: string[]) => JSON.stringify([...names].sort())

  const schedule = () => {
    if (signal.aborted || timer) return
    timer = setTimeout(async () => {
      const version = listing
      try {
        const next = signature(await readNames())
        if (!signal.aborted && version === listing && next !== last) {
          await notify()
          if (version === listing) last = next
        }
      } catch {
        // A transient bridge or transport failure is retried while the connection is live.
      } finally {
        timer = undefined
        schedule()
      }
    }, 1000)
    timer.unref()
  }
  signal.addEventListener('abort', () => clearTimeout(timer), { once: true })

  return (names: string[]) => {
    listing++
    last = signature(names)
    schedule()
  }
}
