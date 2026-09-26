const MAX_AGE_MS = 2000

/**
 * Right-clicks on the document body answered by the React menu. The renderer
 * claims each one synchronously from its DOM handler, i.e. before Blink
 * requests a context menu for it, so the main process sees claims and
 * `context-menu` events in the same order and can pair them FIFO without
 * coordinates (a keyboard-invoked menu reports different points on the two
 * sides). A claim whose event never comes expires.
 */
export class ClickClaims {
  private claims: Array<{ seq: number; at: number }> = []

  claim(seq: number, now: number): void {
    this.prune(now)
    this.claims.push({ seq, at: now })
  }

  /** Consume the oldest live claim for an incoming context-menu event; null when the click was not claimed. */
  take(now: number): number | null {
    this.prune(now)
    return this.claims.shift()?.seq ?? null
  }

  private prune(now: number): void {
    this.claims = this.claims.filter((c) => now - c.at <= MAX_AGE_MS)
  }
}
