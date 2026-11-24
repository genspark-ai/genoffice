/** how many times one layout input may be laid out again for its footnote numbers */
export const NOTE_SETTLE_CAP = 6

export type NoteSettleStop = 'cycle' | 'cap'

/**
 * Stopping rule of the paginate → renumber → paginate loop behind footnotes
 * that restart on each page. A narrower mark can pull a reference onto the
 * page before, which changes its number again, so the loop need not converge:
 * it ends when the numbers a layout asks for are the ones it was made with,
 * when it asks for numbers this run already tried, or at the cap. In the last
 * two cases the marks keep the numbers the page is drawn with.
 */
export class NotePageSettle {
  private input: unknown = null
  private tried = new Set<string>()
  private rounds = 0
  private stopped = false
  /** the most re-layouts one run took, and why the last run that did not settle ended */
  readonly stats: { maxRounds: number; stop: NoteSettleStop | null } = { maxRounds: 0, stop: null }

  /**
   * Whether the marks take `proposed`. `input` identifies what was laid out
   * apart from the numbers (the document as edited): a new one opens a new run.
   */
  step(input: unknown, current: string, proposed: string): boolean {
    if (input !== this.input) this.open(input)
    if (proposed === current) {
      this.open(input)
      return false
    }
    if (this.stopped) return false
    const stop = this.tried.has(proposed) ? 'cycle' : this.rounds >= NOTE_SETTLE_CAP ? 'cap' : null
    if (stop) {
      this.stopped = true
      this.stats.stop = stop
      return false
    }
    this.tried.add(current)
    this.rounds++
    this.stats.maxRounds = Math.max(this.stats.maxRounds, this.rounds)
    return true
  }

  /** the renumbering this loop dispatched made a new document: it continues the run */
  carry(input: unknown): void {
    this.input = input
  }

  private open(input: unknown): void {
    this.input = input
    this.tried.clear()
    this.rounds = 0
    this.stopped = false
  }
}
