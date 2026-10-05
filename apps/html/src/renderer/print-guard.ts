import type { PrintResult } from '../shared/ipc'

/**
 * The in-flight gate for File > Print.
 *
 * The system print dialog is modal, so a second CmdOrCtrl+P while it is open
 * would spawn a second hidden window and a second dialog on top of it — two
 * copies of the same document competing for the same printer. The shell menu
 * owns the accelerator, so this guard is the only thing between one keystroke
 * and two printouts.
 *
 * Pure, so the gate and its release can be tested without mounting the editor:
 * App.tsx is the only caller and holds the flag itself.
 */
export interface PrintGuard {
  /** true while a print is in flight. Shaped like a React ref so App.tsx can
   * hand its own `useRef` straight in. */
  current: boolean
}

export async function runGuardedPrint(
  guard: PrintGuard,
  print: () => Promise<PrintResult>,
  onFailure: (error: string) => void,
): Promise<boolean> {
  if (guard.current) return false
  guard.current = true
  try {
    const result = await print()
    if (!result.ok) {
      onFailure(result.error)
      return false
    }
    // Cancelled in the system dialog: the user's own outcome, so it stays
    // silent and reports "did not print" the same way a failure does.
    return !('canceled' in result)
  } catch (err) {
    // Serialization and IPC rejections land here. They are failures the user
    // needs to see, not silent no-ops.
    onFailure(err instanceof Error ? err.message : String(err))
    return false
  } finally {
    // Every exit path releases the gate: success, failure, cancellation and a
    // thrown print(). A gate that leaked here would wedge printing for the
    // rest of the session, which is worse than having no gate at all.
    guard.current = false
  }
}
