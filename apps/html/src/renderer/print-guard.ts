import type { PrintResult } from '../shared/ipc'

// The system print dialog is modal; a second CmdOrCtrl+P while it is open
// would spawn a second window and dialog. The shell menu owns the accelerator,
// so this gate is the only thing between one keystroke and two printouts.
export interface PrintGuard {
  /** shaped like a React ref so App.tsx hands its own useRef in */
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
    return !('canceled' in result)
  } catch (err) {
    onFailure(err instanceof Error ? err.message : String(err))
    return false
  } finally {
    guard.current = false
  }
}
