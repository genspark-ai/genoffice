/** Toast event bus, split from the ToastHost component so this module has no
 * component exports: React Fast Refresh replaces mixed-export modules wholesale,
 * which strands callers (file-actions) on a stale emitter during dev HMR. */

export interface ToastData {
  text: string
  kind: 'success' | 'error'
}

let emit: ((toast: ToastData) => void) | null = null

/** Registered by ToastHost on mount; null while unmounted. */
export function setToastEmitter(fn: ((toast: ToastData) => void) | null): void {
  emit = fn
}

export function showToast(text: string, kind: 'success' | 'error' = 'success'): void {
  emit?.({ text, kind })
}
