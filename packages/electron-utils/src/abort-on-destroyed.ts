/** Structural subset of Electron's WebContents, so the helper stays Electron-free and unit-testable. */
export interface DestroyableSender {
  isDestroyed(): boolean
  once(event: 'destroyed', listener: () => void): unknown
  removeListener(event: 'destroyed', listener: () => void): unknown
}

/**
 * Aborts `controller` when the renderer that started a stream goes away
 * (tab or window closed mid-stream), so the provider request stops instead
 * of running to completion for a listener that no longer exists.
 * Returns a disposer; call it when the stream settles so the listener does
 * not accumulate on long-lived senders.
 */
export function abortOnDestroyed(
  sender: DestroyableSender,
  controller: AbortController,
): () => void {
  if (sender.isDestroyed()) {
    controller.abort()
    return () => {}
  }
  const onDestroyed = () => controller.abort()
  sender.once('destroyed', onDestroyed)
  return () => sender.removeListener('destroyed', onDestroyed)
}
