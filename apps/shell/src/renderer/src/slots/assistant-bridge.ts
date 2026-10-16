/**
 * A slot's press can ask for the assistant: the request is parked here, Home
 * switches to the assistant, and the panel picks the text up as its draft. The
 * indirection keeps a drawn element from knowing how Home is wired.
 */

const listeners = new Set<(prompt: string) => void>()
let pending: string | null = null

/** Queue a prompt and tell Home to show the assistant. */
export function requestAssistant(prompt: string): void {
  pending = prompt
  for (const listener of listeners) listener(prompt)
}

/** The panel takes the parked prompt, if any, exactly once. */
export function takePendingPrompt(): string {
  const prompt = pending ?? ''
  pending = null
  return prompt
}

export function onAssistantRequest(listener: (prompt: string) => void): () => void {
  listeners.add(listener)
  return () => {
    listeners.delete(listener)
  }
}
