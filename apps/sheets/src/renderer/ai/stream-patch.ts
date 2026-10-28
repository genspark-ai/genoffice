import type { AiChatMessage } from './AiChatPanel'

/** Assistant-message patch for one streamed text update: the live text lands
 * only when the user streams replies; with the pref off the bubble keeps its
 * pending state (the typing indicator) and onDone reveals the completed text.
 * The error flag clears either way — a successful retry must not stay red. */
export function streamedTextPatch(
  entry: AiChatMessage,
  text: string,
  showStreaming: boolean,
): AiChatMessage {
  return showStreaming ? { ...entry, text, isError: false } : { ...entry, isError: false }
}
