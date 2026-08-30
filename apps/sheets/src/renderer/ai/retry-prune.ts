import type { AiChatMessage } from './AiChatPanel'

/**
 * Drop a failed (undelivered) user bubble and its paired error reply, so that
 * Retry re-sends the message in place instead of stacking a duplicate
 * transcript ("Not sent" bubble + raw error + the same text again).
 */
export function pruneFailedExchange(
  chat: readonly AiChatMessage[],
  index: number,
): readonly AiChatMessage[] {
  const entry = chat[index]
  if (!entry || entry.role !== 'user' || !entry.undelivered) return chat
  const partner = chat[index + 1]
  const count = partner && partner.role === 'assistant' && partner.isError ? 2 : 1
  return [...chat.slice(0, index), ...chat.slice(index + count)]
}
