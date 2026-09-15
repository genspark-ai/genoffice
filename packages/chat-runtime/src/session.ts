/**
 * Session factory + load helpers.
 *
 * The `ChatSession` is the durable artefact the renderer mounts against;
 * a session owns its message history and metadata. The runtime takes a
 * session and adds runs on top.
 */

import type { ChatApp, ChatSession, ChatMessage, ChatPersistence, ChatCapability } from './types'

export interface CreateSessionInput {
  id: string
  app: ChatApp
  title?: string
  projectId?: string
  chatId?: string
  capability?: ChatCapability
  messages?: ChatMessage[]
}

export function createSession(input: CreateSessionInput): ChatSession {
  const now = Date.now()
  return {
    id: input.id,
    app: input.app,
    title: input.title ?? 'New Chat',
    createdAt: now,
    updatedAt: now,
    messages: input.messages ? [...input.messages] : [],
    capability: input.capability,
    projectId: input.projectId,
    chatId: input.chatId,
  }
}

export async function loadOrCreateSession(
  persistence: ChatPersistence | undefined,
  input: CreateSessionInput,
): Promise<ChatSession> {
  if (persistence) {
    const existing = await persistence.load(input.id)
    if (existing) return existing
  }
  return createSession(input)
}

export function appendMessage(session: ChatSession, message: ChatMessage): ChatSession {
  return {
    ...session,
    messages: [...session.messages, message],
    updatedAt: message.timestamp,
  }
}

export function replaceMessages(session: ChatSession, messages: ChatMessage[]): ChatSession {
  return { ...session, messages, updatedAt: Date.now() }
}
