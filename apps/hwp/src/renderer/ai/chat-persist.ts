export interface PersistedTool {
  name: string
  summary: string
  isError?: boolean
}

export interface PersistedChatMessage {
  role: 'user' | 'assistant'
  text: string
  tools?: PersistedTool[]
}

export function shouldApplyRestoredChat(opts: {
  messageCount: number
  chatAlreadyHasMessages: boolean
  loopBusy: boolean
  skipped: boolean
}): boolean {
  return (
    !opts.skipped && opts.messageCount > 0 && !opts.chatAlreadyHasMessages && !opts.loopBusy
  )
}

/** After consumePending sets a real path, reload only if we just left unsaved-* and the UI is still empty. */
export function shouldLoadAfterRebind(opts: {
  previousChatId: string
  nextChatId: string
  chatAlreadyHasMessages: boolean
  skipped: boolean
}): boolean {
  return (
    !opts.skipped &&
    !opts.chatAlreadyHasMessages &&
    opts.nextChatId.length > 0 &&
    opts.nextChatId !== opts.previousChatId
  )
}

export function mapPersistedChat(msgs: readonly PersistedChatMessage[]): PersistedChatMessage[] {
  return msgs.map((m) => ({
    role: m.role,
    text: m.text,
    tools: m.tools?.map((tool) => ({
      name: tool.name,
      summary: tool.summary,
      isError: tool.isError,
    })),
  }))
}
