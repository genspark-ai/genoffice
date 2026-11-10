import { useEffect, useState } from 'react'

/** how much tool/step detail AI chat panels show; chosen in the shell's Settings › Appearance */
export type ConversationSteps = 'default' | 'verbose'

// same key the shell writes (apps/shell appearance.ts); editors share it through localStorage
const STORAGE_KEY = 'genoffice.appearance'

export function readConversationSteps(): ConversationSteps {
  try {
    const raw = window.localStorage.getItem(STORAGE_KEY)
    if (!raw) return 'default'
    const data = JSON.parse(raw) as { conversationSteps?: unknown }
    return data.conversationSteps === 'verbose' ? 'verbose' : 'default'
  } catch {
    return 'default'
  }
}

/** live value: updates when Settings changes it in the shell (storage events fire across windows) */
export function useConversationSteps(): ConversationSteps {
  const [value, setValue] = useState<ConversationSteps>(readConversationSteps)
  useEffect(() => {
    const onStorage = (e: StorageEvent): void => {
      if (e.key === null || e.key === STORAGE_KEY) setValue(readConversationSteps())
    }
    window.addEventListener('storage', onStorage)
    // a panel opened after the setting changed, or a missed event
    setValue(readConversationSteps())
    return () => window.removeEventListener('storage', onStorage)
  }, [])
  return value
}
