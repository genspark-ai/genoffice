import type { zh } from './zh'

export const en = {
  aiOpenAssistant: 'Open AI assistant',
  aiCollapsePanel: 'Collapse panel',
  aiComposerPlaceholder: 'Ask about this Hangul document…',
  aiEditorNotReady: 'Hangul editor is still loading…',
  aiEmptyTitle: 'Ask about this Hangul document',
  aiEmptyBody:
    'Summarize, translate, or ask questions. Editing is not available in this first release.',
  aiNewChat: 'New chat',
  aiSend: 'Send',
  aiStop: 'Stop',
  aiHintIdle: 'Enter to send, Shift+Enter for a new line',
  aiHintBusy: 'Replying…',
  aiThinking: 'Thinking',
  aiReplying: 'Replying',
  aiWorking: 'Working',
  aiStopped: 'Stopped',
  aiNoReply: '(no reply)',
  aiTruncatedNote: '(The reply was cut off by the length limit and may be incomplete.)',
  aiTurnLimit: 'Reached the per-run step limit — send another instruction to continue.',
  aiUndelivered: 'Not sent',
  aiRetry: 'Retry',
  aiGskLoginBtn: 'Sign in to Genspark',
  aiUnknownError: 'AI request failed, please retry',
  aiTimeoutError: 'AI response timed out',
  aiOverloadedError: 'The AI service is busy right now — please try again in a moment',
  aiNetworkError:
    'Network problem: could not reach the AI service. Check your connection and try again',
  aiCreditsExhausted: 'Out of credits — top up at genspark.ai',
} satisfies Record<keyof typeof zh, string>
