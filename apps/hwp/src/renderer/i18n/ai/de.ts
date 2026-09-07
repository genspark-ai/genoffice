import type { zh } from './zh'

export const de = {
  aiOpenAssistant: 'KI-Assistent öffnen',
  aiCollapsePanel: 'Panel einklappen',
  aiComposerPlaceholder: 'Fragen Sie zu diesem Hangul-Dokument…',
  aiEditorNotReady: 'Hangul-Editor wird noch geladen…',
  aiEmptyTitle: 'Fragen Sie zu diesem Hangul-Dokument',
  aiEmptyBody:
    'Zusammenfassen, übersetzen oder Fragen stellen. Bearbeiten ist in dieser Version nicht möglich.',
  aiNewChat: 'Neuer Chat',
  aiSend: 'Senden',
  aiStop: 'Stopp',
  aiHintIdle: 'Enter zum Senden, Umschalt+Enter für neue Zeile',
  aiHintBusy: 'Antwortet…',
  aiThinking: 'Denkt nach',
  aiReplying: 'Antwortet',
  aiWorking: 'Arbeitet',
  aiStopped: 'Gestoppt',
  aiNoReply: '(keine Antwort)',
  aiTruncatedNote:
    '(Die Antwort wurde durch das Längenlimit abgeschnitten und kann unvollständig sein.)',
  aiTurnLimit: 'Schrittlimit pro Lauf erreicht — senden Sie eine weitere Anweisung.',
  aiUndelivered: 'Nicht gesendet',
  aiRetry: 'Erneut versuchen',
  aiGskLoginBtn: 'Bei Genspark anmelden',
  aiUnknownError: 'KI-Anfrage fehlgeschlagen, bitte erneut versuchen',
  aiTimeoutError: 'Zeitüberschreitung der KI-Antwort',
  aiOverloadedError: 'Der KI-Dienst ist derzeit überlastet — bitte gleich erneut versuchen',
  aiNetworkError:
    'Netzwerkproblem: Der KI-Dienst ist nicht erreichbar. Prüfe deine Verbindung und versuche es erneut',
  aiCreditsExhausted: 'Guthaben aufgebraucht — bei genspark.ai aufladen',
} satisfies Record<keyof typeof zh, string>
