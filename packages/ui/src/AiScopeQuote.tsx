import React from 'react'

export interface AiScopeQuoteData {
  label: string
  text?: string
}

/** Quote block on a user bubble: the document selection the message targeted */
export function AiScopeQuote({ scope }: { scope: AiScopeQuoteData }): React.JSX.Element {
  return (
    <div className="ai-msg-scope">
      <span className="ai-msg-scope-label">{scope.label}</span>
      {scope.text && (
        <div className="ai-msg-scope-text" dir="auto">
          {scope.text}
        </div>
      )}
    </div>
  )
}
