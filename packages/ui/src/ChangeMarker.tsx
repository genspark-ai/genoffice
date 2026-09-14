/**
 * Visual anchor rendered on top of an AI-edited range so the user can see
 * "this was changed by AI" and undo with one click.
 *
 * Each app wraps the marker in its own editor-aware positioning layer
 * (Tiptap ProseMirror Decoration for docs, CellBox for sheets, Shape outline
 * for slides, PDF text-layer overlay for pdf). This component renders the
 * chrome (border + tooltip + undo button) and forwards `onUndo`.
 */

import React from 'react'
import type { ChatChangePlan } from '@genoffice/chat-runtime/types'

export interface ChangeMarkerProps {
  /** The change plan this marker represents (used to label the tooltip). */
  plan: ChatChangePlan
  /** Auto-expire window in ms (default 5 min). When elapsed the undo button is disabled. */
  windowMs?: number
  /** Localised strings. */
  strings: { original: string; translated: string; undo: string; undoFailed: string }
  /** Undo click handler. */
  onUndo: (plan: ChatChangePlan) => void
}

export function ChangeMarker(props: ChangeMarkerProps): React.JSX.Element {
  const { plan, windowMs = 5 * 60 * 1000, strings, onUndo } = props
  const firstOp = plan.ops[0]
  const isTranslate = firstOp?.kind === 'translate'
  const targetLang = isTranslate && firstOp.kind === 'translate' ? firstOp.ops[0]?.targetLang : ''
  const age = Date.now() - plan.createdAt
  const expired = age > windowMs

  const tipParts: string[] = []
  if (isTranslate) tipParts.push(`${strings.translated}${targetLang ? ` → ${targetLang}` : ''}`)
  else tipParts.push(plan.title)
  const tip = tipParts.join(' · ')

  return (
    <div className="ai-change-marker" data-plan-id={plan.id} data-kind={firstOp?.kind ?? 'freeform'}>
      <div className="ai-change-marker-stripe" aria-hidden />
      <div className="ai-change-marker-tip" role="tooltip">
        <span>{tip}</span>
        <button
          type="button"
          className="ai-change-marker-undo"
          onClick={() => {
            if (expired) return
            try {
              onUndo(plan)
            } catch {
              /* host shows toast via separate channel */
            }
          }}
          disabled={expired}
          title={expired ? strings.undoFailed : strings.undo}
        >
          {strings.undo}
        </button>
      </div>
    </div>
  )
}
