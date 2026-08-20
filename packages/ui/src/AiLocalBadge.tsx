/** Subtle \"local AI\" chip for the AI panel header: shown while the active
 *  provider is Ollama so the user always sees their requests stay local.
 *  The provider label is a brand name (like the Genspark mark already in the
 *  panel headers), so it is not localized. */
export function AiLocalBadge({ model }: { model?: string }) {
  return (
    <span className="ai-local-badge">
      Ollama{model ? ` · ${model}` : ''}
    </span>
  )
}
