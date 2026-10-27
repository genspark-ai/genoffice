# Voice input is a composer slot, not part of the composer

The composer is shared by six apps and hand-copied by a seventh, so it is the most
contended component in the renderer. We add a single `footerEnd` slot to `AiComposer`
and ship the microphone as a separate `AiVoiceButton` driven by a `useVoiceInput`
hook in `@genoffice/ui`, rather than building recording state into the composer
itself. The composer stays a controlled textarea with one optional extra child; an
app that does not adopt voice renders exactly what it renders today.

## Considered Options

- **Build voice into `AiComposer` behind a prop.** Centralised, but grows a shared
  component that six apps depend on with audio state, permissions and 22 locales of
  strings, and it deepens the divergence from the slides clone.
- **Reuse the existing `footerStart` slot.** Zero composer change, but `footerStart`
  is left-aligned: the microphone would sit left of the model picker instead of next
  to send, where the product wants it.

## Consequences

- The slot is conditional: without `footerEnd`, `AiComposer` renders its send button
  exactly as before, so apps adopt voice one at a time.
- Slides carries a hand-written copy of the composer and needs its own edit; it does
  not inherit this.
