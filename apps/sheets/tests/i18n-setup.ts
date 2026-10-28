import { loadStrings } from '../src/renderer/i18n/strings'

// Tests never run the renderer bootstrap that loads the locale chunks, so the
// default module language (zh) and the English fallback are made resident once
// per file, exactly as main.tsx does before the first render.
await loadStrings('zh')
await loadStrings('en')
