// Stub: AiPanel component
import type { Editor } from '@tiptap/react'

export interface AiPreset {
  label: string
  instruction: string
}

export interface MarkdownAiDeps {
  editor: Editor
  // placeholder
}

export function GensparkMark() {
  return null
}

export function AiPanel(_props: MarkdownAiDeps & { presets: AiPreset[] }): null {
  return null
}
