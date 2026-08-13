// Stub: slashCommand module
import type { Editor } from '@tiptap/react'

export interface SlashItem {
  label: string
  icon: string
  action: (editor: Editor) => void
}

export interface SlashController {
  items: SlashItem[]
  query: string
  filteredItems: SlashItem[]
  onKeyDown(event: KeyboardEvent): boolean
  selectItem(index: number): void
}

export interface SlashMenuState {
  open: boolean
  query: string
  items: SlashItem[]
}

export function buildSlashItems(): SlashItem[] {
  return []
}
