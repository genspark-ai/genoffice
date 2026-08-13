// Stub: SlashMenu component
import { forwardRef, useImperativeHandle } from 'react'
import type { SlashMenuState } from '../editor/slashCommand'

export interface SlashMenuHandle {
  onKeyDown(event: KeyboardEvent): boolean
}

export const SlashMenu = forwardRef(function SlashMenu(
  _props: { state: SlashMenuState; onSelect: (index: number) => void },
  ref: React.ForwardedRef<SlashMenuHandle>,
) {
  useImperativeHandle(ref, () => ({
    onKeyDown: () => false,
  }))
  return null
})
