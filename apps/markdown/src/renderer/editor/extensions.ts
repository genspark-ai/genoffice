import type { AnyExtension } from '@tiptap/core'
import StarterKit from '@tiptap/starter-kit'
import { Markdown } from '@tiptap/markdown'
import { TableKit } from '@tiptap/extension-table'
import { TaskItem, TaskList } from '@tiptap/extension-list'
import { Highlight } from '@tiptap/extension-highlight'
import { CodeBlock } from '@tiptap/extension-code-block'
import { ReactNodeViewRenderer } from '@tiptap/react'
import { Placeholder } from '@tiptap/extensions'
import { CodeBlockView } from './CodeBlockView'
import { Callout } from './callout'
import { Toggle } from './toggle'
import { LocalImage } from './localImage'
import { BlockDragHandle } from './blockDragHandle'
import { BlockKeymap } from './blockKeymap'
import { AiHighlight } from './aiHighlight'
import { SlashCommand } from './slashCommand'
import type { SlashController, SlashItem } from './slashCommand'
import { t } from '../i18n/locale'

export interface BuildExtensionsOptions {
  slashController: SlashController
  slashItems: () => SlashItem[]
}

export function buildExtensions(options: BuildExtensionsOptions): AnyExtension[] {
  return [
    // StarterKit v3 bundles bold/italic/strike/code/UNDERLINE/link/lists/history —
    // toggleUnderline etc. need no extra extension package
    StarterKit.configure({
      // LocalImage replaces the plain image; links open externally via main-process guard
      link: { openOnClick: false },
      // replaced by the NodeView-enhanced variant below (language picker + copy)
      codeBlock: false,
    }),
    CodeBlock.extend({
      addNodeView() {
        return ReactNodeViewRenderer(CodeBlockView)
      },
    }),
    Highlight,
    Markdown,
    TableKit.configure({ table: { resizable: false } }),
    TaskList,
    TaskItem.configure({ nested: true }),
    LocalImage,
    Callout,
    Toggle,
    BlockDragHandle,
    BlockKeymap,
    AiHighlight,
    Placeholder.configure({ placeholder: () => t('placeholder') }),
    SlashCommand.configure({
      controller: options.slashController,
      items: options.slashItems,
    }),
  ]
}
