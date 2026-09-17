import type { AnyExtension } from '@tiptap/core'
import StarterKit from '@tiptap/starter-kit'
import { Table, TableKit } from '@tiptap/extension-table'
import { OrderedList, TaskList } from '@tiptap/extension-list'
import { CodeBlock } from '@tiptap/extension-code-block'
import { ReactNodeViewRenderer } from '@tiptap/react'
import { Placeholder } from '@tiptap/extensions'
import { CodeBlockView } from './CodeBlockView'
import { LocalImage } from './localImage'
import { BlockDragHandle } from './blockDragHandle'
import { BlockKeymap } from './blockKeymap'
import { AiHighlight } from './aiHighlight'
import { AiQueueAnchors } from './aiQueueAnchors'
import { InactiveSelection } from './inactiveSelection'
import { SearchHighlight } from './searchHighlight'
import { buildMathExtensions } from './math'
import { SelectiveEscapeMarkdown } from './markdownEscape'
import { SlashCommand } from './slashCommand'
import {
  renderFencedCode,
  StyledBold,
  StyledHardBreak,
  StyledHeading,
  StyledHorizontalRule,
  StyledItalic,
  StyledListItem,
  StyledTaskItem,
} from './markdownStyleRenderers'
import { boundOrderedList, boundTable, boundTaskList } from './boundedTokenizers'
import type { SlashController, SlashItem } from './slashCommand'
import { t } from '../i18n/locale'

export interface BuildExtensionsOptions {
  slashController: SlashController
  slashItems: () => SlashItem[]
}

export function buildExtensions(options: BuildExtensionsOptions): AnyExtension[] {
  return [
    StarterKit.configure({
      // LocalImage replaces the plain image; links open externally via main-process guard
      link: { openOnClick: false },
      // replaced by the NodeView-enhanced variant below (language picker + copy)
      codeBlock: false,
      // underline would serialize as `++text++` — not part of GFM
      underline: false,
      // re-added below with a linear-time markdown tokenizer
      orderedList: false,
      // re-added below with renderers that follow the document's own conventions
      bold: false,
      italic: false,
      heading: false,
      horizontalRule: false,
      hardBreak: false,
      listItem: false,
    }),
    StyledBold,
    StyledItalic,
    StyledHeading,
    StyledHorizontalRule,
    StyledHardBreak,
    StyledListItem,
    OrderedList.extend({
      markdownTokenizer: boundOrderedList(OrderedList.config.markdownTokenizer!),
    }),
    CodeBlock.extend({
      addNodeView() {
        return ReactNodeViewRenderer(CodeBlockView)
      },
      renderMarkdown: (node, h) =>
        renderFencedCode(
          String(node.attrs?.language ?? ''),
          node.content ? h.renderChildren(node.content) : null,
        ),
    }),
    // 4-space nesting: the default 2 spaces is below the content column of
    // ordered items ("1. " = 3), so strict CommonMark parsers (GitHub) would
    // flatten sub-lists in the saved file. 4 is safe for every marker width.
    SelectiveEscapeMarkdown.configure({ indentation: { style: 'space', size: 4 } }),
    // column widths are not expressible in GFM tables — no resizable columns;
    // the wrapper div gives wide tables a horizontal scrollbar
    TableKit.configure({ table: false }),
    Table.extend({ markdownTokenizer: boundTable(Table.config.markdownTokenizer!) }).configure({
      resizable: false,
      renderWrapper: true,
    }),
    TaskList.extend({ markdownTokenizer: boundTaskList(TaskList.config.markdownTokenizer!) }),
    StyledTaskItem.configure({ nested: true }),
    // KaTeX-rendered $...$ / $$...$$ formulas (issue #100)
    ...buildMathExtensions(),
    LocalImage,
    BlockDragHandle,
    BlockKeymap,
    AiHighlight,
    AiQueueAnchors,
    InactiveSelection,
    SearchHighlight,
    Placeholder.configure({ placeholder: () => t('placeholder') }),
    SlashCommand.configure({
      controller: options.slashController,
      items: options.slashItems,
    }),
  ]
}
