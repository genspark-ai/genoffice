import type { AnyExtension } from '@tiptap/core'
import StarterKit from '@tiptap/starter-kit'
import { Table, TableKit } from '@tiptap/extension-table'
import { OrderedList, TaskList } from '@tiptap/extension-list'
import { LooseBulletList, LooseOrderedList, LooseTaskList } from './looseLists'
import { CodeBlock } from '@tiptap/extension-code-block'
import { ReactNodeViewRenderer } from '@tiptap/react'
import { Placeholder } from '@tiptap/extensions'
import { CodeBlockView } from './CodeBlockView'
import { ImageAwareLink, LocalImage } from './localImage'
import { BlockDragHandle } from './blockDragHandle'
import { BlockKeymap } from './blockKeymap'
import { AiHighlight } from './aiHighlight'
import { AiQueueAnchors } from './aiQueueAnchors'
import { InactiveSelection } from './inactiveSelection'
import { SearchHighlight } from './searchHighlight'
import { buildMathExtensions } from './math'
import {
  BlockStartEscapedParagraph,
  SelectiveEscapeMarkdown,
  withTableRendering,
} from './markdownEscape'
import { SlashCommand } from './slashCommand'
import {
  CodeSpan,
  renderFencedCode,
  StyledBold,
  StyledHardBreak,
  StyledHeading,
  StyledHorizontalRule,
  StyledItalic,
  StyledListItem,
  StyledTaskItem,
  renderTable,
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
      // re-added below: the image is inline, the link parser knows about it and the
      // paragraph wraps lone images and escapes block-start syntax
      link: false,
      paragraph: false,
      // replaced by the NodeView-enhanced variant below (language picker + copy)
      codeBlock: false,
      // underline would serialize as `++text++` — not part of GFM
      underline: false,
      // re-added below with a linear-time markdown tokenizer and a `loose` attribute
      orderedList: false,
      bulletList: false,
      // re-added below with renderers that follow the document's own conventions
      bold: false,
      italic: false,
      heading: false,
      horizontalRule: false,
      hardBreak: false,
      listItem: false,
      // re-added below: code after the other marks so it serializes innermost
      code: false,
    }),
    BlockStartEscapedParagraph,
    StyledBold,
    StyledItalic,
    CodeSpan,
    StyledHeading,
    StyledHorizontalRule,
    StyledHardBreak,
    StyledListItem,
    LooseBulletList,
    LooseOrderedList.extend({
      markdownTokenizer: boundOrderedList(OrderedList.config.markdownTokenizer!),
    }),
    CodeBlock.extend({
      addNodeView() {
        return ReactNodeViewRenderer(CodeBlockView)
      },
      // the stock handler only takes a fence at column 0; CommonMark allows 1-3 spaces
      parseMarkdown: (token, h) => {
        if (
          !/^ {0,3}(?:`{3,}|~{3,})/.test(String(token.raw ?? '')) &&
          token.codeBlockStyle !== 'indented'
        ) {
          return []
        }
        return h.createNode(
          'codeBlock',
          { language: token.lang || null },
          token.text ? [h.createTextNode(String(token.text))] : [],
        )
      },
      renderMarkdown: (node, h) =>
        renderFencedCode(
          String(node.attrs?.language ?? ''),
          node.content ? h.renderChildren(node.content) : null,
        ),
    }),
    SelectiveEscapeMarkdown,
    // column widths are not expressible in GFM tables — no resizable columns;
    // the wrapper div gives wide tables a horizontal scrollbar
    TableKit.configure({ table: false }),
    Table.extend({
      markdownTokenizer: boundTable(Table.config.markdownTokenizer!),
      renderMarkdown: (node, h, ctx) => withTableRendering(() => renderTable(node, h, ctx)),
    }).configure({
      resizable: false,
      renderWrapper: true,
    }),
    LooseTaskList.extend({ markdownTokenizer: boundTaskList(TaskList.config.markdownTokenizer!) }),
    StyledTaskItem.configure({ nested: true }),
    // KaTeX-rendered $...$ / $$...$$ formulas (issue #100)
    ...buildMathExtensions(),
    LocalImage,
    // links open externally via main-process guard
    ImageAwareLink.configure({ openOnClick: false }),
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
