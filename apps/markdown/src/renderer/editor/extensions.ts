import { Node, mergeAttributes } from '@tiptap/core'
import StarterKit from '@tiptap/starter-kit'
import { Table } from '@tiptap/extension-table'
import { TableRow } from '@tiptap/extension-table-row'
import { TableCell } from '@tiptap/extension-table-cell'
import { TableHeader } from '@tiptap/extension-table-header'
import { Link } from '@tiptap/extension-link'
import { Image } from '@tiptap/extension-image'
import { TaskList } from '@tiptap/extension-task-list'
import { TaskItem } from '@tiptap/extension-task-item'
import { Placeholder } from '@tiptap/extension-placeholder'
import { TextAlign } from '@tiptap/extension-text-align'
import { Typography } from '@tiptap/extension-typography'
import { Highlight } from '@tiptap/extension-highlight'
import { Underline } from '@tiptap/extension-underline'

export interface FrontmatterOptions {
  HTMLAttributes: Record<string, unknown>
}

declare module '@tiptap/core' {
  interface Commands<ReturnType> {
    frontmatter: {
      setFrontmatter: (data: Record<string, unknown>) => ReturnType
      toggleFrontmatter: () => ReturnType
      removeFrontmatter: () => ReturnType
    }
  }
}

export const FrontmatterNode = Node.create<FrontmatterOptions>({
  name: 'frontmatter',
  group: 'block',
  atom: true,
  selectable: false,
  draggable: false,

  addOptions() {
    return { HTMLAttributes: {} }
  },

  addAttributes() {
    return {
      data: { default: {} },
    }
  },

  parseHTML() {
    return [{ tag: 'div[data-frontmatter]' }]
  },

  renderHTML({ node, HTMLAttributes }) {
    return [
      'div',
      mergeAttributes(this.options.HTMLAttributes, HTMLAttributes, {
        'data-frontmatter': '',
        class: 'markdown-frontmatter',
      }),
      JSON.stringify(node.attrs.data, null, 2),
    ]
  },

  addCommands() {
    return {
      setFrontmatter:
        (data) =>
        ({ commands }) => {
          return commands.insertContent({ type: this.name, attrs: { data } })
        },
      toggleFrontmatter:
        () =>
        ({ commands }) => {
          return commands.toggleNode(this.name, 'paragraph')
        },
      removeFrontmatter:
        () =>
        ({ commands }) => {
          return commands.deleteNode(this.name)
        },
    }
  },
})

// ── All extensions ───────────────────────────────────────────────────────────

export const markdownExtensions = [
  StarterKit.configure({
    heading: { levels: [1, 2, 3, 4, 5, 6] },
  }),
  FrontmatterNode,
  Table.configure({ resizable: true }),
  TableRow,
  TableCell,
  TableHeader,
  Link.configure({
    openOnClick: false,
    autolink: true,
    HTMLAttributes: { class: 'md-link' },
  }),
  Image.configure({
    HTMLAttributes: { class: 'md-image' },
  }),
  TaskList,
  TaskItem.configure({ nested: true }),
  Placeholder.configure({
    placeholder: 'Start writing...',
  }),
  TextAlign.configure({
    types: ['heading', 'paragraph'],
  }),
  Typography,
  Highlight.configure({ multicolor: false }),
  Underline,
]
