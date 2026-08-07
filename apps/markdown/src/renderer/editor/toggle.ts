import { Node, createBlockMarkdownSpec, mergeAttributes } from '@tiptap/core'
import { ReactNodeViewRenderer } from '@tiptap/react'
import { ToggleView } from './ToggleView'

declare module '@tiptap/core' {
  interface Commands<ReturnType> {
    toggleBlock: {
      /** wrap the selection in a collapsible toggle block */
      setToggle: (attrs?: { summary?: string }) => ReturnType
      /** unwrap the toggle block around the selection */
      unsetToggle: () => ReturnType
    }
  }
}

/**
 * Collapsible block serialized as a fenced div:
 *
 *   :::toggle {summary="Section title"}
 *   hidden body…
 *   :::
 *
 * `open` is editor-only display state and never reaches the markdown.
 */
export const Toggle = Node.create({
  name: 'toggle',
  group: 'block',
  content: 'block+',
  defining: true,

  addAttributes() {
    return {
      summary: {
        default: '',
        parseHTML: (el) => el.getAttribute('data-summary') ?? '',
        renderHTML: (attrs) => ({ 'data-summary': attrs.summary }),
      },
      open: {
        default: true,
        parseHTML: (el) => el.getAttribute('data-open') !== 'false',
        renderHTML: (attrs) => ({ 'data-open': attrs.open ? 'true' : 'false' }),
      },
    }
  },

  parseHTML() {
    return [{ tag: 'div[data-toggle]' }]
  },

  renderHTML({ HTMLAttributes }) {
    return ['div', mergeAttributes(HTMLAttributes, { 'data-toggle': '', class: 'md-toggle' }), 0]
  },

  addNodeView() {
    return ReactNodeViewRenderer(ToggleView)
  },

  addCommands() {
    return {
      setToggle:
        (attrs) =>
        ({ commands, editor }) => {
          if (editor.isActive(this.name)) return true
          return commands.wrapIn(this.name, { summary: attrs?.summary ?? '', open: true })
        },
      unsetToggle:
        () =>
        ({ commands }) =>
          commands.lift(this.name),
    }
  },

  ...createBlockMarkdownSpec({
    nodeName: 'toggle',
    defaultAttributes: { summary: '', open: true },
    allowedAttributes: ['summary'],
    serializeAttributes: (attrs) =>
      attrs.summary ? `summary="${String(attrs.summary).replace(/"/g, '\\"')}"` : '',
    content: 'block',
  }),
})
