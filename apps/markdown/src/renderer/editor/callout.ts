import { Node, createBlockMarkdownSpec, mergeAttributes } from '@tiptap/core'

export const CALLOUT_TYPES = ['info', 'tip', 'warning', 'danger'] as const
export type CalloutType = (typeof CALLOUT_TYPES)[number]

declare module '@tiptap/core' {
  interface Commands<ReturnType> {
    callout: {
      /** wrap the selection in a callout (no-op when already inside one) */
      setCallout: (attrs?: { type?: CalloutType }) => ReturnType
      /** unwrap the callout around the selection */
      unsetCallout: () => ReturnType
    }
  }
}

/**
 * Admonition block serialized as a Pandoc-style fenced div:
 *
 *   :::callout {type="warning"}
 *   body…
 *   :::
 */
export const Callout = Node.create({
  name: 'callout',
  group: 'block',
  content: 'block+',
  defining: true,

  addAttributes() {
    return {
      type: {
        default: 'info',
        parseHTML: (el) => el.getAttribute('data-type') || 'info',
        renderHTML: (attrs) => ({ 'data-type': attrs.type }),
      },
    }
  },

  parseHTML() {
    return [{ tag: 'div[data-callout]' }]
  },

  renderHTML({ HTMLAttributes }) {
    return ['div', mergeAttributes(HTMLAttributes, { 'data-callout': '', class: 'md-callout' }), 0]
  },

  addCommands() {
    return {
      setCallout:
        (attrs) =>
        ({ commands, editor }) => {
          if (editor.isActive(this.name)) return true
          return commands.wrapIn(this.name, { type: attrs?.type ?? 'info' })
        },
      unsetCallout:
        () =>
        ({ commands }) =>
          commands.lift(this.name),
    }
  },

  ...createBlockMarkdownSpec({
    nodeName: 'callout',
    defaultAttributes: { type: 'info' },
    allowedAttributes: ['type'],
    serializeAttributes: (attrs) =>
      attrs.type && attrs.type !== 'info' ? `type="${attrs.type}"` : '',
    content: 'block',
  }),
})
