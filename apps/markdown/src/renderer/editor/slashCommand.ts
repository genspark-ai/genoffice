import { Extension } from '@tiptap/core'
import type { Editor, Range } from '@tiptap/core'
import { Suggestion } from '@tiptap/suggestion'
import type { SuggestionProps } from '@tiptap/suggestion'
import type { StringKey } from '../i18n/locale'
import { t } from '../i18n/locale'

export interface SlashItem {
  id: string
  labelKey: StringKey
  /** extra match terms besides id and the localized label */
  keywords: string[]
  run: (editor: Editor, range: Range) => void
}

export interface SlashMenuState {
  items: SlashItem[]
  clientRect: DOMRect | null
  command: (item: SlashItem) => void
}

/** App-side sink for the suggestion lifecycle; the React menu renders from it */
export interface SlashController {
  onOpen(state: SlashMenuState): void
  onUpdate(state: SlashMenuState): void
  onKeyDown(event: KeyboardEvent): boolean
  onClose(): void
}

function chain(editor: Editor, range: Range) {
  return editor.chain().focus().deleteRange(range)
}

/**
 * Block-type conversions (heading/callout/toggle/…) are illegal inside a list
 * item — wrapIn/setHeading would silently fail after the range was already
 * deleted. Lift the current item out of its list(s) first (Notion behavior).
 */
export function liftFromList(editor: Editor): void {
  for (let guard = 0; guard < 10; guard++) {
    const itemName = editor.isActive('taskItem')
      ? 'taskItem'
      : editor.isActive('listItem')
        ? 'listItem'
        : null
    if (!itemName) return
    if (!editor.chain().focus().liftListItem(itemName).run()) return
  }
}

function blockChain(editor: Editor, range: Range) {
  editor.chain().focus().deleteRange(range).run()
  liftFromList(editor)
  return editor.chain().focus()
}

export function buildSlashItems(extra?: { insertImage?: () => void }): SlashItem[] {
  const items: SlashItem[] = [
    {
      id: 'paragraph',
      labelKey: 'styleParagraph',
      keywords: ['text', 'p'],
      run: (e, r) => void blockChain(e, r).setParagraph().run(),
    },
    {
      id: 'h1',
      labelKey: 'styleH1',
      keywords: ['heading', '#'],
      run: (e, r) => void blockChain(e, r).setHeading({ level: 1 }).run(),
    },
    {
      id: 'h2',
      labelKey: 'styleH2',
      keywords: ['heading', '##'],
      run: (e, r) => void blockChain(e, r).setHeading({ level: 2 }).run(),
    },
    {
      id: 'h3',
      labelKey: 'styleH3',
      keywords: ['heading', '###'],
      run: (e, r) => void blockChain(e, r).setHeading({ level: 3 }).run(),
    },
    {
      id: 'bullet',
      labelKey: 'bulletList',
      keywords: ['list', 'ul', '-'],
      run: (e, r) => void chain(e, r).toggleBulletList().run(),
    },
    {
      id: 'ordered',
      labelKey: 'orderedList',
      keywords: ['list', 'ol', '1.'],
      run: (e, r) => void chain(e, r).toggleOrderedList().run(),
    },
    {
      id: 'task',
      labelKey: 'taskList',
      keywords: ['todo', 'checkbox', '[]'],
      run: (e, r) => void chain(e, r).toggleTaskList().run(),
    },
    {
      id: 'quote',
      labelKey: 'styleQuote',
      keywords: ['blockquote', '>'],
      run: (e, r) => void blockChain(e, r).toggleBlockquote().run(),
    },
    {
      id: 'code',
      labelKey: 'styleCodeBlock',
      keywords: ['codeblock', '```'],
      run: (e, r) => void blockChain(e, r).toggleCodeBlock().run(),
    },
    {
      id: 'callout',
      labelKey: 'insertCallout',
      keywords: ['admonition', 'note', 'info'],
      run: (e, r) => void blockChain(e, r).setCallout().run(),
    },
    {
      id: 'toggle',
      labelKey: 'insertToggle',
      keywords: ['details', 'collapse', 'fold'],
      run: (e, r) => void blockChain(e, r).setToggle().run(),
    },
    {
      id: 'table',
      labelKey: 'insertTable',
      keywords: ['grid'],
      run: (e, r) =>
        void blockChain(e, r).insertTable({ rows: 3, cols: 3, withHeaderRow: true }).run(),
    },
    {
      id: 'hr',
      labelKey: 'insertHr',
      keywords: ['divider', 'rule', '---'],
      run: (e, r) => void blockChain(e, r).setHorizontalRule().run(),
    },
  ]
  if (extra?.insertImage) {
    items.push({
      id: 'image',
      labelKey: 'insertImage',
      keywords: ['picture', 'img', 'photo'],
      run: (e, r) => {
        chain(e, r).run()
        extra.insertImage!()
      },
    })
  }
  return items
}

export function filterSlashItems(items: SlashItem[], query: string): SlashItem[] {
  const q = query.trim().toLowerCase()
  if (!q) return items
  return items.filter(
    (item) =>
      item.id.includes(q) ||
      t(item.labelKey).toLowerCase().includes(q) ||
      item.keywords.some((k) => k.includes(q)),
  )
}

interface SlashOptions {
  controller: SlashController | null
  items: () => SlashItem[]
}

export const SlashCommand = Extension.create<SlashOptions>({
  name: 'slashCommand',

  addOptions() {
    return { controller: null, items: () => [] }
  },

  addProseMirrorPlugins() {
    const getController = () => this.options.controller
    const toState = (props: SuggestionProps<SlashItem>): SlashMenuState => ({
      items: props.items,
      clientRect: props.clientRect?.() ?? null,
      command: (item) => props.command(item),
    })
    return [
      Suggestion<SlashItem>({
        editor: this.editor,
        char: '/',
        pluginKey: undefined,
        allowSpaces: false,
        command: ({ editor, range, props }) => props.run(editor, range),
        items: ({ query }) => filterSlashItems(this.options.items(), query),
        render: () => ({
          onStart: (props) => getController()?.onOpen(toState(props)),
          onUpdate: (props) => getController()?.onUpdate(toState(props)),
          onKeyDown: (props) => getController()?.onKeyDown(props.event) ?? false,
          onExit: () => getController()?.onClose(),
        }),
      }),
    ]
  },
})
