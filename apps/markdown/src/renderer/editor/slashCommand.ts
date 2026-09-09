import { Extension } from '@tiptap/core'
import type { Editor, Range } from '@tiptap/core'
import { Suggestion } from '@tiptap/suggestion'
import type { SuggestionProps } from '@tiptap/suggestion'
import type { StringKey } from '../i18n/locale'
import { t } from '../i18n/locale'
import { openMathCreate } from './mathEdit'
import { MERMAID_TEMPLATE } from './mermaid'
import { uiOp, type MdOp } from './ops'

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

/** drop the typed "/query", then run the op on the caret's block */
function slashOp(op: MdOp) {
  return (editor: Editor, range: Range): void => {
    editor.chain().focus().deleteRange(range).run()
    uiOp(editor, op)
  }
}

function clear(editor: Editor, range: Range): void {
  editor.chain().focus().deleteRange(range).run()
}

export function buildSlashItems(extra?: { insertImage?: () => void }): SlashItem[] {
  const items: SlashItem[] = [
    {
      id: 'paragraph',
      labelKey: 'styleParagraph',
      keywords: ['text', 'p'],
      run: slashOp({ op: 'setBlockType', target: 'selection', type: 'paragraph' }),
    },
    {
      id: 'h1',
      labelKey: 'styleH1',
      keywords: ['heading', '#'],
      run: slashOp({ op: 'setBlockType', target: 'selection', type: 'heading', level: 1 }),
    },
    {
      id: 'h2',
      labelKey: 'styleH2',
      keywords: ['heading', '##'],
      run: slashOp({ op: 'setBlockType', target: 'selection', type: 'heading', level: 2 }),
    },
    {
      id: 'h3',
      labelKey: 'styleH3',
      keywords: ['heading', '###'],
      run: slashOp({ op: 'setBlockType', target: 'selection', type: 'heading', level: 3 }),
    },
    {
      id: 'bullet',
      labelKey: 'bulletList',
      keywords: ['list', 'ul', '-'],
      run: slashOp({ op: 'toggleList', target: 'selection', list: 'bullet' }),
    },
    {
      id: 'ordered',
      labelKey: 'orderedList',
      keywords: ['list', 'ol', '1.'],
      run: slashOp({ op: 'toggleList', target: 'selection', list: 'ordered' }),
    },
    {
      id: 'task',
      labelKey: 'taskList',
      keywords: ['todo', 'checkbox', '[]'],
      run: slashOp({ op: 'toggleList', target: 'selection', list: 'task' }),
    },
    {
      id: 'quote',
      labelKey: 'styleQuote',
      keywords: ['blockquote', '>'],
      run: slashOp({ op: 'setBlockType', target: 'selection', type: 'blockquote' }),
    },
    {
      id: 'code',
      labelKey: 'styleCodeBlock',
      keywords: ['codeblock', '```'],
      run: slashOp({ op: 'setBlockType', target: 'selection', type: 'codeBlock' }),
    },
    {
      id: 'table',
      labelKey: 'insertTable',
      keywords: ['grid'],
      run: slashOp({ op: 'insertTable', after: 'selection' }),
    },
    {
      id: 'hr',
      labelKey: 'insertHr',
      keywords: ['divider', 'rule', '---'],
      run: slashOp({ op: 'insertHorizontalRule', after: 'selection' }),
    },
    {
      id: 'math',
      labelKey: 'insertMath',
      keywords: ['formula', 'equation', 'latex', 'katex', '$$'],
      run: (e, r) => {
        clear(e, r)
        openMathCreate(e)
      },
    },
    {
      id: 'diagram',
      labelKey: 'insertDiagram',
      keywords: ['mermaid', 'chart', 'flowchart', 'graph'],
      run: slashOp({ op: 'insertContent', after: 'selection', markdown: MERMAID_TEMPLATE }),
    },
  ]
  if (extra?.insertImage) {
    items.push({
      id: 'image',
      labelKey: 'insertImage',
      keywords: ['picture', 'img', 'photo'],
      run: (e, r) => {
        clear(e, r)
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
