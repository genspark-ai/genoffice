/**
 * The `[[` suggester. Typing the opening brackets offers the vault's notes
 * (the same breadth-first listing the resolver matches against), arrow keys
 * and Enter complete the link in place, `![[` completes the same list for an
 * embed. Completion inserts REAL wiki syntax — a marked run for links, an
 * embed node for `![[` — because live typing does not re-run the markdown
 * tokenizers: plain text would stay plain text until the next reparse.
 */
import { Extension } from '@tiptap/core'
import { PluginKey } from '@tiptap/pm/state'
import { Suggestion } from '@tiptap/suggestion'
import type { SuggestionProps } from '@tiptap/suggestion'
import type { Editor, Range } from '@tiptap/core'
import type { WikiNoteEntry } from '../../shared/ipc'

/** the query part after `[[` up to an alias `|` or a fragment `#` */
function queryNamePart(query: string): string {
  const cut = Math.min(
    ...[query.indexOf('|'), query.indexOf('#')].map((at) => (at < 0 ? query.length : at)),
  )
  return query.slice(0, cut)
}

/** what the completion keeps of the query after the name: `|alias`, `#frag` */
function querySuffixPart(query: string): string {
  const name = queryNamePart(query)
  const suffix = query.slice(name.length)
  // a query the user closed by hand carries the brackets; the completion adds
  // its own, so they are dropped from the suffix (and the name)
  return suffix.replace(/\]\]/, '')
}

export function filterWikiNotes(entries: WikiNoteEntry[], query: string): WikiNoteEntry[] {
  const q = queryNamePart(query).trim().toLowerCase()
  if (!q) return entries
  return entries.filter(
    (entry) => entry.name.toLowerCase().includes(q) || entry.file.toLowerCase().includes(q),
  )
}

/** the text an item completes to, keeping whatever alias/fragment was typed */
function completionFor(entry: WikiNoteEntry, query: string, embed: boolean): string {
  const open = embed ? '![[' : '[['
  const suffix = querySuffixPart(query)
  return `${open}${entry.name}${suffix}]]`
}

function insertCompletion(editor: Editor, range: Range, text: string, embed: boolean): void {
  if (embed) {
    editor
      .chain()
      .focus()
      .deleteRange(range)
      .insertContentAt(range.from, [{ type: 'wikiEmbed', attrs: { raw: text } }])
      .run()
    return
  }
  editor
    .chain()
    .focus()
    .deleteRange(range)
    .insertContentAt(range.from, [{ type: 'text', text, marks: [{ type: 'wikiLink' }] }])
    .run()
}

export interface WikiSuggestState {
  items: WikiNoteEntry[]
  clientRect: DOMRect | null
  command: (item: WikiNoteEntry) => void
}

/** App-side sink for the suggestion lifecycle; the React menu renders from it */
export interface WikiSuggestController {
  onOpen(state: WikiSuggestState): void
  onUpdate(state: WikiSuggestState): void
  onKeyDown(event: KeyboardEvent): boolean
  onClose(): void
}

/**
 * The vault listing, fetched once per editor and refreshed on demand. The
 * suggester's `items` callback is synchronous, so it reads this cache; a
 * refresh started at trigger time lands for the next keystroke.
 */
const cache = new Map<string, WikiNoteEntry>()
let cacheAt = 0
let inflight: Promise<void> | null = null

const CACHE_TTL_MS = 15_000

export function refreshWikiNotes(): void {
  if (inflight) return
  if (cache.size > 0 && Date.now() - cacheAt < CACHE_TTL_MS) return
  // The browser test harness has no Electron channel; an unlistable vault
  // just means an empty suggestion list.
  if (typeof window.markdownApi?.listWikiNotes !== 'function') return
  inflight = window.markdownApi
    .listWikiNotes()
    .then((entries) => {
      cache.clear()
      for (const entry of entries) cache.set(entry.path, entry)
      cacheAt = Date.now()
    })
    .catch(() => undefined)
    .finally(() => {
      inflight = null
    })
}

export function wikiSuggestItems(query: string): WikiNoteEntry[] {
  return filterWikiNotes([...cache.values()], query)
}

interface WikiSuggestOptions {
  controller: WikiSuggestController | null
}

export const WikiSuggest = Extension.create<WikiSuggestOptions>({
  name: 'wikiSuggest',

  addOptions() {
    return { controller: null }
  },

  onBeforeCreate() {
    refreshWikiNotes()
  },

  addProseMirrorPlugins() {
    const getController = () => this.options.controller
    // the key-down props carry no items; the render callbacks last saw them
    let visibleItems: WikiNoteEntry[] = []
    const toState = (props: SuggestionProps<WikiNoteEntry>): WikiSuggestState => ({
      items: props.items,
      clientRect: props.clientRect?.() ?? null,
      command: (item) => props.command(item),
    })
    return [
      Suggestion<WikiNoteEntry>({
        editor: this.editor,
        // the suggester must not swallow arrow keys or Enter when there is
        // nothing to pick (the query closed itself with `]]`)
        char: '[[',
        // a key of its own: the slash suggester shares this editor, and two
        // plugins on the default key cannot coexist in one state
        pluginKey: new PluginKey('wikiSuggest'),
        // the char before the brackets: a space (mid-sentence) or the bang that
        // turns the completion into an embed — a letter prefix stays plain text
        allowedPrefixes: ['!', ' '],
        allowSpaces: true,
        minQueryLength: 0,
        items: ({ query }) => {
          refreshWikiNotes()
          return wikiSuggestItems(query)
        },
        command: ({ editor, range, props: item }) => {
          const before = editor.state.doc.textBetween(
            Math.max(0, range.from - 1),
            range.from,
            undefined,
            '\ufffc',
          )
          const embed = before === '!'
          const query = editor.state.doc.textBetween(range.from, range.to, undefined, '\ufffc')
          insertCompletion(editor, range, completionFor(item, query, embed), embed)
        },
        render: () => ({
          onStart: (props) => {
            visibleItems = props.items
            getController()?.onOpen(toState(props))
          },
          onUpdate: (props) => {
            visibleItems = props.items
            getController()?.onUpdate(toState(props))
          },
          onKeyDown: (props) => {
            if (visibleItems.length === 0) return false
            return getController()?.onKeyDown(props.event) ?? false
          },
          onExit: () => {
            visibleItems = []
            getController()?.onClose()
          },
        }),
      }),
    ]
  },
})
