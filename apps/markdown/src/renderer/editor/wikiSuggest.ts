/**
 * The `[[` suggester. Typing the opening brackets offers the vault's notes
 * (the same breadth-first listing the resolver matches against), arrow keys
 * and Enter complete the link in place, `![[` completes the same list for an
 * embed. A query nothing matches still offers to create that note, and the
 * completion then carries the created file's stem — the spelling that
 * resolves. Completion inserts REAL wiki syntax — a marked run for links, an
 * embed node for `![[` — because live typing does not re-run the markdown
 * tokenizers: plain text would stay plain text until the next reparse.
 */
import { Extension } from '@tiptap/core'
import { PluginKey } from '@tiptap/pm/state'
import { Suggestion } from '@tiptap/suggestion'
import type { SuggestionProps } from '@tiptap/suggestion'
import type { Editor, Range } from '@tiptap/core'
import type { WikiCreateNoteResult, WikiNoteEntry } from '../../shared/ipc'

/** the query part after `[[` up to an alias `|` or a fragment `#` */
function queryNamePart(query: string): string {
  const cut = Math.min(
    ...[query.indexOf('|'), query.indexOf('#')].map((at) => (at < 0 ? query.length : at)),
  )
  return query.slice(0, cut)
}

/** what a created note would be named: the query's file part, trimmed */
function createNameOf(props: SuggestionProps<WikiNoteEntry>): string {
  return queryNamePart(props.query).trim()
}

/** the character just before `range`: the bang that makes a completion an embed */
function embedPrefixBefore(editor: Editor, range: Range): string {
  return editor.state.doc.textBetween(Math.max(0, range.from - 1), range.from, undefined, '\ufffc')
}

/** the file name without its directory or extension: `/vault/Folder/Note.md` -> `Note` */
function stemOf(path: string): string {
  const file = path.split(/[\\/]/).pop() ?? ''
  const dot = file.lastIndexOf('.')
  return dot > 0 ? file.slice(0, dot) : file
}

/**
 * Create the note the query names, then finish the link in place with the
 * created file's stem — the spelling that resolves, even where a sanitized
 * character changed the name. The vault listing is dropped so the next
 * keystroke offers the new note like any other.
 */
async function createNoteAndComplete(
  props: SuggestionProps<WikiNoteEntry>,
  embed: boolean,
): Promise<void> {
  // the browser test harness has no Electron channel: nothing to create with
  if (typeof window.markdownApi?.createWikiNote !== 'function') return
  const name = createNameOf(props)
  if (!name) return
  let result: WikiCreateNoteResult
  try {
    result = await window.markdownApi.createWikiNote({ target: name })
  } catch {
    return // an unavailable creator must not break typing
  }
  if (!result.ok || props.editor.isDestroyed) return
  invalidateWikiNotes()
  const open = embed ? '![[' : '[['
  const text = `${open}${stemOf(result.path)}${querySuffixPart(props.query)}]]`
  insertCompletion(props.editor, props.range, text, embed)
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

/**
 * A query the user already closed carries its own `]]`. Without this the
 * `allowSpaces` match keeps running to the end of the paragraph, so every
 * later keystroke re-opens an empty "No results" box that trails the caret.
 * (Embeds are unaffected — the wikiEmbed atom breaks the text match.)
 */
export function wikiQueryOpen(query: string): boolean {
  return !query.includes(']]')
}

/** the text an item completes to, keeping whatever alias/fragment was typed */
function completionFor(entry: WikiNoteEntry, query: string, embed: boolean): string {
  const open = embed ? '![[' : '[['
  const suffix = querySuffixPart(query)
  return `${open}${entry.name}${suffix}]]`
}

function insertCompletion(editor: Editor, range: Range, text: string, embed: boolean): void {
  if (embed) {
    // `range` covers `[[query` only; the bang that made this an embed sits just
    // before it. completionFor adds its own `!`, so deleting only `range` would
    // leave the typed one behind and save `!![[Beta]]`.
    const from = range.from - 1
    editor
      .chain()
      .focus()
      .deleteRange({ from, to: range.to })
      .insertContentAt(from, [{ type: 'wikiEmbed', attrs: { raw: text } }])
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
  /** the query's file part, '' when there is nothing a new note could be named */
  createName: string
  /** create the note `createName` names and complete the link in place */
  create: () => void
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

/** Forget the cached listing: a note just created must join the next suggestion */
export function invalidateWikiNotes(): void {
  cache.clear()
  cacheAt = 0
  refreshWikiNotes()
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
    /** whether the last render offered a create row; a key-down has no query either */
    let visibleCreateName = ''
    const toState = (props: SuggestionProps<WikiNoteEntry>): WikiSuggestState => {
      const embed = embedPrefixBefore(props.editor, props.range) === '!'
      return {
        items: props.items,
        clientRect: props.clientRect?.() ?? null,
        command: (item) => props.command(item),
        createName: createNameOf(props),
        create: () => void createNoteAndComplete(props, embed),
      }
    }
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
        // a query the user closed by hand is no longer a query: let the plugin
        // exit instead of following the caret with an empty box
        allow: ({ state, range }) =>
          wikiQueryOpen(state.doc.textBetween(range.from, range.to, ' ')),
        minQueryLength: 0,
        items: ({ query }) => {
          refreshWikiNotes()
          return wikiSuggestItems(query)
        },
        command: ({ editor, range, props: item }) => {
          const embed = embedPrefixBefore(editor, range) === '!'
          const query = editor.state.doc.textBetween(range.from, range.to, undefined, '\ufffc')
          insertCompletion(editor, range, completionFor(item, query, embed), embed)
        },
        render: () => ({
          onStart: (props) => {
            visibleItems = props.items
            visibleCreateName = createNameOf(props)
            getController()?.onOpen(toState(props))
          },
          onUpdate: (props) => {
            visibleItems = props.items
            visibleCreateName = createNameOf(props)
            getController()?.onUpdate(toState(props))
          },
          onKeyDown: (props) => {
            // the create row is pickable with nothing matched, so its keys
            // must not fall through to the document either
            if (visibleItems.length === 0 && !visibleCreateName) return false
            return getController()?.onKeyDown(props.event) ?? false
          },
          onExit: () => {
            visibleItems = []
            visibleCreateName = ''
            getController()?.onClose()
          },
        }),
      }),
    ]
  },
})
