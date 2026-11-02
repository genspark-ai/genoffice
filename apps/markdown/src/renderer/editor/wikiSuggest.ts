/**
 * The `[[` suggester. Typing the opening brackets offers the vault's notes
 * (the same breadth-first listing the resolver matches against), arrow keys
 * and Enter complete the link in place, `![[` completes the same list for an
 * embed. A query nothing matches still offers to create that note, and the
 * completion then carries the created file's own target — the spelling
 * that resolves. Completion inserts REAL wiki syntax — a marked run for links, an
 * embed node for `![[` — because live typing does not re-run the markdown
 * tokenizers: plain text would stay plain text until the next reparse.
 *
 * A `#` in the query switches to the fragment list, as in Obsidian: the
 * target note's headings, then its `^block` ids once the fragment starts with
 * a caret. `[[#` lists this document's own. Reading another note's heads is
 * the main process's job (it resolves the name and parses the file); the list
 * is what is there — nothing is ever written into the target.
 */
import { Extension } from '@tiptap/core'
import { PluginKey } from '@tiptap/pm/state'
import { Suggestion } from '@tiptap/suggestion'
import type { SuggestionProps } from '@tiptap/suggestion'
import type { Editor, Range } from '@tiptap/core'
import type { Node as PmNode } from '@tiptap/pm/model'
import type { WikiCreateNoteResult, WikiHeadEntry, WikiNoteEntry } from '../../shared/ipc'

/** the query part after `[[` up to an alias `|` or a fragment `#` */
function queryNamePart(query: string): string {
  const cut = Math.min(
    ...[query.indexOf('|'), query.indexOf('#')].map((at) => (at < 0 ? query.length : at)),
  )
  return query.slice(0, cut)
}

/** the query part after the first `#`, '' while the query has none; a typed
 *  alias is not part of the fragment filter (`[[Note#Head|alias` narrows on `Head`) */
function queryFragmentPart(query: string): string {
  const at = query.indexOf('#')
  if (at < 0) return ''
  const rest = query.slice(at + 1)
  const bar = rest.indexOf('|')
  return bar < 0 ? rest : rest.slice(0, bar)
}

/** one row the suggester offers: a note, or — after a `#` — a heading or block id */
export type WikiSuggestItem =
  | { kind: 'note'; entry: WikiNoteEntry }
  | { kind: 'heading'; text: string; level: number }
  | { kind: 'block'; id: string; text: string }

/** what a created note would be named: the query's file part, trimmed */
function createNameOf(props: SuggestionProps<WikiSuggestItem>): string {
  return queryNamePart(props.query).trim()
}

const TRAILING_ID_RE = /(?:^|\s)\^([A-Za-z0-9-]+)\s*$/

/** a heading whose text carries link syntax cannot sit in a `[[Note#title]]` fragment */
const addressableHeading = (text: string): boolean => !/[[\]|]/.test(text)

/**
 * This document's own headings and block ids, for `[[#`. No IPC: the open
 * document is the state the editor is holding.
 */
export function docFragmentItems(doc: PmNode, fragmentQuery: string): WikiSuggestItem[] {
  const q = fragmentQuery.replace(/^\^/, '').trim().toLowerCase()
  const wantBlocks = fragmentQuery.startsWith('^')
  const out: WikiSuggestItem[] = []
  doc.descendants((node) => {
    if (node.type.name === 'heading' && !wantBlocks) {
      // a trailing `^id` is the heading's own block marker, not its text —
      // the main-side parser strips it too, and a completion must not bake
      // `[[#Title ^id]]` into the link
      const text = node.textContent
        .replace(/\s+/g, ' ')
        .replace(/\s*\^[A-Za-z0-9-]+\s*$/, '')
        .trim()
      if (text && addressableHeading(text) && (!q || text.toLowerCase().includes(q))) {
        out.push({ kind: 'heading', text, level: Number(node.attrs.level ?? 1) })
      }
    }
    // code is code: a `^id` ending a code block is what the author typed
    if (wantBlocks && node.isTextblock && node.type.name !== 'codeBlock') {
      const match = TRAILING_ID_RE.exec(node.textContent)
      if (match) {
        const id = match[1]!
        const text = node.textContent.slice(0, match.index).trim()
        if (!q || id.toLowerCase().includes(q) || text.toLowerCase().includes(q)) {
          out.push({ kind: 'block', id, text })
        }
      }
    }
    return true
  })
  return out.slice(0, 50)
}

/**
 * Another note's heads, fetched through main and cached per target — the
 * suggester's `items` runs per keystroke, and a root walk per keystroke would
 * be the resolver's cost all over again.
 */
const headsCache = new Map<string, Promise<WikiHeadEntry[]>>()

function fetchWikiHeads(getPath: () => string, name: string): Promise<WikiHeadEntry[]> {
  const key = `${getPath()}\u0000${name}`
  const cached = headsCache.get(key)
  if (cached) return cached
  let pending: Promise<WikiHeadEntry[]>
  // The browser test harness has no Electron channel; an unreadable target
  // just means an empty fragment list.
  if (typeof window.markdownApi?.wikiHeads !== 'function') {
    pending = Promise.resolve([])
  } else {
    pending = window.markdownApi
      .wikiHeads({ target: name, fromPath: getPath() })
      .catch(() => [] as WikiHeadEntry[])
  }
  headsCache.set(key, pending)
  return pending
}

/** Forget cached heads: a note just saved may have grown headings or ids */
export function invalidateWikiHeads(): void {
  headsCache.clear()
}

/** narrow fetched heads by what was typed after the `#` */
export function filterWikiHeads(
  entries: WikiHeadEntry[],
  fragmentQuery: string,
): WikiSuggestItem[] {
  const q = fragmentQuery.replace(/^\^/, '').trim().toLowerCase()
  const wantBlocks = fragmentQuery.startsWith('^')
  return entries
    .filter((entry) => (entry.kind === 'heading' ? !wantBlocks : wantBlocks))
    .filter((entry) => entry.kind !== 'heading' || addressableHeading(entry.text))
    .filter((entry) => {
      if (!q) return true
      return entry.kind === 'heading'
        ? entry.text.toLowerCase().includes(q)
        : entry.id.toLowerCase().includes(q) || entry.text.toLowerCase().includes(q)
    })
    .slice(0, 50)
}

/** the character just before `range`: the bang that makes a completion an embed */
function embedPrefixBefore(editor: Editor, range: Range): string {
  return editor.state.doc.textBetween(Math.max(0, range.from - 1), range.from, undefined, '\ufffc')
}

/**
 * Create the note the query names, then finish the link in place with the
 * created file's own target — the spelling that resolves, even where a
 * sanitized character changed the name, directories kept so `[[folder/Note]]`
 * cannot resolve to another `Note.md` nearer in the lookup order. The vault listing is dropped so the next
 * keystroke offers the new note like any other.
 */
async function createNoteAndComplete(
  props: SuggestionProps<WikiSuggestItem>,
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
  const text = `${open}${result.target}${querySuffixPart(props.query)}]]`
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
  items: WikiSuggestItem[]
  clientRect: DOMRect | null
  command: (item: WikiSuggestItem) => void
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
  /** absolute path of the open document; '' while it is untitled. Keys the heads cache. */
  getDocumentPath?: () => string
}

export const WikiSuggest = Extension.create<WikiSuggestOptions>({
  name: 'wikiSuggest',

  addOptions() {
    return { controller: null, getDocumentPath: () => '' }
  },

  onBeforeCreate() {
    refreshWikiNotes()
  },

  addProseMirrorPlugins() {
    const getController = () => this.options.controller
    // the key-down props carry no items; the render callbacks last saw them
    let visibleItems: WikiSuggestItem[] = []
    /** whether the last render offered a create row; a key-down has no query either */
    let visibleCreateName = ''
    // a fragment names something that exists; creation is the notes list's row
    const offeredCreateName = (props: SuggestionProps<WikiSuggestItem>): string =>
      props.query.includes('#') ? '' : createNameOf(props)
    const toState = (props: SuggestionProps<WikiSuggestItem>): WikiSuggestState => {
      const embed = embedPrefixBefore(props.editor, props.range) === '!'
      return {
        items: props.items,
        clientRect: props.clientRect?.() ?? null,
        command: (item) => props.command(item),
        createName: offeredCreateName(props),
        create: () => void createNoteAndComplete(props, embed),
      }
    }
    return [
      Suggestion<WikiSuggestItem>({
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
          if (!query.includes('#')) {
            refreshWikiNotes()
            return wikiSuggestItems(query).map((entry) => ({ kind: 'note', entry }))
          }
          const name = queryNamePart(query)
          const fragment = queryFragmentPart(query)
          // `[[#` addresses this document; anything else goes through main
          if (!name.trim()) return docFragmentItems(this.editor.state.doc, fragment)
          return fetchWikiHeads(() => this.options.getDocumentPath?.() ?? '', name.trim()).then(
            (entries) => filterWikiHeads(entries, fragment),
          )
        },
        command: ({ editor, range, props: item }) => {
          const embed = embedPrefixBefore(editor, range) === '!'
          // the range covers the trigger brackets too, so the name part read
          // from them already carries them; an embed's bang sits before the
          // range and the completion adds its own (as `completionFor` does)
          const query = editor.state.doc.textBetween(range.from, range.to, undefined, '\ufffc')
          let text: string
          if (item.kind === 'note') {
            text = completionFor(item.entry, query, embed)
          } else {
            const name = queryNamePart(query)
            const frag = item.kind === 'block' ? `^${item.id}` : item.text
            // a typed alias rides along (`[[Note#Head|al` stays `…#Head|al`),
            // minus a hand-closed `]]` the completion re-adds
            const bar = query.indexOf('|')
            const alias = bar < 0 ? '' : query.slice(bar).replace(/\]\]/, '')
            text = `${embed ? '!' : ''}${name}#${frag}${alias}]]`
          }
          insertCompletion(editor, range, text, embed)
        },
        render: () => ({
          onStart: (props) => {
            visibleItems = props.items
            visibleCreateName = offeredCreateName(props)
            getController()?.onOpen(toState(props))
          },
          onUpdate: (props) => {
            visibleItems = props.items
            visibleCreateName = offeredCreateName(props)
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
