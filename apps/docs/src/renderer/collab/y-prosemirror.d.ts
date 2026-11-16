declare module 'y-prosemirror' {
  import type * as Y from 'yjs'
  import type { Node as PmNode, Schema } from '@tiptap/pm/model'
  import type { EditorView } from '@tiptap/pm/view'
  import type { Plugin, PluginKey } from '@tiptap/pm/state'

  /**
   * The plugin state y-prosemirror's binding reads: the fragment it is bound
   * to and the flags the sync plugin keeps.
   */
  export interface YSyncState {
    type: Y.XmlFragment | null
    doc: Y.Doc | null
    binding: unknown
    snapshot: unknown
    prevSnapshot: unknown
    isChangeOrigin: boolean
    isUndoRedoOperation: boolean
    addToHistory: boolean
  }

  export class ProsemirrorBinding {
    constructor(yXmlFragment: Y.XmlFragment, mapping?: Map<unknown, unknown>)
    type: Y.XmlFragment
    doc: Y.Doc
    prosemirrorView: EditorView | null
    mapping: Map<unknown, unknown>
    initView(view: EditorView): void
    destroy(): void
    mux(fn: () => void): void
    _forceRerender(): void
    _prosemirrorChanged(doc: PmNode): void
  }

  export const ySyncPluginKey: PluginKey<YSyncState>
  export const yUndoPluginKey: PluginKey<unknown>
  export const yCursorPluginKey: PluginKey<unknown>
  /** the awareness surface y-prosemirror's cursor plugin uses */
  export interface AwarenessSurface {
    clientID: number
    getStates(): Map<number, Record<string, unknown>>
    getLocalState(): Record<string, unknown> | null
    setLocalStateField(field: string, value: unknown): void
    on(event: 'update', listener: (changes: unknown, origin: unknown) => void): void
    off(event: 'update', listener: (changes: unknown, origin: unknown) => void): void
  }

  export function yCursorPlugin(awareness: AwarenessSurface, options?: unknown): Plugin
  export function prosemirrorToYXmlFragment(doc: PmNode, xmlFragment?: Y.XmlFragment): Y.XmlFragment
  export function yXmlFragmentToProsemirrorJSON(schema: Schema, xmlFragment: Y.XmlFragment): unknown
}
