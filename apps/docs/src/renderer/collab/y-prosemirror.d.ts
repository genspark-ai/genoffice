declare module 'y-prosemirror' {
  import type * as Y from 'yjs'
  import type { Node as PmNode, Schema } from '@tiptap/pm/model'
  import type { EditorView } from '@tiptap/pm/view'
  import type { Plugin, PluginKey } from '@tiptap/pm/state'
  import type { Awareness } from 'y-protocols/awareness'

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
  export function yCursorPlugin(awareness: Awareness, options?: unknown): Plugin
  export function prosemirrorToYXmlFragment(doc: PmNode, xmlFragment?: Y.XmlFragment): Y.XmlFragment
  export function yXmlFragmentToProsemirrorJSON(schema: Schema, xmlFragment: Y.XmlFragment): unknown
}
