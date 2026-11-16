import { Extension } from '@tiptap/core'
import { Plugin } from '@tiptap/pm/state'
import type { Node as PmNode } from '@tiptap/pm/model'
import type { EditorView } from '@tiptap/pm/view'
import * as Y from 'yjs'
import { CollabAwareness } from './awareness'
import { ProsemirrorBinding, yCursorPlugin, ySyncPluginKey, type YSyncState } from 'y-prosemirror'
import { COLLAB_FRAGMENT } from '../../shared/collab'

/**
 * Binds the editor to a collaboration room's yjs fragment.
 *
 * y-prosemirror's own sync plugin is bound to one fragment for the lifetime of
 * the editor, but a session may open a different document later — and each
 * document needs its own yjs document, because a yjs document carries the
 * history of everything ever typed into it. The controller therefore owns the
 * binding and can move it to another fragment without recreating the editor.
 *
 * The plugin registers under y-prosemirror's own key: its binding reads the
 * plugin state (the fragment, the history flags) from there.
 */
export class CollabSyncController {
  readonly plugin: Plugin<YSyncState>
  /** the awareness of this editor: one editor, one participant */
  readonly awareness: CollabAwareness

  private readonly mapping = new Map<unknown, unknown>()
  /** the cursor plugin reads the mapping through the sync plugin's state */
  private readonly bindingPlaceholder = { mapping: this.mapping }
  private view: EditorView | null = null
  private binding: ProsemirrorBinding | null = null
  private fragment: Y.XmlFragment | null = null
  /** true while the binding takes over: local content must not be pushed */
  private syncing = false

  constructor() {
    // an idle fragment so the cursor plugin always has a room to point at:
    // the editor is not bound to it, it only keeps the plugin state valid
    const idleDoc = new Y.Doc()
    const idleFragment = idleDoc.getXmlFragment(COLLAB_FRAGMENT)
    this.awareness = new CollabAwareness()
    this.plugin = new Plugin<YSyncState>({
      key: ySyncPluginKey,
      props: {
        editable: (state) => {
          const syncState = ySyncPluginKey.getState(state)
          return syncState?.snapshot == null && syncState?.prevSnapshot == null
        },
      },
      state: {
        init: () => ({
          type: idleFragment,
          doc: idleDoc,
          binding: this.bindingPlaceholder,
          snapshot: null,
          prevSnapshot: null,
          isChangeOrigin: false,
          isUndoRedoOperation: false,
          addToHistory: true,
        }),
        apply: (tr, state) => {
          const change = tr.getMeta(ySyncPluginKey)
          if (change !== undefined) state = { ...state, ...change }
          state.addToHistory = tr.getMeta('addToHistory') !== false
          state.isChangeOrigin = change !== undefined && !!change.isChangeOrigin
          return state
        },
      },
      view: (view) => {
        this.view = view
        this.bind()
        return {
          update: () => this.onViewUpdate(),
          destroy: () => {
            this.view = null
            this.unbind()
          },
        }
      },
    })
  }

  /** true while the editor is bound to a room */
  get isBound(): boolean {
    return this.binding !== null
  }

  /** the editor's current document, or null before the editor exists */
  get currentDoc(): PmNode | null {
    return this.view?.state.doc ?? null
  }

  /** bind the editor to a room's fragment */
  attach(fragment: Y.XmlFragment): void {
    this.unbind()
    this.fragment = fragment
    const view = this.view
    if (!view) return
    const binding = new ProsemirrorBinding(fragment, this.mapping)
    this.binding = binding
    // while the binding takes over, the editor's own content must not be
    // pushed back into the room: the room's content wins
    this.syncing = true
    try {
      // the cursor plugin reads the fragment and the mapping from the sync
      // plugin's state, so the state has to learn about the new binding
      view.dispatch(
        view.state.tr.setMeta(ySyncPluginKey, {
          type: fragment,
          doc: fragment.doc,
          binding,
        }),
      )
      binding.initView(view)
      // the room's content is the editor's content from here on: the binding
      // renders the fragment into the view before the first local edit
      binding._forceRerender()
    } finally {
      this.syncing = false
    }
  }

  /** leave the room: the editor keeps its content, the CRDT stops */
  detach(): void {
    this.unbind()
    this.fragment = null
  }

  private bind(): void {
    const view = this.view
    const fragment = this.fragment
    if (!view || !fragment) return
    this.attach(fragment)
  }

  private unbind(): void {
    this.binding?.destroy()
    this.binding = null
    this.mapping.clear()
  }

  /** push the editor's content into the room after a local change */
  private onViewUpdate(): void {
    const view = this.view
    const binding = this.binding
    if (!view || !binding || this.syncing) return
    const state = this.plugin.getState(view.state)
    binding.mux(() => {
      binding.doc.transact((tr) => {
        tr.meta.set('addToHistory', state?.addToHistory ?? true)
        binding._prosemirrorChanged(view.state.doc)
      }, ySyncPluginKey)
    })
  }
}

let controller: CollabSyncController | null = null

/** the controller the editor's collab plugin created */
export function collabSyncController(): CollabSyncController | null {
  return controller
}

/**
 * The collaboration layer of the editor: the CRDT binding plus the remote
 * cursors. Both stay inert until a room is attached, so a session that works
 * alone behaves exactly as it did before.
 */
export const CollabSync = Extension.create({
  name: 'collabSync',
  addProseMirrorPlugins() {
    const ctrl = new CollabSyncController()
    controller = ctrl
    return [ctrl.plugin, yCursorPlugin(ctrl.awareness)]
  },
})
