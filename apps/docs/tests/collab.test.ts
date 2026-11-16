import { beforeEach, describe, expect, it } from 'vitest'
import { EditorState, TextSelection } from '@tiptap/pm/state'
import { EditorView } from '@tiptap/pm/view'
import { Schema, type Node as PmNode } from '@tiptap/pm/model'
import { yCursorPlugin } from 'y-prosemirror'

import { CollabSession } from '../src/renderer/collab/session'
import { CollabSyncController } from '../src/renderer/collab/sync-extension'
import { CollabAwareness } from '../src/renderer/collab/awareness'
import { CollabRooms, type CollabSender } from '../src/main/collab-rooms'
import { COLLAB_CHANNELS, type CollabBridge, type CollabJoinReply } from '../src/shared/collab'

/**
 * The collaboration loop with two simulated sessions.
 *
 * The sessions talk to the real room registry — only the shell's ipcMain
 * wiring is left out, because that part is a thin translation of the same
 * calls.
 */

const schema = new Schema({
  nodes: {
    doc: { content: 'block+' },
    paragraph: {
      content: 'inline*',
      group: 'block',
      // an attribute proves the CRDT keeps the document's formatting
      attrs: { style: { default: 'body' } },
      toDOM: (node) => ['p', { 'data-style': node.attrs.style }, 0],
      parseDOM: [{ tag: 'p' }],
    },
    text: { group: 'inline' },
  },
  marks: {
    bold: { toDOM: () => ['strong', 0], parseDOM: [{ tag: 'strong' }] },
  },
})

/** one session's link to the rooms */
class TestPeer {
  readonly sender: CollabSender & { destroyed: boolean }

  private readonly updateHandlers: Array<(path: string, update: string) => void> = []
  private readonly awarenessHandlers: Array<(path: string, update: string) => void> = []
  private readonly peersHandlers: Array<
    (path: string, peers: number, owner: boolean, clientIds: number[]) => void
  > = []
  private onDestroy: (() => void) | null = null

  constructor(
    readonly id: number,
    private readonly rooms: CollabRooms,
  ) {
    this.sender = {
      id,
      destroyed: false,
      send: (channel: string, ...args: unknown[]) => this.dispatch(channel, args),
      once: (event: string, listener: () => void) => {
        if (event === 'destroyed') this.onDestroy = listener
      },
      isDestroyed: () => this.sender.destroyed,
    }
  }

  /** the session leaves the room */
  destroy(): void {
    this.sender.destroyed = true
    this.onDestroy?.()
  }

  private dispatch(channel: string, args: unknown[]): void {
    if (channel === COLLAB_CHANNELS.remoteUpdate) {
      for (const h of this.updateHandlers) h(args[0] as string, args[1] as string)
    } else if (channel === COLLAB_CHANNELS.remoteAwareness) {
      for (const h of this.awarenessHandlers) h(args[0] as string, args[1] as string)
    } else if (channel === COLLAB_CHANNELS.peers) {
      for (const h of this.peersHandlers) {
        h(args[0] as string, args[1] as number, args[2] as boolean, args[3] as number[])
      }
    }
  }

  readonly bridge: CollabBridge = {
    join: (path: string, stateVector: string, clientId: number): Promise<CollabJoinReply> =>
      Promise.resolve(this.rooms.join(this.sender, path, stateVector, clientId)),
    update: (path: string, update: string): void => {
      this.rooms.update(this.sender, path, update)
    },
    awareness: (path: string, update: string): void => {
      this.rooms.awareness(this.sender, path, update)
    },
    leave: (path: string): void => {
      this.rooms.leave(this.sender, path)
    },
    onUpdate: (handler: (path: string, update: string) => void): (() => void) => {
      this.updateHandlers.push(handler)
      return () => {
        const i = this.updateHandlers.indexOf(handler)
        if (i >= 0) this.updateHandlers.splice(i, 1)
      }
    },
    onAwareness: (handler: (path: string, update: string) => void): (() => void) => {
      this.awarenessHandlers.push(handler)
      return () => {
        const i = this.awarenessHandlers.indexOf(handler)
        if (i >= 0) this.awarenessHandlers.splice(i, 1)
      }
    },
    onPeers: (
      handler: (path: string, peers: number, owner: boolean, clientIds: number[]) => void,
    ): (() => void) => {
      this.peersHandlers.push(handler)
      return () => {
        const i = this.peersHandlers.indexOf(handler)
        if (i >= 0) this.peersHandlers.splice(i, 1)
      }
    },
  }
}

const DOC_PATH = '/tmp/collab-test.docx'

interface Rig {
  peerA: TestPeer
  peerB: TestPeer
  sessionA: CollabSession
  sessionB: CollabSession
  ctrlA: CollabSyncController
  ctrlB: CollabSyncController
  viewA: EditorView
  viewB: EditorView
}

/** two sessions on the same document, each with its own editor */
function makeRig(rooms: CollabRooms): Rig {
  const peerA = new TestPeer(1, rooms)
  const peerB = new TestPeer(2, rooms)
  const sessionA = new CollabSession({
    bridge: peerA.bridge,
    user: { name: 'A', color: '#e5484d' },
    awareness: new CollabAwareness(),
  })
  const sessionB = new CollabSession({
    bridge: peerB.bridge,
    user: { name: 'B', color: '#30a46c' },
    awareness: new CollabAwareness(),
  })
  const host = document.createElement('div')
  document.body.appendChild(host)
  const ctrlA = new CollabSyncController()
  const ctrlB = new CollabSyncController()
  const viewA = new EditorView(host, {
    state: EditorState.create({
      schema,
      plugins: [ctrlA.plugin, yCursorPlugin(sessionA.awareness)],
    }),
  })
  const viewB = new EditorView(host, {
    state: EditorState.create({
      schema,
      plugins: [ctrlB.plugin, yCursorPlugin(sessionB.awareness)],
    }),
  })
  return { peerA, peerB, sessionA, sessionB, ctrlA, ctrlB, viewA, viewB }
}

/** the document the first session parsed from disk */
function parsedDoc(): PmNode {
  return schema.node('doc', null, [
    schema.node('paragraph', { style: 'title' }, [schema.text('hello')]),
  ])
}

/** the name of the other member's cursor, if it reached this session */
function remoteUser(session: CollabSession): string | undefined {
  for (const [clientId, state] of session.awareness.getStates()) {
    if (clientId === session.awareness.clientID) continue
    const name = (state.user as { name?: string } | undefined)?.name
    if (name) return name
  }
  return undefined
}

/** let the queued renders settle */
async function settle(): Promise<void> {
  await new Promise((resolve) => setTimeout(resolve, 0))
}

/** open the room on both sides and bind both editors to their fragment */
async function joinRoom(rig: Rig): Promise<void> {
  await rig.sessionA.open(DOC_PATH)
  rig.sessionA.publish(parsedDoc())
  await rig.sessionB.open(DOC_PATH)
  rig.ctrlA.attach(rig.sessionA.yFragment!)
  rig.ctrlB.attach(rig.sessionB.yFragment!)
  await settle()
}

describe('document collaboration', () => {
  let rooms: CollabRooms

  beforeEach(() => {
    rooms = new CollabRooms()
  })

  it('gives the room to the first member and lets a late member read it', async () => {
    const rig = makeRig(rooms)
    await rig.sessionA.open(DOC_PATH)
    expect(rig.sessionA.ownsRoom).toBe(true)
    expect(rig.sessionA.hasRoomContent).toBe(false)

    rig.sessionA.publish(parsedDoc())
    expect(rig.sessionA.hasRoomContent).toBe(true)

    await rig.sessionB.open(DOC_PATH)
    expect(rig.sessionB.ownsRoom).toBe(false)
    expect(rig.sessionB.hasRoomContent).toBe(true)

    // the late member sees the whole document, formatting included
    const block = rig.sessionB.yFragment?.toArray()[0] as
      { getAttribute(name: string): unknown; toString(): string } | undefined
    expect(block?.getAttribute('style')).toBe('title')
    expect(block?.toString()).toContain('hello')
  })

  it('syncs a keystroke to the other session', async () => {
    const rig = makeRig(rooms)
    await joinRoom(rig)

    expect(rig.viewA.state.doc.textContent).toBe('hello')
    expect(rig.viewB.state.doc.textContent).toBe('hello')

    // A types at the end of the paragraph
    rig.viewA.dispatch(rig.viewA.state.tr.insertText('!', 6))
    await settle()

    expect(rig.viewA.state.doc.textContent).toBe('hello!')
    expect(rig.viewB.state.doc.textContent).toBe('hello!')
  })

  it('keeps the formatting of the shared document', async () => {
    const rig = makeRig(rooms)
    await joinRoom(rig)

    expect(rig.viewB.state.doc.firstChild?.attrs.style).toBe('title')
  })

  it('shares the cursor through awareness', async () => {
    const rig = makeRig(rooms)
    await joinRoom(rig)

    // A puts the caret inside the paragraph; the other session learns about it
    rig.viewA.dom.focus()
    rig.viewA.dispatch(
      rig.viewA.state.tr.setSelection(TextSelection.create(rig.viewA.state.doc, 3)),
    )
    await settle()

    expect(remoteUser(rig.sessionB)).toBe('A')
  })

  it('drops the cursor of a member that left the room', async () => {
    const rig = makeRig(rooms)
    await joinRoom(rig)
    rig.viewA.dom.focus()
    rig.viewA.dispatch(
      rig.viewA.state.tr.setSelection(TextSelection.create(rig.viewA.state.doc, 3)),
    )
    await settle()
    expect(remoteUser(rig.sessionB)).toBe('A')

    rig.peerA.destroy()
    await settle()

    expect(remoteUser(rig.sessionB)).toBeUndefined()
  })
})
