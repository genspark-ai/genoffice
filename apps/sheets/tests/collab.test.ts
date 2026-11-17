import { beforeEach, describe, expect, it, vi } from 'vitest'
import { CollabRooms, SHEETS_COLLAB_CHANNELS, type CollabSender } from '../src/main/collab-rooms'
import { currentCollabPresence } from '@genoffice/ui'
import { SheetsCollab } from '../src/renderer/collab'
import type { CollabBridge, CollabJoinReply } from '@genoffice/electron-utils/collab-protocol'

/**
 * The sheets collaboration loop with two simulated sessions.
 *
 * The sessions talk to the real room registry — only the shell's ipcMain
 * wiring is left out, because that part is a thin translation of the same
 * calls.
 */

class TestPeer {
  readonly sender: CollabSender & { destroyed: boolean }
  readonly updates: string[] = []
  private readonly peersHandlers: Array<
    (path: string, peers: number, owner: boolean, clientIds: number[]) => void
  > = []
  private readonly awarenessHandlers: Array<(path: string, update: string) => void> = []

  constructor(
    readonly id: number,
    private readonly rooms: CollabRooms,
  ) {
    this.sender = {
      id,
      destroyed: false,
      send: (channel: string, ...args: unknown[]) => {
        if (channel === SHEETS_COLLAB_CHANNELS.remoteUpdate) {
          this.updates.push(args[1] as string)
        } else if (channel === SHEETS_COLLAB_CHANNELS.peers) {
          for (const h of this.peersHandlers) {
            h(args[0] as string, args[1] as number, args[2] as boolean, args[3] as number[])
          }
        } else if (channel === SHEETS_COLLAB_CHANNELS.remoteAwareness) {
          for (const h of this.awarenessHandlers) h(args[0] as string, args[1] as string)
        }
      },
      once: () => undefined,
      isDestroyed: () => this.sender.destroyed,
    }
  }

  readonly bridge: CollabBridge = {
    join: (path: string, _stateVector: string, clientId: number): Promise<CollabJoinReply> =>
      Promise.resolve(this.rooms.join(this.sender, path, clientId)),
    update: (path: string, snapshot: string): void => {
      this.rooms.update(this.sender, path, snapshot)
    },
    awareness: (path: string, update: string): void => {
      this.rooms.awareness(this.sender, path, update)
    },
    leave: (path: string): void => {
      this.rooms.leave(this.sender, path)
    },
    onUpdate: () => () => undefined,
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

const PATH = '/tmp/collab-sheets.xlsx'

describe('sheets collaboration rooms', () => {
  let rooms: CollabRooms

  beforeEach(() => {
    rooms = new CollabRooms()
  })

  it('gives an empty room to its first member', () => {
    const peer = new TestPeer(1, rooms)
    const reply = rooms.join(peer.sender, PATH, 11)
    expect(reply.owner).toBe(true)
    expect(reply.peers).toBe(1)
    expect(reply.update).toBe('')
  })

  it('publishes a snapshot to the other member and keeps it for a late joiner', () => {
    const first = new TestPeer(1, rooms)
    const second = new TestPeer(2, rooms)
    rooms.join(first.sender, PATH, 11)
    rooms.join(second.sender, PATH, 22)

    rooms.update(first.sender, PATH, '{"revision":1}')
    expect(second.updates).toEqual(['{"revision":1}'])

    const late = new TestPeer(3, rooms)
    const reply = rooms.join(late.sender, PATH, 33)
    expect(reply.update).toBe('{"revision":1}')
    expect(reply.owner).toBe(false)
  })

  it('drops a member that leaves and hands the room to the next one', () => {
    const first = new TestPeer(1, rooms)
    const second = new TestPeer(2, rooms)
    rooms.join(first.sender, PATH, 11)
    rooms.join(second.sender, PATH, 22)

    rooms.leave(first.sender, PATH)
    expect(rooms.roomCount).toBe(1)

    const third = new TestPeer(3, rooms)
    const reply = rooms.join(third.sender, PATH, 33)
    expect(reply.peers).toBe(2)
  })
})

describe('sheets collaboration client', () => {
  let rooms: CollabRooms

  beforeEach(() => {
    rooms = new CollabRooms()
  })

  it('publishes nothing while the session works alone', async () => {
    const peer = new TestPeer(1, rooms)
    const apply = vi.fn()
    const collab = new SheetsCollab({
      bridge: peer.bridge,
      snapshot: () => '{"revision":1}',
      apply,
      presence: () => '{}',
    })
    await collab.open(PATH)
    collab.publishNow()
    expect(apply).not.toHaveBeenCalled()
    await collab.close()
  })

  it('publishes once a second member shows up', async () => {
    const first = new TestPeer(1, rooms)
    const second = new TestPeer(2, rooms)
    const collab = new SheetsCollab({
      bridge: first.bridge,
      snapshot: () => '{"revision":1}',
      apply: vi.fn(),
      presence: () => '{}',
    })
    await collab.open(PATH)
    // the second member joins: the room tells the members about it
    rooms.join(second.sender, PATH, 22)
    expect(collab.memberCount).toBe(2)
    collab.publishNow()
    expect(second.updates).toEqual(['{"revision":1}'])
    await collab.close()
  })

  it('tracks the presence of the other members and drops the ones that left', async () => {
    const first = new TestPeer(1, rooms)
    const second = new TestPeer(2, rooms)
    // the other member is in the room, so its presence reaches this one
    rooms.join(first.sender, PATH, 11)
    const collab = new SheetsCollab({
      bridge: second.bridge,
      snapshot: () => '{"revision":1}',
      apply: vi.fn(),
      presence: () => '{}',
    })
    await collab.open(PATH)

    // the other member publishes where it is
    first.bridge.awareness(PATH, JSON.stringify({ id: 11, name: 'A', sheet: 'Sheet1' }))
    expect([...currentCollabPresence().values()]).toEqual([
      { id: 11, name: 'A', location: 'Sheet1' },
    ])

    // and the room reports that it left
    rooms.leave(first.sender, PATH)
    expect(currentCollabPresence().size).toBe(0)
    await collab.close()
  })

  it('applies the room state a late member receives', async () => {
    const first = new TestPeer(1, rooms)
    const second = new TestPeer(2, rooms)
    rooms.join(first.sender, PATH, 11)
    rooms.update(first.sender, PATH, '{"revision":7}')

    const apply = vi.fn()
    const collab = new SheetsCollab({
      bridge: second.bridge,
      snapshot: () => '{"revision":2}',
      apply,
      presence: () => '{}',
    })
    await collab.open(PATH)
    expect(apply).toHaveBeenCalledWith('{"revision":7}')
    await collab.close()
  })
})
