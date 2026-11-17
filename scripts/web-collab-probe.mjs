// Collaboration probe: two real WebSocket clients talk to the running web
// service and check that a room relays between them.
//
// It targets the snapshot rooms (the sheets app): the docs app's room is a
// yjs merge room and only accepts a yjs update, so a plain snapshot is
// rejected there on purpose.
//
//   node scripts/web-collab-probe.mjs <baseURL> <token> [path] [app]
//
// The clients speak the web bridge's frame protocol directly, so the probe
// covers the part the unit tests fake: the WebSocket transport, the ipcMain
// handlers and the room relay.
import { WebSocket } from 'ws'

const base = process.argv[2] ?? 'http://127.0.0.1:8931'
const token = process.argv[3] ?? process.env.WEB_TOKEN ?? ''
const path = process.argv[4] ?? '/tmp/web-collab-probe.xlsx'
const app = process.argv[5] ?? 'sheets'
const channel = (name) => `${app}:collab:${name}`
const { origin } = new URL(base)
const wsBase = `${origin.replace(/^http/, 'ws')}`

let nextId = 1

function connect(label) {
  const socket = new WebSocket(`${wsBase}/${app}/?token=${token}`, {
    headers: { Origin: origin },
  })
  const pending = new Map()
  const events = []
  socket.on('message', (data) => {
    const frame = JSON.parse(String(data))
    if (frame.t === 'reply') {
      const resolve = pending.get(frame.id)
      pending.delete(frame.id)
      resolve?.(frame)
    } else if (frame.t === 'event') {
      events.push(frame)
    }
  })
  const ready = new Promise((resolve, reject) => {
    socket.once('open', resolve)
    socket.once('error', reject)
  })
  return {
    label,
    socket,
    events,
    ready,
    invoke(channel, args) {
      const id = nextId++
      return new Promise((resolve) => {
        pending.set(id, resolve)
        socket.send(JSON.stringify({ t: 'invoke', id, channel, args }))
      })
    },
    send(channel, args) {
      socket.send(JSON.stringify({ t: 'send', channel, args }))
    },
    waitForEvent(channel, timeoutMs = 5000) {
      const found = events.find((event) => event.channel === channel)
      if (found) return Promise.resolve(found)
      return new Promise((resolve, reject) => {
        const timer = setTimeout(
          () => reject(new Error(`${label}: no ${channel} event`)),
          timeoutMs,
        )
        const poll = setInterval(() => {
          const hit = events.find((event) => event.channel === channel)
          if (!hit) return
          clearInterval(poll)
          clearTimeout(timer)
          resolve(hit)
        }, 50)
      })
    },
  }
}

function assert(condition, message) {
  if (!condition) {
    console.error('FAIL:', message)
    process.exitCode = 1
  } else {
    console.log('ok:', message)
  }
}

const a = connect('A')
const b = connect('B')
await a.ready
await b.ready

const joinedA = await a.invoke(channel('join'), [path, '', 11])
const joinedB = await b.invoke(channel('join'), [path, '', 22])
assert(joinedA.ok === true, `A joined the room (${JSON.stringify(joinedA.value ?? joinedA.error)})`)
assert(joinedB.ok === true, `B joined the room (${JSON.stringify(joinedB.value ?? joinedB.error)})`)
assert(joinedB.value?.peers === 2, `the room counts both members (${joinedB.value?.peers})`)

// A publishes a snapshot; B must receive it over the real transport
a.send(channel('update'), [path, '{"revision":42}'])
const relayed = await b.waitForEvent(channel('remote-update'))
assert(
  relayed.args?.[1] === '{"revision":42}',
  `B received the snapshot (${JSON.stringify(relayed.args)})`,
)

// A publishes its presence; B must receive it over the real transport
a.send(channel('awareness'), [
  path,
  JSON.stringify({ id: 11, name: 'A', sheetId: 'sheet-1', sheetName: 'Sheet1', range: 'B2:C4' }),
])
const presence = await b.waitForEvent(channel('remote-awareness'))
assert(
  presence.args?.[1]?.includes('"name":"A"'),
  `B received the presence (${JSON.stringify(presence.args)})`,
)

// a late member joins and gets the room's current state
const c = connect('C')
await c.ready
const joinedC = await c.invoke(channel('join'), [path, '', 33])
assert(
  joinedC.value?.update === '{"revision":42}',
  `a late member got the room state (${joinedC.value?.update})`,
)

a.socket.close()
b.socket.close()
c.socket.close()
console.log(process.exitCode ? 'PROBE FAILED' : 'PROBE PASSED')
