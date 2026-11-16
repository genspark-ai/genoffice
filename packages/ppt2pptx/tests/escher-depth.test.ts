import { describe, expect, it } from 'vitest'
import { ESCHER, parseSpgrContainer } from '../src/escher'
import type { PptRecord } from '../src/records'

const HEADER = 8

function concat(parts: Uint8Array[]): Uint8Array {
  const total = parts.reduce((n, p) => n + p.length, 0)
  const out = new Uint8Array(total)
  let at = 0
  for (const p of parts) {
    out.set(p, at)
    at += p.length
  }
  return out
}

/** One Escher container record wrapping `children`. */
function container(type: number, children: Uint8Array): Uint8Array {
  const head = new Uint8Array(HEADER)
  const dv = new DataView(head.buffer)
  dv.setUint16(0, 0x000f, true) // version 0xf: container
  dv.setUint16(2, type, true)
  dv.setUint32(4, children.length, true)
  return concat([head, children])
}

/** A minimal SpContainer payload (no atoms — the shape parse tolerates it). */
function spContainer(): Uint8Array {
  return container(ESCHER.spContainer, new Uint8Array(0))
}

describe('parseSpgrContainer depth cap', () => {
  it('survives a group chain far deeper than the stack could follow', () => {
    // fxp-equivalent hostile shape: 20k nested spgrContainers, 8 bytes each.
    // The walk is iterative, so only the group descent could overflow; the
    // cap degrades everything past 64 to empty groups.
    let stream: Uint8Array = new Uint8Array(0)
    for (let i = 0; i < 20_000; i++) stream = container(ESCHER.spgrContainer, stream)
    const record: PptRecord = {
      version: 0xf,
      instance: 0,
      type: ESCHER.spgrContainer,
      length: stream.length - HEADER,
      offset: HEADER,
    }
    const shapes = parseSpgrContainer(stream, record)
    expect(Array.isArray(shapes)).toBe(true)
  })

  it('still parses a legitimately nested group with its member', () => {
    // top group > inner group [group-shape, member-shape]: the inner group's
    // first SpContainer is the group itself, its second is the member
    const inner = container(ESCHER.spgrContainer, concat([spContainer(), spContainer()]))
    const stream = container(ESCHER.spgrContainer, inner)
    const record: PptRecord = {
      version: 0xf,
      instance: 0,
      type: ESCHER.spgrContainer,
      length: inner.length,
      offset: HEADER,
    }
    const shapes = parseSpgrContainer(stream, record)
    expect(shapes.length).toBe(1)
    expect(shapes[0]!.children?.length).toBe(1)
  })
})
