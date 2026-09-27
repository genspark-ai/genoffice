import { describe, expect, it } from 'vitest'
import { PassThrough } from 'node:stream'
import {
  attachBoundedRpcStdout,
  MAX_RPC_LINE_BYTES,
  type CodexChildLike,
} from '../src/codex-app-server'

/** A child whose stdout is driven by the test, with a kill spy. */
function fakeChild() {
  const stdout = new PassThrough()
  const killed: number[] = []
  const child: CodexChildLike = {
    stdout: stdout as unknown as CodexChildLike['stdout'],
    kill: () => {
      killed.push(1)
      return true
    },
  }
  return { child, stdout, killed }
}

const settle = () => new Promise((r) => setImmediate(r))

describe('Codex app-server stdout reader', () => {
  it('delivers newline-delimited lines, tolerating CRLF and a final unterminated line', async () => {
    const { child, stdout } = fakeChild()
    const lines: string[] = []
    attachBoundedRpcStdout(
      child,
      (line) => lines.push(line),
      () => undefined,
    )
    stdout.write('{"id":1}\r\n{"id":2}\n{"id":3}')
    await settle()
    expect(lines).toEqual(['{"id":1}', '{"id":2}'])
    stdout.end()
    await settle()
    expect(lines).toEqual(['{"id":1}', '{"id":2}', '{"id":3}'])
  })

  it('reassembles a line split across chunks', async () => {
    const { child, stdout } = fakeChild()
    const lines: string[] = []
    attachBoundedRpcStdout(
      child,
      (line) => lines.push(line),
      () => undefined,
    )
    stdout.write('{"me')
    stdout.write('thod":"x"}\n')
    await settle()
    expect(lines).toEqual(['{"method":"x"}'])
  })

  it('keeps a multi-byte UTF-8 character split across two chunks intact', async () => {
    const { child, stdout } = fakeChild()
    const lines: string[] = []
    attachBoundedRpcStdout(
      child,
      (line) => lines.push(line),
      () => undefined,
    )
    const bytes = Buffer.from('{"text":"\u6587\u6863\u{1f600}"}\n', 'utf8')
    // Cut inside the 3-byte U+6587 and again inside the 4-byte emoji.
    const cutA = bytes.indexOf(Buffer.from('\u6587', 'utf8')) + 1
    const cutB = bytes.indexOf(Buffer.from('😀', 'utf8')) + 2
    stdout.write(bytes.subarray(0, cutA))
    stdout.write(bytes.subarray(cutA, cutB))
    stdout.write(bytes.subarray(cutB))
    await settle()
    expect(lines).toHaveLength(1)
    expect(JSON.parse(lines[0]!)).toEqual({ text: '\u6587\u6863\u{1f600}' })
    expect(lines[0]).not.toContain('\uFFFD')
  })

  it('counts the cap in UTF-8 bytes, not UTF-16 code units', async () => {
    const { child, stdout, killed } = fakeChild()
    const errors: Error[] = []
    attachBoundedRpcStdout(
      child,
      () => undefined,
      (error) => errors.push(error),
    )
    // 3 bytes per character: exceeds the byte cap while staying well under it in code units.
    const chunk = '\u6587'.repeat(64 * 1024)
    for (let written = 0; written <= MAX_RPC_LINE_BYTES; written += chunk.length * 3) {
      stdout.write(chunk)
    }
    await settle()
    expect(errors).toHaveLength(1)
    expect(killed).toHaveLength(1)
  })

  it('kills the child and reports a bounded diagnostic when a line never ends', async () => {
    const { child, stdout, killed } = fakeChild()
    const lines: string[] = []
    const errors: Error[] = []
    attachBoundedRpcStdout(
      child,
      (line) => lines.push(line),
      (error) => errors.push(error),
    )
    const chunk = 'x'.repeat(64 * 1024)
    for (let written = 0; written <= MAX_RPC_LINE_BYTES; written += chunk.length) {
      stdout.write(chunk)
    }
    await settle()
    expect(lines).toEqual([])
    expect(killed).toHaveLength(1)
    expect(errors).toHaveLength(1)
    expect(errors[0]!.message).toMatch(/stdout line exceeded .* bytes without a newline/)
    expect(errors[0]!.message.length).toBeLessThan(200)
  })

  it('ignores everything the child writes after the overflow', async () => {
    const { child, stdout, killed } = fakeChild()
    const lines: string[] = []
    const errors: Error[] = []
    attachBoundedRpcStdout(
      child,
      (line) => lines.push(line),
      (error) => errors.push(error),
    )
    stdout.write('x'.repeat(MAX_RPC_LINE_BYTES + 1))
    await settle()
    stdout.write('{"id":9}\n')
    stdout.end()
    await settle()
    expect(lines).toEqual([])
    expect(errors).toHaveLength(1)
    expect(killed).toHaveLength(1)
  })

  it('survives a kill that throws', async () => {
    const stdout = new PassThrough()
    const child: CodexChildLike = {
      stdout: stdout as unknown as CodexChildLike['stdout'],
      kill: () => {
        throw new Error('ESRCH')
      },
    }
    const errors: Error[] = []
    attachBoundedRpcStdout(
      child,
      () => undefined,
      (error) => errors.push(error),
    )
    stdout.write('x'.repeat(MAX_RPC_LINE_BYTES + 1))
    await settle()
    expect(errors).toHaveLength(1)
  })
})
