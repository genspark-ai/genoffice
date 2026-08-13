/**
 * The RFC 2387 multipart/related body Drive's files.create/files.update expect:
 * a JSON metadata part, then the file bytes, both delimited by the boundary.
 * Electron is mocked for the same reason as google-auth-refresh.test.ts.
 */
import { describe, expect, it, vi } from 'vitest'

vi.mock('electron', () => ({
  app: { getPath: () => '/tmp/genoffice-test-userdata' },
  shell: { openExternal: vi.fn() },
}))

const { buildMultipartUploadBody, buildMoveParentsQuery, buildCopyFileBody } =
  await import('../src/main/google-drive')

function decode(bytes: Uint8Array): string {
  return new TextDecoder().decode(bytes)
}

describe('buildMultipartUploadBody', () => {
  it('wraps metadata JSON and file bytes between the given boundary', () => {
    const fileBytes = new TextEncoder().encode('PK\x03\x04fake-docx-bytes')
    const body = buildMultipartUploadBody(
      'BOUNDARY42',
      { name: 'My Doc', mimeType: 'application/vnd.google-apps.document' },
      fileBytes,
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    )
    const text = decode(body)

    expect(text.startsWith('--BOUNDARY42\r\n')).toBe(true)
    expect(text).toContain('Content-Type: application/json; charset=UTF-8')
    expect(text).toContain('{"name":"My Doc","mimeType":"application/vnd.google-apps.document"}')
    expect(text).toContain(
      'Content-Type: application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    )
    expect(text.trimEnd().endsWith('--BOUNDARY42--')).toBe(true)

    // the file bytes land byte-for-byte at the end of the body (not re-encoded as text)
    const fileFence = body.length - '\r\n--BOUNDARY42--'.length - fileBytes.length
    const fileBytesOut = body.slice(fileFence, fileFence + fileBytes.length)
    expect(Array.from(fileBytesOut)).toEqual(Array.from(fileBytes))
  })

  it('produces an empty JSON object for metadata-less update payloads', () => {
    const fileBytes = new Uint8Array([1, 2, 3])
    const body = buildMultipartUploadBody('B', {}, fileBytes, 'application/octet-stream')
    expect(decode(body)).toContain('{}')
  })
})

describe('buildMoveParentsQuery', () => {
  it('adds the target folder and removes every current parent', () => {
    const query = buildMoveParentsQuery(['old-parent-1'], 'new-folder-id')
    const params = new URLSearchParams(query)
    expect(params.get('addParents')).toBe('new-folder-id')
    expect(params.get('removeParents')).toBe('old-parent-1')
    expect(params.get('fields')).toBe('id,parents')
  })

  it('joins multiple current parents with a comma for removeParents', () => {
    const query = buildMoveParentsQuery(['p1', 'p2', 'p3'], 'target')
    const params = new URLSearchParams(query)
    expect(params.get('removeParents')).toBe('p1,p2,p3')
  })

  it('handles a file with no current parents (empty removeParents)', () => {
    const query = buildMoveParentsQuery([], 'target')
    const params = new URLSearchParams(query)
    expect(params.get('removeParents')).toBe('')
    expect(params.get('addParents')).toBe('target')
  })
})

describe('buildCopyFileBody', () => {
  it('carries only the name when no destination folder is given', () => {
    expect(buildCopyFileBody('Copy of My Doc')).toEqual({ name: 'Copy of My Doc' })
  })

  it('sets parents to the target folder id when one is given', () => {
    expect(buildCopyFileBody('Copy of My Doc', 'folder-123')).toEqual({
      name: 'Copy of My Doc',
      parents: ['folder-123'],
    })
  })
})
