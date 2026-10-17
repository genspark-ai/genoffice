import { readFileSync } from 'node:fs'
import { join, dirname } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'
import { detectRemoteRefs, makeSnippet, parseMimeMessage, singleLine, stripTags } from '../src/mime'

const dir = join(dirname(fileURLToPath(import.meta.url)), 'fixtures')
const load = (name: string): string => readFileSync(join(dir, name), 'utf8')

describe('parseMimeMessage', () => {
  it('parses a plain text message into the view-model', async () => {
    const mail = await parseMimeMessage(load('basic-text.eml'))
    expect(mail.subject).toBe('Hello plain text')
    expect(mail.from).toEqual([{ name: 'Alice Example', address: 'alice@example.com' }])
    expect(mail.to).toEqual([{ name: 'Bob Example', address: 'bob@example.com' }])
    expect(mail.text).toContain('plain text body')
    expect(mail.html).toBeUndefined()
    expect(mail.hasAttach).toBe(false)
    expect(mail.hasRemoteRefs).toBe(false)
    expect(mail.snippet).not.toMatch(/[\r\n]/)
  })

  it('prefers text for the snippet but keeps html for the reader', async () => {
    const mail = await parseMimeMessage(load('multipart-alt.eml'))
    expect(mail.text).toContain('Fallback plain text')
    expect(mail.html).toContain('Rich <b>HTML</b> body')
    expect(mail.cc.map((a) => a.address)).toEqual(['erin@example.com'])
    expect(mail.snippet).toContain('Fallback plain text')
  })

  it('extracts attachments with decoded bytes', async () => {
    const mail = await parseMimeMessage(load('with-attachment.eml'))
    expect(mail.hasAttach).toBe(true)
    expect(mail.attachments).toHaveLength(1)
    const [att] = mail.attachments
    expect(att.filename).toBe('notes.txt')
    expect(Buffer.from(att.content).toString('utf8')).toBe('hello attachment content\n')
    expect(att.size).toBe(att.content.byteLength)
  })

  it('keeps hostile HTML as inert data and flags remote refs', async () => {
    const mail = await parseMimeMessage(load('hostile-mail.eml'))
    // Parser output is data: the script text survives but nothing executes.
    expect(mail.html).toContain('<script>alert(document.cookie)</script>')
    expect(mail.hasRemoteRefs).toBe(true)
    expect(mail.snippet).not.toContain('<script>')
    // Path-traversal filename is preserved raw here; staging sanitizes it.
    expect(mail.attachments[0].filename).toBe('../../evil.exe')
  })

  it('decodes RFC 2047 display names and subjects', async () => {
    const mail = await parseMimeMessage(load('encoded-addresses.eml'))
    expect(mail.from[0].address).toBe('mei@example.com')
    expect(mail.from[0].name).not.toContain('=?UTF-8?')
    expect(mail.subject).not.toContain('=?UTF-8?')
    expect(mail.subject).toContain('subject')
  })
})

describe('singleLine', () => {
  it('folds CRLF header injection into one line', () => {
    expect(singleLine('You won\r\nBcc: evil@example.com')).toBe('You won Bcc: evil@example.com')
    expect(singleLine(undefined)).toBe('')
  })
})

describe('makeSnippet', () => {
  it('folds whitespace and caps length without newlines', () => {
    expect(makeSnippet('a\n\nb\t\tc')).toBe('a b c')
    expect(makeSnippet('x'.repeat(500), 200)).toHaveLength(200)
    expect(makeSnippet('x'.repeat(500), 200).endsWith('…')).toBe(true)
  })
})

describe('stripTags', () => {
  it('drops script and style blocks before snippet fallback', () => {
    expect(stripTags('<p>hi</p><script>evil()</script>')).not.toContain('evil()')
    expect(stripTags('<p>hi</p>')).toContain('hi')
  })
})

describe('detectRemoteRefs', () => {
  it('flags remote images and links, ignores relative and cid refs', () => {
    expect(detectRemoteRefs('<img src="http://tracker.example.com/p.gif">')).toBe(true)
    expect(detectRemoteRefs('<img src="cid:part1">')).toBe(false)
    expect(detectRemoteRefs('<p>plain</p>')).toBe(false)
    expect(detectRemoteRefs(undefined)).toBe(false)
  })
})
