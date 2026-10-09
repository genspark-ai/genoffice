import { describe, expect, it } from 'vitest'

import { applySheetNotes, type SheetNote } from '../src/gateway/xlsx-notes'
import { personIdFor, threadedMirrorText } from '../src/gateway/xlsx-threaded-comments'

const SHEET_PATH = 'xl/worksheets/sheet1.xml'
const ROOT_ID = '{11111111-1111-1111-1111-111111111111}'
const REPLY_ID = '{22222222-2222-2222-2222-222222222222}'
const PERSON_ID = '{AAAAAAAA-AAAA-AAAA-AAAA-AAAAAAAAAAAA}'

const THREAD_PART =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  '<ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">' +
  `<threadedComment ref="A1" dT="2024-01-01T00:00:00.00" personId="${PERSON_ID}" id="${ROOT_ID}">` +
  '<text>Root &amp; text</text></threadedComment>' +
  `<threadedComment ref="A1" dT="2024-01-02T00:00:00.00" personId="${PERSON_ID}" id="${REPLY_ID}" parentId="${ROOT_ID}">` +
  '<text>First reply</text></threadedComment>' +
  '</ThreadedComments>'

const MIRROR = threadedMirrorText('Root & text') + '\nReply:\n    First reply'

const COMMENTS_PART =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  '<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
  `<authors><author>tc=${ROOT_ID}</author></authors>` +
  `<commentList><comment ref="A1" authorId="0"><text><t xml:space="preserve">${MIRROR.replace(/&/g, '&amp;')}</t></text></comment></commentList>` +
  '</comments>'

function fixture(): Map<string, string> {
  return new Map<string, string>([
    [
      '[Content_Types].xml',
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>' +
        '<Override PartName="/xl/comments1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>' +
        '<Override PartName="/xl/threadedComments/threadedComment1.xml" ContentType="application/vnd.ms-excel.threadedcomments+xml"/>' +
        '<Override PartName="/xl/persons/person.xml" ContentType="application/vnd.ms-excel.person+xml"/>' +
        '</Types>',
    ],
    [
      'xl/_rels/workbook.xml.rels',
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId9" Type="http://schemas.microsoft.com/office/2017/10/relationships/person" Target="persons/person.xml"/>' +
        '</Relationships>',
    ],
    [
      'xl/persons/person.xml',
      '<personList xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">' +
        `<person displayName="Ada" id="${PERSON_ID}" userId="ada" providerId="None"/></personList>`,
    ],
    [
      SHEET_PATH,
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
        '<sheetData/><legacyDrawing r:id="rId2"/></worksheet>',
    ],
    [
      'xl/worksheets/_rels/sheet1.xml.rels',
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments1.xml"/>' +
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/vmlDrawing1.vml"/>' +
        '<Relationship Id="rId3" Type="http://schemas.microsoft.com/office/2017/10/relationships/threadedComment" Target="../threadedComments/threadedComment1.xml"/>' +
        '</Relationships>',
    ],
    ['xl/comments1.xml', COMMENTS_PART],
    ['xl/threadedComments/threadedComment1.xml', THREAD_PART],
    [
      'xl/drawings/vmlDrawing1.vml',
      '<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">' +
        '<v:shape id="_x0000_s1025" type="#_x0000_t202"><x:ClientData ObjectType="Note"><x:Row>0</x:Row><x:Column>0</x:Column></x:ClientData></v:shape></xml>',
    ],
  ])
}

function pkgOf(files: Map<string, string>) {
  return {
    paths: async () => [...files.keys()],
    has: async (path: string) => files.has(path),
    readText: async (path: string) => {
      const content = files.get(path)
      if (content === undefined) throw new Error(`missing part: ${path}`)
      return content
    },
    write: (path: string, content: string) => void files.set(path, content),
    add: (path: string, content: string) => void files.set(path, content),
    remove: (path: string) => void files.delete(path),
  }
}

// What the viewer displays for A1 (thread body, not the legacy mirror).
const A1_DISPLAYED: SheetNote = {
  row: 0,
  column: 0,
  author: 'Ada',
  text: 'Root & text\n\nFirst reply',
}

async function run(notes: readonly SheetNote[]) {
  const files = fixture()
  const touched = new Set<string>()
  await applySheetNotes(pkgOf(files), SHEET_PATH, notes, touched)
  return { files, touched }
}

describe('threaded comments stay in step with note edits', () => {
  it('keeps the thread part and the legacy mirror verbatim when the note is untouched', async () => {
    const { files, touched } = await run([A1_DISPLAYED])
    expect(files.get('xl/threadedComments/threadedComment1.xml')).toBe(THREAD_PART)
    expect(touched.has('xl/threadedComments/threadedComment1.xml')).toBe(false)
    const comments = files.get('xl/comments1.xml')!
    expect(comments).toContain(`<author>tc=${ROOT_ID}</author>`)
    expect(comments).toContain('[Threaded comment]')
    expect(comments).toContain('Reply:\n    First reply')
    expect(comments).not.toContain('<author>Ada</author>')
  })

  it('treats the legacy mirror text itself as untouched (legacy-part readers)', async () => {
    const { files } = await run([{ ...A1_DISPLAYED, author: `tc=${ROOT_ID}`, text: MIRROR }])
    expect(files.get('xl/threadedComments/threadedComment1.xml')).toBe(THREAD_PART)
    expect(files.get('xl/comments1.xml')).toContain('Reply:\n    First reply')
  })

  it('edits the root text, folds the reply away and refreshes the mirror', async () => {
    const { files } = await run([{ ...A1_DISPLAYED, text: 'Changed <text> $& $$ $1' }])
    const thread = files.get('xl/threadedComments/threadedComment1.xml')!
    expect(thread).toContain(
      `id="${ROOT_ID}"><text>Changed &lt;text&gt; $&amp; $$ $1</text></threadedComment>`,
    )
    expect(thread).toContain(`personId="${PERSON_ID}"`)
    expect(thread).not.toContain(REPLY_ID)
    const comments = files.get('xl/comments1.xml')!
    expect(comments).toContain(`<author>tc=${ROOT_ID}</author>`)
    expect(comments).toContain('[Threaded comment]')
    expect(comments).toContain('Comment:\n    Changed &lt;text&gt; $&amp; $$ $1')
    expect(comments).not.toContain('First reply')
  })

  it('deleting the note removes the thread, its reply, the rel and the content type', async () => {
    const { files, touched } = await run([])
    expect(files.has('xl/threadedComments/threadedComment1.xml')).toBe(false)
    expect(files.has('xl/comments1.xml')).toBe(false)
    expect(files.get('xl/worksheets/_rels/sheet1.xml.rels')).not.toContain('threadedComment')
    expect(files.get('[Content_Types].xml')).not.toContain('threadedComment1.xml')
    expect(files.get('[Content_Types].xml')).toContain('person.xml')
    expect(files.has('xl/persons/person.xml')).toBe(true)
    expect(touched.has('xl/threadedComments/threadedComment1.xml')).toBe(true)
  })

  it('adding a note on another cell leaves the thread alone and writes the new one legacy-only', async () => {
    const { files } = await run([A1_DISPLAYED, { row: 2, column: 1, author: 'Bob', text: 'plain' }])
    expect(files.get('xl/threadedComments/threadedComment1.xml')).toBe(THREAD_PART)
    const comments = files.get('xl/comments1.xml')!
    expect(comments).toContain('<author>Bob</author>')
    expect(comments).toContain(
      '<comment ref="B3" authorId="1"><text><t xml:space="preserve">plain</t>',
    )
    expect(comments).toContain(`<comment ref="A1" authorId="0">`)
    expect(comments).toContain('[Threaded comment]')
  })

  it('falls back to a plain legacy note when the root person is missing from persons', async () => {
    const files = fixture()
    files.set(
      'xl/persons/person.xml',
      '<personList xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments"/>',
    )
    await applySheetNotes(
      pkgOf(files),
      SHEET_PATH,
      [{ ...A1_DISPLAYED, text: 'edited' }],
      new Set(),
    )
    expect(files.has('xl/threadedComments/threadedComment1.xml')).toBe(false)
    const comments = files.get('xl/comments1.xml')!
    expect(comments).toContain('<author>Ada</author>')
    expect(comments).toContain('<t xml:space="preserve">edited</t>')
    expect(comments).not.toContain('[Threaded comment]')
  })
})

const A1_THREAD: SheetNote = {
  row: 0,
  column: 0,
  author: 'Ada',
  text: 'Root & text',
  thread: {
    id: ROOT_ID,
    personId: PERSON_ID,
    author: 'Ada',
    dT: '2024-01-01T00:00:00.00',
    done: false,
    replies: [
      {
        id: REPLY_ID,
        personId: PERSON_ID,
        author: 'Ada',
        dT: '2024-01-02T00:00:00.00',
        text: 'First reply',
      },
    ],
  },
}
const BOB_ID = personIdFor('Bob')
const NEW_ID = '{33333333-3333-3333-3333-333333333333}'

describe('structured threaded comments', () => {
  it('round-trips an untouched thread byte-identically', async () => {
    const { files, touched } = await run([A1_THREAD])
    expect(files.get('xl/threadedComments/threadedComment1.xml')).toBe(THREAD_PART)
    expect(touched.has('xl/threadedComments/threadedComment1.xml')).toBe(false)
    expect(files.get('xl/persons/person.xml')).toContain(`id="${PERSON_ID}"`)
    expect(files.get('xl/comments1.xml')).toContain('Reply:\n    First reply')
  })

  it('adds a reply from a new person and registers the person', async () => {
    const reply = {
      id: NEW_ID,
      personId: BOB_ID,
      author: 'Bob',
      dT: '2024-02-01T00:00:00.00',
      text: 'Bob <says>',
    }
    const { files } = await run([
      {
        ...A1_THREAD,
        thread: { ...A1_THREAD.thread!, replies: [...A1_THREAD.thread!.replies, reply] },
      },
    ])
    const thread = files.get('xl/threadedComments/threadedComment1.xml')!
    expect(thread).toContain(
      `<threadedComment ref="A1" dT="2024-02-01T00:00:00.00" personId="${BOB_ID}" id="${NEW_ID}" parentId="${ROOT_ID}"><text>Bob &lt;says&gt;</text></threadedComment>`,
    )
    expect(thread.indexOf(REPLY_ID)).toBeLessThan(thread.indexOf(NEW_ID))
    expect(files.get('xl/persons/person.xml')).toContain(
      `<person displayName="Bob" id="${BOB_ID}" userId="Bob" providerId="None"/>`,
    )
    const comments = files.get('xl/comments1.xml')!
    expect(comments).toContain(
      'Comment:\n    Root &amp; text\nReply:\n    First reply\nReply:\n    Bob &lt;says&gt;',
    )
    expect(comments).toContain(`<author>tc=${ROOT_ID}</author>`)
  })

  it('resolves and reopens with done="1" on the root only', async () => {
    const { files } = await run([{ ...A1_THREAD, thread: { ...A1_THREAD.thread!, done: true } }])
    const thread = files.get('xl/threadedComments/threadedComment1.xml')!
    expect(thread).toContain(`id="${ROOT_ID}" done="1">`)
    expect(thread).not.toContain(`id="${REPLY_ID}" parentId="${ROOT_ID}" done`)
    const reopened = await run([A1_THREAD])
    expect(reopened.files.get('xl/threadedComments/threadedComment1.xml')).toBe(THREAD_PART)
  })

  it('deletes a reply and refreshes the mirror', async () => {
    const { files } = await run([{ ...A1_THREAD, thread: { ...A1_THREAD.thread!, replies: [] } }])
    const thread = files.get('xl/threadedComments/threadedComment1.xml')!
    expect(thread).not.toContain(REPLY_ID)
    expect(thread).toContain(ROOT_ID)
    expect(files.get('xl/comments1.xml')).not.toContain('Reply:')
  })

  it('deletes the whole thread when the note is gone', async () => {
    const { files } = await run([])
    expect(files.has('xl/threadedComments/threadedComment1.xml')).toBe(false)
    expect(files.has('xl/comments1.xml')).toBe(false)
  })

  it('converts a thread to a plain note with thread: null', async () => {
    const { files } = await run([{ ...A1_DISPLAYED, text: 'Root & text', thread: null }])
    expect(files.has('xl/threadedComments/threadedComment1.xml')).toBe(false)
    expect(files.get('xl/worksheets/_rels/sheet1.xml.rels')).not.toContain('threadedComment')
    const comments = files.get('xl/comments1.xml')!
    expect(comments).toContain('<author>Ada</author>')
    expect(comments).toContain('<t xml:space="preserve">Root &amp; text</t>')
    expect(comments).not.toContain('[Threaded comment]')
  })

  it('creates the thread part, persons part, rels and content types from scratch', async () => {
    const files = new Map<string, string>([
      [
        '[Content_Types].xml',
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"></Types>',
      ],
      [
        'xl/_rels/workbook.xml.rels',
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>',
      ],
      [
        SHEET_PATH,
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>',
      ],
    ])
    const touched = new Set<string>()
    await applySheetNotes(
      pkgOf(files),
      SHEET_PATH,
      [
        {
          row: 1,
          column: 2,
          author: 'Bob',
          text: 'Fresh',
          thread: {
            id: NEW_ID,
            personId: BOB_ID,
            author: 'Bob',
            dT: '2024-03-01T00:00:00.00',
            done: false,
            replies: [],
          },
        },
      ],
      touched,
    )
    const thread = files.get('xl/threadedComments/threadedComment1.xml')!
    expect(thread).toContain(
      '<ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">',
    )
    expect(thread).toContain(
      `<threadedComment ref="C2" dT="2024-03-01T00:00:00.00" personId="${BOB_ID}" id="${NEW_ID}"><text>Fresh</text></threadedComment>`,
    )
    expect(files.get('xl/persons/person.xml')).toContain(`displayName="Bob" id="${BOB_ID}"`)
    expect(files.get('xl/_rels/workbook.xml.rels')).toContain(
      'relationships/person" Target="persons/person.xml"',
    )
    expect(files.get('xl/worksheets/_rels/sheet1.xml.rels')).toContain(
      'Target="../threadedComments/threadedComment1.xml"',
    )
    const types = files.get('[Content_Types].xml')!
    expect(types).toContain(
      'PartName="/xl/threadedComments/threadedComment1.xml" ContentType="application/vnd.ms-excel.threadedcomments+xml"',
    )
    expect(types).toContain(
      'PartName="/xl/persons/person.xml" ContentType="application/vnd.ms-excel.person+xml"',
    )
    expect(types).toContain('PartName="/xl/comments1.xml"')
    expect(files.get('xl/comments1.xml')).toContain(`<author>tc=${NEW_ID}</author>`)
    expect(files.get('xl/drawings/vmlDrawing1.vml')).toContain(
      '<x:Row>1</x:Row><x:Column>2</x:Column>',
    )
    expect(files.get(SHEET_PATH)).toContain('<legacyDrawing')
  })

  it('derives one stable GUID per display name', () => {
    expect(personIdFor('Bob')).toBe(BOB_ID)
    expect(BOB_ID).toMatch(/^\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}$/)
    expect(personIdFor('Alice')).not.toBe(BOB_ID)
  })
})
