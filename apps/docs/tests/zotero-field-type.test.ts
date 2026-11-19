import { describe, expect, it } from 'vitest'

import { fieldTypeForDisk, fieldTypeForSession } from '../src/renderer/zotero/field-type'

/** Zotero serializes its document preferences as one `<data>` blob; the
 * citation storage rides in the `fieldType` preference beside the others. */
function dataWithFieldType(value: string): string {
  return (
    '<data data-version="3"><prefs>' +
    '<pref name="automaticJournalAbbreviations"><value>true</value></pref>' +
    `<pref name="fieldType"><value>${value}</value></pref>` +
    '<pref name="noteStyle"><value>note</value></pref>' +
    '</prefs></data>'
  )
}

describe('the field type the saved file carries', () => {
  it('writes Field for the storages the LibreOffice bridge reports', () => {
    expect(fieldTypeForDisk(dataWithFieldType('ReferenceMark'))).toContain(
      '<pref name="fieldType"><value>Field</value></pref>',
    )
    // Bookmark is accepted on disk by Word, but nothing writes the bookmarks
    // it names, so it would reopen as citations nothing can find
    expect(fieldTypeForDisk(dataWithFieldType('Bookmark'))).toContain(
      '<pref name="fieldType"><value>Field</value></pref>',
    )
  })

  it('leaves everything but the field type alone', () => {
    const data = dataWithFieldType('ReferenceMark')
    const onDisk = fieldTypeForDisk(data)
    expect(onDisk).toContain(
      '<pref name="automaticJournalAbbreviations"><value>true</value></pref>',
    )
    expect(onDisk).toContain('<pref name="noteStyle"><value>note</value></pref>')
  })

  it('is stable when the session value was already Field', () => {
    expect(fieldTypeForDisk(dataWithFieldType('Field'))).toBe(dataWithFieldType('Field'))
  })
})

describe('the field type the session reads back', () => {
  it('maps Field back to the ReferenceMark the bridge session expects', () => {
    expect(fieldTypeForSession(dataWithFieldType('Field'))).toContain(
      '<pref name="fieldType"><value>ReferenceMark</value></pref>',
    )
  })

  it('leaves the session spellings alone, so a round trip is stable', () => {
    const session = dataWithFieldType('ReferenceMark')
    expect(fieldTypeForSession(session)).toBe(session)
    // disk → session → disk comes back to exactly the stored bytes
    const onDisk = fieldTypeForDisk(session)
    expect(fieldTypeForSession(onDisk)).toBe(session)
  })
})

describe('the attribute spelling some Zotero builds write', () => {
  const attributeForm = (value: string): string =>
    `<data data-version="3"><prefs><pref name="fieldType" value="${value}"/></prefs></data>`

  it('normalizes and maps back through it too', () => {
    expect(fieldTypeForDisk(attributeForm('Bookmark'))).toBe(attributeForm('Field'))
    expect(fieldTypeForSession(attributeForm('Field'))).toBe(attributeForm('ReferenceMark'))
    expect(fieldTypeForSession(attributeForm('Bookmark'))).toBe(attributeForm('Bookmark'))
  })
})
