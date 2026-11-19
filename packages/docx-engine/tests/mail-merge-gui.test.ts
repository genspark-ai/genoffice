import { describe, expect, it } from 'vitest'
import {
  mergeFieldInstruction,
  mergeFieldName,
  mergeOutputFileName,
  type MergeRecord,
} from '../src/mail-merge'

/**
 * The two halves the Docs mail-merge GUI needs from the engine: the field code
 * a template carries (so a template built in the app is interchangeable with
 * one built in Word) and the file name a merged record is written under.
 * genoffice#2097.
 */

const record: MergeRecord = {
  FirstName: 'Ada',
  'Last Name': 'Lovelace',
  Company: 'Analytical Engines Ltd',
}

describe('mergeFieldInstruction', () => {
  it('writes the field code Word writes, with the name quoted', () => {
    expect(mergeFieldInstruction('FirstName')).toBe('MERGEFIELD "FirstName" \\* MERGEFORMAT')
  })

  it('keeps a name with a space readable back', () => {
    const instr = mergeFieldInstruction('Last Name')
    expect(mergeFieldName(instr)).toBe('Last Name')
  })

  it('drops a quote from the name so the field code stays parseable', () => {
    expect(mergeFieldName(mergeFieldInstruction('Say "Hi"'))).toBe('Say Hi')
  })
})

describe('mergeOutputFileName', () => {
  it('uses the first field when no pattern is given', () => {
    expect(mergeOutputFileName('', record, ['FirstName', 'Last Name'])).toBe('Ada')
  })

  it('uses the field the pattern names', () => {
    expect(mergeOutputFileName('Last Name', record, ['FirstName', 'Last Name'])).toBe('Lovelace')
  })

  it('fills {Field} placeholders in a pattern', () => {
    expect(mergeOutputFileName('Letter_{FirstName}', record, ['FirstName'])).toBe('Letter_Ada')
  })

  it('replaces characters a file name cannot carry', () => {
    expect(mergeOutputFileName('a/b:c', record, ['FirstName'])).toBe('a_b_c')
  })

  it('falls back to a usable name when the record holds nothing', () => {
    expect(mergeOutputFileName('FirstName', {}, ['FirstName'])).toBe('document')
  })
})
