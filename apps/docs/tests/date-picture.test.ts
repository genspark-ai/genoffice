/**
 * Word's date-time pictures: the formats the Date and Time dialog lists, as
 * Word showed them for Thursday 12 November 2026, 9:27:44 PM.
 */
import { describe, expect, it } from 'vitest'
import {
  DATE_TIME_PICTURES,
  dateTimePicturesFor,
  datePictureOf,
  formatDatePicture,
  toFullWidth,
  wantsFullWidth,
} from '../src/renderer/editor/date-picture'

const WHEN = new Date(2026, 10, 12, 21, 27, 44)

describe('date-time pictures', () => {
  it('formats every picture of the dialog the way Word lists it', () => {
    expect(DATE_TIME_PICTURES.map((p) => formatDatePicture(WHEN, p))).toEqual([
      '11/12/26',
      'Thursday, November 12, 2026',
      'November 12, 2026',
      '11/12/2026',
      '2026-11-12',
      '12-Nov-26',
      '11.12.26',
      'Nov. 12, 26',
      '12 November 2026',
      'November 26',
      'Nov-26',
      '11/12/26 9:27 PM',
      '11/12/26 9:27:44 PM',
      '9:27 PM',
      '9:27:44 PM',
      '21:27',
      '21:27:44',
    ])
  })

  it('pads, reads midnight and noon on the 12-hour clock, and keeps quoted text', () => {
    const early = new Date(2026, 0, 5, 0, 4, 9)
    expect(formatDatePicture(early, 'MM/dd/yyyy hh:mm:ss am/pm')).toBe('01/05/2026 12:04:09 AM')
    expect(formatDatePicture(early, 'ddd d MMM')).toBe('Mon 5 Jan')
    expect(formatDatePicture(new Date(2026, 0, 5, 12, 0, 0), 'h am/pm')).toBe('12 PM')
    expect(formatDatePicture(early, "'day' d 'of' MMMM")).toBe('day 5 of January')
  })

  it('reads the picture out of a field instruction', () => {
    expect(datePictureOf('DATE \\@ "MMMM d, yyyy"')).toBe('MMMM d, yyyy')
    expect(datePictureOf('TIME \\@ HH:mm')).toBe('HH:mm')
    expect(datePictureOf('DATE')).toBeNull()
  })

  it('names months and days in the language asked for, English when none is', () => {
    expect(formatDatePicture(WHEN, 'dddd, d MMMM yyyy', 'de-DE')).toBe(
      'Donnerstag, 12 November 2026',
    )
    expect(formatDatePicture(WHEN, 'd MMMM yyyy', 'fr-FR')).toBe('12 novembre 2026')
    expect(formatDatePicture(WHEN, 'd MMMM yyyy', 'en-GB')).toBe('12 November 2026')
    expect(formatDatePicture(WHEN, 'd MMMM yyyy', 'not a tag')).toBe('12 November 2026')
  })

  it('lists English as Word does and another language the way it writes dates', () => {
    expect(dateTimePicturesFor('en-US')).toBe(DATE_TIME_PICTURES)
    const german = dateTimePicturesFor('de-DE')
    expect(german.slice(0, 3)).toEqual(['dd.MM.yy', 'dddd, d. MMMM yyyy', 'd. MMMM yyyy'])
    expect(german).toContain('yyyy-MM-dd')
    expect(german).toContain('HH:mm')
    expect(new Set(german).size).toBe(german.length)
    expect(formatDatePicture(WHEN, german[1], 'de-DE')).toBe('Donnerstag, 12. November 2026')
    // words of the language between the parts are kept as text
    const spanish = dateTimePicturesFor('es-ES').find((p) => p.includes('MMMM'))!
    expect(formatDatePicture(WHEN, spanish, 'es-ES')).toMatch(/12 de noviembre de 2026/)
  })

  it('writes a result in full-width characters when the field asks for it', () => {
    expect(toFullWidth('2026/11/12 a')).toBe(
      '\uff12\uff10\uff12\uff16\uff0f\uff11\uff11\uff0f\uff11\uff12 \uff41',
    )
    expect(wantsFullWidth('DATE \\@ "yyyy/M/d" \\* DBCHAR')).toBe(true)
    expect(wantsFullWidth('DATE \\@ "yyyy/M/d"')).toBe(false)
  })
})
