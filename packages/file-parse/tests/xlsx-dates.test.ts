import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import { builtinDateFormat, classifyFormatCode, formatSerial } from '../src/xlsx-dates'
import { xlsxToText } from '../src/xlsx'

describe('classifyFormatCode', () => {
  it('recognises date, time and combined formats', () => {
    expect(classifyFormatCode('yyyy-mm-dd')).toMatchObject({ date: true, time: false })
    expect(classifyFormatCode('[$-409]d\\-mmm;@')).toMatchObject({ date: true, time: false })
    expect(classifyFormatCode('mmm-yy')).toMatchObject({ date: true, time: false })
    expect(classifyFormatCode('mmm')).toMatchObject({ date: true, time: false })
    expect(classifyFormatCode('mmmm')).toMatchObject({ date: true, time: false })
    expect(classifyFormatCode('mm:ss')).toMatchObject({ date: false, time: true, seconds: true })
    expect(classifyFormatCode('h:mm AM/PM')).toMatchObject({
      date: false,
      time: true,
      seconds: false,
    })
    expect(classifyFormatCode('hh:mm:ss')).toMatchObject({ date: false, time: true, seconds: true })
    expect(classifyFormatCode('yyyy-mm-dd hh:mm')).toMatchObject({
      date: true,
      time: true,
      seconds: false,
    })
    expect(classifyFormatCode('[h]:mm:ss')).toMatchObject({
      elapsed: true,
      time: true,
      date: false,
    })
    expect(classifyFormatCode('[h]:mm')).toMatchObject({ elapsed: true, time: true, date: false })
    expect(classifyFormatCode('[mm]:ss')).toMatchObject({ elapsed: true, time: true, date: false })
    expect(classifyFormatCode('[m]')).toMatchObject({ elapsed: true, time: true, date: false })
    expect(classifyFormatCode('[mm]')).toMatchObject({ elapsed: true, time: true, date: false })
    expect(classifyFormatCode('HH:MM')).toMatchObject({ date: false, time: true, seconds: false })
    expect(classifyFormatCode('MM:SS')).toMatchObject({ date: false, time: true, seconds: true })
    expect(classifyFormatCode('YYYY/MM/DD')).toMatchObject({ date: true, time: false })
  })

  it('does not mistake numeric, text or literal-heavy formats for dates', () => {
    expect(classifyFormatCode('"$"#,##0.00')).toBeNull()
    expect(classifyFormatCode('0.00%')).toBeNull()
    expect(classifyFormatCode('General')).toBeNull()
    expect(classifyFormatCode('@')).toBeNull()
    expect(classifyFormatCode('#,##0 "days"')).toBeNull()
    expect(classifyFormatCode('[Red]0.0;[Blue]-0.0')).toBeNull()
    expect(classifyFormatCode('0.00E+00')).toBeNull()
  })
})

describe('builtinDateFormat', () => {
  it('keeps seconds on the built-in mm:ss and mmss.0 formats', () => {
    expect(builtinDateFormat(45)).toMatchObject({ time: true, seconds: true })
    expect(builtinDateFormat(47)).toMatchObject({ time: true, seconds: true })
    expect(builtinDateFormat(20)).toMatchObject({ time: true, seconds: false })
    expect(builtinDateFormat(0)).toBeNull()
  })
})

describe('formatSerial', () => {
  const date = { date: true, time: false, seconds: false, elapsed: false }
  const dateTime = { date: true, time: true, seconds: true, elapsed: false }
  const time = { date: false, time: true, seconds: false, elapsed: false }
  const elapsed = { date: false, time: true, seconds: true, elapsed: true }

  it('decodes 1900-system serials including the phantom 29 Feb 1900', () => {
    expect(formatSerial(45292, date, false)).toBe('2024-01-01')
    expect(formatSerial(34379, date, false)).toBe('1994-02-14')
    expect(formatSerial(1, date, false)).toBe('1900-01-01')
    expect(formatSerial(59, date, false)).toBe('1900-02-28')
    expect(formatSerial(60, date, false)).toBe('1900-02-29')
    expect(formatSerial(61, date, false)).toBe('1900-03-01')
    expect(formatSerial(0, date, false)).toBeNull()
    expect(formatSerial(-3, date, false)).toBeNull()
  })

  it('decodes 1904-system serials without the Lotus gap', () => {
    expect(formatSerial(0, date, true)).toBe('1904-01-01')
    expect(formatSerial(60, date, true)).toBe('1904-03-01')
    expect(formatSerial(43830, date, true)).toBe('2024-01-01') // 1900-system 45292 minus the 1462-day offset
  })

  it('renders time fractions, rounding to the second and carrying past midnight', () => {
    expect(formatSerial(45292.5, dateTime, false)).toBe('2024-01-01 12:00:00')
    expect(formatSerial(0.75, time, false)).toBe('18:00')
    expect(formatSerial(45292.9999999, dateTime, false)).toBe('2024-01-02 00:00:00')
    expect(formatSerial(1.5, elapsed, false)).toBe('36:00:00')
  })
})

describe('xlsxToText date cells', () => {
  function workbook(opts: { date1904?: boolean; sheetXml: string; numFmts?: string; xfs: string }) {
    const zip = new JSZip()
    zip.file(
      'xl/workbook.xml',
      '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
        (opts.date1904 ? '<workbookPr date1904="1"/>' : '<workbookPr/>') +
        '<sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>',
    )
    zip.file(
      'xl/_rels/workbook.xml.rels',
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
        '</Relationships>',
    )
    zip.file(
      'xl/styles.xml',
      '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
        (opts.numFmts ? `<numFmts>${opts.numFmts}</numFmts>` : '') +
        `<cellXfs>${opts.xfs}</cellXfs></styleSheet>`,
    )
    zip.file(
      'xl/worksheets/sheet1.xml',
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
        `<sheetData>${opts.sheetXml}</sheetData></worksheet>`,
    )
    return zip.generateAsync({ type: 'uint8array' })
  }

  it('renders date-styled serials as ISO dates and leaves other numbers alone', async () => {
    const bytes = await workbook({
      numFmts:
        '<numFmt numFmtId="164" formatCode="&quot;$&quot;#,##0.00"/>' +
        '<numFmt numFmtId="165" formatCode="[$-409]d\\-mmm;@"/>' +
        '<numFmt numFmtId="166" formatCode="yyyy-mm-dd hh:mm"/>',
      xfs:
        '<xf numFmtId="0"/><xf numFmtId="164"/><xf numFmtId="165"/>' +
        '<xf numFmtId="14"/><xf numFmtId="166"/><xf numFmtId="21"/>',
      sheetXml:
        '<row r="1">' +
        '<c r="A1"><v>45292</v></c>' + // no style: plain number
        '<c r="B1" s="1"><v>45292</v></c>' + // currency
        '<c r="C1" s="2"><v>34379</v></c>' + // custom d-mmm
        '<c r="D1" s="3"><v>60</v></c>' + // built-in m/d/yyyy, phantom leap day
        '<c r="E1" s="4"><v>45292.75</v></c>' + // custom date+time
        '<c r="F1" s="5"><v>0.5</v></c>' + // built-in h:mm:ss
        '<c r="G1" s="3" t="s"><v>0</v></c>' + // shared string with a date style stays text
        '<c r="H1" s="3"><f>TODAY()</f><v>45356</v></c>' + // formula cached value
        '</row>',
    })
    const text = await xlsxToText(bytes)
    expect(text).toContain(
      '45292 | 45292 | 1994-02-14 | 1900-02-29 | 2024-01-01 18:00 | 12:00:00 |  | 2024-03-05',
    )
  })

  it('honours the 1904 date system', async () => {
    const bytes = await workbook({
      date1904: true,
      xfs: '<xf numFmtId="14"/>',
      sheetXml: '<row r="1"><c r="A1" s="0"><v>0</v></c><c r="B1" s="0"><v>43830</v></c></row>',
    })
    expect(await xlsxToText(bytes)).toContain('1904-01-01 | 2024-01-01')
  })

  it('keeps the raw serial when the style index is out of range or the value is not numeric', async () => {
    const bytes = await workbook({
      xfs: '<xf numFmtId="14"/>',
      sheetXml: '<row r="1"><c r="A1" s="7"><v>45292</v></c><c r="B1" s="0"><v>abc</v></c></row>',
    })
    expect(await xlsxToText(bytes)).toContain('45292 | abc')
  })
})
