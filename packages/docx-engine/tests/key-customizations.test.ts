import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { parseDocx, readKeymaps, saveDocx, writeKeymaps, type SaveBlock } from '../src/index'
import { buildDocx } from './helpers/build-docx'

const WNE = 'xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"'
const R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
const BOLD =
  '<wne:keymap wne:kcmPrimary="0342"><wne:fci wne:fciName="Bold" wne:swArg="0000"/></wne:keymap>'
const MACRO =
  '<wne:keymap wne:kcmPrimary="0A4D"><wne:macro wne:macroName="PROJECT.MODULE1.RUN"/></wne:keymap>'
const SEQUENCE =
  '<wne:keymap wne:kcmPrimary="0342" wne:kcmSecondary="0441"><wne:fci wne:fciIndex="002D"/></wne:keymap>'
const MASK = '<wne:keymap wne:mask="1" wne:kcmPrimary="0242"/>'
const REST =
  '<wne:keymapsBad><wne:keymap wne:chmPrimary="06DF"><wne:fci wne:fciName="Italic"/></wne:keymap></wne:keymapsBad>' +
  '<wne:toolbars><wne:toolbarData r:id="rId1"/></wne:toolbars>'
const part = (keymaps: string) =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<wne:tcg ${R} ${WNE}>${keymaps}${REST}</wne:tcg>`
const FULL = part(`<wne:keymaps>${BOLD}${MACRO}${SEQUENCE}${MASK}</wne:keymaps>`)

const REL_TYPE = 'http://schemas.microsoft.com/office/2006/relationships/keyMapCustomizations'
const CONTENT_TYPE = 'application/vnd.ms-word.keyMapCustomizations+xml'

const allBlocks = (doc: Awaited<ReturnType<typeof parseDocx>>): SaveBlock[] =>
  doc.blocks.map((block) => ({ kind: 'original', docxIndex: block.docxIndex }) as SaveBlock)

describe('the key maps of a Customizations part', () => {
  it('reads the key, the fixed command and the mask of each entry', () => {
    expect(readKeymaps(FULL)).toEqual([
      { raw: BOLD, kcmPrimary: 0x0342, kcmSecondary: null, mask: false, fciName: 'Bold' },
      { raw: MACRO, kcmPrimary: 0x0a4d, kcmSecondary: null, mask: false, fciName: null },
      { raw: SEQUENCE, kcmPrimary: 0x0342, kcmSecondary: 0x0441, mask: false, fciName: null },
      { raw: MASK, kcmPrimary: 0x0242, kcmSecondary: null, mask: true, fciName: null },
    ])
    // key maps for another keyboard are not this keyboard's
    expect(readKeymaps(part(''))).toEqual([])
    expect(readKeymaps('<tcg xmlns="urn:other"/>')).toEqual([])
  })

  it('writes entries it did not make as they were, and leaves the rest of the part alone', () => {
    const [bold, macro, sequence] = readKeymaps(FULL)
    const made = { kcmPrimary: 0x0a4b, kcmSecondary: null, mask: false, fciName: 'Italic' }
    const masked = { kcmPrimary: 0x0849, kcmSecondary: null, mask: true, fciName: null }
    const out = writeKeymaps(FULL, [macro, sequence, made, masked])!
    expect(out).toBe(
      part(
        `<wne:keymaps>${MACRO}${SEQUENCE}` +
          '<wne:keymap wne:kcmPrimary="0A4B"><wne:fci wne:fciName="Italic" wne:swArg="0000"/></wne:keymap>' +
          '<wne:keymap wne:mask="1" wne:kcmPrimary="0849"/></wne:keymaps>',
      ),
    )
    expect(writeKeymaps(FULL, readKeymaps(FULL))).toBe(FULL)
    // a keymaps element cannot be empty: it goes, and comes back first in the tcg
    expect(writeKeymaps(FULL, [])).toBe(part(''))
    expect(writeKeymaps(part(''), [bold])).toBe(part(`<wne:keymaps>${BOLD}</wne:keymaps>`))
    expect(writeKeymaps(`<wne:tcg ${WNE}/>`, [bold])).toBe(
      `<wne:tcg ${WNE}><wne:keymaps>${BOLD}</wne:keymaps></wne:tcg>`,
    )
  })

  it('makes a part for a document without one, and none for nothing', () => {
    expect(writeKeymaps(null, [])).toBeNull()
    const out = writeKeymaps(null, [
      { kcmPrimary: 0x0342, kcmSecondary: null, mask: false, fciName: 'Bold' },
    ])!
    expect(out).toContain(`<wne:tcg ${R} ${WNE}><wne:keymaps>${BOLD}</wne:keymaps></wne:tcg>`)
    expect(readKeymaps(out)).toHaveLength(1)
    expect(writeKeymaps('<tcg xmlns="urn:other"/>', [])).toBeNull()
  })

  it('follows the prefix the part uses', () => {
    const xml =
      '<k:tcg xmlns:k="http://schemas.microsoft.com/office/word/2006/wordml"><k:keymaps>' +
      '<k:keymap k:kcmPrimary="0242"><k:fci k:fciName="Bold"/></k:keymap></k:keymaps></k:tcg>'
    const [entry] = readKeymaps(xml)
    expect(entry).toMatchObject({ kcmPrimary: 0x0242, fciName: 'Bold' })
    expect(
      writeKeymaps(xml, [
        entry,
        { kcmPrimary: 0x0849, kcmSecondary: null, mask: true, fciName: null },
      ]),
    ).toContain('</k:keymap><k:keymap k:mask="1" k:kcmPrimary="0849"/></k:keymaps>')
  })
})

describe('a document that carries keyboard customizations', () => {
  const withPart = () =>
    buildDocx({
      bodyXml: '<w:p><w:r><w:t>Keys</w:t></w:r></w:p>',
      extraRels: `<Relationship Id="rIdKeys" Type="${REL_TYPE}" Target="customizations.xml"/>`,
      extraParts: [{ path: 'word/customizations.xml', xml: FULL, contentType: CONTENT_TYPE }],
    })

  it('hands the part to the host, and saves it byte for byte when nothing asks otherwise', async () => {
    const bytes = await withPart()
    const doc = await parseDocx(bytes)
    expect(doc.keyCustomizations).toBe(FULL)
    const out = await JSZip.loadAsync(
      await saveDocx(doc, allBlocks(doc), { zoteroDocumentData: 'prefs' }),
    )
    expect(await out.file('word/customizations.xml')!.async('string')).toBe(FULL)
  })

  it('replaces the part in place', async () => {
    const doc = await parseDocx(await withPart())
    const next = writeKeymaps(FULL, readKeymaps(FULL).slice(1))!
    const out = await JSZip.loadAsync(
      await saveDocx(doc, allBlocks(doc), { keyCustomizations: next }),
    )
    expect(await out.file('word/customizations.xml')!.async('string')).toBe(next)
    const rels = await out.file('word/_rels/document.xml.rels')!.async('string')
    expect(rels.match(/keyMapCustomizations/g)).toHaveLength(1)
  })

  it('adds the part, its relationship and its content type to a document without one', async () => {
    const doc = await parseDocx(
      await buildDocx({ bodyXml: '<w:p><w:r><w:t>Plain</w:t></w:r></w:p>' }),
    )
    expect(doc.keyCustomizations).toBeUndefined()
    const xml = writeKeymaps(null, [
      { kcmPrimary: 0x0a4b, kcmSecondary: null, mask: false, fciName: 'Italic' },
    ])!
    const saved = await saveDocx(doc, allBlocks(doc), { keyCustomizations: xml })
    const out = await JSZip.loadAsync(saved)
    expect(await out.file('word/customizations.xml')!.async('string')).toBe(xml)
    expect(await out.file('word/_rels/document.xml.rels')!.async('string')).toMatch(
      new RegExp(`Type="${REL_TYPE}" Target="customizations.xml"`),
    )
    expect(await out.file('[Content_Types].xml')!.async('string')).toContain(
      `<Override PartName="/word/customizations.xml" ContentType="${CONTENT_TYPE}"/>`,
    )
    expect((await parseDocx(saved)).keyCustomizations).toBe(xml)
  })
})
