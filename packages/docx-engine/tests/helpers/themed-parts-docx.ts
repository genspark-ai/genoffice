import { buildDocx } from './build-docx'
import { OFFICE_THEME_XML, THEMED_STYLES_XML } from './themed-docx'

const NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
const DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const CT = 'application/vnd.openxmlformats-officedocument.wordprocessingml'

const MAJOR = '<w:rFonts w:asciiTheme="majorHAnsi" w:hAnsiTheme="majorHAnsi"/>'
const tint = (hex: string, slot: string) => `<w:color w:val="${hex}" w:themeColor="${slot}"/>`
const run = (text: string, rPr: string) =>
  `<w:r><w:rPr>${rPr}</w:rPr><w:t xml:space="preserve">${text}</w:t></w:r>`
const para = (...runs: string[]) => `<w:p>${runs.join('')}</w:p>`

const cell = (text: string) => `<w:tc><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`
const HEADER_TABLE =
  '<w:tbl><w:tblPr><w:tblStyle w:val="Banded"/><w:tblW w:w="0" w:type="auto"/>' +
  '<w:tblLook w:val="0420" w:firstRow="1" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>' +
  '</w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>' +
  `<w:tr>${cell('HRegion')}${cell('HSales')}</w:tr><w:tr>${cell('HNorth')}${cell('10')}</w:tr></w:tbl>`

export const THEMED_PARTS = {
  'word/header1.xml':
    `${DECL}<w:hdr ${NS}>` +
    para(
      run('Head one', MAJOR + tint('ED7D31', 'accent2')),
      run(' fixed', '<w:color w:val="FF0000"/>'),
    ) +
    HEADER_TABLE +
    '<w:p/></w:hdr>',
  'word/header2.xml': `${DECL}<w:hdr ${NS}>${para(run('First head', tint('4472C4', 'accent1')))}</w:hdr>`,
  'word/header3.xml': `${DECL}<w:hdr ${NS}>${para(run('Head two', MAJOR + tint('70AD47', 'accent6')))}</w:hdr>`,
  'word/footer1.xml': `${DECL}<w:ftr ${NS}>${para(run('Page foot', MAJOR + tint('44546A', 'text2')))}</w:ftr>`,
  'word/footnotes.xml':
    `${DECL}<w:footnotes ${NS}>` +
    '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>' +
    '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>' +
    '<w:footnote w:id="1"><w:p><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteRef/></w:r>' +
    run(' Foot note', MAJOR + tint('ED7D31', 'accent2')) +
    '</w:p></w:footnote></w:footnotes>',
  'word/endnotes.xml':
    `${DECL}<w:endnotes ${NS}>` +
    '<w:endnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:endnote>' +
    '<w:endnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:endnote>' +
    '<w:endnote w:id="1"><w:p><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:endnoteRef/></w:r>' +
    run(' End note', MAJOR + tint('4472C4', 'accent1')) +
    '</w:p></w:endnote></w:endnotes>',
} as const

const PAGE =
  '<w:pgSz w:w="11906" w:h="16838"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/>'
const SUPER = '<w:rPr><w:vertAlign w:val="superscript"/></w:rPr>'
const BODY =
  '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Overview</w:t></w:r></w:p>' +
  `<w:p><w:r><w:t>plain</w:t></w:r><w:r>${SUPER}<w:footnoteReference w:id="1"/></w:r></w:p>` +
  '<w:p><w:r><w:br w:type="page"/></w:r></w:p><w:p><w:r><w:t>second page</w:t></w:r></w:p>' +
  '<w:p><w:pPr><w:sectPr><w:headerReference w:type="default" r:id="rIdHd1"/>' +
  '<w:headerReference w:type="first" r:id="rIdHd2"/><w:footerReference w:type="default" r:id="rIdFt1"/>' +
  `${PAGE}<w:titlePg/></w:sectPr></w:pPr></w:p>` +
  `<w:p><w:r><w:t>last section</w:t></w:r><w:r>${SUPER}<w:endnoteReference w:id="1"/></w:r></w:p>`

const part = (id: string, kind: string, target: string) =>
  `<Relationship Id="${id}" Type="${REL}/${kind}" Target="${target}"/>`

/**
 * Two sections on the Office theme: the first with a default header (a run on a
 * theme face and colour, a table on a themed style), a first-page header and a
 * footer; the second with a header of its own. A footnote and an endnote carry
 * themed runs too.
 */
export function themedPartsDocx(): Promise<Uint8Array> {
  return buildDocx({
    bodyXml: BODY,
    stylesXml: THEMED_STYLES_XML,
    sectPrXml: `<w:sectPr><w:headerReference w:type="default" r:id="rIdHd3"/>${PAGE}</w:sectPr>`,
    extraRels:
      part('rIdHd1', 'header', 'header1.xml') +
      part('rIdHd2', 'header', 'header2.xml') +
      part('rIdHd3', 'header', 'header3.xml') +
      part('rIdFt1', 'footer', 'footer1.xml') +
      part('rIdFn', 'footnotes', 'footnotes.xml') +
      part('rIdEn', 'endnotes', 'endnotes.xml') +
      part('rIdTheme', 'theme', 'theme/theme1.xml'),
    extraParts: [
      ...Object.entries(THEMED_PARTS).map(([path, xml]) => ({
        path,
        xml,
        contentType: `${CT}.${path.replace(/^word\/|\d*\.xml$/g, '')}+xml`,
      })),
      {
        path: 'word/theme/theme1.xml',
        xml: OFFICE_THEME_XML,
        contentType: 'application/vnd.openxmlformats-officedocument.theme+xml',
      },
    ],
  })
}
