/**
 * Hangul headers, footers and page numbers as Word header/footer parts.
 *
 * hp:header / hp:footer carry their own paragraphs; hp:pageNum draws the page
 * number at a page position (BOTTOM_CENTER "- 1 -" in most government
 * notices) with an optional side character on both sides; hp:pageHiding on
 * the first page hides it there (a cover), which Word spells as a different
 * first-page footer; hp:newNum restarts the count. Header and footer sit in
 * Hangul's header/footer bands: Word measures them from the page edge by the
 * Hangul top/bottom margins (the bands themselves are already folded into
 * the body margins, sectionSettings).
 */

import type JSZip from 'jszip'
import { addDocumentPart } from './parts'

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
const CT = 'application/vnd.openxmlformats-officedocument.wordprocessingml.'

export interface PageNumber {
  /** hp:pageNum pos: BOTTOM_CENTER, TOP_LEFT, … */
  pos: string
  formatType: string
  sideChar: string
}

export interface HeaderFooterParts {
  header: string[] | null
  footer: string[] | null
  pageNumber: PageNumber | null
  /** the page number is hidden on the first page (hp:pageHiding hidePageNum) */
  hideFirstPageNumber: boolean
  /** hp:newNum numType="PAGE" on the first section */
  pageStart: number | null
  /** Hangul top / bottom margins in twips: where the bands begin from the page edge */
  headerDistTw: number | null
  footerDistTw: number | null
}

const FIELD_FORMAT: Record<string, string> = {
  ROMAN_CAPITAL: ' \\* ROMAN',
  ROMAN_SMALL: ' \\* roman',
  LATIN_CAPITAL: ' \\* ALPHABETIC',
  LATIN_SMALL: ' \\* alphabetic',
}

const escapeText = (text: string): string =>
  text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')

/** a page number paragraph: side character, PAGE field, side character */
export function pageNumberParagraphXml(number: PageNumber, runPrXml: string): string {
  const align = /LEFT/.test(number.pos) ? 'left' : /RIGHT/.test(number.pos) ? 'right' : 'center'
  const side = number.sideChar ? escapeText(number.sideChar) : ''
  const text = (t: string) => `<w:r>${runPrXml}<w:t xml:space="preserve">${t}</w:t></w:r>`
  const field = (type: string) => `<w:r>${runPrXml}<w:fldChar w:fldCharType="${type}"/></w:r>`
  return (
    `<w:p><w:pPr><w:spacing w:before="0" w:after="0"/><w:jc w:val="${align}"/></w:pPr>` +
    (side ? text(`${side} `) : '') +
    field('begin') +
    `<w:r>${runPrXml}<w:instrText xml:space="preserve"> PAGE${FIELD_FORMAT[number.formatType] ?? ''} </w:instrText></w:r>` +
    field('separate') +
    text('1') +
    field('end') +
    (side ? text(` ${side}`) : '') +
    '</w:p>'
  )
}

/** header/footer parts may not reference the document's pictures: their runs are kept, drawings dropped */
function withoutDrawings(blocks: string[]): string[] {
  return blocks.map((xml) =>
    xml.replace(/<w:r>(?:(?!<\/w:r>).)*<w:drawing>.*?<\/w:drawing><\/w:r>/gs, ''),
  )
}

function partXml(root: 'hdr' | 'ftr', blocks: string[]): string {
  const body = blocks.filter((b) => /^<w:(p|tbl)[\s>/]/.test(b))
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    `<w:${root} xmlns:w="${W_NS}" xmlns:r="${R_NS}">${body.length > 0 ? body.join('') : '<w:p/>'}</w:${root}>`
  )
}

export async function addHeaderFooter(
  zip: JSZip,
  parts: HeaderFooterParts,
  pageNumberRunPr: string,
): Promise<void> {
  const number = parts.pageNumber && parts.pageNumber.pos !== 'NONE' ? parts.pageNumber : null
  const numberXml = number ? pageNumberParagraphXml(number, pageNumberRunPr) : null
  const header = [...withoutDrawings(parts.header ?? [])]
  const footer = [...withoutDrawings(parts.footer ?? [])]
  if (numberXml && /^TOP/.test(number!.pos)) header.push(numberXml)
  if (numberXml && !/^TOP/.test(number!.pos)) footer.push(numberXml)
  if (header.length === 0 && footer.length === 0 && parts.pageStart === null) return

  const addPart = async (name: string, kind: 'header' | 'footer', xml: string): Promise<string> =>
    (await addDocumentPart(zip, name, xml, `${REL}${kind}`, `${CT}${kind}+xml`, true)) ?? ''

  const refs: string[] = []
  if (header.length > 0) {
    refs.push(
      `<w:headerReference w:type="default" r:id="${await addPart('header1.xml', 'header', partXml('hdr', header))}"/>`,
    )
  }
  if (footer.length > 0) {
    refs.push(
      `<w:footerReference w:type="default" r:id="${await addPart('footer1.xml', 'footer', partXml('ftr', footer))}"/>`,
    )
  }
  // a cover page without its number: the first page's footer keeps the
  // footer paragraphs, only the page number goes
  const firstPage = parts.hideFirstPageNumber && numberXml !== null
  if (firstPage) {
    const firstHeader = withoutDrawings(parts.header ?? [])
    const firstFooter = withoutDrawings(parts.footer ?? [])
    if (header.length > 0) {
      refs.push(
        `<w:headerReference w:type="first" r:id="${await addPart('header2.xml', 'header', partXml('hdr', firstHeader))}"/>`,
      )
    }
    refs.push(
      `<w:footerReference w:type="first" r:id="${await addPart('footer2.xml', 'footer', partXml('ftr', firstFooter))}"/>`,
    )
  }
  const docPath = 'word/document.xml'
  let document = await zip.file(docPath)!.async('string')
  // the first section carries the references; later ones inherit them
  const at = document.indexOf('<w:sectPr')
  if (at >= 0) {
    const end = document.indexOf('</w:sectPr>', at)
    let sect = document.slice(at, end + '</w:sectPr>'.length)
    sect = sect.replace(/^<w:sectPr([^>]*)>/, (open) => `${open}${refs.join('')}`)
    sect = sect.replace(/<w:pgMar ([^>]*)\/>/, (tag, attrs: string) => {
      let next = attrs
      if (parts.headerDistTw !== null)
        next = next.replace(/w:header="\d+"/, `w:header="${parts.headerDistTw}"`)
      if (parts.footerDistTw !== null)
        next = next.replace(/w:footer="\d+"/, `w:footer="${parts.footerDistTw}"`)
      const pgNum = parts.pageStart !== null ? `<w:pgNumType w:start="${parts.pageStart}"/>` : ''
      return `<w:pgMar ${next}/>${pgNum}`
    })
    // CT_SectPr: titlePg precedes textDirection / bidi / docGrid
    if (firstPage) {
      sect = /<w:(?:textDirection|bidi|rtlGutter|docGrid)\b/.test(sect)
        ? sect.replace(/<w:(textDirection|bidi|rtlGutter|docGrid)\b/, '<w:titlePg/><w:$1')
        : sect.replace('</w:sectPr>', '<w:titlePg/></w:sectPr>')
    }
    document = document.slice(0, at) + sect + document.slice(end + '</w:sectPr>'.length)
    zip.file(docPath, document)
  }
}
