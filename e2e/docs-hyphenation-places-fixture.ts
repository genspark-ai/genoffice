/** A narrow-column document with long words in every place text can sit outside the body. */
import type { Page } from '@playwright/test'
import JSZip from 'jszip'

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
const NS =
  `${W} ` +
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml'

const LOWER =
  'internationalization of organizational responsibilities demonstrates characteristically ' +
  'disproportionate institutional representation throughout intercontinental telecommunication ' +
  'infrastructure notwithstanding considerable administrative misunderstanding concerning ' +
  'environmental sustainability recommendations'
const CAPS =
  'INTERNATIONALIZATION OF ORGANIZATIONAL RESPONSIBILITIES DEMONSTRATES CHARACTERISTICALLY ' +
  'DISPROPORTIONATE INSTITUTIONAL REPRESENTATION THROUGHOUT INTERCONTINENTAL TELECOMMUNICATION'

/** each place opens its lower-case paragraph with its own name and its capitals with the name in capitals */
export const PLACES = [
  'body',
  'cell',
  'boxed',
  'frozen',
  'header',
  'footer',
  'footnote',
  'endnote',
] as const
export type Place = (typeof PLACES)[number]

// a face named on the run: every copy of a strip wraps in the same one
const run = (t: string) =>
  `<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/></w:rPr><w:t xml:space="preserve">${t}</w:t></w:r>`
const para = (t: string) => `<w:p>${run(t)}</w:p>`
const SHORT = LOWER.split(' ').slice(0, 9).join(' ')
const SHORT_CAPS = CAPS.split(' ').slice(0, 7).join(' ')
/** headers, footers and notes hold a few lines, so the page keeps room for its body */
const brief = (place: Place) =>
  para(`${place} ${SHORT}`) + para(`${place.toUpperCase()} ${SHORT_CAPS}`)

const both = (place: Place) => para(`${place} ${LOWER}`) + para(`${place.toUpperCase()} ${CAPS}`)

const emu = (px: number) => px * 9525

function textBox(id: number, y: number, h: number, content: string): string {
  const w = 250
  return (
    `<w:p><w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" ` +
    `relativeHeight="${251659264 + id}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">` +
    `<wp:simplePos x="0" y="0"/>` +
    `<wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>` +
    `<wp:positionV relativeFrom="paragraph"><wp:posOffset>${emu(y)}</wp:posOffset></wp:positionV>` +
    `<wp:extent cx="${emu(w)}" cy="${emu(h)}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>` +
    `<wp:wrapTopAndBottom/><wp:docPr id="${id}" name="Text Box ${id}"/><wp:cNvGraphicFramePr/>` +
    `<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">` +
    `<wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${emu(w)}" cy="${emu(h)}"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln w="9525"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></wps:spPr>` +
    `<wps:txbx><w:txbxContent>${content}</w:txbxContent></wps:txbx>` +
    `<wps:bodyPr rot="0" vert="horz" wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="t"/>` +
    `</wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p>`
  )
}

const table = (content: string, width = 3600) =>
  `<w:tbl><w:tblPr><w:tblW w:w="${width}" w:type="dxa"/><w:tblBorders>` +
  ['top', 'left', 'bottom', 'right']
    .map((side) => `<w:${side} w:val="single" w:sz="4" w:space="0" w:color="000000"/>`)
    .join('') +
  `</w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="${width}"/></w:tblGrid>` +
  `<w:tr><w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/></w:tcPr>${content}</w:tc></w:tr></w:tbl>`

const note = (kind: 'footnote' | 'endnote', id: number, content: string) =>
  `<w:${kind} w:id="${id}">${content}</w:${kind}>`
const noteParts = (kind: 'footnote' | 'endnote') =>
  `<?xml version="1.0" encoding="UTF-8"?><w:${kind}s ${W}>` +
  `<w:${kind} w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:${kind}>` +
  `<w:${kind} w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:${kind}>` +
  note(
    kind,
    1,
    `<w:p><w:r><w:${kind}Ref/></w:r>${run(` ${kind} ${SHORT}`)}</w:p>` +
      `<w:p>${run(`${kind.toUpperCase()} ${SHORT_CAPS}`)}</w:p>`,
  ) +
  `</w:${kind}s>`

export interface PlacesOptions {
  /** children of w:settings */
  settings: string
  /** how many extra editable text boxes follow, each on its own paragraph */
  extraBoxes?: number
  /** body paragraphs after the fixed part, to give the document several pages */
  filler?: number
}

export async function hyphenationPlacesDocx(opts: PlacesOptions): Promise<Buffer> {
  const sectPr =
    '<w:sectPr><w:headerReference w:type="default" r:id="rIdH"/><w:footerReference w:type="default" r:id="rIdF"/>' +
    '<w:pgSz w:w="12240" w:h="15840"/>' +
    '<w:pgMar w:top="1440" w:right="4100" w:bottom="1440" w:left="4100" w:header="720" w:footer="720" w:gutter="0"/>' +
    '<w:cols w:space="720"/></w:sectPr>'
  const body =
    `<w:p>${run(`body ${LOWER}`)}<w:r><w:footnoteReference w:id="1"/></w:r></w:p>` +
    `<w:p>${run(`BODY ${CAPS}`)}<w:r><w:endnoteReference w:id="1"/></w:r></w:p>` +
    table(both('cell')) +
    para('') +
    textBox(1, 0, 420, both('boxed')) +
    textBox(2, 0, 460, table(both('frozen'), 3200) + para('')) +
    Array.from({ length: opts.extraBoxes ?? 0 }, (_, i) =>
      textBox(10 + i, 0, 420, both('boxed')),
    ).join('') +
    Array.from({ length: opts.filler ?? 0 }, (_, i) => para(`filler ${i + 1} ${LOWER}`)).join('') +
    sectPr
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      `<Override PartName="/word/document.xml" ContentType="${TYPE}.document.main+xml"/>` +
      `<Override PartName="/word/settings.xml" ContentType="${TYPE}.settings+xml"/>` +
      `<Override PartName="/word/footnotes.xml" ContentType="${TYPE}.footnotes+xml"/>` +
      `<Override PartName="/word/endnotes.xml" ContentType="${TYPE}.endnotes+xml"/>` +
      `<Override PartName="/word/header1.xml" ContentType="${TYPE}.header+xml"/>` +
      `<Override PartName="/word/footer1.xml" ContentType="${TYPE}.footer+xml"/></Types>`,
  )
  zip.file(
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/></Relationships>`,
  )
  zip.file(
    'word/_rels/document.xml.rels',
    `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rIdH" Type="${REL}/header" Target="header1.xml"/>` +
      `<Relationship Id="rIdF" Type="${REL}/footer" Target="footer1.xml"/>` +
      `<Relationship Id="rIdS" Type="${REL}/settings" Target="settings.xml"/>` +
      `<Relationship Id="rIdFn" Type="${REL}/footnotes" Target="footnotes.xml"/>` +
      `<Relationship Id="rIdEn" Type="${REL}/endnotes" Target="endnotes.xml"/></Relationships>`,
  )
  zip.file(
    'word/settings.xml',
    `<?xml version="1.0" encoding="UTF-8"?><w:settings ${W}>${opts.settings}</w:settings>`,
  )
  zip.file(
    'word/header1.xml',
    `<?xml version="1.0" encoding="UTF-8"?><w:hdr ${W}>${brief('header')}</w:hdr>`,
  )
  zip.file(
    'word/footer1.xml',
    `<?xml version="1.0" encoding="UTF-8"?><w:ftr ${W}>${brief('footer')}</w:ftr>`,
  )
  zip.file('word/footnotes.xml', noteParts('footnote'))
  zip.file('word/endnotes.xml', noteParts('endnote'))
  zip.file(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8"?><w:document ${NS}><w:body>${body}</w:body></w:document>`,
  )
  return zip.generateAsync({ type: 'nodebuffer' })
}

export interface PlaceLines {
  /** how many elements on screen hold this text */
  copies: number
  lines: number
  /** words laid out across a line end */
  broken: number
  /** the longest run of lines in a row ending in such a word */
  longest: number
  /** break points of ours in the paragraph */
  candidates: number
  /** the paragraph or an ancestor carries the class of our engine */
  ours: boolean
  /** height of the element holding the text: the strip, the box, the cell, the note */
  height: number
  /** the computed CSS hyphens and the language Chromium breaks under */
  hyphens: string
  lang: string | null
}

/** where each place is drawn on screen */
export const PLACE_SELECTOR: Record<Place, string> = {
  body: '.ProseMirror.doc-page > p',
  cell: '.ProseMirror.doc-page td',
  boxed: '.doc-textbox-editor',
  frozen: '.doc-textbox-row',
  header: '.page-hf-header',
  footer: '.page-hf-footer',
  footnote:
    '.page-gap-notes .page-gap-note, .page-notes:not(.page-endnotes) .page-note, .pv-footnote',
  endnote: '.page-endnotes .page-note, .pv-footnote',
}

/**
 * How the text opening with `start` is laid out inside the `nth` element
 * matching `selector`: the words from `start` up to the next paragraph.
 */
export const placeLines = (
  page: Page,
  start: string,
  selector: string,
  nth = 0,
): Promise<PlaceLines | null> =>
  page.evaluate(
    ({ start, selector, nth }) => {
      const hosts = Array.from(document.querySelectorAll<HTMLElement>(selector)).filter(
        (el) =>
          el.getClientRects().length > 0 &&
          (el.textContent ?? '').replace(/\u00ad/g, '').includes(`${start} `),
      )
      const host = hosts[nth]
      if (!host) return null
      const chars: Array<{ node: Text; offset: number; ch: string }> = []
      const walker = document.createTreeWalker(host, NodeFilter.SHOW_TEXT)
      for (let t = walker.nextNode() as Text | null; t; t = walker.nextNode() as Text | null)
        for (let i = 0; i < t.data.length; i++) chars.push({ node: t, offset: i, ch: t.data[i] })
      const flat = chars.map((c) => c.ch).join('')
      const from = flat.replace(/\u00ad/g, '\u0000').indexOf(`${start} `)
      const lower = /\p{Ll}/u.test(start)
      // the paragraph ends where the letter case turns
      let to = from
      while (
        to < chars.length &&
        !(lower ? /\p{Lu}/u : /\p{Ll}/u).test(chars[to].ch) &&
        !/\d/.test(chars[to].ch)
      )
        to++
      const range = document.createRange()
      const topOf = (i: number) => {
        range.setStart(chars[i].node, chars[i].offset)
        range.setEnd(chars[i].node, chars[i].offset + 1)
        return range.getBoundingClientRect().top
      }
      let broken = 0
      const brokenTops: number[] = []
      const tops: number[] = []
      for (let i = from; i < to;) {
        if (!/\p{L}/u.test(chars[i].ch)) {
          i++
          continue
        }
        let j = i
        while (j + 1 < to && /[\p{L}\u00ad]/u.test(chars[j + 1].ch)) j++
        const a = topOf(i)
        const b = topOf(j)
        if (Math.abs(a - b) > 4) {
          broken++
          brokenTops.push(a)
        }
        tops.push(a, b)
        i = j + 1
      }
      const lineTops = tops
        .sort((a, b) => a - b)
        .filter((top, i, sorted) => i === 0 || top - sorted[i - 1] > 4)
      const lines = lineTops.length
      const brokenLines = brokenTops.map((top) => lineTops.findIndex((t) => Math.abs(t - top) <= 4))
      let longest = 0
      let streak = 0
      brokenLines.forEach((line, i) => {
        streak = i > 0 && line === brokenLines[i - 1] + 1 ? streak + 1 : 1
        longest = Math.max(longest, streak)
      })
      const el = chars[from].node.parentElement!
      const block = el.closest<HTMLElement>('.doc-hyph')
      return {
        copies: hosts.length,
        lines,
        broken,
        longest,
        candidates: (block ?? host).querySelectorAll('.doc-shy').length,
        ours: block !== null,
        height: Math.round(host.getBoundingClientRect().height * 10) / 10,
        hyphens: getComputedStyle(el).hyphens,
        lang: el.closest('[lang]')?.getAttribute('lang') ?? null,
      }
    },
    { start, selector, nth },
  )
