/** A three-page document with what Draft view leaves out: a header, a footer and a floating shape. */
import JSZip from 'jszip'
import { shapeParagraph } from './docs-objects-fixture'

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
const NS =
  `${W} ` +
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml'

const para = (t: string) => `<w:p><w:r><w:t>${t}</w:t></w:r></w:p>`
const sectPr = (extra = '') =>
  `<w:sectPr><w:headerReference w:type="default" r:id="rIdH"/><w:footerReference w:type="default" r:id="rIdF"/>${extra}` +
  '<w:pgSz w:w="12240" w:h="15840"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>'

export const DRAFT_HEADER = 'Quarterly report header'
export const DRAFT_FOOTER = 'Quarterly report footer'

/**
 * Page 1 runs over into page 2 on its own (an automatic break), page 2 ends
 * at a manual page break, and a next-page section break follows two lines
 * into page 3.
 */
export async function draftDocx(): Promise<Buffer> {
  const lines = (prefix: string, n: number) =>
    Array.from({ length: n }, (_, i) => para(`${prefix} line ${i + 1}`)).join('')
  const body =
    para('Opening paragraph') +
    shapeParagraph(1, 120, 10, 140, 60) +
    lines('First', 60) +
    '<w:p><w:r><w:t>Before the manual break</w:t></w:r><w:r><w:br w:type="page"/></w:r></w:p>' +
    lines('Second', 2) +
    `<w:p><w:pPr>${sectPr()}</w:pPr><w:r><w:t>End of section one</w:t></w:r></w:p>` +
    lines('Third', 3) +
    sectPr()
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      `<Override PartName="/word/document.xml" ContentType="${TYPE}.document.main+xml"/>` +
      `<Override PartName="/word/header1.xml" ContentType="${TYPE}.header+xml"/>` +
      `<Override PartName="/word/footer1.xml" ContentType="${TYPE}.footer+xml"/></Types>`,
  )
  zip.file(
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/></Relationships>`,
  )
  zip.file(
    'word/_rels/document.xml.rels',
    `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdH" Type="${REL}/header" Target="header1.xml"/><Relationship Id="rIdF" Type="${REL}/footer" Target="footer1.xml"/></Relationships>`,
  )
  zip.file(
    'word/header1.xml',
    `<?xml version="1.0" encoding="UTF-8"?><w:hdr ${W}>${para(DRAFT_HEADER)}</w:hdr>`,
  )
  zip.file(
    'word/footer1.xml',
    `<?xml version="1.0" encoding="UTF-8"?><w:ftr ${W}>${para(DRAFT_FOOTER)}</w:ftr>`,
  )
  zip.file(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8"?><w:document ${NS}><w:body>${body}</w:body></w:document>`,
  )
  return zip.generateAsync({ type: 'nodebuffer' })
}
