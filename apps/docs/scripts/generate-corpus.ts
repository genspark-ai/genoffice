/**
 * Pagination corpus generation script
 * Run: tsx apps/docs/scripts/generate-corpus.ts
 * Output: apps/docs/tests/pagination-corpus/docx/*.docx
 *       apps/docs/tests/pagination-corpus/meta/*.json
 */

import { mkdirSync, writeFileSync, copyFileSync, existsSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import JSZip from 'jszip'

const __dirname = dirname(fileURLToPath(import.meta.url))
const REPO_ROOT = join(__dirname, '../../..')
const DOCX_OUT = join(__dirname, '../tests/pagination-corpus/docx')
const META_OUT = join(__dirname, '../tests/pagination-corpus/meta')

mkdirSync(DOCX_OUT, { recursive: true })
mkdirSync(META_OUT, { recursive: true })

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
const DOC_NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"'

const STYLES_XML =
  XML_DECL +
  '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
  '<w:docDefaults><w:rPrDefault><w:rPr>' +
  '<w:rFonts w:ascii="Calibri" w:eastAsia="宋体" w:hAnsi="Calibri"/>' +
  '<w:sz w:val="24"/><w:szCs w:val="24"/>' +
  '</w:rPr></w:rPrDefault>' +
  '<w:pPrDefault><w:pPr>' +
  '<w:spacing w:after="160" w:line="276" w:lineRule="auto"/>' +
  '</w:pPr></w:pPrDefault></w:docDefaults>' +
  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/>' +
  '<w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>' +
  '<w:pPr><w:keepNext/><w:spacing w:before="240" w:after="120"/><w:outlineLvl w:val="0"/></w:pPr>' +
  '<w:rPr><w:b/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/>' +
  '<w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>' +
  '<w:pPr><w:keepNext/><w:spacing w:before="200" w:after="120"/><w:outlineLvl w:val="1"/></w:pPr>' +
  '<w:rPr><w:b/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="heading 3"/>' +
  '<w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>' +
  '<w:pPr><w:keepNext/><w:spacing w:before="160" w:after="120"/><w:outlineLvl w:val="2"/></w:pPr>' +
  '<w:rPr><w:b/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/>' +
  '<w:basedOn w:val="Normal"/><w:pPr><w:ind w:left="720"/></w:pPr></w:style>' +
  '</w:styles>'

// A4 page + standard margins + docGrid (Chinese document grid)
const A4_SECT_PR =
  '<w:sectPr>' +
  '<w:pgSz w:w="11906" w:h="16838"/>' +
  '<w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="708" w:footer="708" w:gutter="0"/>' +
  '</w:sectPr>'

// A4 + docGrid (standard Chinese document grid: line=360, type=lines)
const A4_SECT_WITH_DOCGRID =
  '<w:sectPr>' +
  '<w:pgSz w:w="11906" w:h="16838"/>' +
  '<w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="708" w:footer="708" w:gutter="0"/>' +
  '<w:docGrid w:type="lines" w:linePitch="312"/>' +
  '</w:sectPr>'

// US Letter landscape
const LETTER_LANDSCAPE_SECT_PR =
  '<w:sectPr>' +
  '<w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/>' +
  '</w:sectPr>'

interface DocSpec {
  name: string
  bodyXml: string
  sectPr?: string
  extraStylesXml?: string
  footnotesXml?: string
  meta: {
    title: string
    tags: string[]
    description: string
  }
}

async function buildDocx(spec: DocSpec): Promise<Uint8Array> {
  const zip = new JSZip()
  const sectPr = spec.sectPr ?? A4_SECT_PR
  const extraStyles = spec.extraStylesXml ?? ''
  const hasFootnotes = !!spec.footnotesXml

  const stylesXml = extraStyles
    ? STYLES_XML.replace('</w:styles>', extraStyles + '</w:styles>')
    : STYLES_XML

  const contentTypeOverrides = [
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>',
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>',
    hasFootnotes
      ? '<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>'
      : '',
  ].join('')

  zip.file(
    '[Content_Types].xml',
    `${XML_DECL}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      contentTypeOverrides +
      '</Types>',
  )
  zip.file(
    '_rels/.rels',
    `${XML_DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
      '</Relationships>',
  )
  const docRels = [
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>',
    hasFootnotes
      ? '<Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>'
      : '',
  ].join('')
  zip.file(
    'word/_rels/document.xml.rels',
    `${XML_DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${docRels}</Relationships>`,
  )
  zip.file('word/styles.xml', stylesXml)
  if (hasFootnotes) zip.file('word/footnotes.xml', spec.footnotesXml!)
  zip.file(
    'word/document.xml',
    `${XML_DECL}<w:document ${DOC_NS}><w:body>${spec.bodyXml}${sectPr}</w:body></w:document>`,
  )
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })
}

// ---- Paragraph helpers ----

const p = (text: string, style?: string, extra?: string): string => {
  const pStyle = style ? `<w:pStyle w:val="${style}"/>` : ''
  const pPr = pStyle || extra ? `<w:pPr>${pStyle}${extra ?? ''}</w:pPr>` : ''
  return `<w:p>${pPr}<w:r><w:t xml:space="preserve">${escXml(text)}</w:t></w:r></w:p>`
}

const pBreak = (): string => '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'

const escXml = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')

// build a paragraph from repeated text
function loremParagraph(n: number, lang: 'en' | 'zh'): string {
  if (lang === 'en') {
    const words =
      'Lorem ipsum dolor sit amet consectetur adipiscing elit sed do eiusmod tempor incididunt ut labore et dolore magna aliqua Ut enim ad minim veniam quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat'.split(
        ' ',
      )
    let text = ''
    for (let i = 0; i < n; i++) text += words[i % words.length] + ' '
    return p(text.trim())
  } else {
    const chars =
      '在中国经济社会发展的重要历史时期，人工智能技术的广泛应用正在深刻改变各行各业的生产方式和商业模式。数字化转型已成为企业保持竞争力的关键战略，各类组织机构纷纷加快推进信息化建设步伐。'
    let text = ''
    for (let i = 0; i < n; i++) text += chars[i % chars.length]
    return p(text)
  }
}

// generate table rows
function tableRows(rows: number, cols: number, withHeader = false): string {
  let xml = ''
  const colWidth = Math.floor(8000 / cols)
  for (let r = 0; r < rows; r++) {
    const isHeader = withHeader && r === 0
    xml += '<w:tr>'
    for (let c = 0; c < cols; c++) {
      const fill = isHeader ? '<w:shd w:val="clear" w:color="auto" w:fill="D9D9D9"/>' : ''
      const cellText = isHeader ? `列${c + 1}标题` : `第${r + 1}行第${c + 1}列`
      xml +=
        `<w:tc><w:tcPr><w:tcW w:w="${colWidth}" w:type="dxa"/>${fill}</w:tcPr>` +
        `<w:p>${isHeader ? '<w:pPr/>' : ''}<w:r>${isHeader ? '<w:rPr><w:b/></w:rPr>' : ''}<w:t>${escXml(cellText)}</w:t></w:r></w:p></w:tc>`
    }
    xml += '</w:tr>'
  }
  return xml
}

function table(rows: number, cols: number, withHeader = false): string {
  const colWidth = Math.floor(8000 / cols)
  const border = (n: string) => `<w:${n} w:val="single" w:sz="4" w:space="0" w:color="auto"/>`
  const borders =
    '<w:tblBorders>' +
    ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map(border).join('') +
    '</w:tblBorders>'
  const gridCols = Array.from({ length: cols }, () => `<w:gridCol w:w="${colWidth}"/>`).join('')
  return (
    `<w:tbl><w:tblPr><w:tblW w:w="8000" w:type="dxa"/>${borders}</w:tblPr>` +
    `<w:tblGrid>${gridCols}</w:tblGrid>` +
    tableRows(rows, cols, withHeader) +
    '</w:tbl>'
  )
}

// ---- Corpus specs ----

const specs: DocSpec[] = []

// 01. Simple English document (2 pages)
{
  let body = p('Simple English Document', 'Heading1')
  for (let i = 0; i < 40; i++) body += loremParagraph(30, 'en')
  specs.push({
    name: '01-simple-english',
    bodyXml: body,
    meta: {
      title: 'Simple English Document',
      tags: ['english', 'multi-page', 'simple'],
      description: 'Pure English long text, ~2 pages, no special formatting',
    },
  })
}

// 02. Long pure-Chinese document (with docGrid)
{
  let body = p('中文长篇文档测试', 'Heading1')
  for (let i = 0; i < 50; i++) body += loremParagraph(50, 'zh')
  specs.push({
    name: '02-chinese-long-docgrid',
    bodyXml: body,
    sectPr: A4_SECT_WITH_DOCGRID,
    meta: {
      title: 'Long Chinese text (docGrid)',
      tags: ['chinese', 'docGrid', 'multi-page'],
      description:
        'Pure Chinese with a w:docGrid lines grid, typical Chinese official-document format',
    },
  })
}

// 03. Mixed Chinese-English
{
  let body = p('中英文混排文档', 'Heading1')
  for (let i = 0; i < 20; i++) {
    body += loremParagraph(20, 'zh')
    body += loremParagraph(20, 'en')
  }
  specs.push({
    name: '03-mixed-lang',
    bodyXml: body,
    meta: {
      title: 'Mixed Chinese-English',
      tags: ['mixed-language', 'multi-page'],
      description: 'Alternating Chinese/English paragraphs, tests mixed-language pagination',
    },
  })
}

// 04. Multilevel headings (keepNext styles)
{
  let body = ''
  for (let h = 1; h <= 5; h++) {
    body += p(`第${h}章 标题层级测试`, 'Heading1')
    body += p(`第${h}.1节`, 'Heading2')
    for (let i = 0; i < 8; i++) body += loremParagraph(30, 'zh')
    body += p(`第${h}.2节`, 'Heading2')
    body += p(`第${h}.2.1小节`, 'Heading3')
    for (let i = 0; i < 6; i++) body += loremParagraph(30, 'zh')
  }
  specs.push({
    name: '04-headings-keepnext',
    bodyXml: body,
    sectPr: A4_SECT_WITH_DOCGRID,
    meta: {
      title: 'Multilevel headings document',
      tags: ['headings', 'keepNext', 'multi-section', 'docGrid'],
      description:
        '5 chapters of multilevel headings (with keepNext styles), tests heading page-break behavior',
    },
  })
}

// 05. Long table across pages
{
  let body = p('长表格跨页测试', 'Heading1')
  body += p('以下表格含 30 行，跨多页：')
  body += table(30, 4, true)
  body += p('表格后续正文段落。')
  specs.push({
    name: '05-long-table',
    bodyXml: body,
    meta: {
      title: 'Long table across pages',
      tags: ['table', 'cross-page', 'table-header'],
      description: '30-row 4-column table, tests table splitting across pages',
    },
  })
}

// 06. With footnotes (full footnotes.xml structure)
{
  // footnote documents need footnotes.xml + a Content-Type override + rels
  // built directly with the customBuildDocxWithFootnotes function
  const FOOTNOTE_COUNT = 20
  let body = p('含脚注文档测试', 'Heading1')
  for (let i = 1; i <= FOOTNOTE_COUNT; i++) {
    body +=
      `<w:p><w:r><w:t xml:space="preserve">第${i}段正文，包含参考引用。</w:t></w:r>` +
      `<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="${i}"/></w:r></w:p>`
  }
  for (let i = 0; i < 20; i++) body += loremParagraph(20, 'zh')
  // generate footnotes.xml (with the separator + each footnote)
  const FOOTNOTES_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
  let footnotesXml = `${XML_DECL}<w:footnotes ${FOOTNOTES_NS}>`
  // two built-in system footnotes: separator and continuationSeparator
  footnotesXml +=
    '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>'
  footnotesXml +=
    '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>'
  for (let i = 1; i <= FOOTNOTE_COUNT; i++) {
    footnotesXml +=
      `<w:footnote w:id="${i}"><w:p>` +
      `<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:t>${i}</w:t></w:r>` +
      `<w:r><w:t xml:space="preserve"> 这是第${i}条脚注内容，详见参考文献。</w:t></w:r>` +
      `</w:p></w:footnote>`
  }
  footnotesXml += '</w:footnotes>'
  specs.push({
    name: '06-with-footnotes',
    bodyXml: body,
    footnotesXml,
    meta: {
      title: 'Document with footnotes',
      tags: ['footnotes', 'multi-page'],
      description: "20 footnote references, tests footnote placeholders' impact on pagination",
    },
  })
}

// 07. With page breaks
{
  let body = p('分页符测试文档', 'Heading1')
  body += p('第一页内容')
  for (let i = 0; i < 5; i++) body += loremParagraph(20, 'zh')
  body += pBreak()
  body += p('第二页（强制分页符后）')
  for (let i = 0; i < 5; i++) body += loremParagraph(20, 'zh')
  body += pBreak()
  body += p('第三页')
  for (let i = 0; i < 8; i++) body += loremParagraph(20, 'zh')
  specs.push({
    name: '07-page-breaks',
    bodyXml: body,
    meta: {
      title: 'Manual page breaks',
      tags: ['page-break', 'forced-break'],
      description: 'Two manual page breaks, tests breakAfter logic',
    },
  })
}

// 08. Multi-section mixed paper (A4 portrait + A4 landscape)
{
  const sect1 =
    '<w:sectPr><w:type w:val="continuous"/>' +
    '<w:pgSz w:w="11906" w:h="16838"/>' +
    '<w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="708" w:footer="708" w:gutter="0"/>' +
    '</w:sectPr>'
  let body = p('多节文档 - 第一节（A4 竖向）', 'Heading1')
  for (let i = 0; i < 15; i++) body += loremParagraph(25, 'zh')
  // insert a section break (sectPr embedded in the paragraph's pPr)
  body += `<w:p><w:pPr>${sect1}</w:pPr></w:p>`
  body += p('第二节（A4 横向）', 'Heading1')
  for (let i = 0; i < 10; i++) body += loremParagraph(25, 'zh')
  specs.push({
    name: '08-multi-section-paper',
    bodyXml: body,
    sectPr: LETTER_LANDSCAPE_SECT_PR,
    meta: {
      title: 'Multi-section mixed paper sizes',
      tags: ['multi-section', 'landscape', 'portrait', 'section-break'],
      description: 'Two sections: A4 portrait + landscape, tests pagination at section switches',
    },
  })
}

// 09. Large paragraph (within one page)
{
  let body = p('大段落测试', 'Heading1')
  // one large paragraph, close to but not exceeding a page
  let bigPara = ''
  for (let i = 0; i < 200; i++) bigPara += '这是一段很长的文字内容，用于测试大段落的分页处理行为。'
  body += `<w:p><w:r><w:t xml:space="preserve">${escXml(bigPara)}</w:t></w:r></w:p>`
  body += p('段落后继续内容')
  for (let i = 0; i < 5; i++) body += loremParagraph(30, 'zh')
  specs.push({
    name: '09-large-paragraph',
    bodyXml: body,
    meta: {
      title: 'Large paragraph (cross-page)',
      tags: ['large-paragraph', 'cross-page'],
      description: 'One oversized paragraph spanning pages, tests intra-block splitting',
    },
  })
}

// 10. Chinese official-document format (docGrid + standard page margins)
{
  let body = ''
  body +=
    `<w:p><w:pPr><w:jc w:val="center"/><w:spacing w:before="480" w:after="240"/></w:pPr>` +
    `<w:r><w:rPr><w:b/><w:sz w:val="44"/><w:szCs w:val="44"/></w:rPr><w:t>关于深化体制改革的通知</w:t></w:r></w:p>`
  body += `<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>各省、自治区、直辖市人民政府：</w:t></w:r></w:p>`
  for (let i = 0; i < 8; i++) {
    body +=
      `<w:p><w:pPr><w:ind w:firstLine="480"/></w:pPr>` +
      `<w:r><w:t xml:space="preserve">${escXml('　　' + '为贯彻落实党中央、国务院关于深化体制改革的重要决策部署，切实推进政府职能转变，优化营商环境，提升治理效能，现就有关事项通知如下：'.repeat(3))}</w:t></w:r></w:p>`
  }
  body += `<w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:t>国务院办公厅</w:t></w:r></w:p>`
  body += `<w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:t>二○二六年七月二十六日</w:t></w:r></w:p>`
  // attachments
  for (let i = 1; i <= 3; i++) {
    body += p(`附件${i}：相关事项说明`, 'Heading2')
    for (let j = 0; j < 10; j++) body += loremParagraph(40, 'zh')
  }
  specs.push({
    name: '10-cn-official-doc',
    bodyXml: body,
    sectPr: A4_SECT_WITH_DOCGRID,
    meta: {
      title: 'Chinese official-document format',
      tags: ['chinese', 'official-doc', 'docGrid', 'firstLine-indent', 'center-justify'],
      description:
        'Standard government document format: centered large title + first-line-indented body + signature block, docGrid',
    },
  })
}

// 11. Table-dense (multiple tables)
{
  let body = p('多表格文档', 'Heading1')
  for (let t = 0; t < 5; t++) {
    body += p(`表格 ${t + 1}`)
    body += table(8, 3, true)
    body += p('表格说明文字：以上数据来自系统统计。')
  }
  specs.push({
    name: '11-multi-tables',
    bodyXml: body,
    meta: {
      title: 'Multi-table document',
      tags: ['table', 'multi-table', 'table-header'],
      description: '5 tables interleaved with text, tests pagination of text between tables',
    },
  })
}

// 12. Orphan heading test (heading right at the page bottom)
{
  let body = p('标题孤悬测试', 'Heading1')
  // fill until close to the page end
  for (let i = 0; i < 30; i++) body += loremParagraph(15, 'en')
  // the heading here should sit at the page bottom, with the following content on the next page
  body += p('这是一个可能孤悬的二级标题', 'Heading2')
  body += p('标题后面的正文段落，标题不应孤悬。')
  for (let i = 0; i < 10; i++) body += loremParagraph(20, 'en')
  specs.push({
    name: '12-orphan-heading',
    bodyXml: body,
    meta: {
      title: 'Orphan heading test',
      tags: ['headings', 'keepNext', 'orphan', 'widow'],
      description: 'Heading lands right at the page bottom, tests keepNext orphan prevention',
    },
  })
}

// 13. Long pure-English document (3+ pages)
{
  let body = p('Long English Report', 'Heading1')
  body += p('Abstract', 'Heading2')
  for (let i = 0; i < 15; i++) body += loremParagraph(40, 'en')
  body += p('Introduction', 'Heading2')
  for (let i = 0; i < 20; i++) body += loremParagraph(35, 'en')
  body += p('Methodology', 'Heading2')
  for (let i = 0; i < 20; i++) body += loremParagraph(35, 'en')
  body += p('Results', 'Heading2')
  for (let i = 0; i < 15; i++) body += loremParagraph(30, 'en')
  body += p('Conclusion', 'Heading2')
  for (let i = 0; i < 10; i++) body += loremParagraph(30, 'en')
  specs.push({
    name: '13-english-report',
    bodyXml: body,
    meta: {
      title: 'English Report (3+ pages)',
      tags: ['english', 'headings', 'multi-page', 'long'],
      description: 'English report format, 4-section heading structure, 3+ pages',
    },
  })
}

// 14. Long Chinese document (no docGrid)
{
  let body = p('中文文档（无网格）', 'Heading1')
  for (let i = 0; i < 60; i++) body += loremParagraph(40, 'zh')
  specs.push({
    name: '14-chinese-no-docgrid',
    bodyXml: body,
    meta: {
      title: 'Long Chinese text (no docGrid)',
      tags: ['chinese', 'no-docGrid', 'multi-page'],
      description: 'Pure Chinese without docGrid, contrast test against 02',
    },
  })
}

// 15. Mixed elements (headings + tables + body)
{
  let body = p('混合内容文档', 'Heading1')
  for (let section = 1; section <= 4; section++) {
    body += p(`第${section}节 内容概述`, 'Heading2')
    for (let i = 0; i < 5; i++) body += loremParagraph(25, 'zh')
    body += p(`第${section}.1 数据汇总`, 'Heading3')
    body += table(6, 3, true)
    for (let i = 0; i < 3; i++) body += loremParagraph(20, 'zh')
    body += p(`第${section}.2 分析结论`, 'Heading3')
    for (let i = 0; i < 6; i++) body += loremParagraph(25, 'zh')
  }
  specs.push({
    name: '15-mixed-content',
    bodyXml: body,
    sectPr: A4_SECT_WITH_DOCGRID,
    meta: {
      title: 'Mixed content (headings + tables + body)',
      tags: ['mixed', 'table', 'headings', 'docGrid'],
      description: 'Headings, tables and body text mixed — the most common report structure',
    },
  })
}

// 16. Numbered-list document
{
  const _NUMBERING_XML =
    XML_DECL +
    '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/>' +
    '<w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum>' +
    '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>' +
    '</w:numbering>'
  // inline numbering into extra styles (no separate file; simplified approach)
  let body = p('编号列表文档', 'Heading1')
  for (let i = 1; i <= 50; i++) {
    body +=
      `<w:p><w:pPr><w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="0"/></w:numPr></w:pPr>` +
      `<w:r><w:t xml:space="preserve">${i}. 列表项目：${escXml('这是第' + i + '条列表内容，用于测试编号列表的分页行为。')}</w:t></w:r></w:p>`
  }
  specs.push({
    name: '16-numbered-list',
    bodyXml: body,
    meta: {
      title: 'Numbered-list document',
      tags: ['numbered-list', 'list-paragraph', 'multi-page'],
      description: '50 numbered list items, tests list pagination across pages',
    },
  })
}

// 17. Short single-page document
{
  let body = p('单页短文档', 'Heading1')
  body += p('这是一份很短的文档，所有内容都在第一页内。')
  body += p('第二段内容。')
  body += p('第三段内容，文档到此结束。')
  specs.push({
    name: '17-single-page',
    bodyXml: body,
    meta: {
      title: 'Short single-page document',
      tags: ['single-page', 'short'],
      description: 'All content fits on one page, baseline case',
    },
  })
}

// 18. Paragraphs with the pageBreakBefore property
{
  let body = p('pageBreakBefore 测试', 'Heading1')
  for (let i = 0; i < 5; i++) body += loremParagraph(20, 'zh')
  body += `<w:p><w:pPr><w:pageBreakBefore/></w:pPr><w:r><w:t>此段落设置了 pageBreakBefore 属性</w:t></w:r></w:p>`
  for (let i = 0; i < 5; i++) body += loremParagraph(20, 'zh')
  body += `<w:p><w:pPr><w:pageBreakBefore/></w:pPr><w:r><w:t>第二个 pageBreakBefore 段落</w:t></w:r></w:p>`
  for (let i = 0; i < 5; i++) body += loremParagraph(20, 'zh')
  specs.push({
    name: '18-page-break-before',
    bodyXml: body,
    meta: {
      title: 'pageBreakBefore property',
      tags: ['pageBreakBefore', 'forced-break'],
      description: 'Two paragraphs with the pageBreakBefore property, tests breakBefore logic',
    },
  })
}

// 19. Large table (over one page)
{
  let body = p('超大表格测试', 'Heading1')
  body += p('以下表格超过一页：')
  body += table(60, 3, true)
  body += p('超大表格后续内容。')
  specs.push({
    name: '19-oversized-table',
    bodyXml: body,
    meta: {
      title: 'Oversized table (>1 page)',
      tags: ['table', 'cross-page', 'large-table', 'pixel-cut'],
      description: '60-row table exceeding one page, triggers the pixel hard-cut logic',
    },
  })
}

// 20. Chinese official document (shorter, ~1.5 pages)
{
  let body =
    `<w:p><w:pPr><w:jc w:val="center"/></w:pPr>` +
    `<w:r><w:rPr><w:b/><w:sz w:val="36"/></w:rPr><w:t>人工智能办公系统建设方案</w:t></w:r></w:p>`
  body += `<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>信息化建设工作小组</w:t></w:r></w:p>`
  body += `<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>2026年7月</w:t></w:r></w:p>`
  for (let i = 0; i < 12; i++) {
    body +=
      `<w:p><w:pPr><w:ind w:firstLine="480"/></w:pPr>` +
      `<w:r><w:t xml:space="preserve">${escXml('　　本方案旨在系统推进办公自动化智能升级，通过引入先进的人工智能技术，全面提升行政办公效率，降低运营成本，改善用户体验，为现代化治理体系建设提供有力支撑。'.repeat(2))}</w:t></w:r></w:p>`
  }
  specs.push({
    name: '20-cn-short-report',
    bodyXml: body,
    sectPr: A4_SECT_WITH_DOCGRID,
    meta: {
      title: 'Short Chinese report (~1.5 pages)',
      tags: ['chinese', 'docGrid', 'short-report', 'center-title'],
      description: '~1.5-page Chinese report, centered title, first-line-indented body',
    },
  })
}

// 21. Multi-section continuation (continuous section)
{
  const sect1Continuous =
    '<w:sectPr><w:type w:val="continuous"/>' +
    '<w:pgSz w:w="11906" w:h="16838"/>' +
    '<w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="708" w:footer="708" w:gutter="0"/>' +
    '</w:sectPr>'
  let body = p('连续节分隔符测试', 'Heading1')
  for (let i = 0; i < 10; i++) body += loremParagraph(25, 'zh')
  body += `<w:p><w:pPr>${sect1Continuous}</w:pPr></w:p>`
  body += p('续页节（内容紧接上节）', 'Heading2')
  for (let i = 0; i < 15; i++) body += loremParagraph(25, 'zh')
  specs.push({
    name: '21-continuous-section',
    bodyXml: body,
    meta: {
      title: 'Continuous section break',
      tags: ['multi-section', 'continuous-section'],
      description:
        'Contains a continuous section break, tests section switching without a forced page break',
    },
  })
}

// 22. Multi-column document
{
  const twoColSect =
    '<w:sectPr>' +
    '<w:pgSz w:w="11906" w:h="16838"/>' +
    '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/>' +
    '<w:cols w:num="2" w:space="720"/>' +
    '</w:sectPr>'
  let body = p('双栏布局文档', 'Heading1')
  for (let i = 0; i < 30; i++) body += loremParagraph(20, 'zh')
  specs.push({
    name: '22-two-columns',
    bodyXml: body,
    sectPr: twoColSect,
    meta: {
      title: 'Two-column layout',
      tags: ['two-column', 'columns', 'layout'],
      description: 'Two-column layout, tests pagination with columns',
    },
  })
}

// 23. tblHeader repeated headers across pages
{
  const border = (n: string) => `<w:${n} w:val="single" w:sz="4" w:space="0" w:color="auto"/>`
  const borders =
    '<w:tblBorders>' +
    ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map(border).join('') +
    '</w:tblBorders>'
  let rows =
    '<w:tr><w:trPr><w:tblHeader/></w:trPr>' +
    ['列A表头', '列B表头']
      .map(
        (t) =>
          `<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="D9D9D9"/></w:tcPr>` +
          `<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>${t}</w:t></w:r></w:p></w:tc>`,
      )
      .join('') +
    '</w:tr>'
  for (let r = 1; r <= 45; r++) {
    const cantSplit = r % 10 === 0 ? '<w:trPr><w:cantSplit/></w:trPr>' : ''
    rows +=
      `<w:tr>${cantSplit}` +
      [1, 2]
        .map(
          (c) =>
            `<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>` +
            `<w:p><w:r><w:t>第${r}行${c === 1 ? 'A' : 'B'}</w:t></w:r></w:p></w:tc>`,
        )
        .join('') +
      '</w:tr>'
  }
  const body =
    p('tblHeader 跨页表头重复测试', 'Heading1') +
    `<w:tbl><w:tblPr>${borders}</w:tblPr>${rows}</w:tbl>` +
    p('表后段落。')
  specs.push({
    name: '23-tblheader-repeat',
    bodyXml: body,
    meta: {
      title: 'tblHeader repeated header',
      tags: ['table', 'tblHeader', 'cantSplit'],
      description:
        'Cross-page table with a tblHeader first row (repeated header placeholder on following pages) + some cantSplit rows',
    },
  })
}

// write the files
async function writeAll() {
  for (const spec of specs) {
    const bytes = await buildDocx(spec)
    const docxPath = join(DOCX_OUT, `${spec.name}.docx`)
    writeFileSync(docxPath, bytes)
    const metaPath = join(META_OUT, `${spec.name}.json`)
    writeFileSync(metaPath, JSON.stringify({ ...spec.meta, file: `${spec.name}.docx` }, null, 2))
    console.log(`✓ ${spec.name}.docx (${(bytes.length / 1024).toFixed(1)} KB)`)
  }

  // copy existing fixtures (simple files, already in the docx directory)
  const existingFiles = [
    { src: join(REPO_ROOT, 'fixtures/generated/simple.docx'), name: 'fixture-simple.docx' },
    {
      src: join(REPO_ROOT, 'fixtures/generated/kitchen-sink.docx'),
      name: 'fixture-kitchen-sink.docx',
    },
  ]
  for (const f of existingFiles) {
    if (existsSync(f.src)) {
      const dst = join(DOCX_OUT, f.name)
      copyFileSync(f.src, dst)
      console.log(`✓ ${f.name} (copied from fixtures)`)
      // meta
      const title = f.name.replace(/^fixture-/, '').replace('.docx', '')
      const metaPath = join(META_OUT, f.name.replace('.docx', '.json'))
      writeFileSync(
        metaPath,
        JSON.stringify(
          {
            title,
            file: f.name,
            tags: ['fixture', 'real-world'],
            description: `Real test file copied from repo fixtures`,
          },
          null,
          2,
        ),
      )
    }
  }

  console.log(
    `\ntotal ${specs.length + existingFiles.length} corpus files (${specs.length} programmatically built, ${existingFiles.length} from fixtures)`,
  )
}

writeAll().catch((err) => {
  console.error(err)
  process.exit(1)
})
