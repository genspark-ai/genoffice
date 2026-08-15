import JSZip from 'jszip'

/** 1x1 red PNG */
export const TINY_PNG_BASE64 =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=='

export interface BuildDocxOptions {
  /** raw XML of <w:body> children (without w:sectPr; one is appended automatically) */
  bodyXml: string
  /** extra <w:style> entries appended to the default styles.xml */
  extraStylesXml?: string
  withNumbering?: boolean
  /** full word/numbering.xml content (implies withNumbering) */
  numberingXml?: string
  withImage?: boolean
  /** extra Relationship entries for word/_rels/document.xml.rels */
  extraRels?: string
  /** extra zip parts (charts etc.) with their [Content_Types].xml override */
  extraParts?: Array<{ path: string; xml: string; contentType: string }>
  /** extra binary media parts with their [Content_Types].xml Default extension */
  binaryParts?: Array<{ path: string; base64: string; extension: string; contentType: string }>
  /** extra sectPr children inserted first (headerReference/footerReference…) */
  sectPrExtra?: string
}

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'

const DOC_NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" ' +
  'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"'

const STYLES_XML_OPEN =
  XML_DECL +
  '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/>' +
  '<w:pPr><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/>' +
  '<w:pPr><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:b/><w:sz w:val="28"/></w:rPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/></w:style>' +
  '<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/>' +
  '<w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr></w:style>'

const NUMBERING_XML =
  XML_DECL +
  '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
  '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlText w:val="&#61623;"/></w:lvl></w:abstractNum>' +
  '<w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum>' +
  '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>' +
  '<w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>' +
  '</w:numbering>'

export async function buildDocx(options: BuildDocxOptions): Promise<Uint8Array> {
  const zip = new JSZip()
  const withNumbering = options.withNumbering || options.numberingXml !== undefined

  const contentTypes = [
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
    '<Default Extension="xml" ContentType="application/xml"/>',
    options.withImage ? '<Default Extension="png" ContentType="image/png"/>' : '',
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>',
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>',
    withNumbering
      ? '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>'
      : '',
    ...(options.extraParts ?? []).map(
      (p) => `<Override PartName="/${p.path}" ContentType="${p.contentType}"/>`,
    ),
    ...(options.binaryParts ?? []).map(
      (p) => `<Default Extension="${p.extension}" ContentType="${p.contentType}"/>`,
    ),
  ].join('')
  zip.file(
    '[Content_Types].xml',
    `${XML_DECL}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">${contentTypes}</Types>`,
  )

  zip.file(
    '_rels/.rels',
    `${XML_DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
      '</Relationships>',
  )

  const docRels = [
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>',
    withNumbering
      ? '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>'
      : '',
    options.withImage
      ? '<Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>'
      : '',
    options.extraRels ?? '',
  ].join('')
  zip.file(
    'word/_rels/document.xml.rels',
    `${XML_DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${docRels}</Relationships>`,
  )

  zip.file('word/styles.xml', `${STYLES_XML_OPEN}${options.extraStylesXml ?? ''}</w:styles>`)
  if (withNumbering) zip.file('word/numbering.xml', options.numberingXml ?? NUMBERING_XML)
  if (options.withImage) zip.file('word/media/image1.png', TINY_PNG_BASE64, { base64: true })
  for (const part of options.extraParts ?? []) zip.file(part.path, part.xml)
  for (const part of options.binaryParts ?? []) zip.file(part.path, part.base64, { base64: true })

  const sectPr =
    `<w:sectPr>${options.sectPrExtra ?? ''}<w:pgSz w:w="11906" w:h="16838"/>` +
    '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/>' +
    '</w:sectPr>'
  zip.file(
    'word/document.xml',
    `${XML_DECL}<w:document ${DOC_NS}><w:body>${options.bodyXml}${sectPr}</w:body></w:document>`,
  )

  // Pin zip entry mtimes (JSZip defaults them to "now") so regenerating the
  // committed fixtures is byte-deterministic and CI can diff for drift.
  zip.forEach((_path, entry) => {
    entry.date = new Date(Date.UTC(2026, 0, 1))
  })
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })
}

export const IMAGE_PARAGRAPH_XML =
  '<w:p><w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="914400"/>' +
  '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
  '<pic:pic><pic:blipFill><a:blip r:embed="rId10"/></pic:blipFill></pic:pic>' +
  '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>'

export const MATH_PARAGRAPH_XML = '<w:p><m:oMath><m:r><m:t>E=mc^2</m:t></m:r></m:oMath></w:p>'

export const CHART_PARAGRAPH_XML =
  '<w:p><w:r><w:drawing><wp:inline><wp:extent cx="5486400" cy="3200400"/>' +
  '<wp:docPr id="1" name="图表 1"/>' +
  '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">' +
  '<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rId30"/>' +
  '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>'

const chartStrCache = (values: string[], f: string) =>
  `<c:strRef><c:f>${f}</c:f><c:strCache><c:ptCount val="${values.length}"/>` +
  values.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join('') +
  '</c:strCache></c:strRef>'

const chartNumCache = (values: (number | null)[], f: string) =>
  `<c:numRef><c:f>${f}</c:f><c:numCache><c:formatCode>General</c:formatCode>` +
  `<c:ptCount val="${values.length}"/>` +
  values.map((v, i) => (v === null ? '' : `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`)).join('') +
  '</c:numCache></c:numRef>'

/** bar chart part: title, 2 series over 3 categories, with a gap in series 2 */
export const CHART_PART_XML =
  XML_DECL +
  '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
  '<c:chart><c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>销售</a:t></a:r><a:r><a:t>统计</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>' +
  '<c:plotArea><c:layout/><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>' +
  '<c:ser><c:idx val="0"/><c:order val="0"/>' +
  `<c:tx>${chartStrCache(['华东'], 'Sheet1!$B$1')}</c:tx>` +
  `<c:cat>${chartStrCache(['一月', '二月', '三月'], 'Sheet1!$A$2:$A$4')}</c:cat>` +
  `<c:val>${chartNumCache([120, 88.5, 96], 'Sheet1!$B$2:$B$4')}</c:val>` +
  '</c:ser>' +
  '<c:ser><c:idx val="1"/><c:order val="1"/>' +
  `<c:tx>${chartStrCache(['华南'], 'Sheet1!$C$1')}</c:tx>` +
  `<c:cat>${chartStrCache(['一月', '二月', '三月'], 'Sheet1!$A$2:$A$4')}</c:cat>` +
  `<c:val>${chartNumCache([70, null, 110], 'Sheet1!$C$2:$C$4')}</c:val>` +
  '</c:ser>' +
  '<c:axId val="111"/><c:axId val="222"/></c:barChart></c:plotArea></c:chart></c:chartSpace>'

export const CHART_RELS =
  '<Relationship Id="rId30" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="charts/chart1.xml"/>'

/** a document whose only body content is one embedded bar chart */
export async function buildChartDocx(bodyPrefixXml = ''): Promise<Uint8Array> {
  return buildDocx({
    bodyXml: bodyPrefixXml + CHART_PARAGRAPH_XML,
    extraRels: CHART_RELS,
    extraParts: [
      {
        path: 'word/charts/chart1.xml',
        xml: CHART_PART_XML,
        contentType: 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml',
      },
    ],
  })
}

export const TABLE_XML =
  '<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/></w:tblPr>' +
  '<w:tr><w:tc><w:p><w:r><w:t>A1</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>B1</w:t></w:r></w:p></w:tc></w:tr>' +
  '<w:tr><w:tc><w:p><w:r><w:t>A2</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>B2</w:t></w:r></w:p></w:tc></w:tr>' +
  '</w:tbl>'

/** A representative document exercising every block type the engine supports. */
export function kitchenSinkBody(): string {
  return [
    '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>第一章 概述</w:t></w:r></w:p>',
    '<w:p><w:r><w:t xml:space="preserve">普通段落,包含</w:t></w:r>' +
      '<w:r><w:rPr><w:b/></w:rPr><w:t>加粗</w:t></w:r>' +
      '<w:r><w:rPr><w:i/></w:rPr><w:t>斜体</w:t></w:r>' +
      '<w:r><w:rPr><w:u w:val="single"/></w:rPr><w:t>下划线</w:t></w:r>' +
      '<w:r><w:rPr><w:strike/></w:rPr><w:t>删除线</w:t></w:r>' +
      '<w:r><w:rPr><w:color w:val="FF0000"/><w:sz w:val="28"/></w:rPr><w:t>红色14号</w:t></w:r>' +
      '<w:r><w:t xml:space="preserve">与特殊字符 &amp; &lt; &gt; 结束。</w:t></w:r></w:p>',
    '<w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:t>1.1 列表</w:t></w:r></w:p>',
    '<w:p><w:pPr><w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>无序项一</w:t></w:r></w:p>',
    '<w:p><w:pPr><w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>无序项二</w:t></w:r></w:p>',
    '<w:p><w:pPr><w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr></w:pPr><w:r><w:t>有序项一</w:t></w:r></w:p>',
    '<w:p><w:r><w:t xml:space="preserve">链接示例:</w:t></w:r>' +
      '<w:hyperlink r:id="rId20"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>示例网站</w:t></w:r></w:hyperlink></w:p>',
    TABLE_XML,
    IMAGE_PARAGRAPH_XML,
    MATH_PARAGRAPH_XML,
    '<w:p><w:r><w:t>尾段。</w:t></w:r></w:p>',
  ].join('')
}

export async function buildKitchenSinkDocx(): Promise<Uint8Array> {
  return buildDocx({
    bodyXml: kitchenSinkBody(),
    withNumbering: true,
    withImage: true,
    extraRels:
      '<Relationship Id="rId20" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/" TargetMode="External"/>',
  })
}

/** 2x3 table: row 2 has a whole-row delete revision (trPr w:del + delText), row 3 cell A has cellIns */
export const REVISION_TABLE_XML =
  '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>' +
  '<w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>' +
  '<w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>A1</w:t></w:r></w:p></w:tc>' +
  '<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>B1</w:t></w:r></w:p></w:tc></w:tr>' +
  '<w:tr><w:trPr><w:del w:id="11" w:author="Alice" w:date="2026-01-01T00:00:00Z"/></w:trPr>' +
  '<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr><w:p><w:del w:id="13" w:author="Alice" w:date="2026-01-01T00:00:00Z"><w:r><w:delText>A2del</w:delText></w:r></w:del></w:p></w:tc>' +
  '<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr><w:p><w:del w:id="14" w:author="Alice" w:date="2026-01-01T00:00:00Z"><w:r><w:delText>B2del</w:delText></w:r></w:del></w:p></w:tc></w:tr>' +
  '<w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/><w:cellIns w:id="15" w:author="Alice" w:date="2026-01-01T00:00:00Z"/></w:tcPr><w:p><w:r><w:t>A3</w:t></w:r></w:p></w:tc>' +
  '<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>B3</w:t></w:r></w:p></w:tc></w:tr>' +
  '</w:tbl>'

/** outer 1x2 table, cell A contains a 2x1 nested table (InnerA/InnerB) + trailing empty paragraph */
export const NESTED_TABLE_XML =
  '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>' +
  '<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>' +
  '<w:tr><w:tc><w:p><w:r><w:t>Outer</w:t></w:r></w:p>' +
  '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/></w:tblGrid>' +
  '<w:tr><w:tc><w:p><w:r><w:t>InnerA</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>InnerB</w:t></w:r></w:p></w:tc></w:tr>' +
  '</w:tbl><w:p/></w:tc>' +
  '<w:tc><w:p><w:r><w:t>RightCell</w:t></w:r></w:p></w:tc></w:tr>' +
  '</w:tbl>'
