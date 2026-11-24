import { buildDocx } from './build-docx'

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
const A = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
const C = 'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"'

export const THEMED_STYLES_XML =
  `<w:styles ${W}><w:docDefaults><w:rPrDefault><w:rPr>` +
  '<w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi"/><w:sz w:val="22"/>' +
  '</w:rPr></w:rPrDefault></w:docDefaults>' +
  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/>' +
  '<w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:hAnsiTheme="majorHAnsi"/><w:sz w:val="56"/></w:rPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/>' +
  '<w:pPr><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:hAnsiTheme="majorHAnsi"/>' +
  '<w:color w:val="4472C4" w:themeColor="accent1"/><w:sz w:val="32"/></w:rPr></w:style>' +
  '<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/></w:style>' +
  '<w:style w:type="table" w:styleId="Banded"><w:name w:val="Grid Table 4 Accent 1"/><w:basedOn w:val="TableNormal"/>' +
  '<w:tblPr><w:tblStyleRowBandSize w:val="1"/><w:tblBorders>' +
  ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']
    .map(
      (side) =>
        `<w:${side} w:val="single" w:sz="4" w:space="0" w:color="8EAADB" w:themeColor="accent1" w:themeTint="99"/>`,
    )
    .join('') +
  '</w:tblBorders></w:tblPr>' +
  '<w:tblStylePr w:type="firstRow"><w:rPr><w:b/><w:color w:val="FFFFFF" w:themeColor="background1"/></w:rPr>' +
  '<w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="4472C4" w:themeFill="accent1"/></w:tcPr></w:tblStylePr>' +
  '<w:tblStylePr w:type="band1Horz"><w:tcPr>' +
  '<w:shd w:val="clear" w:color="auto" w:fill="D9E2F3" w:themeFill="accent1" w:themeFillTint="33"/>' +
  '</w:tcPr></w:tblStylePr></w:style></w:styles>'

const slot = (tag: string, hex: string) => `<a:${tag}><a:srgbClr val="${hex}"/></a:${tag}>`
const fontGroup = (tag: string, face: string) =>
  `<a:${tag}><a:latin typeface="${face}"/><a:ea typeface=""/><a:cs typeface=""/></a:${tag}>`
const solid = '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'

/** the Office 2013 theme, the way Word writes a theme part */
export const OFFICE_THEME_XML =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<a:theme ${A} name="Office Theme"><a:themeElements>` +
  '<a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>' +
  '<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>' +
  slot('dk2', '44546A') +
  slot('lt2', 'E7E6E6') +
  slot('accent1', '4472C4') +
  slot('accent2', 'ED7D31') +
  slot('accent3', 'A5A5A5') +
  slot('accent4', 'FFC000') +
  slot('accent5', '5B9BD5') +
  slot('accent6', '70AD47') +
  slot('hlink', '0563C1') +
  slot('folHlink', '954F72') +
  '</a:clrScheme><a:fontScheme name="Office">' +
  fontGroup('majorFont', 'Calibri Light') +
  fontGroup('minorFont', 'Calibri') +
  '</a:fontScheme><a:fmtScheme name="Office">' +
  `<a:fillStyleLst>${solid}${solid}${solid}</a:fillStyleLst>` +
  `<a:lnStyleLst><a:ln w="6350">${solid}</a:ln><a:ln w="12700">${solid}</a:ln><a:ln w="19050">${solid}</a:ln></a:lnStyleLst>` +
  '<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle>' +
  '<a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>' +
  `<a:bgFillStyleLst>${solid}${solid}${solid}</a:bgFillStyleLst>` +
  '</a:fmtScheme></a:themeElements></a:theme>'

const cell = (text: string, tcPr = '') =>
  `<w:tc>${tcPr}<w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`
const ACCENT6_CELL =
  '<w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="70AD47" w:themeFill="accent6"/></w:tcPr>'
const TABLE =
  '<w:tbl><w:tblPr><w:tblStyle w:val="Banded"/><w:tblW w:w="0" w:type="auto"/>' +
  '<w:tblLook w:val="0420" w:firstRow="1" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>' +
  '</w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>' +
  ['Region', 'North', 'South']
    .map(
      (name, i) =>
        `<w:tr>${cell(name)}${cell(i === 0 ? 'Sales' : String(i * 10), i === 2 ? ACCENT6_CELL : '')}</w:tr>`,
    )
    .join('') +
  '</w:tbl>'

const scheme = (name: string) => `<a:solidFill><a:schemeClr val="${name}"/></a:solidFill>`
const SHAPE =
  '<w:p><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">' +
  '<wp:extent cx="1905000" cy="571500"/><wp:docPr id="7" name="Box 7"/><wp:cNvGraphicFramePr/>' +
  '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
  '<wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:cNvSpPr txBox="1"/>' +
  '<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1905000" cy="571500"/></a:xfrm>' +
  `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>${scheme('accent2')}<a:ln w="19050">${scheme('accent1')}</a:ln></wps:spPr>` +
  '<wps:txbx><w:txbxContent><w:p><w:r><w:rPr><w:color w:val="44546A" w:themeColor="text2"/></w:rPr>' +
  '<w:t>boxed</w:t></w:r></w:p></w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp>' +
  '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>'

const cache = (tag: string, kind: 'str' | 'num', values: (string | number)[]) =>
  `<c:${tag}><c:${kind}Ref><c:f>Sheet1!$A$1</c:f><c:${kind}Cache><c:ptCount val="${values.length}"/>` +
  values.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join('') +
  `</c:${kind}Cache></c:${kind}Ref></c:${tag}>`
const series = (i: number, values: number[], spPr = '') =>
  `<c:ser><c:idx val="${i}"/><c:order val="${i}"/>${cache('tx', 'str', [`Series ${i + 1}`])}${spPr}` +
  `${cache('cat', 'str', ['A', 'B'])}${cache('val', 'num', values)}</c:ser>`
const CHART_PART =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<c:chartSpace ${C} ${A}><c:chart><c:plotArea><c:layout/>` +
  '<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>' +
  series(0, [3, 5], `<c:spPr>${scheme('accent1')}</c:spPr>`) +
  series(1, [4, 2]) +
  '<c:axId val="1"/><c:axId val="2"/></c:barChart>' +
  '<c:catAx><c:axId val="1"/><c:crossAx val="2"/></c:catAx><c:valAx><c:axId val="2"/><c:crossAx val="1"/></c:valAx>' +
  '</c:plotArea></c:chart></c:chartSpace>'
const CHART =
  '<w:p><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">' +
  '<wp:extent cx="3810000" cy="2286000"/><wp:docPr id="9" name="Chart 9"/><wp:cNvGraphicFramePr/>' +
  '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">' +
  `<c:chart ${C} r:id="rIdChart1"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`

const BODY =
  '<w:p><w:pPr><w:pStyle w:val="Title"/></w:pPr><w:r><w:t>Report</w:t></w:r></w:p>' +
  '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Overview</w:t></w:r></w:p>' +
  '<w:p><w:r><w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:hAnsiTheme="majorHAnsi"/></w:rPr><w:t>display face</w:t></w:r>' +
  '<w:r><w:rPr><w:color w:val="ED7D31" w:themeColor="accent2"/></w:rPr><w:t> tinted</w:t></w:r>' +
  '<w:r><w:rPr><w:color w:val="FF0000"/></w:rPr><w:t> literal</w:t></w:r></w:p>' +
  TABLE +
  '<w:p><w:r><w:t>between</w:t></w:r></w:p>' +
  SHAPE +
  CHART +
  '<w:p><w:r><w:t>plain</w:t></w:r></w:p>'

/**
 * A document on the Office theme whose heading, a run's face and colour, a table
 * style (shading, borders, banding), a cell's own shading, a text box (fill, line, text) and a chart
 * (one series on an accent, one on the automatic palette) all reference the theme.
 */
export function themedDocx(withThemePart = true): Promise<Uint8Array> {
  return buildDocx({
    bodyXml: BODY,
    stylesXml: THEMED_STYLES_XML,
    extraRels:
      '<Relationship Id="rIdChart1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="charts/chart1.xml"/>' +
      (withThemePart
        ? '<Relationship Id="rIdTheme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>'
        : ''),
    extraParts: [
      {
        path: 'word/charts/chart1.xml',
        xml: CHART_PART,
        contentType: 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml',
      },
      ...(withThemePart
        ? [
            {
              path: 'word/theme/theme1.xml',
              xml: OFFICE_THEME_XML,
              contentType: 'application/vnd.openxmlformats-officedocument.theme+xml',
            },
          ]
        : []),
    ],
  })
}
