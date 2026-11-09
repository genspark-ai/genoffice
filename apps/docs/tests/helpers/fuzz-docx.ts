/** A seeded generator of varied-but-valid DOCX packages for the round-trip
 *  fuzz suite. The external genoffice#2381 trial measured one hand-built hostile
 *  document; this generator composes the same feature vocabulary (rich runs,
 *  CJK, hyperlinks, tables with gridSpan/vMerge shapes including orphan
 *  continuations, nested tables, images, math, tracked changes, simple fields,
 *  block content controls) into random combinations, so the byte-identical
 *  untouched round trip is exercised over many documents instead of one.
 *  Same seed -> same bytes, so a failing seed reproduces exactly. */
import {
  buildDocx,
  IMAGE_PARAGRAPH_XML,
  MATH_PARAGRAPH_XML,
  TINY_PNG_BASE64,
} from '../../../../packages/docx-engine/tests/helpers/build-docx'

/** mulberry32: small, fast, fully deterministic across runs and platforms. */
function rng(seed: number): () => number {
  let a = seed >>> 0
  return () => {
    a |= 0
    a = (a + 0x6d2b79f5) | 0
    let t = Math.imul(a ^ (a >>> 15), 1 | a)
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296
  }
}

const CJK = [
  '\u5b63\u5ea6',
  '\u8425\u6536\u603b\u7ed3',
  '\u9644\u5f55\u7532',
  '\u5ba1\u9605',
  '\u5408\u5e76\u5355\u5143\u683c',
  '\u7b2c\u4e09\u8282',
  '\u9879\u76ee\u91cc\u7a0b\u7891',
  '\u98ce\u9669\u767b\u8bb0',
]
const LATIN = ['Quarterly review', 'Appendix A', 'risk register', 'milestone 12', 'draft v2']
const REV_AT = 'w:author="Fuzz" w:date="2026-01-01T00:00:00Z"'

interface RunStyle {
  bold?: boolean
  italic?: boolean
  underline?: boolean
  strike?: boolean
  color?: string
  size?: number
  vertAlign?: 'superscript' | 'subscript'
}

function runXml(text: string, r: () => number, styleProb = 0.45): string {
  const style: RunStyle = {}
  if (r() < styleProb) style.bold = true
  if (r() < styleProb) style.italic = true
  if (r() < styleProb) style.underline = true
  if (r() < 0.2) style.strike = true
  if (r() < 0.2) style.color = ['FF0000', '0563C1', '0088FF', 'C00000'][Math.floor(r() * 4)]
  if (r() < 0.2) style.size = [20, 24, 28, 32][Math.floor(r() * 4)]
  if (r() < 0.15) style.vertAlign = r() < 0.5 ? 'superscript' : 'subscript'
  const props = Object.entries(style)
    .map(([k, v]) => {
      if (k === 'bold') return '<w:b/>'
      if (k === 'italic') return '<w:i/>'
      if (k === 'underline') return '<w:u w:val="single"/>'
      if (k === 'strike') return '<w:strike/>'
      if (k === 'color') return `<w:color w:val="${v}"/>`
      if (k === 'size') return `<w:sz w:val="${v}"/>`
      return `<w:vertAlign w:val="${v}"/>`
    })
    .join('')
  const rPr = props ? `<w:rPr>${props}</w:rPr>` : ''
  return `<w:r>${rPr}<w:t xml:space="preserve">${text}</w:t></w:r>`
}

function paraXml(r: () => number, seed: number, index: number): string {
  const kind = r()
  const text = () =>
    r() < 0.5
      ? `${CJK[Math.floor(r() * CJK.length)]}-${seed}-${index}`
      : `${LATIN[Math.floor(r() * LATIN.length)]} ${index} &amp; &lt;ev&gt;`
  if (kind < 0.12)
    return `<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>${text()}</w:t></w:r></w:p>`
  if (kind < 0.2)
    return `<w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:t>${text()}</w:t></w:r></w:p>`
  if (kind < 0.32) {
    const numId = r() < 0.5 ? 1 : 2
    return `<w:p><w:pPr><w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="${numId}"/></w:numPr></w:pPr>${runXml(text(), r)}</w:p>`
  }
  if (kind < 0.4)
    return (
      '<w:p><w:r><w:t>see </w:t></w:r>' +
      '<w:hyperlink r:id="rId20"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr>' +
      `<w:t>${text()}</w:t></w:r></w:hyperlink></w:p>`
    )
  if (kind < 0.46)
    // simple field
    return (
      `<w:p><w:fldSimple w:instr=" PAGE "><w:r><w:t>${index}</w:t></w:r></w:fldSimple>` +
      `<w:r><w:t xml:space="preserve"> / ${text()}</w:t></w:r></w:p>`
    )
  if (kind < 0.54)
    // tracked insertion + deletion in one paragraph
    return (
      '<w:p>' +
      `<w:ins w:id="${100 + index}" ${REV_AT}>${runXml(text(), r)}</w:ins>` +
      `<w:del w:id="${200 + index}" ${REV_AT}><w:r><w:delText>deleted ${index}</w:delText></w:r></w:del>` +
      `${runXml(text(), r)}</w:p>`
    )
  if (kind < 0.58)
    // block-level content control wrapping one paragraph
    return (
      `<w:sdt><w:sdtPr><w:id w:val="${index + 1}"/><w:tag w:val="fuzz-${seed}-${index}"/></w:sdtPr>` +
      `<w:sdtContent><w:p>${runXml(text(), r)}</w:p></w:sdtContent></w:sdt>`
    )
  if (r() < 0.08) return MATH_PARAGRAPH_XML
  const runCount = 1 + Math.floor(r() * 3)
  let runs = ''
  for (let i = 0; i < runCount; i++) runs += runXml(text(), r)
  return `<w:p>${runs}</w:p>`
}

/** One table: gridSpan merges, column-tracked vMerge (restart -> continuations,
 *  occasionally an orphan continuation with no restart) and random cell props. */
function tableXml(r: () => number, seed: number, tableIndex: number): string {
  const cols = 1 + Math.floor(r() * 3)
  const rows = 2 + Math.floor(r() * 3)
  const colWidth = 2000 + Math.floor(r() * 1400)
  // per column: where a vMerge restart sits (or none); a later row may then continue it
  const restartAt = new Map<number, number>()
  for (let c = 0; c < cols; c++) if (r() < 0.35) restartAt.set(c, Math.floor(r() * rows))
  // an orphan continuation in one column of one row (the genoffice#2381/genoffice#2405 shape)
  const orphanCol = r() < 0.3 ? Math.floor(r() * cols) : -1
  const orphanRow = orphanCol >= 0 ? Math.floor(r() * rows) : -1

  let grid = '<w:tblGrid>'
  for (let c = 0; c < cols; c++) grid += `<w:gridCol w:w="${colWidth}"/>`
  grid += '</w:tblGrid>'

  let body = ''
  for (let row = 0; row < rows; row++) {
    let tr = ''
    let col = 0
    while (col < cols) {
      let span = 1
      if (r() < 0.25 && cols - col >= 2) span = 2
      const tcPr: string[] = [`<w:tcW w:w="${colWidth * span}" w:type="dxa"/>`]
      if (r() < 0.25)
        tcPr.push(
          `<w:shd w:val="clear" w:color="auto" w:fill="F${Math.floor(r() * 90000 + 9999)}"/>`,
        )
      const isRestart = restartAt.get(col) === row
      const isOrphan = col === orphanCol && row === orphanRow
      if (span === 1 && isRestart) tcPr.push('<w:vMerge w:val="restart"/>')
      else if (
        span === 1 &&
        (isOrphan || (restartAt.has(col) && row > (restartAt.get(col) ?? rows)))
      ) {
        tcPr.push(
          isOrphan && r() < 0.5 ? '<w:textDirection w:val="btLr"/><w:vMerge/>' : '<w:vMerge/>',
        )
      }
      const cellText =
        span === 2
          ? `span ${row}:${col}`
          : tcPr.some((p) => p.includes('vMerge'))
            ? 'cont'
            : `t${tableIndex}r${row}c${col}`
      tr +=
        `<w:tc><w:tcPr>${tcPr.join('')}</w:tcPr>` +
        `<w:p>${runXml(cellText === 'cont' ? `${CJK[Math.floor(r() * CJK.length)]}\u7eed` : cellText, r)}</w:p></w:tc>`
      col += span
    }
    body += `<w:tr>${tr}</w:tr>`
  }
  return `<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/></w:tblPr>${grid}${body}</w:tbl>`
}

/** An outer table whose first cell hosts a nested table (random inner shape). */
function nestedTableXml(r: () => number): string {
  const innerCols = 1 + Math.floor(r() * 2)
  let innerGrid = ''
  let innerRow = ''
  for (let c = 0; c < innerCols; c++) {
    innerGrid += '<w:gridCol w:w="1200"/>'
    innerRow += `<w:tc><w:p><w:r><w:t>Inner${c}</w:t></w:r></w:p></w:tc>`
  }
  return (
    '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>' +
    '<w:tblGrid><w:gridCol w:w="3400"/><w:gridCol w:w="2600"/></w:tblGrid>' +
    '<w:tr><w:tc><w:p><w:r><w:t>Outer cell</w:t></w:r></w:p>' +
    `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid>${innerGrid}</w:tblGrid>` +
    `<w:tr>${innerRow}</w:tr></w:tbl><w:p/></w:tc>` +
    '<w:tc><w:p><w:r><w:t>Right cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
  )
}

/** Build one fuzz document: the anchor paragraph (always the edit target of
 *  the suite) followed by 3-9 random feature blocks. The package flags follow
 *  what the body actually references — an unreferenced media/numbering part or
 *  hyperlink relationship would be pruned by the save and is not a fidelity
 *  signal, so the generator never produces orphans. */
/** A section-properties roll for the final sectPr: portrait/landscape,
 *  column count, titlePg — all optional, always a valid minimal sectPr. */
function sectionRoll(r: () => number): string {
  const landscape = r() < 0.3
  const w = landscape ? 16838 : 11906
  const h = landscape ? 11906 : 16838
  const orient = landscape ? ' w:orient="landscape"' : ''
  const cols = r() < 0.25 ? `<w:cols w:num="${1 + Math.floor(r() * 2)}" w:space="708"/>` : ''
  const titlePg = r() < 0.2 ? '<w:titlePg/>' : ''
  return (
    `<w:sectPr><w:pgSz w:w="${w}" w:h="${h}"${orient}/>` +
    '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/>' +
    `${cols}${titlePg}</w:sectPr>`
  )
}

/** A mid-body section break: its own sectPr inside a paragraph's pPr, with a
 *  page size that may differ from the final section's. */
function sectionBreakXml(r: () => number): string {
  const landscape = r() < 0.4
  const w = landscape ? 16838 : 11906
  const h = landscape ? 11906 : 16838
  const orient = landscape ? ' w:orient="landscape"' : ''
  return (
    `<w:p><w:pPr><w:sectPr><w:pgSz w:w="${w}" w:h="${h}"${orient}/>` +
    '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/>' +
    '</w:sectPr></w:pPr></w:p>'
  )
}

/** A rich header/footer body: 1-2 paragraphs, formatted runs, a PAGE or
 *  NUMPAGES simple field, a tab separator, and optionally a logo image
 *  (the caller wires the media relationship when `image` comes back true). */
function richHfXml(
  r: () => number,
  seed: number,
  kind: 'header' | 'footer',
): { xml: string; image: boolean } {
  const paras: string[] = []
  const paraCount = 1 + (r() < 0.4 ? 1 : 0)
  let image = false
  for (let i = 0; i < paraCount; i++) {
    const pieces: string[] = []
    if (i === 0 && r() < 0.7) {
      const instr = kind === 'footer' || r() < 0.5 ? ' PAGE ' : ' NUMPAGES '
      pieces.push(`<w:fldSimple w:instr="${instr}"><w:r><w:t>1</w:t></w:r></w:fldSimple>`)
    }
    if (r() < 0.5) pieces.push('<w:r><w:tab/></w:r>')
    pieces.push(runXml(`${CJK[Math.floor(r() * CJK.length)]} ${kind} ${seed}-${i}`, r))
    if (i === paraCount - 1 && r() < 0.5) {
      image = true
      pieces.push(
        inlineImageRunXml(kind === 'header' ? 'rIdHImg' : 'rIdFImg', 90 + (seed % 7), 50, 50),
      )
    }
    paras.push(`<w:p>${pieces.join('')}</w:p>`)
  }
  return { xml: paras.join(''), image }
}

/** One inline drawing run: the same shape as IMAGE_PARAGRAPH_XML, with a
 *  caller-chosen relationship id, docPr id and pixel size. */
function inlineImageRunXml(rid: string, docPrId: number, wPx: number, hPx: number): string {
  const cx = wPx * 9525
  const cy = hPx * 9525
  return (
    `<w:r><w:drawing><wp:inline><wp:extent cx="${cx}" cy="${cy}"/>` +
    `<wp:docPr id="${docPrId}" name="fuzz${docPrId}"/>` +
    '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
    `<pic:pic><pic:blipFill><a:blip r:embed="${rid}"/></pic:blipFill></pic:pic>` +
    '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>'
  )
}

/** A body paragraph anchoring one comment: range markers around a plain run,
 *  then the reference run. Never used on the anchor block (the suite's edit
 *  target must stay free of markers the regenerated XML would rewrite). */
function commentAnchorParaXml(
  r: () => number,
  seed: number,
  index: number,
  commentId: number,
): string {
  return (
    `<w:p><w:commentRangeStart w:id="${commentId}"/>` +
    runXml(`commented ${seed}-${index} \u6279\u6ce8`, r) +
    `<w:commentRangeEnd w:id="${commentId}"/>` +
    `<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="${commentId}"/></w:r></w:p>`
  )
}

interface FuzzComment {
  id: number
  paraId: string
  author: string
  text: string
  paras: number
  done: boolean
  parentId?: number
}

/** word/comments.xml: one or two paragraphs per comment, the paraId on the
 *  last one (the commentsExtended link key), optionally an inline picture. */
function commentsPartXml(comments: FuzzComment[], imageCommentId: number | undefined): string {
  const body = comments
    .map((c) => {
      const paras: string[] = []
      for (let p = 0; p < c.paras; p++) {
        const isLast = p === c.paras - 1
        const pieces: string[] = []
        if (isLast && c.id === imageCommentId)
          pieces.push(inlineImageRunXml('rIdCImg', 200 + c.id, 64, 48))
        pieces.push(
          `<w:r><w:t xml:space="preserve">${c.text}${p > 0 ? ` (${p + 1})` : ''}</w:t></w:r>`,
        )
        // only the last paragraph's paraId is the link key; earlier ones carry
        // deterministic neighbors
        const pid = isLast
          ? c.paraId
          : (parseInt(c.paraId, 16) + p).toString(16).toUpperCase().padStart(8, '0')
        paras.push(
          `<w:p w14:paraId="${pid}" w14:textId="${paraIdFor(pid)}">${pieces.join('')}</w:p>`,
        )
      }
      return (
        `<w:comment w:id="${c.id}" w:author="${c.author}" w:date="2026-01-01T00:00:00Z" w:initials="${c.author[0]}">` +
        paras.join('') +
        '</w:comment>'
      )
    })
    .join('')
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
    'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" ' +
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
    'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
    body +
    '</w:comments>'
  )
}

/** the link paraId of a comment: 8 hex digits, deterministic per seed and id. */
function commentParaId(seed: number, id: number): string {
  return (0x20000000 + seed * 0x1000 + id * 0x37).toString(16).toUpperCase().padStart(8, '0')
}

/** w14:paraId from an existing one: rotate deterministically, 8 hex digits. */
function paraIdFor(from: string): string {
  return ((parseInt(from, 16) ^ 0x5a5a5a) >>> 0).toString(16).toUpperCase().padStart(8, '0')
}

/** word/commentsExtended.xml: reply/resolution records keyed by paraId. */
function commentsExtendedPartXml(comments: FuzzComment[]): string {
  const entries = comments
    .map((c) => {
      const parent = c.parentId ? comments.find((p) => p.id === c.parentId) : undefined
      return (
        `<w15:commentEx w15:paraId="${c.paraId}"` +
        (parent ? ` w15:paraIdParent="${parent.paraId}"` : '') +
        ` w15:done="${c.done ? 1 : 0}"/>`
      )
    })
    .join('')
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">' +
    entries +
    '</w15:commentsEx>'
  )
}

const COMMENT_AUTHORS = ['Fuzz', 'Reviewer', 'Editor']

export async function buildFuzzDocx(seed: number): Promise<Uint8Array> {
  const r = rng(seed * 7919 + 17)
  const blocks: string[] = [`<w:p><w:r><w:t>Anchor seed ${seed}</w:t></w:r></w:p>`]
  const count = 3 + Math.floor(r() * 7)
  for (let i = 0; i < count; i++) {
    const pick = r()
    if (pick < 0.42) blocks.push(paraXml(r, seed, i))
    else if (pick < 0.72) blocks.push(tableXml(r, seed, i))
    else if (pick < 0.78) blocks.push(nestedTableXml(r))
    else if (pick < 0.84) blocks.push(IMAGE_PARAGRAPH_XML)
    else if (pick < 0.9) blocks.push(MATH_PARAGRAPH_XML)
    else blocks.push(paraXml(r, seed, i))
  }
  // section breaks (0-2), each after a random block (never before the anchor)
  const breaks = Math.floor(r() * 3)
  for (let b = 0; b < breaks; b++) {
    const at = 1 + Math.floor(r() * (blocks.length - 1))
    blocks.splice(at, 0, sectionBreakXml(r))
  }
  const extraStyles: string[] = []
  const extraParts: Array<{ path: string; xml: string; contentType: string }> = []
  const binaryParts: Array<{
    path: string
    base64: string
    extension: string
    contentType: string
  }> = []
  const extraRels: string[] = []

  // header/footer knob: either, both, or neither; simple or rich (fields,
  // multi-paragraph, formatted runs, tab); a rich part may carry a logo image
  // (relationship + media wired through headerRelsXml / footerRelsXml)
  const hfRoll = (kind: 'header' | 'footer'): string | undefined => {
    if (r() >= 0.4) return undefined
    if (r() >= 0.6)
      return `<w:p>${runXml(`${CJK[Math.floor(r() * CJK.length)]} ${kind} ${seed}`, r)}</w:p>`
    const rich = richHfXml(r, seed, kind)
    if (rich.image)
      binaryParts.push({
        path: `word/media/image${kind === 'header' ? 'H' : 'F'}1.png`,
        base64: TINY_PNG_BASE64,
        extension: 'png',
        contentType: 'image/png',
      })
    return rich.xml
  }
  const headerXml = hfRoll('header')
  const footerXml = hfRoll('footer')

  // comments knob: 0-3 anchored comments, optionally commentsExtended, and
  // one comment may carry an inline picture (comments.xml.rels + media part)
  if (r() < 0.45) {
    const commentCount = 1 + Math.floor(r() * 3)
    const withExt = r() < 0.6
    const imageCommentId = r() < 0.4 ? 1 : undefined
    const comments: FuzzComment[] = []
    for (let i = 0; i < commentCount; i++) {
      const id = i + 1
      const at = 1 + Math.floor(r() * (blocks.length - 1))
      blocks.splice(at, 0, commentAnchorParaXml(r, seed, i, id))
      comments.push({
        id,
        paraId: commentParaId(seed, id),
        author: COMMENT_AUTHORS[Math.floor(r() * COMMENT_AUTHORS.length)]!,
        text:
          r() < 0.5
            ? `${CJK[Math.floor(r() * CJK.length)]}\u6279\u6ce8 ${id}`
            : `note ${id} on ${seed}`,
        paras: r() < 0.3 ? 2 : 1,
        done: r() < 0.5,
        parentId: i > 0 && r() < 0.4 ? 1 : undefined,
      })
    }
    if (imageCommentId !== undefined) {
      extraParts.push({
        path: 'word/_rels/comments.xml.rels',
        xml: `<Relationship Id="rIdCImg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/imageC1.png"/>`,
        contentType: '', // .rels: no Override (filtered in buildDocx)
      })
      binaryParts.push({
        path: 'word/media/imageC1.png',
        base64: TINY_PNG_BASE64,
        extension: 'png',
        contentType: 'image/png',
      })
    }
    extraParts.push({
      path: 'word/comments.xml',
      xml: commentsPartXml(comments, imageCommentId),
      contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml',
    })
    extraRels.push(
      '<Relationship Id="rIdCm" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>',
    )
    if (withExt) {
      extraParts.push({
        path: 'word/commentsExtended.xml',
        xml: commentsExtendedPartXml(comments),
        contentType:
          'application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml',
      })
      extraRels.push(
        '<Relationship Id="rIdCmx" Type="http://schemas.microsoft.com/office/2011/relationships/commentsExtended" Target="commentsExtended.xml"/>',
      )
    }
    // the reference-run rStyle lives in styles.xml in real Word files
    extraStyles.push(
      '<w:style w:type="character" w:styleId="CommentReference"><w:name w:val="annotation reference"/></w:style>',
    )
  }
  // joined after the comment anchors are spliced in, or they never reach the body
  const body = blocks.join('')
  if (body.includes('rId20'))
    extraRels.push(
      '<Relationship Id="rId20" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/" TargetMode="External"/>',
    )
  return buildDocx({
    bodyXml: body,
    withNumbering: body.includes('<w:numPr>'),
    withImage: body.includes('r:embed="rId10"'),
    headerXml,
    headerRelsXml: headerXml?.includes('rIdHImg')
      ? '<Relationship Id="rIdHImg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/imageH1.png"/>'
      : undefined,
    footerXml,
    footerRelsXml: footerXml?.includes('rIdFImg')
      ? '<Relationship Id="rIdFImg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/imageF1.png"/>'
      : undefined,
    sectPrXml: r() < 0.45 ? sectionRoll(r) : undefined,
    extraStylesXml: extraStyles.length ? extraStyles.join('') : undefined,
    extraParts: extraParts.length ? extraParts : undefined,
    binaryParts: binaryParts.length ? binaryParts : undefined,
    extraRels: extraRels.length ? extraRels.join('') : undefined,
  })
}
