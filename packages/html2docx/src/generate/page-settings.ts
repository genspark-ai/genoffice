// @ts-nocheck — generation layer ported verbatim from untyped JS; it is typed
// file by file without logic changes, and until then strict consumers
// (apps/html, apps/shell) must not fail on it.
import {
  Document,
  Footer,
  Header,
  ImageRun,
  LineRuleType,
  Paragraph,
  Table,
  TextWrappingType,
} from 'docx'
import { bidiLangOf, detectCJK, eastAsiaLangOf } from './document-meta'
import { buildNumberingConfig } from './numbering'
import { spacerParagraph } from './word-utils'
import zlib from 'node:zlib'

function partitionIr(ir) {
  const docSettings = ir.find((node) => node.type === 'docsettings')
  const pageBgNode = ir.find((node) => node.type === 'pagebg')
  const headerNode = ir.find((node) => node.type === 'headerpart')
  const footerNode = ir.find((node) => node.type === 'footerpart')
  const bodyIr = ir.filter(
    (node) =>
      node.type !== 'docsettings' &&
      node.type !== 'pagebg' &&
      node.type !== 'headerpart' &&
      node.type !== 'footerpart',
  )
  return { bodyIr, docSettings, footerNode, headerNode, pageBgNode }
}

// Word ignores <w:background> when printing / exporting PDF, so a plain
// page color must also become a behind-text float. A 1×1 PNG stretched to
// page size costs ~70 bytes.
function solidColorPng(hex) {
  const rgb = [0, 2, 4].map((i) => parseInt(hex.slice(i, i + 2), 16) || 0)
  const chunk = (type, data) => {
    const body = Buffer.concat([Buffer.from(type, 'ascii'), data])
    const out = Buffer.alloc(body.length + 8)
    out.writeUInt32BE(data.length, 0)
    body.copy(out, 4)
    out.writeUInt32BE(zlib.crc32(body) >>> 0, body.length + 4)
    return out
  }
  const ihdr = Buffer.alloc(13)
  ihdr.writeUInt32BE(1, 0)
  ihdr.writeUInt32BE(1, 4)
  ihdr[8] = 8 // bit depth
  ihdr[9] = 2 // color type: RGB
  const idat = zlib.deflateSync(Buffer.from([0, ...rgb]))
  return Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    chunk('IHDR', ihdr),
    chunk('IDAT', idat),
    chunk('IEND', Buffer.alloc(0)),
  ])
}

function addPageBackgroundFloat(generator, pageBgNode) {
  const shotData = pageBgNode?.shotId ? generator.images[pageBgNode.shotId] : null
  const data = shotData || (pageBgNode?.color ? solidColorPng(pageBgNode.color) : null)
  if (!data) return
  const context = generator.context
  generator.pageBackgroundFloat = new ImageRun({
    type: 'png',
    data,
    transformation: {
      width: Math.round(context.pageWidthDxa / 15),
      height: Math.round(context.pageHeightDxa / 15),
    },
    floating: {
      horizontalPosition: { relative: 'page', offset: 0 },
      verticalPosition: { relative: 'page', offset: 0 },
      behindDocument: true,
      wrap: { type: TextWrappingType.NONE },
      margins: { top: 0, bottom: 0, left: 0, right: 0 },
    },
  })
}

function renderSection(generator, { bodyIr, footerNode, headerNode }) {
  const headerChildren = []
  if (generator.pageBackgroundFloat) {
    headerChildren.push(
      new Paragraph({
        children: [generator.pageBackgroundFloat],
        spacing: { before: 0, after: 0, line: 1, lineRule: LineRuleType.EXACT },
      }),
    )
  }
  if (headerNode) {
    headerChildren.push(...generator.render(withPageNumberFields(headerNode.children || [])))
  }
  const footerChildren = footerNode
    ? generator.render(withPageNumberFields(footerNode.children || []))
    : []
  const children = generator.render(bodyIr)
  // Trailing spacers render nothing at document end but can overflow onto a
  // phantom background-only last page when the content fills its final page.
  // A table must not end the body either: Word then auto-inserts a
  // default-height paragraph mark, which overflows the same way. End on one
  // minimal explicit paragraph instead.
  while (children.length > 1 && children[children.length - 1]?.__h2dSpacerPx) {
    children.pop()
  }
  if (!(children[children.length - 1] instanceof Paragraph)) {
    children.push(
      new Paragraph({
        children: [],
        spacing: { before: 0, after: 0, line: 1, lineRule: LineRuleType.EXACT },
      }),
    )
  }
  // Final-card top margin: a closing footer card stranded on a fresh page
  // still shows its authored top gap, so the gap buys nothing — collapse the
  // whole run of spacers before it to improve the chance the card fits the
  // previous page's tail instead.
  let lastBlock = children.length - 1
  while (lastBlock >= 0 && children[lastBlock] instanceof Paragraph) lastBlock--
  if (children[lastBlock] instanceof Table) {
    let firstSpacer = lastBlock
    while (firstSpacer > 0 && children[firstSpacer - 1]?.__h2dSpacerPx) firstSpacer--
    const runPx = children
      .slice(firstSpacer, lastBlock)
      .reduce((sum, spacer) => sum + spacer.__h2dSpacerPx, 0)
    if (runPx > 16) {
      children.splice(firstSpacer, lastBlock - firstSpacer, spacerParagraph(generator.context, 8))
    }
  }
  if (generator.floats.length) {
    children.unshift(
      new Paragraph({
        children: generator.floats,
        spacing: { before: 0, after: 0, line: 1, lineRule: LineRuleType.EXACT },
      }),
    )
  }
  return { children, footerChildren, headerChildren }
}

// Literal page numbers in a real header/footer ("Page 1 of 3" and the CJK
// "page N / total M" form) are frozen at whatever the browser rendered — replace the
// digits with native PAGE / NUMPAGES fields so Word keeps them true.
const LATIN_PAGE_RE = /\b((?:page|pg\.?|p\.)\s*)(\d+)((?:\s*(?:of|\/|de)\s*)(\d+))?/i
const CJK_PAGE_RE =
  /(\u7b2c\s*)(\d+)(\s*[\u9875\u9801])((?:\s*[,\uff0c/]?\s*\u5171\s*)(\d+)(\s*[\u9875\u9801]))?/
function pageFieldSegments(run) {
  const text = run.text || ''
  const match = LATIN_PAGE_RE.exec(text) || CJK_PAGE_RE.exec(text)
  if (!match) return null
  const isCjk = match.length > 5
  const segments = []
  const push = (t) => t && segments.push({ ...run, text: t })
  push(text.slice(0, match.index) + match[1])
  segments.push({ ...run, text: '', pageNumberField: 'current' })
  if (isCjk) {
    push(match[3])
    if (match[5]) {
      push(match[4].slice(0, match[4].length - match[5].length - match[6].length))
      segments.push({ ...run, text: '', pageNumberField: 'total' })
      push(match[6])
    }
    push(text.slice(match.index + match[0].length))
  } else {
    if (match[4]) {
      push(match[3].slice(0, match[3].length - match[4].length))
      segments.push({ ...run, text: '', pageNumberField: 'total' })
    }
    push(text.slice(match.index + match[0].length))
  }
  return segments
}
function withPageNumberFields(nodes) {
  const mapCells = (cells) =>
    (cells || []).map((cell) => ({
      ...cell,
      runs: cell.runs ? cell.runs.flatMap((run) => pageFieldSegments(run) || [run]) : cell.runs,
      children: cell.children ? withPageNumberFields(cell.children) : cell.children,
    }))
  return (nodes || []).map((node) => {
    const mapped = { ...node }
    if (node.runs) {
      mapped.runs = node.runs.flatMap((run) => pageFieldSegments(run) || [run])
    }
    if (node.children) mapped.children = withPageNumberFields(node.children)
    if (node.rows) {
      mapped.rows = node.rows.map((row) => ({ ...row, cells: mapCells(row.cells) }))
    }
    if (node.cells) mapped.cells = mapCells(node.cells)
    if (node.items) mapped.items = withPageNumberFields(node.items)
    return mapped
  })
}

function createDocument(context, parts, rendered, generator) {
  const isCJK = detectCJK(parts.bodyIr)
  const docRTL = Boolean(parts.docSettings?.rtl)
  const bidiLang = docRTL ? bidiLangOf(parts.docSettings?.lang) : undefined
  // A solid-color pagebg (no shotId) still becomes a header-anchored float
  // via addPageBackgroundFloat's solidColorPng fallback — the header-distance
  // reservation must be zeroed for that case too, or Word pads a ~1.25cm gap
  // above the body on every page (the header area is reserved regardless of
  // how empty its content looks).
  const hasPageBackgroundHeader = Boolean(generator?.pageBackgroundFloat) && !parts.headerNode
  const needsTopBleed = hasPageBackgroundHeader && context.pageMargins.top === 0
  const pageMargins = hasPageBackgroundHeader
    ? {
        ...context.pageMargins,
        top: needsTopBleed ? -context.pxToTwips(14) : context.pageMargins.top,
        header: 0,
      }
    : context.pageMargins

  const meta = parts.docSettings?.meta || {}
  return new Document({
    title: meta.title,
    creator: meta.author || 'html2docx',
    description: meta.description,
    keywords: meta.keywords,
    background: parts.pageBgNode?.color ? { color: parts.pageBgNode.color } : undefined,
    numbering: { config: buildNumberingConfig() },
    styles: {
      default: {
        document: {
          // OOXML default is OFF unless a style sets it; Word's own new-doc
          // template enables it — match that so no stranded single lines
          paragraph: { widowControl: true } as any,
          run: isCJK
            ? {
                // ascii/hAnsi stay Latin-proportional: a CJK font on Latin
                // runs renders SimSun-style monospaced glyphs.
                font: {
                  ascii: 'Arial',
                  hAnsi: 'Arial',
                  eastAsia:
                    { 'ja-JP': 'Yu Gothic', 'ko-KR': 'Malgun Gothic' }[
                      eastAsiaLangOf(parts.docSettings?.lang, parts.bodyIr)
                    ] || 'Microsoft YaHei',
                },
                size: Math.round(21 * context.fontScale),
                // w:lang w:val must stay Latin: tagging it zh-CN makes Word
                // apply CJK line breaking to Latin words ("Rep/ort").
                language: {
                  value: 'en-US',
                  eastAsia: eastAsiaLangOf(parts.docSettings?.lang, parts.bodyIr),
                },
              }
            : {
                font: 'Arial',
                size: Math.round(21 * context.fontScale),
                language: docRTL
                  ? { value: bidiLang, bidirectional: bidiLang }
                  : { value: 'en-US' },
              },
        },
      },
    },
    sections: [
      {
        properties: {
          page: {
            size: { width: context.pageWidthDxa, height: context.pageHeightDxa },
            margin: pageMargins,
          },
        },
        headers: rendered.headerChildren.length
          ? { default: new Header({ children: rendered.headerChildren }) }
          : undefined,
        footers: rendered.footerChildren.length
          ? { default: new Footer({ children: rendered.footerChildren }) }
          : undefined,
        children: rendered.children,
      },
    ],
  })
}

export { addPageBackgroundFloat, createDocument, partitionIr, renderSection }
