import JSZip from 'jszip'
import { XMLParser } from 'fast-xml-parser'

// Text fidelity: no trim (xml:space="preserve" runs carry the spaces between words),
// no numeric coercion of tag values (otherwise <a:t>02139</a:t> becomes a number and loses characters)
const parser = new XMLParser({ ignoreAttributes: true, trimValues: false, parseTagValue: false })

function slideNumber(path: string): number {
  const m = /slide(\d+)\.xml$/.exec(path)
  return m ? Number(m[1]) : 0
}

/** collect the text of all a:t descendants of one node */
function collectText(node: unknown, out: string[], isText = false): void {
  if (node == null) return
  if (typeof node === 'string' || typeof node === 'number') {
    // untrimmed, a non-a:t element written across lines carries its whitespace as a value
    if (isText) out.push(String(node))
    return
  }
  if (Array.isArray(node)) {
    for (const item of node) collectText(item, out, isText)
    return
  }
  if (typeof node === 'object') {
    for (const [key, value] of Object.entries(node)) {
      if (key === 'a:t') collectText(value, out, true)
      else if (typeof value === 'object') collectText(value, out)
    }
  }
}

/** walk the slide tree; each a:p paragraph becomes one output line */
function collectParagraphs(node: unknown, out: string[]): void {
  if (node == null || typeof node !== 'object') return
  if (Array.isArray(node)) {
    for (const item of node) collectParagraphs(item, out)
    return
  }
  for (const [key, value] of Object.entries(node)) {
    if (key === 'a:p') {
      for (const para of Array.isArray(value) ? value : [value]) {
        const texts: string[] = []
        collectText(para, texts)
        const line = texts.join('')
        if (line.trim()) out.push(line)
      }
    } else {
      collectParagraphs(value, out)
    }
  }
}

/** extract slide text from a pptx: one "## Slide N" section per slide, one line per paragraph */
export async function pptxToText(bytes: Uint8Array): Promise<string> {
  const zip = await JSZip.loadAsync(bytes)
  const slidePaths = Object.keys(zip.files)
    .filter((p) => /^ppt\/slides\/slide\d+\.xml$/.test(p))
    .sort((a, b) => slideNumber(a) - slideNumber(b))
  const sections: string[] = []
  for (const path of slidePaths) {
    const xml = await zip.files[path]!.async('text')
    const paras: string[] = []
    collectParagraphs(parser.parse(xml), paras)
    sections.push([`## Slide ${slideNumber(path)}`, ...paras].join('\n'))
  }
  return sections.join('\n\n')
}
