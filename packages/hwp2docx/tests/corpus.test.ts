/**
 * Real-document sweep: every .hwpx under HWPX_CORPUS_DIR converts, the .docx
 * re-parses through docx-engine, and the body text survives. Skipped when the
 * variable is unset (the corpus is not checked in — government forms and
 * press releases carry their own terms).
 */

import { readdirSync, readFileSync, statSync } from 'node:fs'
import { join } from 'node:path'
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { parseDocx } from '../../docx-engine/src/index'
import { convertHwpToDocx, convertHwpxToDocx, hwpToHwpx } from '../src/index'

const dir = process.env.HWPX_CORPUS_DIR

function walk(root: string): string[] {
  const out: string[] = []
  for (const name of readdirSync(root)) {
    const path = join(root, name)
    if (statSync(path).isDirectory()) out.push(...walk(path))
    else if (/\.hwpx?$/i.test(name)) out.push(path)
  }
  return out
}

const squash = (text: string): string => text.replace(/\s+/g, '')

async function hwpxText(bytes: Uint8Array): Promise<string> {
  const zip = await JSZip.loadAsync(bytes)
  let text = ''
  for (const name of Object.keys(zip.files).sort()) {
    if (!/^Contents\/section\d+\.xml$/.test(name)) continue
    const xml = await zip.file(name)!.async('string')
    // body text only: header/footer/footnote bodies are not converted
    const stripped = xml.replace(/<hp:(header|footer|footNote|endNote)\b[\s\S]*?<\/hp:\1>/g, '')
    for (const match of stripped.matchAll(/<hp:t(?:\s[^>]*)?>([\s\S]*?)<\/hp:t>/g)) {
      text += match[1].replace(/<[^>]+>/g, '')
    }
  }
  return squash(
    text
      .replace(/&lt;/g, '<')
      .replace(/&gt;/g, '>')
      .replace(/&amp;/g, '&')
      .replace(/&quot;/g, '"'),
  )
}

async function docxText(docx: Uint8Array): Promise<string> {
  const zip = await JSZip.loadAsync(docx)
  const xml = await zip.file('word/document.xml')!.async('string')
  let text = ''
  for (const match of xml.matchAll(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g)) text += match[1]
  return squash(
    text
      .replace(/&lt;/g, '<')
      .replace(/&gt;/g, '>')
      .replace(/&amp;/g, '&')
      .replace(/&quot;/g, '"'),
  )
}

describe.skipIf(!dir)('hwpx corpus', () => {
  const files = dir ? walk(dir) : []
  for (const file of files) {
    it(file.slice(dir!.length + 1), async () => {
      const bytes = new Uint8Array(readFileSync(file))
      const binary = bytes[0] === 0xd0 && bytes[1] === 0xcf
      if (binary) {
        // the strict HWPX entry point refuses it; the content-routed one converts it
        await expect(convertHwpxToDocx(bytes)).rejects.toMatchObject({ code: 'hwp-binary' })
      }
      const result = await convertHwpToDocx(bytes)
      const parsed = await parseDocx(result.docx)
      expect(parsed.blocks.length).toBeGreaterThan(0)
      const source = await hwpxText(binary ? await hwpToHwpx(bytes) : bytes)
      const converted = await docxText(result.docx)
      // character multiset overlap: how much of the source text the .docx carries
      // (order-insensitive, so one missing caption cannot cascade into a miss)
      const counts = new Map<string, number>()
      for (const ch of converted) counts.set(ch, (counts.get(ch) ?? 0) + 1)
      let hit = 0
      for (const ch of source) {
        const left = counts.get(ch) ?? 0
        if (left > 0) {
          hit++
          counts.set(ch, left - 1)
        }
      }
      const coverage = source.length === 0 ? 1 : hit / source.length
      const report = {
        file: file.split('/').pop(),
        kb: Math.round(result.docx.length / 1024),
        blocks: result.blocks,
        coverage: Math.round(coverage * 1000) / 10,
        dropped: result.dropped.map((d) => `${d.what}:${d.count}`).join(' '),
      }
      process.stdout.write(`CORPUS ${JSON.stringify(report)}\n`)
      expect(coverage).toBeGreaterThan(0.98)
    })
  }
})
