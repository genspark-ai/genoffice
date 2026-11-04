/** Hostile-input fuzz for the DOCX import surface — the #2381 trial's Report D
 *  axis, run continuously instead of once. The round-trip suite
 *  (roundtrip-fuzz.test.ts) feeds the pipeline only valid documents; this file
 *  feeds it corrupted ones and asserts the one property that matters for an
 *  importer of untrusted bytes: **bounded, clean failure** — parseDocx either
 *  refuses with a thrown Error or succeeds, the editor/save stages never hang,
 *  and nothing escapes as a process crash or unhandled rejection. No fidelity
 *  assertions here: a corrupted package has no "right" output.
 *
 *  Corruption is seeded (same seed → same bytes, failing cases reproduce
 *  exactly) over the same corpus the round-trip suite generates. CI runs the
 *  default counts; HOSTILE_SEEDS / HOSTILE_PIPELINE_SEEDS sweep larger spaces.
 */
import { Editor } from '@tiptap/core'
import JSZip from 'jszip'
import { parseDocx, saveDocx } from '@genoffice/docx-engine'
import { describe, expect, it } from 'vitest'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { buildFuzzDocx } from './helpers/fuzz-docx'

const SEEDS = Math.max(1, Number(process.env.HOSTILE_SEEDS ?? 10))
const PIPELINE_SEEDS = Math.max(1, Number(process.env.HOSTILE_PIPELINE_SEEDS ?? 4))

/** Same generator discipline as helpers/fuzz-docx: small, deterministic. */
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

type Corrupt = (bytes: Uint8Array, r: () => number) => Promise<Uint8Array>

/** Flip a handful of random bytes anywhere in the package. */
const bitflip: Corrupt = async (bytes, r) => {
  const out = Uint8Array.from(bytes)
  const flips = 1 + Math.floor(r() * 8)
  for (let i = 0; i < flips; i++) {
    const at = Math.floor(r() * out.length)
    out[at] ^= 1 << Math.floor(r() * 8)
  }
  return out
}

/** Cut the file at a random offset. */
const truncate: Corrupt = async (bytes, r) => {
  const at = 100 + Math.floor(r() * Math.max(1, bytes.length - 100))
  return bytes.slice(0, at)
}

/** Remove a random package entry (a missing document.xml or styles part). */
const entryDrop: Corrupt = async (bytes, r) => {
  const zip = await JSZip.loadAsync(bytes)
  const names = Object.keys(zip.files).filter((name) => !zip.files[name].dir)
  if (names.length === 0) return bytes
  zip.remove(names[Math.floor(r() * names.length)])
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })
}

/** Replace a random entry's content with garbage bytes. */
const entryGarbage: Corrupt = async (bytes, r) => {
  const zip = await JSZip.loadAsync(bytes)
  const names = Object.keys(zip.files).filter((name) => !zip.files[name].dir)
  if (names.length === 0) return bytes
  const victim = names[Math.floor(r() * names.length)]
  const junk = new Uint8Array(64 + Math.floor(r() * 512))
  for (let i = 0; i < junk.length; i++) junk[i] = Math.floor(r() * 256)
  zip.file(victim, junk)
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })
}

/** Trailing junk after the end of central directory. */
const trailingJunk: Corrupt = async (bytes, r) => {
  const junk = new Uint8Array(16 + Math.floor(r() * 256))
  for (let i = 0; i < junk.length; i++) junk[i] = Math.floor(r() * 256)
  const out = new Uint8Array(bytes.length + junk.length)
  out.set(bytes)
  out.set(junk, bytes.length)
  return out
}

const CORRUPTIONS: Array<[string, Corrupt]> = [
  ['bitflip', bitflip],
  ['truncate', truncate],
  ['entrydrop', entryDrop],
  ['entrygarbage', entryGarbage],
  ['trailingjunk', trailingJunk],
]

/** Run the app's own import (and optionally its editor/save stages) over the
 *  corrupted bytes. Everything is caught: a thrown Error is a legitimate
 *  refusal and the point of the exercise is that refusal stays clean and
 *  bounded — hangs and process crashes fail the test on their own. */
async function attemptPipeline(bytes: Uint8Array, withEditor: boolean): Promise<string> {
  try {
    const parsed = await parseDocx(bytes)
    if (!withEditor) return 'parsed'
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    try {
      const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
      const saved = await saveDocx(parsed, plan.saveBlocks)
      // whatever came out must itself be a package a later open can read
      await parseDocx(saved)
      return 'saved'
    } finally {
      editor.destroy()
    }
  } catch {
    return 'refused'
  }
}

describe('hostile docx import fuzz', () => {
  for (const [mode, corrupt] of CORRUPTIONS) {
    for (let seed = 1; seed <= SEEDS; seed++) {
      it(`seed ${seed} ${mode}: parse refuses cleanly or succeeds, never hangs`, async () => {
        const valid = await buildFuzzDocx(seed)
        const bytes = await corrupt(valid, rng(seed * 7919 + 31))
        const outcome = await attemptPipeline(bytes, false)
        expect(['refused', 'parsed']).toContain(outcome)
      }, 20000)
    }
    for (let seed = 1; seed <= PIPELINE_SEEDS; seed++) {
      it(`seed ${seed} ${mode}: editor+save stages stay bounded on a damaged package`, async () => {
        const valid = await buildFuzzDocx(seed)
        const bytes = await corrupt(valid, rng(seed * 7919 + 53))
        const outcome = await attemptPipeline(bytes, true)
        expect(['refused', 'parsed', 'saved']).toContain(outcome)
      }, 30000)
    }
  }

  // Report D's resource classes, as fixed constructions: each has a named
  // ceiling in the engine, and the ceiling must answer with a clean refusal.
  describe('resource bombs', () => {
    it('an entry flood hits the named ceiling without hanging', async () => {
      const zip = new JSZip()
      for (let i = 0; i < 5000; i++) zip.file(`flood/${i}.xml`, '<a/>')
      const bytes = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })
      await attemptPipeline(bytes, false)
    }, 30000)

    it('deeply nested XML hits the tag-depth ceiling without hanging', async () => {
      const depth = 20000
      const body = '<w:p><w:r><w:t>x</w:t></w:r></w:p>'
      const bytes = await buildDocxWithDocumentXml(
        `${'<w:evil>'.repeat(depth)}${body}${'</w:evil>'.repeat(depth)}`,
      )
      await attemptPipeline(bytes, false)
    }, 30000)

    it('an attribute bomb hits the attribute ceiling without hanging', async () => {
      const big = 'A'.repeat(8 * 1024 * 1024)
      const bytes = await buildDocxWithDocumentXml(`<w:p><w:r><w:t x="${big}">x</w:t></w:r></w:p>`)
      await attemptPipeline(bytes, false)
    }, 60000)
  })
})

/** A minimal valid package whose document.xml body is replaced wholesale. */
async function buildDocxWithDocumentXml(bodyXml: string): Promise<Uint8Array> {
  const zip = new JSZip()
  const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
  zip.file(
    '[Content_Types].xml',
    `${XML_DECL}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
      '</Types>',
  )
  zip.file(
    '_rels/.rels',
    `${XML_DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
      '</Relationships>',
  )
  zip.file(
    'word/document.xml',
    `${XML_DECL}<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${bodyXml}</w:body></w:document>`,
  )
  zip.forEach((_path, entry) => {
    entry.date = new Date(Date.UTC(2026, 0, 1))
  })
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })
}
