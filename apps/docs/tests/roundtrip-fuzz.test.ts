/** Seeded round-trip fuzz for the DOCX save pipeline — the in-repo
 *  continuation of the genoffice#2381 editors_trial methodology. The external trial
 *  measured one hand-built document; this suite asserts the same properties
 *  over many random combinations (see helpers/fuzz-docx.ts):
 *
 *  1. untouched: parse -> editor -> save plan -> saveDocx is byte-identical to
 *     the input, and a second pass over the output is a byte-identical fixed
 *     point (the trial's "verbatim untouched blocks" + idempotence claims, as
 *     a general property);
 *  2. edit: regenerating one paragraph keeps every other package part
 *     byte-identical and every untouched block's original XML slice intact.
 *
 *  CI runs the default 24 seeds; `scripts/roundtrip-sweep.mjs N` sweeps N
 *  seeds locally. A failing seed reproduces exactly via buildFuzzDocx(seed). */
import { Editor } from '@tiptap/core'
import JSZip from 'jszip'
import { parseDocx, saveDocx } from '@genoffice/docx-engine'
import { describe, expect, it } from 'vitest'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { buildFuzzDocx } from './helpers/fuzz-docx'

const SEEDS = Math.max(1, Number(process.env.FUZZ_SEEDS ?? 24))
// the first block of every fuzz document, the guaranteed edit target
const anchorText = (seed: number) => `Anchor seed ${seed}`

async function partBytes(bytes: Uint8Array): Promise<Map<string, Uint8Array>> {
  const zip = await JSZip.loadAsync(bytes)
  const out = new Map<string, Uint8Array>()
  for (const [name, entry] of Object.entries(zip.files)) {
    if (!entry.dir) out.set(name, await entry.async('uint8array'))
  }
  return out
}

/** The app's own import/export path, exactly what the external trial drove. */
async function saveThroughEditor(bytes: Uint8Array): Promise<Uint8Array> {
  const parsed = await parseDocx(bytes)
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  try {
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    return await saveDocx(parsed, plan.saveBlocks)
  } finally {
    editor.destroy()
  }
}

describe('seeded docx round-trip fuzz', () => {
  for (let seed = 1; seed <= SEEDS; seed++) {
    it(`seed ${seed}: untouched save is a byte-identical fixed point`, async () => {
      const bytes = await buildFuzzDocx(seed)
      const saved = await saveThroughEditor(bytes)
      expect(Buffer.from(saved).equals(Buffer.from(bytes))).toBe(true)
      // and the output is itself a fixed point of the same pipeline
      const resaved = await saveThroughEditor(saved)
      expect(Buffer.from(resaved).equals(Buffer.from(saved))).toBe(true)
    }, 60000)

    it(`seed ${seed}: regenerating the anchor keeps every other part byte-identical`, async () => {
      const bytes = await buildFuzzDocx(seed)
      const parsed = await parseDocx(bytes)
      const editor = new Editor({
        element: document.createElement('div'),
        extensions: editorExtensions,
        content: blocksToPmDoc(parsed.blocks) as never,
      })
      try {
        const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
        const visible = parsed.blocks.filter((b) => !b.hidden)
        const anchor = visible.find((b) =>
          (b.runs ?? []).some((run) => run.text === anchorText(seed)),
        )
        expect(anchor, 'anchor block exists').toBeTruthy()
        const finalBlocks = plan.saveBlocks.map((block, i) =>
          block.kind === 'original' && block.docxIndex === anchor!.docxIndex
            ? {
                kind: 'generated' as const,
                block: { type: 'paragraph' as const, runs: [{ text: `rewritten ${seed}` }] },
              }
            : plan.saveBlocks[i],
        )
        const saved = await saveDocx(parsed, finalBlocks)

        // every package part except the body keeps its exact bytes
        const before = await partBytes(bytes)
        const after = await partBytes(saved)
        expect([...after.keys()].sort()).toEqual([...before.keys()].sort())
        for (const [name, content] of before) {
          if (name === 'word/document.xml') continue
          expect(after.get(name), name).toEqual(content)
        }

        // untouched blocks keep their original XML slices (text retention)
        const docXml = new TextDecoder().decode(after.get('word/document.xml')!)
        for (const block of visible) {
          if (block.docxIndex === anchor!.docxIndex) continue
          expect(docXml).toContain(block.originalXml!)
        }
        expect(docXml).toContain(`rewritten ${seed}`)
      } finally {
        editor.destroy()
      }
    }, 60000)
  }
})
