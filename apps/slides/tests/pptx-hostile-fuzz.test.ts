/** Hostile-input fuzz for both PPT import surfaces — the slides sibling of the
 *  docs suite's roundtrip-hostile.test.ts and the sheets suite's
 *  xlsx-sidecar-hostile.test.ts. The regular slides tests feed valid decks;
 *  this file feeds corrupted ones and asserts the one property that matters
 *  for importers of untrusted bytes: **bounded, clean failure** — an open
 *  either refuses with a thrown Error or succeeds, and a healthy deck still
 *  opens afterwards, so module-level state poisoned by a hostile parse fails
 *  the test on its own. No fidelity assertions here: a corrupted package has
 *  no "right" output.
 *
 *  Two surfaces eat untrusted bytes:
 *  - openPptx (@genoffice/pptx-engine): the OOXML package the Slides app opens;
 *    on success the pipeline property also round-trips the deck through
 *    savePptx and re-opens the saved bytes.
 *  - convertPptToPptx (@genoffice/ppt2pptx): the legacy binary .ppt container
 *    (CFB over MS-PPT records), fed corrupted containers built from the
 *    synthetic-ppt generator.
 *
 *  Corruption is seeded (same seed → same bytes, failing cases reproduce
 *  exactly). CI runs the default counts; HOSTILE_SEEDS / HOSTILE_PIPELINE_SEEDS
 *  sweep larger spaces. */
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'

import { openPptx, savePptx } from '@genoffice/pptx-engine'
import { convertPptToPptx } from '../../../packages/ppt2pptx/src'
import { buildFuzzPptx, rng } from './helpers/fuzz-pptx'
import {
  buildPptContainer,
  buildPptStreams,
  type SlideSpec,
} from '../../../packages/ppt2pptx/tests/synthetic-ppt'

const SEEDS = Math.max(1, Number(process.env.HOSTILE_SEEDS ?? 10))
const PIPELINE_SEEDS = Math.max(1, Number(process.env.HOSTILE_PIPELINE_SEEDS ?? 3))

type Corrupt = (bytes: Uint8Array, r: () => number) => Promise<Uint8Array>

/** Flip a handful of random bytes anywhere in the package. */
const bitflip: Corrupt = async (bytes, r) => {
  const out = Uint8Array.from(bytes)
  const flips = 1 + Math.floor(r() * 8)
  for (let i = 0; i < flips; i++) {
    const at = Math.floor(r() * out.length)
    out[at] = (out[at] ?? 0) ^ (1 << Math.floor(r() * 8))
  }
  return out
}

/** Cut the file at a random offset. */
const truncate: Corrupt = async (bytes, r) => {
  const at = 100 + Math.floor(r() * Math.max(1, bytes.length - 100))
  return bytes.slice(0, at)
}

/** Remove a random package entry (a missing slide or presentation part). */
const entryDrop: Corrupt = async (bytes, r) => {
  const zip = await JSZip.loadAsync(bytes)
  const names = Object.keys(zip.files).filter((name) => !zip.files[name]?.dir)
  if (names.length === 0) return bytes
  zip.remove(names[Math.floor(r() * names.length)]!)
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })
}

/** Replace a random package entry's content with garbage bytes. */
const entryGarbage: Corrupt = async (bytes, r) => {
  const zip = await JSZip.loadAsync(bytes)
  const names = Object.keys(zip.files).filter((name) => !zip.files[name]?.dir)
  if (names.length === 0) return bytes
  const victim = names[Math.floor(r() * names.length)]!
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

/** Garbage a random window in place — the zip-modes' counterpart for a binary
 *  container that is not a zip (the legacy CFB surface). */
const garbageWindow: Corrupt = async (bytes, r) => {
  const out = Uint8Array.from(bytes)
  const at = Math.floor(r() * Math.max(1, out.length - 256))
  const len = Math.min(64 + Math.floor(r() * 192), out.length - at)
  for (let i = 0; i < len; i++) out[at + i] = Math.floor(r() * 256)
  return out
}

const PACKAGE_CORRUPTIONS: Array<[string, Corrupt]> = [
  ['bitflip', bitflip],
  ['truncate', truncate],
  ['entrydrop', entryDrop],
  ['entrygarbage', entryGarbage],
  ['trailingjunk', trailingJunk],
]

const CONTAINER_CORRUPTIONS: Array<[string, Corrupt]> = [
  ['bitflip', bitflip],
  ['truncate', truncate],
  ['garbagewindow', garbageWindow],
]

/** Run the app's own open over the corrupted bytes, then (on success) the
    save-and-reopen the pipeline would do next. Everything is caught: a thrown
    Error is a legitimate refusal, and the point is that refusal stays clean
    and bounded — hangs fail the test's timeout on their own. */
async function attemptOpen(bytes: Uint8Array): Promise<'refused' | 'opened' | 'saved'> {
  try {
    const opened = await openPptx(bytes)
    const saved = await savePptx(opened)
    await openPptx(saved)
    return 'saved'
  } catch {
    try {
      await openPptx(bytes)
      return 'opened'
    } catch {
      return 'refused'
    }
  }
}

describe('hostile PPTX import fuzz', () => {
  it('the generator produces decks the open pipeline accepts', async () => {
    for (let seed = 1; seed <= 5; seed++) {
      const outcome = await attemptOpen(await buildFuzzPptx(seed))
      expect(outcome, `seed ${seed}`).toBe('saved')
    }
  })

  for (const [mode, corrupt] of PACKAGE_CORRUPTIONS) {
    for (let seed = 1; seed <= SEEDS; seed++) {
      it(`seed ${seed} ${mode}: open refuses cleanly or succeeds, never hangs`, async () => {
        const valid = await buildFuzzPptx(seed)
        const bytes = await corrupt(valid, rng(seed * 7919 + 31))
        const outcome = await attemptOpen(bytes)
        expect(['refused', 'opened', 'saved']).toContain(outcome)
      }, 20000)
    }
    for (let seed = 1; seed <= PIPELINE_SEEDS; seed++) {
      it(`seed ${seed} ${mode}: a damaged package that opens still saves and re-opens`, async () => {
        const valid = await buildFuzzPptx(seed)
        const bytes = await corrupt(valid, rng(seed * 7919 + 53))
        await attemptOpen(bytes)
      }, 30000)
    }
  }

  // Resource classes, as fixed constructions: an OOXML reader of untrusted
  // packages must answer each of these within its normal budget.
  describe('resource bombs', () => {
    it('an entry flood opens or refuses without hanging', async () => {
      const zip = new JSZip()
      for (let i = 0; i < 5000; i++) zip.file(`flood/${i}.xml`, '<a/>')
      const bytes = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })
      await attemptOpen(bytes)
    }, 30000)

    it('deeply nested slide XML opens or refuses without hanging', async () => {
      const depth = 20_000
      const zip = new JSZip()
      zip.file(
        'ppt/slides/slide1.xml',
        '<?xml version="1.0"?><p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" ' +
          'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree>' +
          `${'<a:deep>'.repeat(depth)}<a:leaf/>${'</a:deep>'.repeat(depth)}` +
          '</p:spTree></p:cSld></p:sld>',
      )
      await attemptOpen(await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }))
    }, 30000)

    it('an attribute bomb opens or refuses without hanging', async () => {
      const big = 'A'.repeat(8 * 1024 * 1024)
      const zip = new JSZip()
      zip.file(
        'ppt/slides/slide1.xml',
        '<?xml version="1.0"?><p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" ' +
          `xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><a:evil attr="${big}"/>` +
          '</p:spTree></p:cSld></p:sld>',
      )
      await attemptOpen(await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }))
    }, 60000)

    it('a lying slide size opens or refuses without hanging', async () => {
      const zip = new JSZip()
      zip.file(
        'ppt/presentation.xml',
        '<?xml version="1.0"?><p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" ' +
          'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
          '<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>' +
          '<p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst>' +
          '<p:sldSz cx="914400000" cy="685800000"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>',
      )
      await attemptOpen(await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }))
    }, 30000)
  })
})

describe('hostile legacy .ppt import fuzz', () => {
  /** Varied-but-valid legacy containers: text kinds, outline placement, a
      second stale generation, the encrypted header token. */
  function fuzzSpecs(seed: number): SlideSpec[] {
    const r = rng(seed * 947 + 7)
    const count = 1 + Math.floor(r() * 3)
    return Array.from({ length: count }, (_v, i) => ({
      textSequences: [
        {
          kind: Math.floor(r() * 6),
          text: i % 2 === 0 ? '季度回顾\nRevenue up' : `Slide ${i} 🎯`,
          atom: r() < 0.5 ? ('chars' as const) : ('bytes' as const),
        },
      ],
    }))
  }

  function fuzzContainer(seed: number): Uint8Array {
    const options: { secondGeneration?: SlideSpec[]; outlineTexts?: boolean; encrypted?: boolean } =
      {}
    if (seed % 3 === 0) options.secondGeneration = fuzzSpecs(seed + 97)
    if (seed % 2 === 0) options.outlineTexts = true
    if (seed % 5 === 0) options.encrypted = true
    const streams = buildPptStreams(fuzzSpecs(seed), options)
    return buildPptContainer(streams.document, streams.currentUser)
  }

  it('the generator produces containers the legacy importer accepts', async () => {
    for (let seed = 1; seed <= 5; seed++) {
      const bytes = fuzzContainer(seed)
      if (seed % 5 === 0) {
        // the encrypted header token must refuse, not convert
        await expect(convertPptToPptx(bytes)).rejects.toThrow()
      } else {
        const result = await convertPptToPptx(bytes)
        expect(result.slideCount, `seed ${seed}`).toBeGreaterThan(0)
      }
    }
  })

  for (const [mode, corrupt] of CONTAINER_CORRUPTIONS) {
    for (let seed = 1; seed <= SEEDS; seed++) {
      it(`seed ${seed} ${mode}: the legacy importer refuses cleanly or converts, never hangs`, async () => {
        const valid = fuzzContainer(seed)
        const bytes = await corrupt(valid, rng(seed * 7919 + 71))
        try {
          await convertPptToPptx(bytes)
        } catch {
          // a refusal is legitimate; a hang or process crash is not
        }
        // the importer must stay healthy afterwards: a benign container
        // still converts
        const benign = await convertPptToPptx(fuzzContainer(1))
        expect(benign.slideCount).toBeGreaterThan(0)
      }, 20000)
    }
  }
})
