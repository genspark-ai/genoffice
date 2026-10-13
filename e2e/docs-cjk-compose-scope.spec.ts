import { test, expect } from '@playwright/test'
import { mkdirSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import JSZip from 'jszip'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'

/**
 * Fidelity measurement for the composition-scoped EA font rule
 * (.ProseMirror-composing .doc-east-asian-font in styles.css).
 *
 * Scenario from the review: a document whose EA face is not installed on the
 * host under the metric strut the app emits for exactly that case (the
 * "standIn" path in doc-style-css.ts). The span's rendered declaration stays
 * EA-primary, so the rendered line geometry must equal main's byte-for-byte;
 * only a live composition may reflow the line under the caret. This spec
 * measures per-line box heights in the real editor and pins: default == a
 * composition round-trip, default == main simulated (our rule
 * overridden by the EA-primary declaration main's spans carried, composition
 * live), and reports the composing-line delta the rule may cause.
 *
 * Compositions are driven by dispatching real compositionstart/compositionend
 * events, never by touching classList: the measurement must go through the
 * ScriptFonts plumbing that puts .ProseMirror-composing on the editor root,
 * or it would pass with that plumbing missing.
 *
 * The EA face is picked at runtime: a probe session checks Word's CJK faces
 * with the same canvas width trick line-metrics' isFontAvailable uses, and
 * the first one the host is missing names the fixture's w:eastAsia.
 */

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

const HAN_TEXT =
  '\u5728\u7cfb\u7edf\u5de5\u7a0b\u5b9e\u8df5\u4e2d\uff0c\u53ef\u9760\u6027\u6307\u6807\u9700\u8981\u8d2f\u7a7f\u8bbe\u8ba1\u3001\u7814\u5236\u3001\u8bd5\u9a8c\u4e0e\u670d\u5f79\u7684\u5168\u5bff\u547d\u5468\u671f\uff0c\u5e76\u901a\u8fc7\u5206\u9636\u6bb5\u8bc4\u5ba1\u6301\u7eed\u9a8c\u8bc1\u5176\u6709\u6548\u6027\u3002'

/** Word-bundled CJK faces, most-missing-first on a plain macOS/Linux box */
const EA_FACE_CANDIDATES = [
  'MS Mincho',
  'Yu Mincho',
  'DFKai-SB',
  'FangSong',
  'SimSun-ExtB',
] as const

function cjkDocx(eastAsiaFace: string): Promise<Buffer> {
  const xml =
    `<?xml version="1.0" encoding="UTF-8"?><w:document ${W}><w:body>` +
    `<w:p><w:r><w:t>${HAN_TEXT}</w:t></w:r></w:p>` +
    `<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>` +
    `<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>` +
    '</w:body></w:document>'
  const styles =
    `<?xml version="1.0" encoding="UTF-8"?><w:styles ${W}>` +
    '<w:docDefaults><w:rPrDefault><w:rPr>' +
    `<w:rFonts w:ascii="Calibri" w:eastAsia="${eastAsiaFace}" w:hAnsi="Calibri"/>` +
    '<w:sz w:val="22"/><w:szCs w:val="22"/>' +
    '</w:rPr></w:rPrDefault></w:docDefaults></w:styles>'
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>',
  )
  zip.file(
    '_rels/.rels',
    '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
  )
  zip.file(
    'word/_rels/document.xml.rels',
    '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>',
  )
  zip.file('word/document.xml', xml)
  zip.file('word/styles.xml', styles)
  return zip.generateAsync({ type: 'nodebuffer' })
}

/** the docs URL can load with no editor mounted (renderer birth race): reload recovers */
async function awaitDocsEditor(page: import('@playwright/test').Page): Promise<void> {
  const editor = page.locator('.ProseMirror').first()
  try {
    await editor.waitFor({ timeout: 45_000 })
  } catch {
    await page.reload({ waitUntil: 'domcontentloaded' })
    await editor.waitFor({ timeout: 90_000 })
  }
}

/** the app's own availability check (line-metrics isFontAvailable): a face the
 *  fallback widths cannot distinguish from is not installed */
async function firstMissingEaFace(page: import('@playwright/test').Page): Promise<string | null> {
  return page.evaluate(
    (candidates) => {
      const canvas = document.createElement('canvas')
      const ctx = canvas.getContext('2d')
      if (!ctx) return null
      const sample = '\ud55c\uae00\u3042\u30a2\u4e2d\u6587abcWXYmm123'
      const widthWith = (family: string) => {
        ctx.font = `32px ${family}`
        return ctx.measureText(sample).width
      }
      for (const face of candidates) {
        const quoted = `"${face}"`
        const available =
          widthWith(`${quoted}, monospace`) !== widthWith('monospace') ||
          widthWith(`${quoted}, serif`) !== widthWith('serif')
        if (!available) return face
      }
      return null
    },
    EA_FACE_CANDIDATES as unknown as string[],
  )
}

test('the composing font rule leaves rendered lines byte-identical and reports its delta', async () => {
  test.setTimeout(360_000)
  const artifacts = join(__dirname, 'artifacts')
  mkdirSync(artifacts, { recursive: true })
  const docPath = join(artifacts, 'cjk-compose-fixture.docx')

  // probe pass: pick an EA face this host is missing so the standIn strut
  // activates, exactly as on a machine without the document's fonts. Home
  // has no '://home/' URL — launchShell's own page IS the Home window.
  const probe = await launchShell({ onboardingSeen: true, videoDir: 'docs-cjk-compose-probe' })
  // no initialiser: the try either assigns it or throws, so an initial `null`
  // could never be read (no-useless-assignment)
  let eaFace: string | null
  try {
    await probe.page.waitForTimeout(1500)
    eaFace = await firstMissingEaFace(probe.page)
  } finally {
    await closeAndSaveVideo(probe, 'docs-cjk-compose-probe')
  }
  console.log('[cjk-compose] EA face for the fixture:', eaFace)
  test.skip(eaFace === null, 'every candidate EA face is installed on this host')
  const face = eaFace as string
  writeFileSync(docPath, await cjkDocx(face))

  const launched = await launchShell({
    onboardingSeen: true,
    openFile: docPath,
    videoDir: 'docs-cjk-compose-scope',
  })
  const { app } = launched
  try {
    const page = await waitForPageWithUrl(app, '://docs/')
    await awaitDocsEditor(page)
    await page.waitForTimeout(2500)

    /** per-line box heights of the decorated paragraph, zoom-normalized */
    const measure = (): Promise<number[]> =>
      page.evaluate(() => {
        const p = document.querySelector('.doc-page p')
        if (!p) return []
        const zoom =
          parseFloat(
            getComputedStyle(document.querySelector<HTMLElement>('.doc-zoom')!).zoom ?? '1',
          ) || 1
        const range = document.createRange()
        range.selectNodeContents(p)
        const lines = [...range.getClientRects()]
          .filter((r) => r.height > 1 && r.width > 1)
          .map((r) => Math.round((r.height / zoom) * 100) / 100)
        const span = p.querySelector('.doc-east-asian-font')
        return {
          lines,
          paraH: Math.round((p.getBoundingClientRect().height / zoom) * 100) / 100,
          spans: document.querySelectorAll('.doc-east-asian-font').length,
          spanFamily: span ? getComputedStyle(span).fontFamily.slice(0, 60) : '(none)',
          composing:
            document.querySelector('.ProseMirror')?.classList.contains('ProseMirror-composing') ??
            false,
        }
      })

    type Read = {
      lines: number[]
      paraH: number
      spans: number
      spanFamily: string
      composing: boolean
    }
    const sameRead = (a: Read, b: Read) =>
      a.lines.length > 0 &&
      a.lines.length === b.lines.length &&
      a.lines.every((h, i) => h === b.lines[i]) &&
      a.paraH === b.paraH &&
      a.spans === b.spans &&
      a.spanFamily === b.spanFamily

    /**
     * Open or close a real composition on the editor root and report whether
     * .ProseMirror-composing came out the other side. The event is all this
     * touches: the ScriptFonts plugin owns the class, so this is the only way
     * to measure (or prove) that the plumbing works.
     */
    const setComposing = (
      page: import('@playwright/test').Page,
      on: boolean,
    ): Promise<boolean | null> =>
      page.evaluate((composing) => {
        const dom = document.querySelector('.ProseMirror')
        if (!dom) return null
        dom.dispatchEvent(
          new CompositionEvent(composing ? 'compositionstart' : 'compositionend', {
            bubbles: true,
            data: '',
          }),
        )
        return dom.classList.contains('ProseMirror-composing')
      }, on)

    /** first read that agrees with its predecessor (fonts/pagination settled) */
    const settledRead = async (applyState: () => Promise<void>): Promise<Read> => {
      let previous: Read | null = null
      await applyState()
      await expect
        .poll(
          async () => {
            const next = await measure()
            const stable = previous !== null && sameRead(previous, next)
            previous = next
            return stable ? next : null
          },
          { timeout: 30_000 },
        )
        .not.toBeNull()
      return previous as Read
    }

    await page.evaluate(() => document.fonts.ready)
    const rendered = await settledRead(async () => {})
    console.log(
      `[cjk-compose] rendered: lines ${rendered.lines.join(', ')} px, paraH ${rendered.paraH}px, spans ${rendered.spans}, family "${rendered.spanFamily}", composing=${rendered.composing}`,
    )
    test.info().annotations.push({
      type: 'measurement',
      description: `${face}: rendered lines ${rendered.lines.join(', ')} px, paragraph ${rendered.paraH} px, ${rendered.spans} spans, family ${rendered.spanFamily}`,
    })
    // the decorations must exist for the scenario to mean anything
    expect(rendered.spans).toBeGreaterThan(0)

    // a live composition puts .ProseMirror-composing on the editor root, which
    // is what makes our rule apply; measure the delta it may cause while typing
    // (reviewer: confined to the composition)
    const composing = await settledRead(async () => {
      expect(await setComposing(page, true)).toBe(true)
    })
    const delta = composing.lines.map((h, i) => Math.round((h - rendered.lines[i]) * 100) / 100)
    const paraDelta = Math.round((composing.paraH - rendered.paraH) * 100) / 100
    console.log(
      `[cjk-compose] composing: lines ${composing.lines.join(', ')} px (line delta ${delta.join(', ')}), paraH ${composing.paraH}px (para delta ${paraDelta}px), family "${composing.spanFamily}"`,
    )
    test.info().annotations.push({
      type: 'measurement',
      description: `${face}: composing-class paragraph ${composing.paraH} px (delta ${paraDelta}px), family ${composing.spanFamily}`,
    })

    // a composition round-trip restores the rendered geometry exactly
    const restored = await settledRead(async () => {
      expect(await setComposing(page, false)).toBe(false)
    })
    expect(sameRead(restored, rendered)).toBe(true)

    // main simulated: a live composition, but our rule overridden by the
    // EA-primary declaration main's spans carried — the geometry must be
    // indistinguishable from the default read
    const mainSimulated = await settledRead(async () => {
      await page.evaluate(() => {
        const style = document.createElement('style')
        style.dataset.mainSim = 'true'
        // !important and appended last: the live rule is !important too, and
        // among equals the later stylesheet wins — this reproduces main, where
        // the inline EA-primary declaration was the only word
        style.textContent =
          '.ProseMirror-composing .doc-east-asian-font { font-family: var(--doc-east-asian-font, inherit) !important }'
        document.head.append(style)
      })
      expect(await setComposing(page, true)).toBe(true)
    })
    expect(sameRead(mainSimulated, rendered)).toBe(true)
  } finally {
    await closeAndSaveVideo(launched, 'docs-cjk-compose-scope')
  }
})
