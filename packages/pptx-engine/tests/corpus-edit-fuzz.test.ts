/**
 * Opt-in corpus sweep: apply every engine mutator to every element of each deck in CORPUS_DIRS,
 * save, and report schema violations the edit introduced (tools/ooxml-validate, differential).
 *   CORPUS_DIRS=/path/a,/path/b FUZZ_OUT=/tmp/fuzz.json npx vitest run tests/corpus-edit-fuzz.test.ts
 * Skipped without CORPUS_DIRS; the report is written to FUZZ_OUT, never asserted.
 */
import { it } from 'vitest'
import fs from 'node:fs'
import path from 'node:path'
import { validatePptx, newProblems } from '../../../tools/ooxml-validate/validate-pptx.mjs'
import * as E from '../src/index'
import type { TextElement, GroupElement } from '../src/types'

const DIRS = (process.env.CORPUS_DIRS ?? '').split(',').filter(Boolean)
const OUT = process.env.FUZZ_OUT ?? '/tmp/fuzz-report.json'
const MAX_BYTES = 30 * 1024 * 1024

const files = DIRS.flatMap((d) =>
  fs
    .readdirSync(d)
    .filter((f) => f.endsWith('.pptx') && !f.startsWith('~$'))
    .map((f) => path.join(d, f)),
).filter((f) => fs.statSync(f).size <= MAX_BYTES)

type Rec = {
  deck: string
  step: string
  error?: string
  problems?: { part: string; message: string }[]
}
const report: Rec[] = []
const step = (deck: string, name: string, fn: () => unknown) => {
  try {
    fn()
  } catch (e) {
    report.push({ deck, step: name, error: String((e as Error)?.message ?? e).slice(0, 300) })
  }
}

it.skipIf(!DIRS.length)(
  'corpus edit fuzz',
  async () => {
    let n = 0
    for (const file of files) {
      const deck = path.basename(file)
      n++
      const bytes = fs.readFileSync(file)
      let before: Awaited<ReturnType<typeof validatePptx>>
      try {
        before = await validatePptx(bytes)
      } catch (e) {
        report.push({ deck, step: 'validate-original', error: String(e).slice(0, 300) })
        continue
      }
      let opened: E.OpenedPptx
      try {
        opened = await E.openPptx(bytes)
      } catch (e) {
        report.push({ deck, step: 'open', error: String((e as Error).message).slice(0, 300) })
        continue
      }
      const slides = opened.deck.slides.slice(0, 8)
      slides.forEach((slide, si) => {
        const ids = slide.elements.map((e) => e.id)
        for (const id of ids) {
          const el = slide.elements.find((e) => e.id === id)
          if (!el) continue
          const isText = el.type === 'text' || el.type === 'shape'
          if (isText && (el as TextElement).text?.paragraphs.length) {
            step(deck, 'setFont', () =>
              E.setElementFont(slide, id, { color: '#FFFFFF', fontSizePt: 14, bold: true }),
            )
            step(deck, 'setParagraphFormat', () =>
              E.setElementParagraphFormat(slide, id, {
                align: 'center',
                bullet: 'char',
                lineSpacing: 1.5,
              } as any),
            )
            step(deck, 'setTextBodyProps', () =>
              E.setElementTextBodyProps(slide, id, {
                autofit: 'shrink',
                insets: { l: 91440, t: 45720 },
                wrap: true,
              }),
            )
            step(deck, 'setTextAnchor', () => E.setElementTextAnchor(slide, id, 'middle'))
            step(deck, 'setLink', () =>
              E.setElementLink(opened, si, id, {
                kind: 'url',
                url: 'https://example.com/?a=1&b=2',
              } as any),
            )
            // force the rebuild path: run count changes
            step(deck, 'rebuildText', () => {
              const t = (el as TextElement).text!
              t.paragraphs[0]!.runs.push({ text: ' <added & "quoted">' })
              if (t.paragraphs.length > 1) t.paragraphs.pop()
              el.dirty = true
            })
          }
          if (el.type === 'shape') {
            step(deck, 'setFill', () => E.setElementFill(opened, slide, id, '#FF880080'))
            step(deck, 'setEffects', () =>
              E.setElementEffects(slide, id, {
                shadow: { color: '#00000080', blurRad: 50800, dist: 38100, dirDeg: 45 },
                glow: { color: '#FFFF00', radius: 63500 },
                softEdge: 12700,
              } as any),
            )
            step(deck, 'setPresetGeometry', () => E.setShapePresetGeometry(slide, id, 'roundRect'))
            step(deck, 'setAdjust', () => E.setShapeAdjustValues(slide, id, { adj: 25000 }))
          }
          if (el.type === 'shape' || el.type === 'text' || el.type === 'picture') {
            step(deck, 'setStroke', () => {
              ;(el as TextElement).stroke = E.strokePatchToModel({
                color: '#FF0000',
                widthEmu: 12700,
                dash: 'dash',
                cap: 'rnd',
              } as any)
              el.dirtyStroke = true
            })
          }
          if (el.type === 'picture') {
            step(deck, 'srcRect', () =>
              E.editPictureSrcRect(slide, id, { l: 0.1, t: 0.1, r: 0.1, b: 0.1 }),
            )
            step(deck, 'opacity', () => E.setPictureOpacity(slide, id, 0.5))
          }
          if (el.type === 'group') {
            const grp = el as GroupElement
            for (const child of (grp.children ?? []).slice(0, 4)) {
              const cid = child.id
              step(deck, 'groupChildFont', () =>
                E.setGroupChildFont(slide, id, cid, { color: '#00FF00' }),
              )
              step(deck, 'groupChildFill', () => E.editGroupChildFill(slide, id, cid, '#112233'))
              step(deck, 'groupChildStroke', () =>
                E.editGroupChildStroke(slide, id, cid, { color: '#000000', widthEmu: 6350 }),
              )
              step(deck, 'groupChildPara', () =>
                E.setGroupChildParagraphFormat(slide, id, cid, { align: 'right' }),
              )
              step(deck, 'groupChildTransform', () =>
                E.editGroupChildTransform(
                  slide,
                  id,
                  cid,
                  {
                    x: child.transform.offset.x + 100,
                    y: child.transform.offset.y,
                    cx: child.transform.offset.cx,
                    cy: child.transform.offset.cy,
                  },
                  0,
                ),
              )
            }
          }
          if (el.type === 'table') {
            step(deck, 'tableCellText', () =>
              E.editTableCellText(slide, id, 0, 0, [{ runs: [{ text: 'cell & <x>' }] }]),
            )
            step(deck, 'tableColWidth', () => E.setTableColWidth(slide, id, 0, 1234567))
            step(deck, 'tableRowHeight', () => E.setTableRowHeight(slide, id, 0, 456789))
            step(deck, 'tableCellAnchor', () => E.setTableCellAnchor(slide, id, 0, 0, 'middle'))
          }
          step(deck, 'transform', () => {
            el.transform.offset.x += 100
            el.transform.rot = Math.round(15 * 60000)
            el.dirtyTransform = true
          })
        }
        step(deck, 'reorder', () => ids[0] && E.reorderElement(slide, ids[0], 'front'))
        step(deck, 'background', () => E.setSlideBackground(opened, slide, '#0000FF'))
        step(deck, 'transition', () => E.setSlideTransition(slide, 'fade' as any))
        step(deck, 'notes', () => E.setSlideNotes(opened, si, 'note & <text>'))
      })
      step(deck, 'duplicateSlide', () => E.duplicateSlide(opened, 0))
      step(deck, 'addElement', () =>
        E.addElement(opened.deck.slides[0]!, {
          kind: 'rect',
          offset: { x: 0, y: 0, cx: 100000, cy: 100000 },
          paragraphs: [{ runs: [{ text: 'new' }] }],
        }),
      )
      step(deck, 'addTable', () =>
        E.addTable(opened, 0, {
          rows: 2,
          cols: 2,
          offset: { x: 0, y: 0, cx: 1000000, cy: 500000 },
        }),
      )
      let saved: Uint8Array
      try {
        saved = await E.savePptx(opened)
      } catch (e) {
        report.push({ deck, step: 'save', error: String((e as Error).message).slice(0, 300) })
        continue
      }
      try {
        const after = await validatePptx(saved)
        const fresh = newProblems(before, after)
        if (fresh.length) report.push({ deck, step: 'schema', problems: fresh.slice(0, 40) })
      } catch (e) {
        report.push({ deck, step: 'validate-edited', error: String(e).slice(0, 300) })
      }
      if (n % 10 === 0)
        fs.writeFileSync(OUT, JSON.stringify({ done: n, total: files.length, report }, null, 1))
    }
    fs.writeFileSync(OUT, JSON.stringify({ done: n, total: files.length, report }, null, 1))
  },
  3_600_000,
)
