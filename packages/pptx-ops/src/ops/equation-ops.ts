/**
 * Math in a slide: PowerPoint 2010+ stores an equation as a paragraph-level
 * <mc:AlternateContent> whose Choice (a14) carries OMML and whose Fallback is
 * a plain run. The engine keeps that block as a TextRun.rawXml with the
 * linearized text as its display text, so older readers and our renderer see
 * the fallback while PowerPoint shows the typeset formula.
 */
import { latexToOmml } from '@genoffice/docx-engine/math'
import { addElement, type Paragraph, type TextElement, type TextRun } from '@genoffice/pptx-engine'
import { ommlToText } from './omml-linear'
import {
  GuidedError,
  register,
  resolveElement,
  resolveSlide,
  type Op,
  type OpRecord,
} from './registry'

const MC_NS = 'http://schemas.openxmlformats.org/markup-compatibility/2006'
const A14_NS = 'http://schemas.microsoft.com/office/drawing/2010/main'
const M_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/math'

function escapeText(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
}

/** The run PowerPoint writes for an inline equation, with `text` as its fallback. */
export function equationRun(latex: string): TextRun {
  const inner = latexToOmml(latex)
  const oMath = `<m:oMath xmlns:m="${M_NS}">${inner}</m:oMath>`
  const text = ommlToText(oMath) || latex.trim()
  const rawXml =
    `<mc:AlternateContent xmlns:mc="${MC_NS}"><mc:Choice xmlns:a14="${A14_NS}" Requires="a14">` +
    `<a14:m><m:oMathPara xmlns:m="${M_NS}"><m:oMathParaPr><m:jc m:val="centerGroup"/></m:oMathParaPr>${oMath}</m:oMathPara></a14:m>` +
    `</mc:Choice><mc:Fallback><a:r><a:rPr lang="en-US" dirty="0"/><a:t>${escapeText(text)}</a:t></a:r></mc:Fallback></mc:AlternateContent>`
  return { text, rawXml }
}

function requireLatex(op: Op): string {
  if (typeof op.latex !== 'string' || !op.latex.trim()) {
    throw new GuidedError(
      `op "${op.op}" needs "latex": the formula, e.g. "E = mc^2" or "\\\\frac{a}{b}".`,
    )
  }
  try {
    latexToOmml(op.latex)
  } catch (e) {
    throw new GuidedError(
      `op "${op.op}": "latex" uses syntax the converter does not cover (${(e as Error).message}). Supported: \\\\frac, \\\\sqrt, ^, _, \\\\sum/\\\\int with limits, \\\\left( \\\\right), matrices, Greek letters, \\\\text{}.`,
    )
  }
  return op.latex
}

function requirePosition(op: Op): 'start' | 'end' {
  if (op.position === undefined || op.position === 'end') return 'end'
  if (op.position === 'start') return 'start'
  throw new GuidedError(`op "${op.op}": "position" must be "start" or "end" (default end).`)
}

function optionalBox(op: Op): { x: number; y: number; cx: number; cy: number } | undefined {
  if (op.box === undefined) return undefined
  const b = op.box as Record<string, unknown> | null
  const n = (k: string) =>
    typeof b?.[k] === 'number' && Number.isFinite(b[k]) ? (b[k] as number) : null
  const x = n('x')
  const y = n('y')
  const cx = n('cx') ?? n('w')
  const cy = n('cy') ?? n('h')
  if (x == null || y == null || cx == null || cy == null || cx <= 0 || cy <= 0) {
    throw new GuidedError(
      `op "${op.op}": "box" must be an EMU rect {x, y, cx, cy} for the new text box.`,
    )
  }
  return { x, y, cx, cy }
}

register({
  name: 'insertEquation',
  validate(op, ctx) {
    requireLatex(op)
    requirePosition(op)
    const box = optionalBox(op)
    if (box) {
      resolveSlide(ctx, op)
      if (op.target?.el !== undefined) {
        throw new GuidedError(
          `op "${op.op}": give either target.el (existing text) or "box" (new text box), not both.`,
        )
      }
      return
    }
    if (op.target?.el === undefined) {
      throw new GuidedError(
        `op "${op.op}" needs target.el (a text or shape element) or "box": {x, y, cx, cy} to create a text box for the formula.`,
      )
    }
    resolveElement(ctx, op, { types: ['text', 'shape'] })
  },
  apply(op, ctx): OpRecord {
    const run = equationRun(requireLatex(op))
    const position = requirePosition(op)
    const box = optionalBox(op)
    const paragraph: Paragraph = { runs: [run], align: 'center' }
    if (box) {
      const { slide } = resolveSlide(ctx, op)
      const el = addElement(slide, { kind: 'textbox', offset: box, paragraphs: [paragraph] })
      return { op, created: [el.id], after: { text: run.text } }
    }
    const { el } = resolveElement(ctx, op, { types: ['text', 'shape'] })
    const te = el as TextElement
    if (!te.text) te.text = { paragraphs: [] }
    const paragraphs = te.text.paragraphs
    const onlyMark =
      paragraphs.length === 1 && paragraphs[0]!.runs.every((r) => !r.text && !r.field && !r.rawXml)
    if (onlyMark) paragraphs.splice(0, 1, paragraph)
    else if (position === 'start') paragraphs.unshift(paragraph)
    else paragraphs.push(paragraph)
    te.dirty = true
    return { op, after: { text: run.text, paragraphs: paragraphs.length } }
  },
})
