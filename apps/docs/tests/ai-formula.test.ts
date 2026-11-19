import { describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import {
  BLANK_BULLET_NUM_ID,
  BLANK_ORDERED_NUM_ID,
  buildBlankDocx,
  parseDocx,
  saveDocx,
} from '@genoffice/docx-engine'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { executeTool } from '../src/renderer/ai/tools'

/**
 * Guard tests for the AI <formula> tag: restricted HTML with LaTeX formulas
 * becomes native OMML through the same protected-block path as the insert
 * dialog, saves to docx, and reparses as formula blocks.
 */

const NUM_IDS = { bullet: BLANK_BULLET_NUM_ID, ordered: BLANK_ORDERED_NUM_ID }

async function createBlankEditor() {
  const { editorExtensions } = await import('../src/renderer/editor/extensions')
  const bytes = await buildBlankDocx()
  const parsed = await parseDocx(bytes)
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
  })
  editor.commands.setContent(blocksToPmDoc(parsed.blocks) as never)
  return { editor, parsed }
}

describe('AI formula generation: <formula> tag', () => {
  it('<formula> inside insert_content generates a native formula block that reparses after save', async () => {
    const { editor, parsed } = await createBlankEditor()
    const exec = await executeTool(
      editor,
      {
        id: 't',
        name: 'insert_content',
        input: {
          html: [
            '<h1>Derivation</h1>',
            '<p>By the Pythagorean theorem:</p>',
            '<formula>c = \\sqrt{a^2 + b^2}</formula>',
          ].join('\n'),
        },
      },
      NUM_IDS,
    )
    expect(exec.isError).toBeFalsy()
    expect(exec.mutated).toBe(true)

    const json = editor.getJSON() as PmNode
    const formula = json.content!.find((n) => n.type === 'docProtected')
    expect(formula).toBeDefined()
    // freshly inserted formula blocks are labeled via t('editorEquation'); tests run under the zh default locale
    expect(formula!.attrs!.label).toBe('公式')
    expect(formula!.attrs!.previewText).toBe('c = \\sqrt{a^2 + b^2}')
    expect(String(formula!.attrs!.genXml)).toContain('<m:oMathPara>')

    const plan = pmDocToSavePlan(json, parsed.blocks)
    const saved = await saveDocx(parsed, plan.saveBlocks)
    const reparsed = await parseDocx(saved)
    const block = reparsed.blocks.find((b) => b.formulaDisplay)
    expect(block).toBeDefined()
    expect(block!.formulaDisplay!.mathml).toContain('<msqrt>')
    expect(block!.formulaDisplay!.tokens.join('')).toContain('=')
    editor.destroy()
  })

  it('<formula> inside <p> becomes an inline formula flowing with the text', async () => {
    const { editor } = await createBlankEditor()
    const exec = await executeTool(
      editor,
      {
        id: 't',
        name: 'insert_content',
        input: { html: '<p>Area formula <formula>A = \\pi r^2</formula> as above.</p>' },
      },
      NUM_IDS,
    )
    expect(exec.isError).toBeFalsy()
    const json = editor.getJSON() as PmNode
    const para = json.content!.find(
      (n) => n.type === 'docParagraph' && JSON.stringify(n).includes('Area formula'),
    )
    expect(para).toBeDefined()
    const inline = para!.content!.find((n) => n.type === 'docInlineMath')
    expect(inline).toBeDefined()
    expect(inline!.attrs!.latex).toBe('A = \\pi r^2')
    expect(String(inline!.attrs!.omml)).toContain('<m:oMath>')
    expect(String(inline!.attrs!.mathml)).toContain('display="inline"')
    // Text before and after the formula stays in the same paragraph
    const texts = para!.content!.filter((n) => n.type === 'text').map((n) => n.text)
    expect(texts.join('')).toContain('Area formula')
    expect(texts.join('')).toContain('as above.')
    editor.destroy()
  })

  it('inline formula save → reparses as a math run in a mixed paragraph; saving again keeps identical bytes', async () => {
    const { editor, parsed } = await createBlankEditor()
    await executeTool(
      editor,
      {
        id: 't',
        name: 'insert_content',
        input: { html: '<p>From <formula>E = mc^2</formula> we get mass-energy equivalence.</p>' },
      },
      NUM_IDS,
    )
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const saved = await saveDocx(parsed, plan.saveBlocks)
    editor.destroy()

    const reparsed = await parseDocx(saved)
    const block = reparsed.blocks.find((b) => b.runs?.some((r) => r.math))
    expect(block).toBeDefined()
    expect(block!.type).toBe('paragraph')
    const mathRun = block!.runs!.find((r) => r.math)!
    expect(mathRun.math!.omml).toContain('<m:sSup>')
    expect(block!.runs!.map((r) => r.text).join('')).toContain('mass-energy equivalence')

    // Reopen without edits, save again = bytes unchanged
    const { editorExtensions } = await import('../src/renderer/editor/extensions')
    const editor2 = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
    })
    editor2.commands.setContent(blocksToPmDoc(reparsed.blocks) as never)
    const plan2 = pmDocToSavePlan(editor2.getJSON() as PmNode, reparsed.blocks)
    expect(plan2.changedCount).toBe(0)
    expect(plan2.deletedCount).toBe(0)
    const saved2 = await saveDocx(reparsed, plan2.saveBlocks)
    expect(Buffer.from(saved2).equals(Buffer.from(saved))).toBe(true)
    editor2.destroy()
  })

  it('read_blocks → replace_blocks in the same session rewrites a formula paragraph; formula structure preserved', async () => {
    const { editor, parsed } = await createBlankEditor()
    await executeTool(
      editor,
      {
        id: 't1',
        name: 'insert_content',
        input: { html: '<p>Original text <formula>a^2 + b^2 = c^2</formula> ending.</p>' },
      },
      NUM_IDS,
    )
    // Same session: inline formula carries latex; read_blocks reads it back as LaTeX source
    const mathIndex = (editor.getJSON() as PmNode).content!.findIndex((n) =>
      JSON.stringify(n).includes('docInlineMath'),
    )
    const read = await executeTool(
      editor,
      {
        id: 't2',
        name: 'read_blocks',
        input: { startBlockIndex: mathIndex, endBlockIndex: mathIndex },
      },
      NUM_IDS,
    )
    expect(read.output).toContain('<formula>a^2 + b^2 = c^2</formula>')
    const exec = await executeTool(
      editor,
      {
        id: 't3',
        name: 'replace_blocks',
        input: {
          startBlockIndex: mathIndex,
          endBlockIndex: mathIndex,
          html: read.output.replace('Original text', 'Rewritten text'),
        },
      },
      NUM_IDS,
    )
    expect(exec.isError).toBeFalsy()

    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const saved = await saveDocx(parsed, plan.saveBlocks)
    editor.destroy()
    const final = await parseDocx(saved)
    const block = final.blocks.find((b) => b.runs?.some((r) => r.math))!
    expect(block.runs!.map((r) => r.text).join('')).toContain('Rewritten text')
    expect(block.runs!.find((r) => r.math)!.math!.omml).toContain('<m:sSup>')
  })

  it('invalid LaTeX fails the whole call and leaves the document unchanged', async () => {
    const { editor } = await createBlankEditor()
    const before = JSON.stringify(editor.getJSON())
    const exec = await executeTool(
      editor,
      {
        id: 't',
        name: 'insert_content',
        input: { html: '<p>Body text.</p><formula>\\notacommand{x}</formula>' },
      },
      NUM_IDS,
    )
    expect(exec.isError).toBe(true)
    expect(exec.output).toContain('Cannot parse the LaTeX')
    expect(JSON.stringify(editor.getJSON())).toBe(before)
    editor.destroy()
  })

  it('read_blocks serializes formula blocks as rewritable <formula> LaTeX', async () => {
    const { editor } = await createBlankEditor()
    await executeTool(
      editor,
      {
        id: 't1',
        name: 'insert_content',
        input: { html: '<formula>E = mc^2</formula>' },
      },
      NUM_IDS,
    )
    const exec = await executeTool(
      editor,
      {
        id: 't2',
        name: 'read_blocks',
        input: { startBlockIndex: 0, endBlockIndex: editor.state.doc.childCount - 1 },
      },
      NUM_IDS,
    )
    expect(exec.output).toContain('<formula>E = mc^2</formula>')
    editor.destroy()
  })

  it('after reopen, LaTeX is recovered by decompiling OMML: formula blocks and inline formulas stay editable/rewritable', async () => {
    const { editor, parsed } = await createBlankEditor()
    await executeTool(
      editor,
      {
        id: 't',
        name: 'insert_content',
        input: {
          html:
            '<formula>x = \\frac{-b \\pm \\sqrt{{b}^{2} - 4ac}}{2a}</formula>' +
            '<p>Inline <formula>{E} = m{c}^{2}</formula> formula.</p>',
        },
      },
      NUM_IDS,
    )
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const saved = await saveDocx(parsed, plan.saveBlocks)
    editor.destroy()

    const reparsed = await parseDocx(saved)
    // Display formula block: parse restores latex
    const display = reparsed.blocks.find((b) => b.formulaDisplay)
    expect(display?.formulaDisplay?.latex).toBeTruthy()
    expect(display!.formulaDisplay!.latex).toContain('\\frac')
    // Inline formula: convert restores latex (still double-click editable after reopening)
    const { editorExtensions } = await import('../src/renderer/editor/extensions')
    const editor2 = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
    })
    editor2.commands.setContent(blocksToPmDoc(reparsed.blocks) as never)
    const json = editor2.getJSON() as PmNode
    const inline = JSON.stringify(json).match(/"latex":"([^"]*\^[^"]*)"/)
    expect(inline).toBeTruthy()
    // read_blocks reads back real LaTeX (display block + inline are no longer placeholder hints)
    const read = await executeTool(
      editor2,
      {
        id: 't2',
        name: 'read_blocks',
        input: { startBlockIndex: 0, endBlockIndex: editor2.state.doc.childCount - 1 },
      },
      NUM_IDS,
    )
    expect(read.output).not.toContain('Protected')
    expect((read.output.match(/<formula>/g) ?? []).length).toBe(2)
    expect(read.output).toContain('\\frac')
    editor2.destroy()
  })

  it('formula document reopened then saved without edits = identical bytes', async () => {
    const { editor, parsed } = await createBlankEditor()
    await executeTool(
      editor,
      {
        id: 't',
        name: 'insert_content',
        input: {
          html: '<p>Derivation:</p><formula>\\sum_{k=1}^{n} k = \\frac{n(n+1)}{2}</formula>',
        },
      },
      NUM_IDS,
    )
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const saved = await saveDocx(parsed, plan.saveBlocks)
    editor.destroy()

    const reparsed = await parseDocx(saved)
    const { editorExtensions } = await import('../src/renderer/editor/extensions')
    const editor2 = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
    })
    editor2.commands.setContent(blocksToPmDoc(reparsed.blocks) as never)
    const plan2 = pmDocToSavePlan(editor2.getJSON() as PmNode, reparsed.blocks)
    expect(plan2.changedCount).toBe(0)
    expect(plan2.deletedCount).toBe(0)
    const saved2 = await saveDocx(reparsed, plan2.saveBlocks)
    expect(Buffer.from(saved2).equals(Buffer.from(saved))).toBe(true)
    editor2.destroy()
  })
})
