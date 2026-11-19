import { latexToOmml, mathParagraphXml, mathTokensOf, ommlToMathML } from '@genoffice/docx-engine'
import type { Editor } from '@tiptap/core'
import type { FormulaDisplay } from '@genoffice/docx-engine'
import { t } from '../i18n/locale'
import type { PmNode } from './convert'
import { mathLimits, ommlFollowsLimitDefaults } from './math-defaults'
import { TRACK_IGNORE } from './revisions'

/**
 * Protected display-equation block built from LaTeX (throws on syntax outside
 * the supported subset). Shared by the insert dialog/gallery and the AI
 * <formula> tag; genXml carries the OMML through the save path.
 */
/** MathML for inline flow (ommlToMathML emits display="block" for equations);
 *  `display` keeps display math-style (large operators, stacked limits) in the flow */
export function inlineMathML(omml: string, display = false): string {
  return ommlToMathML(omml, mathLimits()).replace(
    / display="block"/g,
    display ? ' display="inline" displaystyle="true"' : ' display="inline"',
  )
}

/**
 * Atomic inline formula node built from LaTeX (throws on unsupported syntax).
 * Flows with paragraph text; saves as a Run.math (<m:oMath> emitted verbatim).
 */
export function inlineEquationNodeJson(latex: string): PmNode {
  const inner = latexToOmml(latex, mathLimits())
  const omml = `<m:oMath>${inner}</m:oMath>`
  return {
    type: 'docInlineMath',
    attrs: {
      omml,
      mathml: inlineMathML(omml),
      latex: latex.trim(),
      text: mathTokensOf(omml).join(''),
    },
  }
}

export function equationBlockJson(latex: string): PmNode {
  const omml = latexToOmml(latex, mathLimits())
  return {
    type: 'docProtected',
    attrs: {
      docxIndex: null,
      blockType: 'passthrough',
      label: t('editorEquation'),
      previewText: latex.trim(),
      genXml: mathParagraphXml(omml, 'default'),
      formulaDisplay: {
        tokens: mathTokensOf(omml),
        mathml: ommlToMathML(`<m:oMath>${omml}</m:oMath>`, mathLimits()),
        omml: `<m:oMath>${omml}</m:oMath>`,
        latex: latex.trim(),
        mathPara: true,
        jcDefault: true,
      },
    },
  }
}

/** redraws the n-ary operators that leave their limits to the document, after its defaults changed */
export function refreshLimitDefaults(editor: Editor): void {
  const tr = editor.state.tr
  editor.state.doc.descendants((node, pos) => {
    if (node.type.name === 'docInlineMath') {
      const omml = String(node.attrs.omml ?? '')
      if (!ommlFollowsLimitDefaults(omml)) return
      tr.setNodeMarkup(pos, undefined, {
        ...node.attrs,
        mathml: inlineMathML(omml, !!node.attrs.display),
      })
      return
    }
    const formula = node.attrs.formulaDisplay as FormulaDisplay | null | undefined
    if (!formula?.omml || !formula.mathml || !ommlFollowsLimitDefaults(formula.omml)) return
    tr.setNodeMarkup(pos, undefined, {
      ...node.attrs,
      formulaDisplay: { ...formula, mathml: ommlToMathML(formula.omml, mathLimits()) },
    })
  })
  // a redraw of unchanged content: neither an undo step nor a tracked change
  if (tr.docChanged)
    editor.view.dispatch(tr.setMeta('addToHistory', false).setMeta(TRACK_IGNORE, true))
}
