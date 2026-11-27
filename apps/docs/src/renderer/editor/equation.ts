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
 * Atomic inline formula node built directly from an <m:oMath> fragment (used
 * when converting a complex MathType equation whose OMML is outside the LaTeX
 * subset).
 */
export function inlineEquationNodeFromOmml(
  omml: string,
  latex: string | null = null,
  display = false,
): PmNode {
  const normalized = omml.startsWith('<m:oMath') ? omml : `<m:oMath>${omml}</m:oMath>`
  return {
    type: 'docInlineMath',
    attrs: {
      omml: normalized,
      mathml: inlineMathML(normalized, display),
      display,
      latex: latex ? latex.trim() : null,
      text: mathTokensOf(normalized).join(''),
      oleXml: null,
      oleRawRPr: null,
      oleProgId: null,
      olePositionHalfPoints: null,
    },
  }
}

/**
 * Atomic inline formula node built from LaTeX (throws on unsupported syntax).
 * Flows with paragraph text; saves as a Run.math (<m:oMath> emitted verbatim).
 */
export function inlineEquationNodeJson(latex: string, display = false): PmNode {
  const inner = latexToOmml(latex, mathLimits())
  return inlineEquationNodeFromOmml(`<m:oMath>${inner}</m:oMath>`, latex, display)
}

/** Protected display-equation block built directly from an <m:oMath> fragment */
export function equationBlockFromOmml(omml: string, latex: string | null = null): PmNode {
  const inner = omml.replace(/^<m:oMath(?:\s[^>]*)?>/, '').replace(/<\/m:oMath>$/, '')
  const wrapped = `<m:oMath>${inner}</m:oMath>`
  const tokens = mathTokensOf(wrapped)
  return {
    type: 'docProtected',
    attrs: {
      docxIndex: null,
      blockType: 'passthrough',
      label: t('editorEquation'),
      previewText: latex ? latex.trim() : tokens.join(''),
      genXml: mathParagraphXml(inner, 'default'),
      formulaDisplay: {
        tokens,
        mathml: ommlToMathML(wrapped, mathLimits()),
        omml: wrapped,
        ...(latex ? { latex: latex.trim() } : {}),
        mathPara: true,
        jcDefault: true,
      },
    },
  }
}

export function equationBlockJson(latex: string): PmNode {
  const omml = latexToOmml(latex, mathLimits())
  return equationBlockFromOmml(`<m:oMath>${omml}</m:oMath>`, latex)
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
