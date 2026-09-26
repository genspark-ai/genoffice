import type { AnyExtension } from '@tiptap/core'
import { Extension } from '@tiptap/core'
import { Plugin } from '@tiptap/pm/state'
import { DOMParser as ProseMirrorDOMParser, DOMSerializer } from '@tiptap/pm/model'
import { BlockMath, InlineMath } from '@tiptap/extension-mathematics'
import { openMathEditor } from './mathEdit'
import { matchInlineMath } from './mathSyntax'

const StrictInlineMath = InlineMath.extend({
  markdownTokenizer: {
    name: 'inlineMath',
    level: 'inline',
    start: (src: string) => src.indexOf('$'),
    tokenize: (src: string) => {
      const match = matchInlineMath(src)
      return match && { type: 'inlineMath', ...match }
    },
  },
})

/**
 * AI chatbots (Gemini, Claude) emit LaTeX with \(...\) inline and \[...\]
 * block delimiters, but the editor only tokenizes $...$ / $$...$$. Normalize
 * pasted text so pasted formulas render instead of staying plain text.
 * Currency ($5, $10) is untouched: only backslash delimiters convert.
 *
 * Guard against Markdown escapes: `\[1\]` is also an escaped bracket
 * (citations like `see \[1\]`), and `\(note\)` an escaped paren. Only convert
 * when the inner text looks like LaTeX (a `\command`, `^`, `_`, `=`, or an
 * operator next to a digit); `\[...\]` must additionally sit alone on its
 * line(s), which is how display math is actually pasted.
 */
function isLatexLike(inner: string): boolean {
  if (inner.includes('^') || inner.includes('_') || inner.includes('=')) return true
  if (/\\[a-zA-Z]+/.test(inner)) return true
  if (/\d\s*[+\-*/<>|]|[+\-*/<>|]\s*\d/.test(inner)) return true
  return false
}

/** `plain` is ProseMirror's flag for code-block / Shift+paste targets: LaTeX source must stay verbatim there. */
export function normalizePastedMath(text: string, plain = false): string {
  if (plain) return text
  // Leave fenced and inline code untouched even when formulas surround it.
  return text
    .split(
      /(^[ \t]*`{3,}[^\n]*\n[\s\S]*?^[ \t]*`{3,}[ \t]*$|^[ \t]*~{3,}[^\n]*\n[\s\S]*?^[ \t]*~{3,}[ \t]*$|`+[^`]*`+)/m,
    )
    .map((part, index) => (index % 2 ? part : normalizeMathDelimiters(part)))
    .join('')
}

function normalizeMathDelimiters(text: string): string {
  const withBlocks = text.replace(
    /\\\[(.+?)\\\]/gs,
    (match: string, latex: string, offset: number, full: string) => {
      const inner = latex.trim()
      if (inner === '' || !isLatexLike(inner)) return match
      const before = full.slice(0, offset)
      const after = full.slice(offset + match.length)
      const lineStart = before.lastIndexOf('\n') + 1
      const beforeOnLine = before.slice(lineStart, offset)
      const nextNl = after.indexOf('\n')
      const afterOnLine = nextNl === -1 ? after : after.slice(0, nextNl)
      if (beforeOnLine.trim() !== '' || afterOnLine.trim() !== '') return match
      return '$$' + inner + '$$'
    },
  )
  // inline delimiters convert only outside display blocks: `$…$` nested in
  // `$$…$$` would split one formula into three tokens
  return withBlocks
    .split(/(\$\$[\s\S]*?\$\$)/)
    .map((segment, index) =>
      index % 2 === 1
        ? segment
        : segment.replace(/\\\((.+?)\\\)/gs, (match: string, latex: string) => {
            const inner = latex.trim()
            if (inner === '' || inner.includes('\n')) return match
            if (!isLatexLike(inner)) return match
            return `$${inner}$`
          }),
    )
    .join('')
}

/** Convert only formula carriers; preserve the surrounding rich clipboard HTML. */
function normalizeMathHtml(html: string): string {
  const root = document.createElement('div')
  root.innerHTML = html
  const mathElement = (latex: string, block: boolean) => {
    const element = document.createElement(block ? 'div' : 'span')
    element.setAttribute('data-type', block ? 'block-math' : 'inline-math')
    element.setAttribute('data-latex', latex)
    return element
  }
  // KaTeX includes both accessible MathML and a visual HTML copy. Replace the
  // entire carrier, otherwise its formula appears twice after pasting.
  for (const annotation of root.querySelectorAll('annotation[encoding="application/x-tex"]')) {
    if (annotation.closest('pre, code, [data-type]')) continue
    const carrier =
      annotation.closest('.katex-display') ??
      annotation.closest('.katex') ??
      annotation.closest('.mwe-math-element') ??
      annotation.closest('math')
    const latex = annotation.textContent?.trim()
    if (!carrier || !latex) continue
    carrier.replaceWith(
      mathElement(
        latex,
        carrier.matches('.katex-display, math[display="block"]') ||
          !!carrier.querySelector('math[display="block"], .mwe-math-fallback-image-display'),
      ),
    )
  }
  // Display formulas copied from web pages often use <br> and inline spans.
  // Join only a standalone text block, leaving neighboring rich text intact.
  for (const block of root.querySelectorAll('p, div')) {
    if (
      block.closest('pre, code, [data-type]') ||
      block.querySelector('pre, code, [data-type], p, div, ul, ol, table, blockquote')
    )
      continue
    const source = block.cloneNode(true) as HTMLElement
    for (const br of source.querySelectorAll('br')) br.replaceWith('\n')
    const normalized = normalizePastedMath(source.textContent ?? '').trim()
    const match = /^\$\$([^$]+)\$\$$/.exec(normalized)
    if (match) block.replaceWith(mathElement(match[1].trim(), true))
  }
  const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT)
  const texts: Text[] = []
  while (walker.nextNode()) texts.push(walker.currentNode as Text)
  for (const text of texts) {
    if (text.parentElement?.closest('pre, code, [data-type], script, style')) continue
    const normalized = normalizePastedMath(text.data)
    const fragment = document.createDocumentFragment()
    // Backticks protect code in plain-looking HTML as well.
    const pattern =
      /`+[^`]*`+|(?<![\\$])\$\$([^$]+)\$\$|(?<![\\$])\$(?![\s$])([^$\n]*[^\s$])\$(?![\d$])/g
    let start = 0
    let changed = false
    for (const match of normalized.matchAll(pattern)) {
      if (match[0].startsWith('`')) continue
      fragment.append(normalized.slice(start, match.index))
      fragment.append(mathElement((match[1] ?? match[2]).trim(), match[1] !== undefined))
      start = match.index! + match[0].length
      changed = true
    }
    if (changed) {
      fragment.append(normalized.slice(start))
      text.replaceWith(fragment)
    }
  }
  return root.innerHTML
}

/** Math nodes are atoms — clicking one opens the LaTeX edit popover. */
const MathClickEdit = Extension.create({
  name: 'mathClickEdit',

  addProseMirrorPlugins() {
    const editor = this.editor
    return [
      new Plugin({
        props: {
          clipboardTextParser: (text, context, plain) => {
            // Keep ProseMirror's plain-text paragraphs and active marks. Only
            // formula tokens become nodes; surrounding text is never Markdown-parsed.
            const root = document.createElement('div')
            const serializer = DOMSerializer.fromSchema(editor.schema)
            let paragraph: HTMLElement | undefined
            const appendText = (value: string) => {
              if (!paragraph && root.lastElementChild?.getAttribute('data-type') === 'block-math')
                value = value.replace(/^(?:\r\n?|\n)+/, '')
              if (!value) return
              const lines = value.split(/(?:\r\n?|\n)+/)
              lines.forEach((line, index) => {
                if (index) paragraph = undefined
                if (!paragraph) paragraph = root.appendChild(document.createElement('p'))
                if (line)
                  paragraph.appendChild(
                    serializer.serializeNode(editor.schema.text(line, context.marks())),
                  )
              })
            }
            const parts = plain
              ? [text]
              : text.split(
                  /(^[ \t]*`{3,}[^\n]*\n[\s\S]*?^[ \t]*`{3,}[ \t]*$|^[ \t]*~{3,}[^\n]*\n[\s\S]*?^[ \t]*~{3,}[ \t]*$|`+[^`]*`+)/m,
                )
            parts.forEach((part, index) => {
              if (plain || index % 2) {
                appendText(part)
                return
              }
              const normalized = normalizeMathDelimiters(part)
              const pattern =
                /(?<![\\$])\$\$([^$]+)\$\$|(?<![\\$])\$(?![\s$])([^$\n]*[^\s$])\$(?![\d$])/g
              let start = 0
              for (const match of normalized.matchAll(pattern)) {
                appendText(normalized.slice(start, match.index))
                const block = match[1] !== undefined
                const node = editor.schema.nodes[block ? 'blockMath' : 'inlineMath'].create({
                  latex: (match[1] ?? match[2]).trim(),
                })
                if (block) {
                  if (paragraph && !paragraph.hasChildNodes()) paragraph.remove()
                  root.appendChild(serializer.serializeNode(node))
                  paragraph = undefined
                } else {
                  if (!paragraph) paragraph = root.appendChild(document.createElement('p'))
                  paragraph.appendChild(serializer.serializeNode(node))
                }
                start = match.index! + match[0].length
              }
              appendText(normalized.slice(start))
            })
            return ProseMirrorDOMParser.fromSchema(editor.schema).parseSlice(root, {
              preserveWhitespace: true,
              context,
            })
          },
          transformPastedHTML: normalizeMathHtml,
          handleClickOn: (view, _pos, node, nodePos, event) => {
            if (node.type.name !== 'blockMath' && node.type.name !== 'inlineMath') return false
            if (!view.editable) return false
            const target = event.target as HTMLElement | null
            const anchor = target?.closest?.('.tiptap-mathematics-render') ?? target
            if (!anchor) return false
            openMathEditor(editor, { pos: nodePos, anchor: anchor.getBoundingClientRect() })
            return true
          },
        },
      }),
    ]
  },
})

export function buildMathExtensions(): AnyExtension[] {
  return [BlockMath, StrictInlineMath, MathClickEdit]
}
