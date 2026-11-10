import katex from 'katex'
import type { Editor } from '@tiptap/core'
import { t } from '../i18n/locale'

/** Anchor rectangle in viewport coordinates (a DOMRect or a caret box) */
interface AnchorRect {
  left: number
  top: number
  bottom: number
}

interface OpenMathOptions {
  /** position of an existing math node to edit; null → insert a new block formula */
  pos: number | null
  anchor: AnchorRect
}

let activePop: { el: HTMLElement; dispose: () => void } | null = null

function closeActive(): void {
  activePop?.dispose()
  activePop = null
}

/**
 * Floating LaTeX editor for math nodes: textarea + live KaTeX preview.
 * Apply with the button or Cmd/Ctrl+Enter, dismiss with Esc or an outside
 * click. Applying an empty input deletes the node.
 */
export function openMathEditor(editor: Editor, options: OpenMathOptions): void {
  closeActive()
  const { pos, anchor } = options
  const node = pos === null ? null : editor.state.doc.nodeAt(pos)
  const isBlock = node ? node.type.name === 'blockMath' : true
  const expectedType = node?.type.name ?? null

  /** the node at `pos`, but only while it is still the math node we opened on */
  const targetNode = () => {
    if (pos === null) return null
    const target = editor.state.doc.nodeAt(pos)
    return target && target.type.name === expectedType ? target : null
  }

  const pop = document.createElement('div')
  pop.className = 'md-math-pop'

  const input = document.createElement('textarea')
  input.className = 'md-math-input'
  input.rows = isBlock ? 4 : 2
  input.placeholder = t('mathPlaceholder')
  input.value = String(node?.attrs.latex ?? '')
  input.spellcheck = false

  const preview = document.createElement('div')
  preview.className = 'md-math-preview'

  const row = document.createElement('div')
  row.className = 'md-math-actions'
  const apply = document.createElement('button')
  apply.className = 'md-math-apply'
  apply.textContent = t('linkApply')
  row.appendChild(apply)
  if (pos !== null) {
    const del = document.createElement('button')
    del.className = 'md-math-delete'
    del.textContent = t('blockDelete')
    del.addEventListener('click', () => {
      const target = targetNode()
      if (target) {
        editor
          .chain()
          .focus()
          .deleteRange({ from: pos, to: pos + target.nodeSize })
          .run()
      }
      close()
    })
    row.appendChild(del)
  }

  const renderPreview = (): void => {
    const latex = input.value.trim()
    preview.classList.toggle('md-math-preview-empty', latex === '')
    if (latex === '') {
      preview.textContent = ''
      return
    }
    katex.render(latex, preview, { throwOnError: false, displayMode: isBlock })
  }
  input.addEventListener('input', renderPreview)

  const applyLatex = (): void => {
    const latex = input.value.trim()
    if (pos === null) {
      if (latex) editor.chain().focus().insertBlockMath({ latex }).run()
    } else {
      const target = targetNode()
      if (target) {
        if (!latex) {
          editor
            .chain()
            .focus()
            .deleteRange({ from: pos, to: pos + target.nodeSize })
            .run()
        } else if (target.type.name === 'blockMath') {
          editor.chain().focus().updateBlockMath({ latex, pos }).run()
        } else {
          editor.chain().focus().updateInlineMath({ latex, pos }).run()
        }
      }
    }
    close()
  }
  apply.addEventListener('click', applyLatex)

  pop.addEventListener('keydown', (event) => {
    if (event.key === 'Escape') {
      event.stopPropagation()
      close()
      editor.commands.focus()
    } else if (event.key === 'Enter' && (event.metaKey || event.ctrlKey)) {
      event.preventDefault()
      applyLatex()
    }
  })

  const onPointerDown = (event: PointerEvent): void => {
    if (!pop.contains(event.target as Node)) close()
  }
  document.addEventListener('pointerdown', onPointerDown, true)

  // any document change (typing in the editor, an AI edit, …) invalidates
  // `pos` — close instead of ever applying against a stale position
  const onTransaction = ({ transaction }: { transaction: { docChanged: boolean } }): void => {
    if (transaction.docChanged) close()
  }
  editor.on('transaction', onTransaction)

  const close = (): void => {
    document.removeEventListener('pointerdown', onPointerDown, true)
    editor.off('transaction', onTransaction)
    pop.remove()
    if (activePop?.el === pop) activePop = null
  }

  pop.append(input, preview, row)
  document.body.appendChild(pop)
  activePop = { el: pop, dispose: close }

  // position below the anchor, clamped to the viewport — measured only after
  // the preview has rendered, so a tall formula flips above the anchor
  renderPreview()
  const width = pop.offsetWidth
  const height = pop.offsetHeight
  const margin = 8
  const left = Math.max(margin, Math.min(anchor.left, window.innerWidth - width - margin))
  let top = anchor.bottom + 6
  if (top + height > window.innerHeight - margin) {
    top = Math.max(margin, anchor.top - height - 6)
  }
  pop.style.left = `${left}px`
  pop.style.top = `${top}px`

  input.focus()
  input.setSelectionRange(input.value.length, input.value.length)
}

/** Open the popover in insert mode, anchored at the current caret. */
export function openMathCreate(editor: Editor): void {
  const coords = editor.view.coordsAtPos(editor.state.selection.from)
  openMathEditor(editor, { pos: null, anchor: coords })
}
