import { isMathTypeProgId } from './mtef'
import { parseSeqInstruction } from './seq-fields'
import type { Run } from './types'
import {
  attrsOf,
  childrenOf,
  findChild,
  nameOf,
  serializeXNode,
  textOf,
  type XNode,
} from './xml-utils'

interface NumberField {
  code: string
  result?: string
}

function displayedNumber({ code, result }: NumberField): string | undefined {
  const seq = parseSeqInstruction(code)
  if (seq) return seq.hidden ? '' : (result ?? '')
  const macro = /^\s*MACROBUTTON\s+(MTPlaceRef|MTEditEquationSection\d*)\b\s*/i.exec(code)
  if (!macro) return undefined
  if (result !== undefined) return result
  if (macro[1].toLowerCase() !== 'mtplaceref') return ''
  // MTPlaceRef commonly has a prompt containing nested SEQ fields and no cached result.
  return code
    .slice(macro[0].length)
    .replace(/\\\*\s+MERGEFORMAT\b/gi, '')
    .trim()
}

function isEquation(node: XNode): boolean {
  if (nameOf(node) === 'm:oMath' || nameOf(node) === 'm:oMathPara') return true
  const object = nameOf(node) === 'w:r' && findChild(node, 'w:object')
  const ole = object && findChild(object, 'o:OLEObject')
  return !!ole && isMathTypeProgId(attrsOf(ole).ProgID)
}

function runContent(run: XNode): XNode[] {
  return childrenOf(run).filter((child) => nameOf(child) && nameOf(child) !== 'w:rPr')
}

function isText(child: XNode): boolean {
  return nameOf(child) === 'w:t' || nameOf(child) === 'w:instrText'
}

/** Replace each supported field span with one inline atom, keeping the field and
 * its surrounding bookmarks together. If any field cannot be represented intact,
 * the caller uses the ordinary field paragraph path for the whole paragraph. */
export function equationNumberFields(paragraph: XNode): ReadonlyMap<XNode, Run | null> | undefined {
  const nodes = childrenOf(paragraph).filter((child) => nameOf(child))
  if (!nodes.some(isEquation)) return
  const fields = new Map<XNode, Run | null>()
  const stack: NumberField[] = []
  const bookmarks: XNode[] = []
  let start = 0

  function append(field: NumberField, value: string): void {
    if (field.result === undefined) field.code += value
    else field.result += value
  }

  function finish(end: number, text: string): void {
    const parent = stack.at(-1)
    if (parent) {
      append(parent, text)
      return
    }
    let first = start
    let last = end
    while (first > 0 && last + 1 < nodes.length) {
      const open = nodes[first - 1]
      const close = nodes[last + 1]
      const id = attrsOf(open)['w:id']
      if (
        nameOf(open) !== 'w:bookmarkStart' ||
        nameOf(close) !== 'w:bookmarkEnd' ||
        !id ||
        attrsOf(close)['w:id'] !== id
      )
        break
      first--
      last++
    }
    const span = nodes.slice(first, last + 1)
    for (const child of span) fields.set(child, null)
    fields.set(span[0], { text, mathTypeFieldXml: span.map(serializeXNode).join('') })
  }

  for (const [i, node] of nodes.entries()) {
    const name = nameOf(node)
    const current = stack.at(-1)
    if (name === 'w:fldSimple') {
      const runs = childrenOf(node).filter((child) => nameOf(child))
      if (!runs.every((run) => nameOf(run) === 'w:r' && runContent(run).every(isText))) return
      const text = displayedNumber({
        code: attrsOf(node)['w:instr'] ?? '',
        result: runs.length ? runs.flatMap(runContent).map(textOf).join('') : undefined,
      })
      if (text === undefined) return
      if (!current) start = i
      finish(i, text)
      continue
    }
    if (name !== 'w:r') {
      if (current) return
      if (name === 'w:bookmarkStart' || name === 'w:bookmarkEnd') bookmarks.push(node)
      else if (name !== 'w:pPr' && !isEquation(node)) return
      continue
    }
    const contents = runContent(node)
    const marker = contents.find((child) => nameOf(child) === 'w:fldChar')
    if (!marker) {
      if (current) {
        if (!contents.every(isText)) return
        append(current, contents.map(textOf).join(''))
      } else if (
        !(contents.length === 1 && isEquation(node)) &&
        contents.some((child) => !['w:t', 'w:tab', 'w:br', 'w:cr'].includes(nameOf(child)!))
      )
        return
      continue
    }
    if (contents.length !== 1) return
    switch (attrsOf(marker)['w:fldCharType']) {
      case 'begin':
        if (!current) start = i
        stack.push({ code: '' })
        break
      case 'separate':
        if (!current || current.result !== undefined) return
        current.result = ''
        break
      case 'end': {
        if (!current) return
        const text = displayedNumber(current)
        if (text === undefined) return
        stack.pop()
        finish(i, text)
        break
      }
      default:
        return
    }
  }
  if (!stack.length && fields.size && bookmarks.every((node) => fields.has(node))) return fields
}
