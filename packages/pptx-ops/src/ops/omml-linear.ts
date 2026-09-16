/**
 * Plain-text reading of an OMML fragment for the equation run's fallback text
 * (what older readers, the GenOffice preview and `slides read` show):
 * fractions as a/b, scripts as x^2 / x_i, radicals as √(x), n-ary operators
 * with their limits. Structure-only: anything unknown flattens to its text.
 */
interface Node {
  name: string
  attrs: string
  children: Node[]
  text: string
}

function parse(xml: string): Node {
  const root: Node = { name: '', attrs: '', children: [], text: '' }
  const stack: Node[] = [root]
  const re = /<\/([\w:]+)\s*>|<([\w:]+)([^>]*?)(\/?)>|([^<]+)/g
  let m: RegExpExecArray | null
  while ((m = re.exec(xml)) !== null) {
    const top = stack[stack.length - 1]!
    if (m[1]) {
      if (stack.length > 1) stack.pop()
    } else if (m[2]) {
      const node: Node = { name: m[2], attrs: m[3] ?? '', children: [], text: '' }
      top.children.push(node)
      if (!m[4]) stack.push(node)
    } else if (m[5]) {
      top.text += m[5]
    }
  }
  return root
}

function unescape(s: string): string {
  return s
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&')
}

const child = (n: Node, name: string) => n.children.find((c) => c.name === name)
const attr = (n: Node | undefined, name: string) =>
  n ? new RegExp(`\\b${name}="([^"]*)"`).exec(n.attrs)?.[1] : undefined

/** Wrap multi-token operands so a/b+c and (a+b)/c stay distinguishable. */
function group(s: string): string {
  return /^[\p{L}\p{N}.]+$/u.test(s) || s === '' ? s : `(${s})`
}

function lin(n: Node): string {
  const all = () => n.children.map(lin).join('')
  const part = (name: string) => {
    const c = child(n, name)
    return c ? lin(c) : ''
  }
  switch (n.name) {
    case 'm:t':
      return unescape(n.text)
    case 'm:rPr':
    case 'm:ctrlPr':
    case 'm:fPr':
    case 'm:sSupPr':
    case 'm:sSubPr':
    case 'm:sSubSupPr':
    case 'm:radPr':
    case 'm:dPr':
    case 'm:naryPr':
    case 'm:accPr':
    case 'm:barPr':
    case 'm:mPr':
    case 'm:funcPr':
    case 'm:limLowPr':
    case 'm:limUppPr':
    case 'm:boxPr':
    case 'm:borderBoxPr':
    case 'm:groupChrPr':
    case 'm:eqArrPr':
    case 'm:oMathParaPr':
      return ''
    case 'm:f':
      return `${group(part('m:num'))}/${group(part('m:den'))}`
    case 'm:sSup':
      return `${part('m:e')}^${group(part('m:sup'))}`
    case 'm:sSub':
      return `${part('m:e')}_${group(part('m:sub'))}`
    case 'm:sSubSup':
      return `${part('m:e')}_${group(part('m:sub'))}^${group(part('m:sup'))}`
    case 'm:sPre':
      return `${group(part('m:sub'))}${group(part('m:sup'))}${part('m:e')}`
    case 'm:rad': {
      const deg = part('m:deg')
      return `${deg ? `${deg}√` : '√'}(${part('m:e')})`
    }
    case 'm:d': {
      const pr = child(n, 'm:dPr')
      const beg = pr && child(pr, 'm:begChr') ? (attr(child(pr, 'm:begChr'), 'm:val') ?? '') : '('
      const end = pr && child(pr, 'm:endChr') ? (attr(child(pr, 'm:endChr'), 'm:val') ?? '') : ')'
      return `${beg}${n.children
        .filter((c) => c.name === 'm:e')
        .map(lin)
        .join(', ')}${end}`
    }
    case 'm:nary': {
      const pr = child(n, 'm:naryPr')
      const chr = (pr && attr(child(pr, 'm:chr'), 'm:val')) ?? '∫'
      const sub = part('m:sub')
      const sup = part('m:sup')
      return `${chr}${sub ? `_${group(sub)}` : ''}${sup ? `^${group(sup)}` : ''} ${part('m:e')}`
    }
    case 'm:func':
      return `${part('m:fName')} ${part('m:e')}`
    case 'm:limLow':
      return `${part('m:e')}_${group(part('m:lim'))}`
    case 'm:limUpp':
      return `${part('m:e')}^${group(part('m:lim'))}`
    case 'm:bar':
    case 'm:acc': {
      const pr = child(n, 'm:accPr')
      return `${part('m:e')}${(pr && attr(child(pr, 'm:chr'), 'm:val')) ?? '\u0304'}`
    }
    case 'm:m':
      return `[${n.children
        .filter((c) => c.name === 'm:mr')
        .map((r) => r.children.map(lin).join(', '))
        .join('; ')}]`
    case 'm:eqArr':
      return n.children
        .filter((c) => c.name === 'm:e')
        .map(lin)
        .join('; ')
    default:
      return all()
  }
}

/** `<m:oMath>…</m:oMath>` (or bare OMML children) → one line of plain text. */
export function ommlToText(xml: string): string {
  return lin(parse(xml)).replace(/\s+/g, ' ').trim()
}
