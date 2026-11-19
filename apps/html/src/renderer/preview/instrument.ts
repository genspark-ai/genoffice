import type { ParseMap } from '../document/parse-map'

export const SID_ATTR = 'data-gx-sid'

interface StartAttribute {
  name: string
  start: number
  end: number
}

function startAttributes(tag: string): {
  prefixEnd: number
  closeStart: number
  attributes: StartAttribute[]
} {
  const prefix = /^<[A-Za-z][^\s/>]*/.exec(tag)
  const close = /\s*\/?>$/.exec(tag)
  const prefixEnd = prefix?.[0].length ?? 1
  const closeStart = close?.index ?? tag.length
  const attributes: StartAttribute[] = []
  let index = prefixEnd
  const whitespace = (at: number) => at < closeStart && /\s/.test(tag[at]!)
  while (index < closeStart) {
    while (whitespace(index)) index += 1
    if (index >= closeStart) break
    const start = index
    while (index < closeStart && !whitespace(index) && tag[index] !== '=' && tag[index] !== '>') {
      index += 1
    }
    if (index === start) {
      index += 1
      continue
    }
    const name = tag.slice(start, index)
    while (whitespace(index)) index += 1
    if (tag[index] === '=') {
      index += 1
      while (whitespace(index)) index += 1
      const quote = tag[index] === '"' || tag[index] === "'" ? tag[index] : ''
      if (quote) {
        const end = tag.indexOf(quote, index + 1)
        index = end >= 0 && end < closeStart ? end + 1 : closeStart
      } else {
        while (index < closeStart && !whitespace(index)) index += 1
      }
    }
    attributes.push({ name, start, end: index })
  }
  return { prefixEnd, closeStart, attributes }
}

function replaceAuthoredSid(tag: string, sid: number): string {
  const { prefixEnd, closeStart, attributes } = startAttributes(tag)
  const authored = attributes.filter((attribute) => attribute.name.toLowerCase() === SID_ATTR)
  let withoutAuthored = tag
  if (authored.length > 0) {
    const kept = attributes
      .filter((attribute) => !authored.includes(attribute))
      .map((attribute) => ` ${tag.slice(attribute.start, attribute.end)}`)
    withoutAuthored = tag.slice(0, prefixEnd) + kept.join('') + tag.slice(closeStart)
  }
  const close = /\s*\/?>$/.exec(withoutAuthored)
  const at = close?.index ?? withoutAuthored.length
  return `${withoutAuthored.slice(0, at)} ${SID_ATTR}="${sid}"${withoutAuthored.slice(at)}`
}

/**
 * The preview copy of the document: every element with a source location gets
 * a `data-gx-sid` so the inspector can map clicks straight back to the parse map,
 * and the inspector script is appended. The saved text never sees either.
 */
export function instrumentForPreview(text: string, map: ParseMap, inspectorSource: string): string {
  const edits = map.elements
    .map((entry) => ({
      start: entry.startTag[0],
      end: entry.startTag[1],
      text: replaceAuthoredSid(text.slice(entry.startTag[0], entry.startTag[1]), entry.sid),
    }))
    .sort((a, b) => b.start - a.start)
  let out = text
  for (const edit of edits) out = out.slice(0, edit.start) + edit.text + out.slice(edit.end)
  // the frame reports this version with every message so the app can drop input from a stale reload
  const script = `<script data-gx-inspector>${inspectorSource.replace('__GX_VERSION__', String(map.version))}</script>`
  const bodyClose = out.search(/<\/body\s*>/i)
  return bodyClose >= 0 ? out.slice(0, bodyClose) + script + out.slice(bodyClose) : out + script
}
