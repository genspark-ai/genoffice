import type { ParseMap } from '../document/parse-map'

export const SID_ATTR = 'data-sid'

/**
 * The preview copy of the document: every element with a source location gets
 * a `data-sid` so the inspector can map clicks straight back to the parse map,
 * and the inspector script is appended. The saved text never sees either.
 */
export function instrumentForPreview(text: string, map: ParseMap, inspectorSource: string): string {
  const inserts = map.elements
    .map((e) => {
      const tag = text.slice(e.startTag[0], e.startTag[1])
      const close = /\s*\/?>$/.exec(tag)
      const at = e.startTag[0] + (close ? close.index : tag.length)
      return { at, text: ` ${SID_ATTR}="${e.sid}"` }
    })
    .sort((a, b) => b.at - a.at)
  let out = text
  for (const ins of inserts) out = out.slice(0, ins.at) + ins.text + out.slice(ins.at)
  // the frame reports this version with every message so the app can drop input from a stale reload
  const script = `<script data-gx-inspector>${inspectorSource.replace('__GX_VERSION__', String(map.version))}</script>`
  const bodyClose = out.search(/<\/body\s*>/i)
  return bodyClose >= 0 ? out.slice(0, bodyClose) + script + out.slice(bodyClose) : out + script
}
