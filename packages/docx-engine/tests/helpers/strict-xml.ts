import JSZip from 'jszip'
import { SaxesParser } from 'saxes'

const MC_NS = 'http://schemas.openxmlformats.org/markup-compatibility/2006'
/** attributes of the compatibility namespace whose value is a list of prefixes */
const PREFIX_LISTS = new Set(['Ignorable', 'MustUnderstand', 'ProcessContent'])

/**
 * What a strict, namespace-aware reader makes of one part: null when it parses,
 * otherwise the first error. A prefix named by `mc:Ignorable` / `Requires` must
 * be declared as well; Word refuses a part where it is not.
 */
export function xmlProblem(xml: string): string | null {
  const parser = new SaxesParser({ xmlns: true })
  let problem: string | null = null
  const scopes: Array<Record<string, string>> = []
  const declared = (prefix: string) => scopes.some((ns) => prefix in ns)
  parser.on('error', (e) => {
    problem ??= e.message
  })
  parser.on('opentag', (tag) => {
    scopes.push(tag.ns ?? {})
    for (const attr of Object.values(tag.attributes)) {
      // Requires sits unprefixed on mc:Choice; the others are mc: attributes on any element
      const list =
        attr.local === 'Requires'
          ? attr.prefix === '' && tag.uri === MC_NS
          : attr.uri === MC_NS && PREFIX_LISTS.has(attr.local)
      if (!list) continue
      for (const token of attr.value.split(/\s+/).filter(Boolean)) {
        const prefix = token.split(':')[0]!
        if (!declared(prefix)) problem ??= `${attr.name} names the undeclared prefix ${prefix}`
      }
    }
    if (tag.isSelfClosing) scopes.pop()
  })
  parser.on('closetag', (tag) => {
    if (!tag.isSelfClosing) scopes.pop()
  })
  try {
    parser.write(xml.charCodeAt(0) === 0xfeff ? xml.slice(1) : xml).close()
  } catch (e) {
    problem ??= (e as Error).message
  }
  return problem
}

/**
 * Every XML part of a saved package that a strict reader rejects, as
 * `path: error`. A part the source package already carried broken is not the
 * save's doing and is left out.
 */
export async function packageXmlProblems(
  bytes: Uint8Array | ArrayBuffer,
  source?: Uint8Array | ArrayBuffer,
): Promise<string[]> {
  // copied: bytes that crossed a structured clone are another realm's Uint8Array
  const load = (data: Uint8Array | ArrayBuffer) =>
    JSZip.loadAsync(new Uint8Array(data as ArrayBuffer))
  const zip = await load(bytes)
  const before = source ? await load(source) : null
  const problems: string[] = []
  for (const [name, entry] of Object.entries(zip.files)) {
    if (entry.dir || !/\.(xml|rels|vml)$/i.test(name)) continue
    const problem = xmlProblem(await entry.async('string'))
    if (!problem) continue
    const original = await before?.file(name)?.async('string')
    if (original !== undefined && xmlProblem(original)) continue
    problems.push(`${name}: ${problem.replace(/\s+/g, ' ')}`)
  }
  return problems
}
