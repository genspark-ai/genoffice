/// OpenXML SDK-family producers bind the spreadsheetml namespace to a prefix
/// (<x:workbook><x:sheets>...), invisible to the text-based patch pipeline's
/// unprefixed element patterns. Parts are normalized at read time instead:
/// the prefix is stripped and re-declared as the default namespace. Only
/// parts the planner rewrites reach the file, so untouched entries keep
/// their bytes; anything ambiguous returns the input unchanged (the save
/// then fails exactly as it did before this pass existed).

const NORMALIZABLE_NAMESPACES = new Set([
  'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
  'http://schemas.openxmlformats.org/package/2006/relationships',
  'http://schemas.openxmlformats.org/package/2006/content-types',
])

const SPREADSHEETML_NAMESPACE = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
const OFFICE_RELATIONSHIPS_NAMESPACE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

/// Root start tag; the quote alternation keeps '>' inside attribute values
/// (legal in XML) from ending the match early.
const ROOT_START_TAG = /<(?![?!])(?:[^>"']|"[^"]*"|'[^']*')*?\/?>/

/// Elements the pipeline appends carry r:-prefixed attributes; some
/// producers bind r on individual elements only, so the host root gets the
/// conventional binding before such an element is inserted.
export function ensureRelationshipNamespace(partXml: string): string {
  const root = ROOT_START_TAG.exec(partXml)?.[0]
  if (!root || /\bxmlns:r\s*=/.test(root)) return partXml
  const bound = root.replace(/^<([^\s/>]+)/, `<$1 xmlns:r="${OFFICE_RELATIONSHIPS_NAMESPACE}"`)
  return partXml.replace(root, bound)
}

export function normalizeOoxmlPartPrefix(xml: string): string {
  const root = /<(?![?!])([^\s/>]+)((?:[^>"']|"[^"]*"|'[^']*')*?)\/?>/.exec(xml)
  if (!root?.[1]) return xml
  const colon = root[1].indexOf(':')
  if (colon <= 0) return xml
  const prefix = root[1].slice(0, colon)
  if (!/^[A-Za-z_][A-Za-z0-9_]*$/.test(prefix)) return xml

  const rootTag = root[0]
  const declarationPattern = new RegExp(`xmlns:${prefix}\\s*=\\s*"([^"]*)"`)
  const declaration = declarationPattern.exec(rootTag)
  const uri = declaration?.[1]
  if (declaration === null || uri === undefined || !NORMALIZABLE_NAMESPACES.has(uri)) return xml

  // Sections where plain text substitution could corrupt content.
  if (xml.includes('<![CDATA[') || xml.includes('<!--')) return xml
  // An existing default binding would capture the de-prefixed elements.
  if (/\sxmlns\s*=\s*["']/.test(xml)) return xml
  // The prefix rebound to another namespace deeper in the tree.
  for (const other of xml.matchAll(new RegExp(`\\sxmlns:${prefix}\\s*=\\s*"([^"]*)"`, 'g'))) {
    if (other[1] !== uri) return xml
  }

  const stripped = xml.replaceAll(`<${prefix}:`, '<').replaceAll(`</${prefix}:`, '</')
  // Attributes may still use the prefix; keep its binding alongside then.
  const bindings = new RegExp(`\\s${prefix}:`).test(stripped)
    ? `xmlns="${uri}" ${declaration[0]}`
    : `xmlns="${uri}"`
  let result = stripped.replace(declaration[0], bindings)

  // The pipeline appends elements carrying r:-prefixed attributes (new
  // <sheet r:id=...>); these producers bind r per element, not on the root,
  // so a root-level binding keeps such additions well-formed.
  if (
    uri === SPREADSHEETML_NAMESPACE &&
    !new RegExp(`xmlns:r\\s*=\\s*"`).test(ROOT_START_TAG.exec(result)?.[0] ?? '') &&
    [...result.matchAll(/\sxmlns:r\s*=\s*"([^"]*)"/g)].every(
      (other) => other[1] === OFFICE_RELATIONSHIPS_NAMESPACE,
    )
  ) {
    result = result.replace(
      `xmlns="${uri}"`,
      `xmlns="${uri}" xmlns:r="${OFFICE_RELATIONSHIPS_NAMESPACE}"`,
    )
  }
  return result
}
