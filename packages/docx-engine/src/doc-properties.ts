import JSZip from 'jszip'
import {
  CUSTOM_PROPERTIES_CONTENT_TYPE,
  CUSTOM_PROPERTIES_PATH,
  CUSTOM_PROPERTIES_REL_TYPE,
  PROPERTY_RE,
  attrValue,
  decodeXmlText,
  propertyText,
} from './zotero-doc-props'
import { escapeXmlAttr, escapeXmlText } from './xml-utils'

/** docProps/core.xml — the core properties part. */
export const CORE_PROPERTIES_PATH = 'docProps/core.xml'
/** docProps/app.xml — the extended properties part. */
export const APP_PROPERTIES_PATH = 'docProps/app.xml'

export const CORE_PROPERTIES_CONTENT_TYPE =
  'application/vnd.openxmlformats-package.core-properties+xml'
export const CORE_PROPERTIES_REL_TYPE =
  'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties'
export const APP_PROPERTIES_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.extended-properties+xml'
export const APP_PROPERTIES_REL_TYPE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties'

/** the fmtid every Word custom property carries */
const CUSTOM_PROPERTIES_FMTID = '{D5CDD505-2E9C-101B-9397-08002B2CF9AE}'

/** friendly key -> the element of docProps/core.xml that holds it */
const CORE_PROPERTIES: Record<string, { prefix: string; local: string }> = {
  title: { prefix: 'dc', local: 'title' },
  subject: { prefix: 'dc', local: 'subject' },
  creator: { prefix: 'dc', local: 'creator' },
  keywords: { prefix: 'cp', local: 'keywords' },
  description: { prefix: 'dc', local: 'description' },
  lastModifiedBy: { prefix: 'cp', local: 'lastModifiedBy' },
  revision: { prefix: 'cp', local: 'revision' },
  created: { prefix: 'dcterms', local: 'created' },
  modified: { prefix: 'dcterms', local: 'modified' },
  category: { prefix: 'cp', local: 'category' },
  contentStatus: { prefix: 'cp', local: 'contentStatus' },
  identifier: { prefix: 'dc', local: 'identifier' },
  language: { prefix: 'dc', local: 'language' },
  version: { prefix: 'cp', local: 'version' },
}

/** friendly key -> the element of docProps/app.xml that holds it */
const APP_PROPERTIES: Record<string, string> = {
  template: 'Template',
  totalTime: 'TotalTime',
  pages: 'Pages',
  words: 'Words',
  characters: 'Characters',
  application: 'Application',
  docSecurity: 'DocSecurity',
  lines: 'Lines',
  paragraphs: 'Paragraphs',
  scaleCrop: 'ScaleCrop',
  company: 'Company',
  linksUpToDate: 'LinksUpToDate',
  charactersWithSpaces: 'CharactersWithSpaces',
  sharedDoc: 'SharedDoc',
  hyperlinksChanged: 'HyperlinksChanged',
  appVersion: 'AppVersion',
  manager: 'Manager',
}

/**
 * The schema order of CT_Properties: app.xml is a sequence, so an element that
 * is not there yet has to be written where the schema puts it.
 */
const APP_PROPERTY_ORDER = [
  'Template',
  'Manager',
  'Company',
  'Pages',
  'Words',
  'Characters',
  'PresentationFormat',
  'Lines',
  'Paragraphs',
  'Slides',
  'Notes',
  'TotalTime',
  'HiddenSlides',
  'MMClips',
  'ScaleCrop',
  'HeadingPairs',
  'TitlesOfParts',
  'LinksUpToDate',
  'CharactersWithSpaces',
  'SharedDoc',
  'HyperlinkBase',
  'HLinks',
  'HyperlinksChanged',
  'DigSig',
  'Application',
  'AppVersion',
  'DocSecurity',
]

/** the core property keys, in the order the part lists them */
export const CORE_PROPERTY_KEYS = Object.keys(CORE_PROPERTIES)
/** the extended (app) property keys */
export const APP_PROPERTY_KEYS = Object.keys(APP_PROPERTIES)

/** a custom-property value of a set_properties op: text, or a typed number/boolean */
export type PropertyValue = string | number | boolean

/** the document properties of a .docx, read from its three docProps parts */
export interface DocumentProperties {
  /** docProps/core.xml by friendly key (title, creator, lastModifiedBy, …) */
  core: Record<string, string>
  /** docProps/app.xml by friendly key (company, template, application, …) */
  app: Record<string, string>
  /** docProps/custom.xml by property name (a typed value is its vt text: "42", "true", …) */
  custom: Record<string, string>
  /**
   * docProps/custom.xml by property name: the vt type the value is stored as
   * (lpwstr, i4, i8, r8, bool, filetime, date, …). A set_properties op that
   * updates such a property keeps its type while the new value fits it.
   */
  customTypes: Record<string, string>
}

/** what a set_properties op asks for: a value per key, or null to remove it */
export interface PropertyChanges {
  core?: Record<string, string | null>
  app?: Record<string, string | null>
  /** custom values: text stays lpwstr, a number/boolean takes a typed vt value */
  custom?: Record<string, PropertyValue | null>
  /** remove every property of all three parts (a full metadata scrub) */
  clear?: 'all'
}

/**
 * `<dc:title>`, `<dc:title/>` and `<title>`: the element by local name, the
 * prefix optional because a part may bind it to another prefix.
 */
function elementRe(local: string): RegExp {
  return new RegExp(
    `<([A-Za-z_][\\w.-]*:)?${local}(?=[\\s/>])([^>]*?)(?:/>|>([\\s\\S]*?)</([A-Za-z_][\\w.-]*:)?${local}\\s*>)`,
    'g',
  )
}

/** the text of an element, null when the part does not carry it */
function elementText(xml: string, local: string): string | null {
  const match = elementRe(local).exec(xml)
  return match ? decodeXmlText(match[3] ?? '') : null
}

/** insert an element before the root's closing tag */
function insertElement(xml: string, element: string): string {
  const at = xml.lastIndexOf('</')
  return at === -1 ? xml + element : xml.slice(0, at) + element + xml.slice(at)
}

/** replace the text of the element, keeping its attributes; insert it when missing */
function setElementText(xml: string, local: string, value: string, element: string): string {
  let found = false
  const out = xml.replace(elementRe(local), (_match, openPrefix, attrs) => {
    found = true
    const name = `${openPrefix ?? ''}${local}`
    return `<${name}${attrs}>${escapeXmlText(value)}</${name}>`
  })
  return found ? out : insertElement(xml, element)
}

/** the core properties of a core.xml part */
export function parseCorePropertiesXml(xml: string): Record<string, string> {
  const out: Record<string, string> = {}
  for (const [key, { local }] of Object.entries(CORE_PROPERTIES)) {
    const text = elementText(xml, local)
    if (text !== null) out[key] = text
  }
  return out
}

/** the extended properties of an app.xml part */
export function parseAppPropertiesXml(xml: string): Record<string, string> {
  const out: Record<string, string> = {}
  for (const [key, local] of Object.entries(APP_PROPERTIES)) {
    const text = elementText(xml, local)
    if (text !== null) out[key] = text
  }
  return out
}

const VT_VALUE_RE =
  /<(?:[A-Za-z_][\w.-]*:)?(lpwstr|lpstr|bstr|i1|i2|i4|i8|ui1|ui2|ui4|ui8|int|uint|r4|r8|bool|filetime|date|decimal)\b[^>]*>([\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?\1\s*>/

/** the vt element a custom property carries; unknown value elements stay untyped */
const VT_TYPE_RE =
  /<(?:[A-Za-z_][\w.-]*:)?(lpwstr|lpstr|bstr|i1|i2|i4|i8|ui1|ui2|ui4|ui8|int|uint|r4|r8|bool|filetime|date|decimal|cy|error|null|clsid|blob|stream|storage|variant|vector)\b/

function customPropertyValue(property: string): string {
  const match = VT_VALUE_RE.exec(property)
  return match ? decodeXmlText(match[2]) : propertyText(property)
}

/** the custom properties of a custom.xml part, by property name */
export function parseCustomPropertiesXml(xml: string): Record<string, string> {
  const out: Record<string, string> = {}
  for (const property of xml.match(PROPERTY_RE) ?? []) {
    const name = attrValue(property, 'name')
    if (name) out[name] = customPropertyValue(property)
  }
  return out
}

/** the vt types of a custom.xml part's properties, by property name */
export function parseCustomPropertyTypesXml(xml: string): Record<string, string> {
  const out: Record<string, string> = {}
  for (const property of xml.match(PROPERTY_RE) ?? []) {
    const name = attrValue(property, 'name')
    const type = VT_TYPE_RE.exec(property)?.[1]
    if (name && type) out[name] = type
  }
  return out
}

/** the three docProps parts of a package, empty when the part is missing */
export async function readDocumentProperties(zip: JSZip): Promise<DocumentProperties> {
  const core = zip.file(CORE_PROPERTIES_PATH)
  const app = zip.file(APP_PROPERTIES_PATH)
  const custom = zip.file(CUSTOM_PROPERTIES_PATH)
  const customXml = custom ? await custom.async('string') : ''
  return {
    core: core ? parseCorePropertiesXml(await core.async('string')) : {},
    app: app ? parseAppPropertiesXml(await app.async('string')) : {},
    custom: custom ? parseCustomPropertiesXml(customXml) : {},
    customTypes: custom ? parseCustomPropertyTypesXml(customXml) : {},
  }
}

/** the properties of a package that carries no docProps part at all */
export function emptyDocumentProperties(): DocumentProperties {
  return { core: {}, app: {}, custom: {}, customTypes: {} }
}

function blankCorePropertiesXml(): string {
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"' +
    ' xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/"' +
    ' xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
    '</cp:coreProperties>'
  )
}

function blankAppPropertiesXml(): string {
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"' +
    ' xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">' +
    '</Properties>'
  )
}

/** Application a new GenOffice document records in docProps/app.xml */
export const NEW_DOCUMENT_APPLICATION = 'GenOffice'
/** Template a new GenOffice document records (Word's own default) */
export const NEW_DOCUMENT_TEMPLATE = 'Normal.dotm'
/**
 * AppVersion of a new document. The docx-engine package version at release
 * time; a blank-template test pins it to package.json so a version bump fails
 * loudly instead of shipping a stale stamp.
 */
export const NEW_DOCUMENT_APP_VERSION = '0.1.0'

/** the W3CDTF stamp Word writes into core properties (seconds, Z) */
export function w3cdtfStamp(date: Date): string {
  return date.toISOString().replace(/\.\d{3}Z$/, 'Z')
}

/**
 * Word-like core part of a freshly created document: no creator or other
 * people data, just the creation/modification stamp and revision 1, so later
 * property sets (and the save path's modified/revision update) have a home.
 */
export function newDocumentCorePropertiesXml(created: Date = new Date()): string {
  const stamp = w3cdtfStamp(created)
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"' +
    ' xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/"' +
    ' xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
    `<dcterms:created xsi:type="dcterms:W3CDTF">${stamp}</dcterms:created>` +
    `<dcterms:modified xsi:type="dcterms:W3CDTF">${stamp}</dcterms:modified>` +
    '<cp:revision>1</cp:revision>' +
    '</cp:coreProperties>'
  )
}

/** Word-like extended part of a freshly created document (no company/manager) */
export function newDocumentAppPropertiesXml(
  opts: { application?: string; appVersion?: string; template?: string } = {},
): string {
  const application = opts.application ?? NEW_DOCUMENT_APPLICATION
  const appVersion = opts.appVersion ?? NEW_DOCUMENT_APP_VERSION
  const template = opts.template ?? NEW_DOCUMENT_TEMPLATE
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"' +
    ' xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">' +
    `<Template>${escapeXmlText(template)}</Template>` +
    `<Application>${escapeXmlText(application)}</Application>` +
    `<AppVersion>${escapeXmlText(appVersion)}</AppVersion>` +
    '</Properties>'
  )
}

function blankCustomPropertiesXml(): string {
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"' +
    ' xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">' +
    '</Properties>'
  )
}

/** every child of the root goes, known or not; the root itself stays */
function clearRootChildrenXml(xml: string, root: string): string {
  const re = new RegExp(
    `(<(?:[A-Za-z_][\\w.-]*:)?${root}(?=[\\s>])[^>]*>)[\\s\\S]*(</(?:[A-Za-z_][\\w.-]*:)?${root}\\s*>)`,
  )
  return xml.replace(re, '$1$2')
}

function clearCorePropertiesXml(xml: string): string {
  return clearRootChildrenXml(xml, 'coreProperties')
}

function clearAppPropertiesXml(xml: string): string {
  return clearRootChildrenXml(xml, 'Properties')
}

/** every custom property goes; the part itself stays */
function clearCustomPropertiesXml(xml: string): string {
  return xml.replace(PROPERTY_RE, '')
}

/** set or remove the core properties of a core.xml part */
export function patchCorePropertiesXml(
  xml: string,
  changes: Record<string, string | null>,
): string {
  let out = xml
  for (const [key, value] of Object.entries(changes)) {
    const spec = CORE_PROPERTIES[key]
    if (!spec) continue
    if (value === null) {
      out = out.replace(elementRe(spec.local), '')
      continue
    }
    const name = out.includes(`xmlns:${spec.prefix}=`) ? `${spec.prefix}:${spec.local}` : spec.local
    const attrs =
      (spec.local === 'created' || spec.local === 'modified') && out.includes('xmlns:xsi=')
        ? ' xsi:type="dcterms:W3CDTF"'
        : ''
    out = setElementText(
      out,
      spec.local,
      value,
      `<${name}${attrs}>${escapeXmlText(value)}</${name}>`,
    )
  }
  return out
}

/** set or remove the extended properties of an app.xml part */
export function patchAppPropertiesXml(xml: string, changes: Record<string, string | null>): string {
  let out = xml
  for (const [key, value] of Object.entries(changes)) {
    const local = APP_PROPERTIES[key]
    if (!local) continue
    if (value === null) {
      out = out.replace(elementRe(local), '')
      continue
    }
    out = setAppElement(out, local, value)
  }
  return out
}

/** an app.xml element goes where the schema sequence puts it */
function setAppElement(xml: string, local: string, value: string): string {
  let found = false
  const out = xml.replace(elementRe(local), (_match, openPrefix, attrs) => {
    found = true
    const name = `${openPrefix ?? ''}${local}`
    return `<${name}${attrs}>${escapeXmlText(value)}</${name}>`
  })
  if (found) return out
  const element = `<${local}>${escapeXmlText(value)}</${local}>`
  const at = APP_PROPERTY_ORDER.indexOf(local)
  for (const other of APP_PROPERTY_ORDER) {
    if (APP_PROPERTY_ORDER.indexOf(other) <= at) continue
    const match = elementRe(other).exec(xml)
    if (match) return xml.slice(0, match.index) + element + xml.slice(match.index)
  }
  return insertElement(xml, element)
}

/** the type a new custom property gets when the value cannot be typed */
const CUSTOM_PROPERTY_DEFAULT_TYPE = 'lpwstr'

/** vt integer types and the range each can hold */
const VT_INT_RANGES: Record<string, readonly [number, number]> = {
  i1: [-128, 127],
  i2: [-32768, 32767],
  i4: [-2147483648, 2147483647],
  i8: [Number.MIN_SAFE_INTEGER, Number.MAX_SAFE_INTEGER],
  ui1: [0, 255],
  ui2: [0, 65535],
  ui4: [0, 4294967295],
  ui8: [0, Number.MAX_SAFE_INTEGER],
  int: [-2147483648, 2147483647],
  uint: [0, 4294967295],
}
const VT_FLOAT_TYPES = new Set(['r4', 'r8', 'decimal', 'cy'])
const VT_TIME_TYPES = new Set(['filetime', 'date'])
const VT_TEXT_TYPES = new Set(['lpwstr', 'lpstr', 'bstr'])

/** the text form of a scalar op value (used for text fallbacks) */
function scalarText(value: PropertyValue): string {
  if (typeof value === 'boolean') return value ? 'true' : 'false'
  return String(value)
}

/** Word's W3CDTF text (a bare date fills the day start, no zone means Z) */
function w3cdtfValue(value: string): string | null {
  const match = /^(\d{4}-\d{2}-\d{2})(?:T(\d{2}:\d{2}:\d{2})(?:\.\d+)?(Z|[+-]\d{2}:\d{2})?)?$/.exec(
    value.trim(),
  )
  if (!match) return null
  return match[2] ? `${match[1]}T${match[2]}${match[3] ?? 'Z'}` : `${match[1]}T00:00:00Z`
}

/**
 * The text a value carries in an existing vt type, null when the value does
 * not fit it: numbers for the integer/float types, true/false for bool, an
 * ISO-8601 stamp for filetime/date, anything for the text types.
 */
function typedCustomText(type: string, value: PropertyValue): string | null {
  if (VT_TEXT_TYPES.has(type)) return scalarText(value)
  if (type === 'bool') {
    if (typeof value === 'boolean') return value ? 'true' : 'false'
    if (typeof value === 'number') {
      if (value === 1) return 'true'
      if (value === 0) return 'false'
    }
    if (typeof value === 'string') {
      const text = value.trim().toLowerCase()
      if (text === 'true' || text === '1') return 'true'
      if (text === 'false' || text === '0') return 'false'
    }
    return null
  }
  const range = VT_INT_RANGES[type]
  if (range) {
    const number =
      typeof value === 'number'
        ? value
        : typeof value === 'string' && /^[+-]?\d+$/.test(value.trim())
          ? Number(value.trim())
          : null
    if (number === null || !Number.isInteger(number)) return null
    if (number < range[0] || number > range[1]) return null
    return String(number)
  }
  if (VT_FLOAT_TYPES.has(type)) {
    if (typeof value === 'number') return Number.isFinite(value) ? String(value) : null
    if (
      typeof value === 'string' &&
      /^[+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?$/.test(value.trim())
    )
      return value.trim()
    return null
  }
  if (VT_TIME_TYPES.has(type)) return typeof value === 'string' ? w3cdtfValue(value) : null
  return null
}

/**
 * The vt type and text of a set value: an existing property keeps its type
 * while the value fits it, otherwise (and for a new property) the type comes
 * from the value — text stays lpwstr, a boolean is bool, an integer i4/i8 and
 * a fraction r8.
 */
function resolveCustomValue(
  existingType: string | null,
  value: PropertyValue,
): { type: string; text: string } {
  if (existingType) {
    const text = typedCustomText(existingType, value)
    if (text !== null) return { type: existingType, text }
    return { type: CUSTOM_PROPERTY_DEFAULT_TYPE, text: scalarText(value) }
  }
  if (typeof value === 'boolean') return { type: 'bool', text: value ? 'true' : 'false' }
  if (typeof value === 'number') {
    if (!Number.isFinite(value)) return { type: CUSTOM_PROPERTY_DEFAULT_TYPE, text: String(value) }
    if (Number.isInteger(value) && value >= -2147483648 && value <= 2147483647)
      return { type: 'i4', text: String(value) }
    if (Number.isSafeInteger(value)) return { type: 'i8', text: String(value) }
    return { type: 'r8', text: String(value) }
  }
  return { type: CUSTOM_PROPERTY_DEFAULT_TYPE, text: value }
}

/** the prefix a custom.xml part binds the docPropsVTypes namespace to */
function vTypesPrefix(xml: string): string {
  return /xmlns:([A-Za-z_][\w.-]*)="[^"]*docPropsVTypes"/.exec(xml)?.[1] ?? 'vt'
}

/** set or remove the custom properties of a custom.xml part, keeping stored vt types */
export function patchCustomPropertiesXml(
  xml: string,
  changes: Record<string, PropertyValue | null>,
): string {
  let out = xml
  let maxPid = 1
  for (const property of out.match(PROPERTY_RE) ?? []) {
    const pid = Number(attrValue(property, 'pid'))
    if (Number.isSafeInteger(pid)) maxPid = Math.max(maxPid, pid)
  }
  const existingTypes = parseCustomPropertyTypesXml(out)
  const prefix = vTypesPrefix(out)
  for (const [name, value] of Object.entries(changes)) {
    if (value === null) {
      out = out.replace(PROPERTY_RE, (property) =>
        attrValue(property, 'name') === name ? '' : property,
      )
      continue
    }
    const { type, text } = resolveCustomValue(existingTypes[name] ?? null, value)
    const valueXml = `<${prefix}:${type}>${escapeXmlText(text)}</${prefix}:${type}>`
    let replaced = false
    out = out.replace(PROPERTY_RE, (property) => {
      if (attrValue(property, 'name') !== name) return property
      replaced = true
      const fmtid = attrValue(property, 'fmtid') ?? CUSTOM_PROPERTIES_FMTID
      const pid = attrValue(property, 'pid') ?? String(maxPid + 1)
      return (
        `<property fmtid="${escapeXmlAttr(fmtid)}" pid="${escapeXmlAttr(pid)}"` +
        ` name="${escapeXmlAttr(name)}">${valueXml}</property>`
      )
    })
    if (replaced) continue
    maxPid += 1
    out = insertElement(
      out,
      `<property fmtid="${CUSTOM_PROPERTIES_FMTID}" pid="${maxPid}"` +
        ` name="${escapeXmlAttr(name)}">${valueXml}</property>`,
    )
  }
  return out
}

function maxRelId(relsXml: string): number {
  let max = 0
  for (const match of relsXml.matchAll(/\bId\s*=\s*["']rId(\d+)["']/g)) {
    const n = parseInt(match[1]!, 10)
    if (n > max) max = n
  }
  return max
}

async function addContentTypeOverride(
  zip: JSZip,
  partName: string,
  contentType: string,
): Promise<void> {
  const file = zip.file('[Content_Types].xml')
  if (!file) return
  const xml = await file.async('string')
  if (xml.includes(`PartName="${partName}"`)) return
  zip.file(
    '[Content_Types].xml',
    xml.replace(
      '</Types>',
      `<Override PartName="${partName}" ContentType="${contentType}"/></Types>`,
    ),
    { date: file.date },
  )
}

async function addRootRelationship(zip: JSZip, relType: string, target: string): Promise<void> {
  const path = '_rels/.rels'
  const file = zip.file(path)
  const xml = file
    ? await file.async('string')
    : '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>'
  if (xml.includes(`Type="${relType}"`)) return
  const id = `rId${maxRelId(xml) + 1}`
  const out = xml.replace(
    '</Relationships>',
    `<Relationship Id="${id}" Type="${relType}" Target="${target}"/></Relationships>`,
  )
  zip.file(path, out, file ? { date: file.date } : undefined)
}

/** write a docProps part, creating it (with its content type and relationship) when missing */
async function writePropertiesPart(
  zip: JSZip,
  path: string,
  xml: string,
  contentType: string,
  relType: string,
): Promise<void> {
  const entry = zip.file(path)
  if (entry) {
    zip.file(path, xml, { date: entry.date })
    return
  }
  zip.file(path, xml)
  await addContentTypeOverride(zip, `/${path}`, contentType)
  await addRootRelationship(zip, relType, path)
}

const hasSetValue = (values: Record<string, unknown> | undefined): boolean =>
  !!values && Object.values(values).some((value) => value !== null)

/**
 * Write the requested property changes into the package: a string sets a
 * property, null removes it, clear: "all" empties all three parts. A part that
 * is not there yet is created, so a document without docProps can still take
 * properties.
 */
export async function applyPropertyChanges(zip: JSZip, changes: PropertyChanges): Promise<void> {
  const clearAll = changes.clear === 'all'
  if (!clearAll && !changes.core && !changes.app && !changes.custom) return

  if (clearAll || changes.core) {
    const entry = zip.file(CORE_PROPERTIES_PATH)
    const values = changes.core ?? {}
    if (entry || hasSetValue(values)) {
      const xml = entry ? await entry.async('string') : blankCorePropertiesXml()
      // the clear comes first: a value given in the same op still lands
      const out = patchCorePropertiesXml(clearAll ? clearCorePropertiesXml(xml) : xml, values)
      await writePropertiesPart(
        zip,
        CORE_PROPERTIES_PATH,
        out,
        CORE_PROPERTIES_CONTENT_TYPE,
        CORE_PROPERTIES_REL_TYPE,
      )
    }
  }

  if (clearAll || changes.app) {
    const entry = zip.file(APP_PROPERTIES_PATH)
    const values = changes.app ?? {}
    if (entry || hasSetValue(values)) {
      const xml = entry ? await entry.async('string') : blankAppPropertiesXml()
      const out = patchAppPropertiesXml(clearAll ? clearAppPropertiesXml(xml) : xml, values)
      await writePropertiesPart(
        zip,
        APP_PROPERTIES_PATH,
        out,
        APP_PROPERTIES_CONTENT_TYPE,
        APP_PROPERTIES_REL_TYPE,
      )
    }
  }

  if (clearAll || changes.custom) {
    const entry = zip.file(CUSTOM_PROPERTIES_PATH)
    const values = changes.custom ?? {}
    if (entry || hasSetValue(values)) {
      const xml = entry ? await entry.async('string') : blankCustomPropertiesXml()
      const out = patchCustomPropertiesXml(clearAll ? clearCustomPropertiesXml(xml) : xml, values)
      await writePropertiesPart(
        zip,
        CUSTOM_PROPERTIES_PATH,
        out,
        CUSTOM_PROPERTIES_CONTENT_TYPE,
        CUSTOM_PROPERTIES_REL_TYPE,
      )
    }
  }
}

/**
 * The set_properties op: core / app / custom hold the values to set (null
 * removes), clear: "all" empties every part. Custom values may be strings,
 * numbers or booleans (a typed value keeps an existing property's vt type);
 * core/app stay strings. Throws on anything else.
 */
export function parsePropertyChanges(input: Record<string, unknown>): PropertyChanges {
  const changes: PropertyChanges = {}
  if (input.clear !== undefined) {
    if (input.clear !== 'all') throw new Error('set_properties: clear must be "all"')
    changes.clear = 'all'
  }
  for (const section of ['core', 'app', 'custom'] as const) {
    const values = input[section]
    if (values === undefined) continue
    if (!values || typeof values !== 'object' || Array.isArray(values)) {
      throw new Error(`set_properties: ${section} must be an object of property values`)
    }
    const known =
      section === 'core' ? CORE_PROPERTY_KEYS : section === 'app' ? APP_PROPERTY_KEYS : []
    const out: Record<string, PropertyValue | null> = {}
    for (const [key, value] of Object.entries(values as Record<string, unknown>)) {
      if (known.length > 0 && !known.includes(key)) {
        throw new Error(
          `set_properties: unknown ${section} property "${key}" (known: ${known.join(', ')})`,
        )
      }
      if (value === null) {
        out[key] = null
        continue
      }
      const typed = typeof value === 'number' || typeof value === 'boolean'
      if (typeof value !== 'string' && !(section === 'custom' && typed)) {
        throw new Error(
          section === 'custom'
            ? `set_properties: custom.${key} must be a string, number, boolean or null`
            : `set_properties: ${section}.${key} must be a string or null`,
        )
      }
      out[key] = value
    }
    if (section === 'custom') changes.custom = out as Record<string, PropertyValue | null>
    else changes[section] = out as Record<string, string | null>
  }
  if (!changes.clear && isEmpty(changes.core) && isEmpty(changes.app) && isEmpty(changes.custom)) {
    throw new Error('set_properties: nothing to change (give core, app, custom or clear)')
  }
  return changes
}

function isEmpty(values: Record<string, unknown> | undefined): boolean {
  return !values || Object.keys(values).length === 0
}
