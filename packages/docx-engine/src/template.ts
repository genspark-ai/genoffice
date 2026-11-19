import JSZip from 'jszip'
import { assertZipInflatesWithinLimits, loadDocxZip } from './zip-load'

const DOCUMENT_MAIN =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'
const TEMPLATE_MAIN =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml'
const MACRO_TEMPLATE_MAIN = 'application/vnd.ms-word.template.macroEnabledTemplate.main+xml'

/** the project, its signatures and the Word data beside it: what makes a package macro-enabled */
const MACRO_PART = /(?:^|\/)vba(?:Project|Data)[^/]*$/i
const MACRO_REL_TYPE = /\/(?:vbaProject\w*|wordVbaData)$/

const THUMBNAIL_PART = /^docProps\/thumbnail\.(jpe?g|png)$/i
const THUMBNAIL_REL =
  'http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail'
const ROOT_RELS = '_rels/.rels'
const RELATIONSHIP = /<Relationship\b[^>]*?\/>/g
const OVERRIDE = /<Override\b[^>]*?\/>/g
const TYPES = '[Content_Types].xml'

const generate = (zip: JSZip): Promise<Uint8Array> =>
  zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })

/**
 * A .dotx is a .docx whose main part is declared as a template; Word refuses a
 * file whose extension and declared type disagree, so the declaration has to
 * follow the file wherever it is saved.
 */
async function withMainType(bytes: Uint8Array, from: string, to: string): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(bytes)
  const types = zip.file('[Content_Types].xml')
  const xml = await types?.async('string')
  if (!xml || !xml.includes(from)) return bytes
  zip.file('[Content_Types].xml', xml.split(from).join(to))
  if (to === DOCUMENT_MAIN) await dropThumbnail(zip)
  return generate(zip)
}

/**
 * The picture of the template's first page is not a picture of the document
 * made from it. Returns whether the package carried one.
 */
async function dropThumbnail(zip: JSZip): Promise<boolean> {
  const rels = await zip.file(ROOT_RELS)?.async('string')
  if (!rels) return false
  const dropped: string[] = []
  const kept = rels.replace(RELATIONSHIP, (rel) => {
    if (/\bType="([^"]*)"/.exec(rel)?.[1] !== THUMBNAIL_REL) return rel
    const target = /\bTarget="\/?([^"]*)"/.exec(rel)?.[1]
    if (target) {
      zip.remove(target)
      dropped.push(target)
    }
    return ''
  })
  if (kept === rels) return false
  zip.file(ROOT_RELS, kept)
  const types = await zip.file(TYPES)?.async('string')
  const declared = types?.replace(OVERRIDE, (override) =>
    dropped.includes(/\bPartName="\/?([^"]*)"/.exec(override)?.[1] ?? '') ? '' : override,
  )
  if (declared !== undefined && declared !== types) zip.file(TYPES, declared)
  return true
}

/**
 * A .docx may not hold macros: Word refuses one that does. The document made
 * from a .dotm keeps everything of the template but its VBA project.
 */
async function macroTemplateAsDocx(zip: JSZip, typesXml: string): Promise<Uint8Array> {
  const dropped = Object.keys(zip.files).filter((name) => MACRO_PART.test(name))
  for (const name of dropped) {
    zip.remove(name)
    const slash = name.lastIndexOf('/')
    zip.remove(`${name.slice(0, slash + 1)}_rels/${name.slice(slash + 1)}.rels`)
  }
  for (const name of Object.keys(zip.files)) {
    if (!/(?:^|\/)_rels\/[^/]+\.rels$/.test(name)) continue
    const rels = await zip.file(name)!.async('string')
    const kept = rels.replace(RELATIONSHIP, (rel) =>
      MACRO_REL_TYPE.test(/\bType="([^"]*)"/.exec(rel)?.[1] ?? '') ? '' : rel,
    )
    if (kept !== rels) zip.file(name, kept)
  }
  const types = typesXml
    .split(MACRO_TEMPLATE_MAIN)
    .join(DOCUMENT_MAIN)
    .replace(/<Override\b[^>]*?\/>/g, (override) => {
      const part = /\bPartName="\/?([^"]*)"/.exec(override)?.[1] ?? ''
      return dropped.includes(part) ? '' : override
    })
  zip.file('[Content_Types].xml', types)
  await dropThumbnail(zip)
  return generate(zip)
}

/** the same package declared as a template (.dotx) */
export const docxAsTemplate = (bytes: Uint8Array): Promise<Uint8Array> =>
  withMainType(bytes, DOCUMENT_MAIN, TEMPLATE_MAIN)

/**
 * A template (.dotx, or a .dotm without its macros) declared as a document,
 * so what is made from it saves as a .docx.
 */
export async function templateAsDocx(bytes: Uint8Array): Promise<Uint8Array> {
  // a template is a file from outside: every part is inflated below, ahead of the parser's gate
  await assertZipInflatesWithinLimits(bytes)
  const zip = await JSZip.loadAsync(bytes)
  const xml = await zip.file('[Content_Types].xml')?.async('string')
  if (xml?.includes(MACRO_TEMPLATE_MAIN)) return macroTemplateAsDocx(zip, xml)
  if (xml?.includes(TEMPLATE_MAIN)) return withMainType(bytes, TEMPLATE_MAIN, DOCUMENT_MAIN)
  // a document used as a template
  return (await dropThumbnail(zip)) ? generate(zip) : bytes
}

/**
 * The preview picture a template carries for a gallery, when it is one a
 * browser shows. A gallery reads files nobody chose to open, so the package
 * goes through the size gate and throws when it is past the limits.
 */
export async function templateThumbnail(
  bytes: Uint8Array,
): Promise<{ mime: 'image/jpeg' | 'image/png'; data: Uint8Array } | null> {
  const zip = await loadDocxZip(bytes)
  const rels = await zip.file(ROOT_RELS)?.async('string')
  for (const rel of rels?.match(RELATIONSHIP) ?? []) {
    if (/\bType="([^"]*)"/.exec(rel)?.[1] !== THUMBNAIL_REL) continue
    const target = /\bTarget="\/?([^"]*)"/.exec(rel)?.[1] ?? ''
    const kind = THUMBNAIL_PART.exec(target)?.[1].toLowerCase()
    const part = kind ? zip.file(target) : null
    if (!part) return null
    return {
      mime: kind === 'png' ? 'image/png' : 'image/jpeg',
      data: await part.async('uint8array'),
    }
  }
  return null
}

/** the package with `jpeg` as its preview picture, in place of any it had */
export async function withTemplateThumbnail(
  bytes: Uint8Array,
  jpeg: Uint8Array,
): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(bytes)
  if (!zip.file(ROOT_RELS) || !zip.file(TYPES)) return bytes
  await dropThumbnail(zip)
  const types = await zip.file(TYPES)!.async('string')
  const current = (await zip.file(ROOT_RELS)!.async('string')).replace(
    '</Relationships>',
    `<Relationship Id="rIdThumbnail" Type="${THUMBNAIL_REL}" Target="docProps/thumbnail.jpeg"/></Relationships>`,
  )
  zip.file(ROOT_RELS, current)
  zip.file('docProps/thumbnail.jpeg', jpeg)
  if (!/<Default\b[^>]*\bExtension="jpeg"/i.test(types))
    zip.file(
      '[Content_Types].xml',
      types.replace(/(<Types\b[^>]*>)/, '$1<Default Extension="jpeg" ContentType="image/jpeg"/>'),
    )
  return generate(zip)
}
