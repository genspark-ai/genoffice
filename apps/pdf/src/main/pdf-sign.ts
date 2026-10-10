import { X509Certificate } from 'node:crypto'
import { deflateSync } from 'node:zlib'
import forge from 'node-forge'
import { PDFArray, PDFDict, PDFDocument, PDFName, PDFNumber, PDFRef, PDFString } from 'pdf-lib'
import type { PDFObject } from 'pdf-lib'
import type {
  CertificateErrorCode,
  CertificateInfo,
  InspectCertificateResult,
  SignaturePosition,
} from '../shared/ipc'
import { useNativePkcs12Crypto } from './forge-native'
import { collectSignatureFields, describeCertificate } from './pdf-signatures'

useNativePkcs12Crypto()

/** Error carrying a code the renderer maps to localized text */
export class SignError extends Error {
  constructor(
    readonly code: CertificateErrorCode | 'pdf-encrypted' | 'sign-failed',
    cause?: unknown,
  ) {
    super(code, { cause })
  }
}

export interface LoadedCertificate {
  privateKey: forge.pki.rsa.PrivateKey
  certificate: forge.pki.Certificate
  /** Signer certificate first, then any other certificate shipped in the file */
  chain: forge.pki.Certificate[]
}

const toX509 = (cert: forge.pki.Certificate): X509Certificate =>
  new X509Certificate(
    Buffer.from(forge.asn1.toDer(forge.pki.certificateToAsn1(cert)).getBytes(), 'binary'),
  )

/** Decrypt a PKCS#12 bundle and pick the certificate that belongs to its private key */
export function loadCertificate(
  p12: Uint8Array,
  password: string,
  wantedSha1?: string,
): LoadedCertificate {
  let bundle: forge.pkcs12.Pkcs12Pfx
  try {
    const der = forge.util.createBuffer(Buffer.from(p12).toString('binary'))
    bundle = forge.pkcs12.pkcs12FromAsn1(forge.asn1.fromDer(der), password)
  } catch (err) {
    const message = err instanceof Error ? err.message : ''
    // forge reports a wrong password as a failed MAC (or a failed decrypt for MAC-less files)
    if (/password|MAC|decrypt/i.test(message)) throw new SignError('cert-password', err)
    if (/unsupported|unknown|OID|curve|private key|RSAPrivateKey/i.test(message))
      throw new SignError('cert-unsupported-key', err)
    throw new SignError('cert-invalid', err)
  }

  const keyBags = [
    ...(bundle.getBags({ bagType: forge.pki.oids.pkcs8ShroudedKeyBag })[
      forge.pki.oids.pkcs8ShroudedKeyBag
    ] ?? []),
    ...(bundle.getBags({ bagType: forge.pki.oids.keyBag })[forge.pki.oids.keyBag] ?? []),
  ]
  const certBags = bundle.getBags({ bagType: forge.pki.oids.certBag })[forge.pki.oids.certBag] ?? []
  const keys = keyBags
    .map((bag) => bag.key as forge.pki.rsa.PrivateKey | undefined)
    .filter((key): key is forge.pki.rsa.PrivateKey => !!key)
  if (keys.length === 0)
    throw new SignError(keyBags.length > 0 ? 'cert-unsupported-key' : 'cert-no-key')
  const certs = certBags
    .map((bag) => bag.cert)
    .filter((cert): cert is forge.pki.Certificate => !!cert)
  if (certs.length === 0) throw new SignError('cert-invalid')

  // A bundle may hold several identities (a whole keychain export); `wantedSha1` picks one
  const candidates = wantedSha1 ? certs.filter((cert) => sha1Of(cert) === wantedSha1) : certs
  for (const signer of candidates) {
    const pub = signer.publicKey as forge.pki.rsa.PublicKey
    const privateKey = keys.find(
      (key) => !!pub.n && !!key.n && pub.n.compareTo(key.n) === 0 && pub.e.compareTo(key.e) === 0,
    )
    if (privateKey) {
      return {
        privateKey,
        certificate: signer,
        chain: [signer, ...certs.filter((c) => c !== signer)],
      }
    }
  }
  throw new SignError('cert-unsupported-key')
}

const sha1Of = (cert: forge.pki.Certificate): string =>
  forge.md.sha1
    .create()
    .update(forge.asn1.toDer(forge.pki.certificateToAsn1(cert)).getBytes())
    .digest()
    .toHex()

/** Decrypt once; the result feeds both the preview and the signature, so the key is read a single time */
export function openCertificate(
  p12: Uint8Array,
  password: string,
  wantedSha1?: string,
):
  | { ok: true; loaded: LoadedCertificate; info: CertificateInfo[] }
  | { ok: false; error: CertificateErrorCode } {
  try {
    const loaded = loadCertificate(p12, password, wantedSha1)
    const info: CertificateInfo[] = loaded.chain.map((cert) => describeCertificate(toX509(cert)))
    return { ok: true, loaded, info }
  } catch (err) {
    const code = err instanceof SignError ? err.code : 'cert-invalid'
    return {
      ok: false,
      error: code === 'sign-failed' || code === 'pdf-encrypted' ? 'cert-invalid' : code,
    }
  }
}

export function inspectCertificate(
  p12: Uint8Array,
  password: string,
  wantedSha1?: string,
): InspectCertificateResult {
  const opened = openCertificate(p12, password, wantedSha1)
  return opened.ok ? { ok: true, signer: opened.info[0]!, chain: opened.info } : opened
}

// ── PDF structure ───────────────────────────────────────────────────────────

export interface VisibleSignature {
  pageIndex: number
  position: SignaturePosition
  width: number
  height: number
  png: Uint8Array
}

export interface SignPdfOptions {
  /** Already decrypted by openCertificate; skips decrypting p12 again */
  loaded?: LoadedCertificate
  p12: Uint8Array
  password: string
  /** Lower-case SHA-1 of the signer certificate when the bundle holds several identities */
  signerSha1?: string
  reason?: string
  location?: string
  contactInfo?: string
  certifyLevel?: 1 | 2 | 3
  visible?: VisibleSignature
  /** Defaults to now; tests pin it */
  signingTime?: Date
}

/** Hex characters reserved for the CMS blob; doubled once if the certificate chain turns out larger */
const PLACEHOLDER_SIZES = [16 * 1024, 48 * 1024]
const BYTE_RANGE_DIGITS = 10
const MAX_SIGNABLE_BYTES = 400 * 1024 * 1024

const latin1 = (text: string): Buffer => Buffer.from(text, 'latin1')

/** PDF text string: plain literal when ASCII, UTF-16BE hex (with BOM) otherwise */
function pdfText(value: string): string {
  if (/^[\x20-\x7e]*$/.test(value)) {
    return `(${value.replace(/[\\()]/g, (c) => `\\${c}`)})`
  }
  const units = [0xfeff]
  for (let i = 0; i < value.length; i++) units.push(value.charCodeAt(i))
  return `<${units
    .map((u) => u.toString(16).padStart(4, '0'))
    .join('')
    .toUpperCase()}>`
}

const pdfDate = (date: Date): string => {
  const p = (n: number) => String(n).padStart(2, '0')
  return `D:${date.getUTCFullYear()}${p(date.getUTCMonth() + 1)}${p(date.getUTCDate())}${p(date.getUTCHours())}${p(date.getUTCMinutes())}${p(date.getUTCSeconds())}Z`
}

const serialize = (ref: PDFRef, object: PDFObject): Buffer => {
  const body = new Uint8Array(object.sizeInBytes())
  object.copyBytesInto(body, 0)
  return Buffer.concat([
    latin1(`${ref.objectNumber} ${ref.generationNumber} obj\n`),
    Buffer.from(body),
    latin1('\nendobj\n'),
  ])
}

interface PreviousXref {
  offset: number
  /** The previous section is an xref stream, so this update must be one too */
  isStream: boolean
  size: number
}

function readPreviousXref(bytes: Buffer): PreviousXref {
  const tail = bytes.subarray(Math.max(0, bytes.length - 2048)).toString('latin1')
  const marker = tail.lastIndexOf('startxref')
  const match = marker >= 0 ? /startxref\s+(\d+)/.exec(tail.slice(marker)) : null
  if (!match) throw new SignError('sign-failed')
  const offset = Number(match[1])
  if (!Number.isSafeInteger(offset) || offset <= 0 || offset >= bytes.length) {
    throw new SignError('sign-failed')
  }
  const head = bytes.subarray(offset, offset + 4096).toString('latin1')
  const isStream = !head.trimStart().startsWith('xref')
  const searchFrom = isStream ? head : head.slice(Math.max(0, head.indexOf('trailer')))
  const size = /\/Size\s+(\d+)/.exec(searchFrom)
  return { offset, isStream, size: size ? Number(size[1]) : 0 }
}

/** Rewritable copy of the object a reference points to, plus the ref to write it back under */
function editableDict(
  pdf: PDFDocument,
  value: PDFObject | undefined,
): { ref?: PDFRef; dict: PDFDict } | undefined {
  if (value instanceof PDFRef) {
    const dict = pdf.context.lookupMaybe(value, PDFDict)
    return dict ? { ref: value, dict: dict.clone(pdf.context) } : undefined
  }
  return value instanceof PDFDict ? { dict: value.clone(pdf.context) } : undefined
}

function editableArray(
  pdf: PDFDocument,
  value: PDFObject | undefined,
): { ref?: PDFRef; array: PDFArray } {
  if (value instanceof PDFRef) {
    const array = pdf.context.lookupMaybe(value, PDFArray)
    if (array) return { ref: value, array: array.clone(pdf.context) }
  }
  return { array: value instanceof PDFArray ? value.clone(pdf.context) : pdf.context.obj([]) }
}

/** Box in default user space for a corner of the page as the viewer displays it (rotation applied) */
function placeBox(
  page: ReturnType<PDFDocument['getPages']>[number],
  position: SignaturePosition,
  width: number,
  height: number,
): { rect: [number, number, number, number]; rotation: number; w: number; h: number } {
  const crop = page.getCropBox()
  const rotation = ((page.getRotation().angle % 360) + 360) % 360
  const sideways = rotation === 90 || rotation === 270
  const viewW = sideways ? crop.height : crop.width
  const viewH = sideways ? crop.width : crop.height
  const w = Math.min(width, viewW)
  const h = Math.min(height, viewH)
  const margin = Math.min(36, viewW / 10, viewH / 10)
  const left = position.endsWith('left')
    ? margin
    : position === 'center'
      ? (viewW - w) / 2
      : viewW - margin - w
  const bottom = position.startsWith('top')
    ? viewH - margin - h
    : position === 'center'
      ? (viewH - h) / 2
      : margin
  // Map the displayed rectangle (origin bottom-left of the rotated page) back to unrotated space
  const corners: [number, number][] = [
    [left, bottom],
    [left + w, bottom + h],
  ]
  const mapped = corners.map(([vx, vy]): [number, number] => {
    if (rotation === 90) return [viewH - vy, vx]
    if (rotation === 180) return [viewW - vx, viewH - vy]
    if (rotation === 270) return [vy, viewW - vx]
    return [vx, vy]
  })
  const [[ax, ay], [bx, by]] = mapped as [[number, number], [number, number]]
  return {
    rect: [
      crop.x + Math.min(ax, bx),
      crop.y + Math.min(ay, by),
      crop.x + Math.max(ax, bx),
      crop.y + Math.max(ay, by),
    ],
    rotation,
    w,
    h,
  }
}

function xrefTable(entries: { num: number; offset: number; gen: number }[]): Buffer {
  const sorted = [...entries].sort((a, b) => a.num - b.num)
  const runs: (typeof sorted)[] = []
  for (const entry of sorted) {
    const run = runs[runs.length - 1]
    if (run && run[run.length - 1]!.num + 1 === entry.num) run.push(entry)
    else runs.push([entry])
  }
  let text = 'xref\n'
  for (const run of runs) {
    text += `${run[0]!.num} ${run.length}\n`
    for (const e of run) {
      text += `${String(e.offset).padStart(10, '0')} ${String(e.gen).padStart(5, '0')} n \n`
    }
  }
  return latin1(text)
}

function xrefStreamData(entries: { num: number; offset: number; gen: number }[]): {
  index: number[]
  data: Buffer
} {
  const sorted = [...entries].sort((a, b) => a.num - b.num)
  const index: number[] = []
  const rows: Buffer[] = []
  let previous = -2
  for (const e of sorted) {
    if (e.num === previous + 1) index[index.length - 1]!++
    else index.push(e.num, 1)
    previous = e.num
    const row = Buffer.alloc(7)
    row[0] = 1
    row.writeUInt32BE(e.offset, 1)
    row.writeUInt16BE(e.gen, 5)
    rows.push(row)
  }
  return { index, data: Buffer.concat(rows) }
}

interface Assembled {
  file: Buffer
  contentsStart: number
  contentsEnd: number
  byteRangeAt: number
}

/**
 * Append a signature revision to `original` (incremental update, so earlier signatures stay
 * valid) with a zero-filled /Contents of `placeholderHex` characters and a fixed-width
 * /ByteRange. The caller fills both in once the byte layout is known.
 */
async function assemble(
  original: Buffer,
  options: SignPdfOptions,
  placeholderHex: number,
): Promise<Assembled> {
  let pdf: PDFDocument
  try {
    pdf = await PDFDocument.load(original, { updateMetadata: false, throwOnInvalidObject: false })
  } catch (err) {
    if (err instanceof Error && /encrypt/i.test(err.message)) throw new SignError('pdf-encrypted')
    throw new SignError('sign-failed')
  }
  const { context } = pdf
  const previous = readPreviousXref(original)
  // Object numbers must not collide with free slots of the previous table either
  context.largestObjectNumber = Math.max(context.largestObjectNumber, previous.size - 1)
  const firstNew = context.largestObjectNumber + 1

  const existing = collectSignatureFields(pdf)
  if (options.certifyLevel && existing.length > 0) throw new SignError('sign-failed')
  const pageIndex = options.visible?.pageIndex ?? 0
  const page = pdf.getPages()[pageIndex]
  if (!page) throw new SignError('sign-failed')

  const sigRef = context.nextRef()
  const widgetRef = context.nextRef()
  const catalogRef = context.trailerInfo.Root
  if (!(catalogRef instanceof PDFRef)) throw new SignError('sign-failed')

  // Visible appearance: the renderer's PNG as an image painted by a form XObject
  let appearanceRef: PDFRef | undefined
  let rect: [number, number, number, number] = [0, 0, 0, 0]
  if (options.visible) {
    const placed = placeBox(
      page,
      options.visible.position,
      options.visible.width,
      options.visible.height,
    )
    rect = placed.rect
    const { rotation } = placed
    const image = await pdf.embedPng(options.visible.png)
    await image.embed()
    // The box lives in unrotated page space; the form's /Matrix turns its upright BBox to match
    const radians = (rotation * Math.PI) / 180
    const cos = Math.round(Math.cos(radians))
    const sin = Math.round(Math.sin(radians))
    const stream = context.stream(`q ${placed.w} 0 0 ${placed.h} 0 0 cm /Sig0 Do Q`, {
      Type: 'XObject',
      Subtype: 'Form',
      BBox: [0, 0, placed.w, placed.h],
      Matrix: [cos, sin, -sin, cos, 0, 0],
      Resources: { XObject: { Sig0: image.ref } },
    })
    appearanceRef = context.register(stream)
  }

  // Fresh objects created above (image, appearance) are everything past `firstNew`; the two
  // reserved refs are written by hand below.
  const modified = new Map<number, { ref: PDFRef; object: PDFObject }>()
  const touch = (ref: PDFRef, object: PDFObject) => modified.set(ref.objectNumber, { ref, object })

  // Page /Annots
  const pageEdit = editableDict(pdf, page.ref)!
  const annots = editableArray(pdf, pageEdit.dict.get(PDFName.of('Annots')))
  annots.array.push(widgetRef)
  if (annots.ref) touch(annots.ref, annots.array)
  else pageEdit.dict.set(PDFName.of('Annots'), annots.array)
  touch(page.ref, pageEdit.dict)

  // AcroForm /Fields and /SigFlags, plus the catalog when the form lives inline or is new
  const catalog = pdf.catalog.clone(context)
  const acroEdit = editableDict(pdf, pdf.catalog.get(PDFName.of('AcroForm'))) ?? {
    dict: context.obj({}) as PDFDict,
  }
  const fields = editableArray(pdf, acroEdit.dict.get(PDFName.of('Fields')))
  fields.array.push(widgetRef)
  if (fields.ref) touch(fields.ref, fields.array)
  else acroEdit.dict.set(PDFName.of('Fields'), fields.array)
  acroEdit.dict.set(PDFName.of('SigFlags'), PDFNumber.of(3))
  if (acroEdit.ref) touch(acroEdit.ref, acroEdit.dict)
  else catalog.set(PDFName.of('AcroForm'), acroEdit.dict)
  if (options.certifyLevel) {
    catalog.set(PDFName.of('Perms'), context.obj({ DocMDP: sigRef }))
  }
  if (!acroEdit.ref || options.certifyLevel) touch(catalogRef, catalog)

  // Widget: unique field name, locked, printable
  const taken = new Set(existing.map((f) => f.name))
  let counter = existing.length + 1
  while (taken.has(`Signature${counter}`)) counter++
  const widget = context.obj({
    Type: 'Annot',
    Subtype: 'Widget',
    FT: 'Sig',
    T: PDFString.of(`Signature${counter}`),
    V: sigRef,
    F: 132,
    Rect: rect,
    P: page.ref,
  }) as PDFDict
  if (appearanceRef) {
    widget.set(PDFName.of('AP'), context.obj({ N: appearanceRef }))
  }
  touch(widgetRef, widget)

  const signingTime = options.signingTime ?? new Date()
  const placeholder = '0'.repeat(placeholderHex)
  const byteRangeText = `[0 ${'0'.repeat(BYTE_RANGE_DIGITS)} ${'0'.repeat(BYTE_RANGE_DIGITS)} ${'0'.repeat(BYTE_RANGE_DIGITS)}]`
  const reference = options.certifyLevel
    ? ` /Reference [<< /Type /SigRef /TransformMethod /DocMDP /DigestMethod /SHA256 /TransformParams << /Type /TransformParams /P ${options.certifyLevel} /V /1.2 >> >>]`
    : ''
  const optional = (key: string, value?: string) => (value ? ` /${key} ${pdfText(value)}` : '')
  const sigPrefix = latin1(
    `<< /Type /Sig /Filter /Adobe.PPKLite /SubFilter /adbe.pkcs7.detached${reference}` +
      optional('Reason', options.reason) +
      optional('Location', options.location) +
      optional('ContactInfo', options.contactInfo) +
      ` /M (${pdfDate(signingTime)}) /ByteRange `,
  )

  // New objects pdf-lib registered (image, appearance) follow the hand-built ones
  const created: { ref: PDFRef; object: PDFObject }[] = []
  for (const [ref, object] of context.enumerateIndirectObjects()) {
    if (ref.objectNumber >= firstNew && ref !== sigRef && ref !== widgetRef)
      created.push({ ref, object })
  }

  const parts: Buffer[] = [original, latin1(original[original.length - 1] === 0x0a ? '' : '\n')]
  let length = parts[0]!.length + parts[1]!.length
  const offsets: { num: number; offset: number; gen: number }[] = []
  const push = (buffer: Buffer) => {
    parts.push(buffer)
    length += buffer.length
  }
  const pushObject = (ref: PDFRef, object: PDFObject) => {
    offsets.push({ num: ref.objectNumber, offset: length, gen: ref.generationNumber })
    push(serialize(ref, object))
  }
  for (const { ref, object } of modified.values()) {
    if (ref !== widgetRef) pushObject(ref, object)
  }
  for (const { ref, object } of created) pushObject(ref, object)
  pushObject(widgetRef, widget)

  // The signature dictionary, laid out by hand so the placeholders have known positions
  offsets.push({ num: sigRef.objectNumber, offset: length, gen: 0 })
  push(latin1(`${sigRef.objectNumber} 0 obj\n`))
  push(sigPrefix)
  const byteRangeAt = length
  push(latin1(byteRangeText))
  push(latin1(' /Contents <'))
  const contentsStart = length - 1
  push(latin1(placeholder))
  push(latin1('>'))
  const contentsEnd = length
  push(latin1(' >>\nendobj\n'))

  // Cross-reference section and trailer
  const size = Math.max(context.largestObjectNumber, ...offsets.map((o) => o.num)) + 1
  const trailer = new Map<string, string>()
  trailer.set('Root', `${catalogRef.objectNumber} ${catalogRef.generationNumber} R`)
  const info = context.trailerInfo.Info
  if (info instanceof PDFRef) trailer.set('Info', `${info.objectNumber} ${info.generationNumber} R`)
  const id = context.trailerInfo.ID
  if (id instanceof PDFArray) trailer.set('ID', id.toString())
  trailer.set('Prev', String(previous.offset))

  const xrefOffset = length
  if (previous.isStream) {
    const xrefNum = size
    const all = [...offsets, { num: xrefNum, offset: xrefOffset, gen: 0 }]
    const { index, data } = xrefStreamData(all)
    const body = deflateSync(data)
    const dict =
      `<< /Type /XRef /Size ${xrefNum + 1} /W [1 4 2] /Index [${index.join(' ')}] /Filter /FlateDecode` +
      [...trailer].map(([k, v]) => ` /${k} ${v}`).join('') +
      ` /Length ${body.length} >>`
    push(latin1(`${xrefNum} 0 obj\n${dict}\nstream\n`))
    push(body)
    push(latin1('\nendstream\nendobj\n'))
  } else {
    push(xrefTable(offsets))
    const dict = `<< /Size ${size}` + [...trailer].map(([k, v]) => ` /${k} ${v}`).join('') + ' >>'
    push(latin1(`trailer\n${dict}\n`))
  }
  push(latin1(`startxref\n${xrefOffset}\n%%EOF\n`))

  return { file: Buffer.concat(parts), contentsStart, contentsEnd, byteRangeAt }
}

/** Detached CMS over the two signed ranges; returns the DER bytes */
function createSignature(data: Buffer, loaded: LoadedCertificate, signingTime: Date): Buffer {
  const p7 = forge.pkcs7.createSignedData()
  p7.content = forge.util.createBuffer(data.toString('binary'))
  for (const cert of loaded.chain) p7.addCertificate(cert)
  p7.addSigner({
    key: loaded.privateKey,
    certificate: loaded.certificate,
    digestAlgorithm: forge.pki.oids.sha256,
    authenticatedAttributes: [
      { type: forge.pki.oids.contentType, value: forge.pki.oids.data },
      { type: forge.pki.oids.messageDigest },
      { type: forge.pki.oids.signingTime, value: signingTime as unknown as string },
    ],
  })
  p7.sign({ detached: true })
  const asn1 = p7.toAsn1()
  restoreSignerIssuer(asn1, loaded.certificate)
  return Buffer.from(forge.asn1.toDer(asn1).getBytes(), 'binary')
}

/**
 * forge rebuilds the signer's issuer name from its decoded attributes and, for UTF-8 values,
 * encodes the already-UTF-8 bytes a second time. The signature does not cover the signer
 * identifier, so put the certificate's own issuer encoding back; otherwise verifiers that
 * match issuer + serial (OpenSSL, Acrobat) cannot find the signing certificate.
 */
function restoreSignerIssuer(root: forge.asn1.Asn1, cert: forge.pki.Certificate): void {
  const A = forge.asn1
  const certAsn1 = A.fromDer(A.toDer(forge.pki.certificateToAsn1(cert)))
  const tbs = (certAsn1.value as forge.asn1.Asn1[])[0]!
  const fields = tbs.value as forge.asn1.Asn1[]
  const issuer = fields[fields[0]!.tagClass === A.Class.CONTEXT_SPECIFIC ? 3 : 2]!
  const signedData = ((root.value as forge.asn1.Asn1[])[1]!.value as forge.asn1.Asn1[])[0]!
  const signerInfos = (signedData.value as forge.asn1.Asn1[]).find(
    (node, index) => index > 1 && node.tagClass === A.Class.UNIVERSAL && node.type === A.Type.SET,
  )
  const signerInfo = (signerInfos?.value as forge.asn1.Asn1[] | undefined)?.[0]
  const idAndSerial = (signerInfo?.value as forge.asn1.Asn1[] | undefined)?.[1]
  if (idAndSerial && Array.isArray(idAndSerial.value)) idAndSerial.value[0] = issuer
}

/** Sign `bytes` with a PKCS#12 certificate; the result is the original bytes plus one signature revision */
export async function signPdf(bytes: Uint8Array, options: SignPdfOptions): Promise<Uint8Array> {
  if (bytes.byteLength > MAX_SIGNABLE_BYTES) throw new SignError('sign-failed')
  const loaded =
    options.loaded ?? loadCertificate(options.p12, options.password, options.signerSha1)
  const original = Buffer.from(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  const signingTime = options.signingTime ?? new Date()
  const withTime = { ...options, signingTime }

  for (const placeholderBytes of PLACEHOLDER_SIZES) {
    const { file, contentsStart, contentsEnd, byteRangeAt } = await assemble(
      original,
      withTime,
      placeholderBytes * 2,
    )
    const range = [0, contentsStart, contentsEnd, file.length - contentsEnd]
    const rangeText = `[${range[0]} ${String(range[1]).padStart(BYTE_RANGE_DIGITS, '0')} ${String(range[2]).padStart(BYTE_RANGE_DIGITS, '0')} ${String(range[3]).padStart(BYTE_RANGE_DIGITS, '0')}]`
    file.write(rangeText, byteRangeAt, 'latin1')
    const signed = Buffer.concat([file.subarray(0, contentsStart), file.subarray(contentsEnd)])
    const cms = createSignature(signed, loaded, signingTime)
    if (cms.length > placeholderBytes) continue
    file.write(cms.toString('hex').toUpperCase(), contentsStart + 1, 'latin1')
    return file
  }
  throw new SignError('sign-failed')
}
