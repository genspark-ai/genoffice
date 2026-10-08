import { X509Certificate, createHash, verify as cryptoVerify } from 'node:crypto'
import type { KeyObject } from 'node:crypto'
import tls from 'node:tls'
import forge from 'node-forge'
import {
  PDFArray,
  PDFDict,
  PDFDocument,
  PDFHexString,
  PDFName,
  PDFNumber,
  PDFRef,
  PDFString,
} from 'pdf-lib'
import type { PDFObject } from 'pdf-lib'
import type { CertificateInfo, PdfSignatureInfo, PdfSignatureProblem } from '../shared/ipc'

const A = forge.asn1

type Asn1 = forge.asn1.Asn1

const OID_SIGNED_DATA = '1.2.840.113549.1.7.2'
const OID_MESSAGE_DIGEST = '1.2.840.113549.1.9.4'
const OID_SIGNING_TIME = '1.2.840.113549.1.9.5'

const DIGEST_OIDS: Record<string, string> = {
  '1.3.14.3.2.26': 'sha1',
  '2.16.840.1.101.3.4.2.4': 'sha224',
  '2.16.840.1.101.3.4.2.1': 'sha256',
  '2.16.840.1.101.3.4.2.2': 'sha384',
  '2.16.840.1.101.3.4.2.3': 'sha512',
}

/** Signature algorithms crypto.verify can check with a plain (hash, key) pair: RSA PKCS#1 v1.5 and ECDSA */
const RSA_OIDS = new Set([
  '1.2.840.113549.1.1.1',
  '1.2.840.113549.1.1.5',
  '1.2.840.113549.1.1.11',
  '1.2.840.113549.1.1.12',
  '1.2.840.113549.1.1.13',
  '1.2.840.113549.1.1.14',
])
const ECDSA_OIDS = new Set([
  '1.2.840.10045.2.1',
  '1.2.840.10045.4.1',
  '1.2.840.10045.4.3.1',
  '1.2.840.10045.4.3.2',
  '1.2.840.10045.4.3.3',
  '1.2.840.10045.4.3.4',
])

/** Signature dictionaries are plain (never inside an object stream), so the raw marker is a cheap pre-filter */
const BYTE_RANGE_MARKER = Buffer.from('/ByteRange')

const lookupDict = (dict: PDFDict, key: string): PDFDict | undefined =>
  dict.lookupMaybe(PDFName.of(key), PDFDict)

const lookupText = (dict: PDFDict, key: string): string | undefined => {
  const value = dict.lookup(PDFName.of(key))
  return value instanceof PDFString || value instanceof PDFHexString
    ? value.decodeText()
    : undefined
}

// ── certificates ────────────────────────────────────────────────────────────

/**
 * Node renders the UTF-8 bytes of a distinguished name one byte per character, so a CJK common
 * name arrives as mojibake. Re-decode when every character is a byte value and the bytes form
 * valid UTF-8; anything else is already text and stays as is.
 */
function fixDn(dn: string): string {
  if ([...dn].some((ch) => ch.charCodeAt(0) > 0xff)) return dn
  const decoded = Buffer.from(dn, 'latin1').toString('utf8')
  return decoded.includes('\ufffd') ? dn : decoded
}

const dnLines = (dn: string): string[] => dn.split('\n').filter(Boolean)

const commonNameOf = (dn: string): string => {
  const cn = dnLines(dn).find((line) => line.startsWith('CN='))
  return cn ? cn.slice(3) : dnLines(dn).join(', ')
}

const keyAlgorithmOf = (key: KeyObject): string => {
  const type = key.asymmetricKeyType ?? 'unknown'
  const details = key.asymmetricKeyDetails
  if (type === 'rsa') return `RSA ${details?.modulusLength ?? ''}`.trim()
  if (type === 'ec') return `EC ${details?.namedCurve ?? ''}`.trim()
  return type.toUpperCase()
}

const isSelfSigned = (cert: X509Certificate): boolean => {
  try {
    // Name match plus a valid signature by its own key. X509Certificate.checkIssued is stricter
    // (it also wants keyCertSign), and rejects the self-signed certificates Windows generates.
    return cert.subject === cert.issuer && cert.verify(cert.publicKey)
  } catch {
    return false
  }
}

export function describeCertificate(cert: X509Certificate): CertificateInfo {
  const subject = fixDn(cert.subject)
  const issuer = fixDn(cert.issuer)
  return {
    commonName: commonNameOf(subject),
    subject,
    issuerCommonName: commonNameOf(issuer),
    issuer,
    serialNumber: cert.serialNumber,
    validFrom: new Date(cert.validFrom).toISOString(),
    validTo: new Date(cert.validTo).toISOString(),
    fingerprint: cert.fingerprint256,
    keyAlgorithm: keyAlgorithmOf(cert.publicKey),
    selfSigned: isSelfSigned(cert),
  }
}

let trustedFingerprints: Set<string> | null = null

/** Roots the OS trusts plus Node's bundled Mozilla set; computed once */
function trustedRootFingerprints(): Set<string> {
  if (trustedFingerprints) return trustedFingerprints
  const pems: string[] = [...tls.rootCertificates]
  try {
    // Node >= 22.15 can read the operating-system store (the one the user manages)
    const getter = (tls as unknown as { getCACertificates?: (kind: string) => string[] })
      .getCACertificates
    if (getter) pems.push(...getter('system'))
  } catch {
    /* the bundled set alone is still a sound baseline */
  }
  const out = new Set<string>()
  for (const pem of pems) {
    try {
      out.add(new X509Certificate(pem).fingerprint256)
    } catch {
      /* skip unparsable store entries */
    }
  }
  trustedFingerprints = out
  return out
}

/** Order the pool into a chain starting at `signer`, following issuer links that verify cryptographically */
function buildChain(signer: X509Certificate, pool: X509Certificate[]): X509Certificate[] {
  const chain = [signer]
  let current = signer
  for (let depth = 0; depth < 10; depth++) {
    if (isSelfSigned(current)) break
    const next = pool.find((candidate) => {
      if (chain.includes(candidate)) return false
      try {
        return current.issuer === candidate.subject && current.verify(candidate.publicKey)
      } catch {
        return false
      }
    })
    if (!next) break
    chain.push(next)
    current = next
  }
  return chain
}

const chainIsTrusted = (chain: X509Certificate[]): boolean => {
  const roots = trustedRootFingerprints()
  return chain.some((cert) => roots.has(cert.fingerprint256))
}

// ── PKCS#7 / CMS ────────────────────────────────────────────────────────────

interface ParsedSignedData {
  certificates: X509Certificate[]
  certificateIssuerDer: string[]
  digestOid: string
  signatureOid: string
  signature: Buffer
  /** DER of the signed attributes re-tagged as a SET, ready to verify against; absent when none */
  signedAttributesDer?: Buffer
  messageDigest?: Buffer
  signingTime?: Date
  signerSerialHex: string
  signerIssuerDer: string
}

const isContext = (node: Asn1, type: number): boolean =>
  node.tagClass === A.Class.CONTEXT_SPECIFIC && node.type === type

const children = (node: Asn1): Asn1[] => (Array.isArray(node.value) ? node.value : [])

const derOf = (node: Asn1): string => A.toDer(node).getBytes()

const oidOf = (node: Asn1): string => A.derToOid(node.value as string)

const stripLeadingZeros = (hex: string): string => hex.replace(/^0+/, '') || '0'

/** Length of the first DER element in `bytes`; trailing zero padding of /Contents must not reach the parser */
function derElementLength(bytes: Buffer): number {
  if (bytes.length < 2) throw new Error('short der')
  const first = bytes[1]!
  if (first < 0x80) return 2 + first
  const lengthBytes = first & 0x7f
  if (lengthBytes === 0 || lengthBytes > 4 || bytes.length < 2 + lengthBytes) {
    throw new Error('bad der length')
  }
  let length = 0
  for (let i = 0; i < lengthBytes; i++) length = length * 256 + bytes[2 + i]!
  return 2 + lengthBytes + length
}

function parseSignedData(contents: Buffer): ParsedSignedData {
  const total = derElementLength(contents)
  if (total > contents.length) throw new Error('truncated')
  const root = A.fromDer(forge.util.createBuffer(contents.subarray(0, total).toString('binary')))
  const [contentType, wrapper] = children(root)
  if (!contentType || !wrapper || oidOf(contentType) !== OID_SIGNED_DATA) {
    throw new Error('not signed data')
  }
  const signedData = children(wrapper)[0]
  if (!signedData) throw new Error('empty signed data')

  const certificates: X509Certificate[] = []
  const certificateIssuerDer: string[] = []
  let signerInfo: Asn1 | undefined
  for (const node of children(signedData)) {
    if (isContext(node, 0)) {
      for (const certNode of children(node)) {
        if (certNode.tagClass !== A.Class.UNIVERSAL || certNode.type !== A.Type.SEQUENCE) continue
        try {
          certificates.push(new X509Certificate(Buffer.from(derOf(certNode), 'binary')))
          // tbsCertificate: [version]? serial sigAlg issuer …
          const tbs = children(certNode)[0]!
          const fields = children(tbs)
          const offset = isContext(fields[0]!, 0) ? 1 : 0
          certificateIssuerDer.push(derOf(fields[offset + 2]!))
        } catch {
          /* an unreadable certificate in the bag is skipped; the signer lookup reports it if needed */
        }
      }
    } else if (
      node.tagClass === A.Class.UNIVERSAL &&
      node.type === A.Type.SET &&
      node !== children(signedData)[1]
    ) {
      signerInfo = children(node)[0]
    }
  }
  if (!signerInfo) throw new Error('no signer info')

  const parts = children(signerInfo)
  const sid = parts[1]!
  const sidParts = children(sid)
  if (sidParts.length < 2) throw new Error('unsupported signer identifier')
  const digestOid = oidOf(children(parts[2]!)[0]!)

  let index = 3
  let signedAttributesDer: Buffer | undefined
  let messageDigest: Buffer | undefined
  let signingTime: Date | undefined
  if (parts[index] && isContext(parts[index]!, 0)) {
    const attrs = parts[index]!
    const asSet = A.create(A.Class.UNIVERSAL, A.Type.SET, true, attrs.value as Asn1[])
    signedAttributesDer = Buffer.from(derOf(asSet), 'binary')
    for (const attr of children(attrs)) {
      const [attrOid, valueSet] = children(attr)
      if (!attrOid || !valueSet) continue
      const value = children(valueSet)[0]
      if (!value) continue
      const id = oidOf(attrOid)
      if (id === OID_MESSAGE_DIGEST) messageDigest = Buffer.from(value.value as string, 'binary')
      else if (id === OID_SIGNING_TIME) {
        signingTime =
          value.type === A.Type.UTCTIME
            ? A.utcTimeToDate(value.value as string)
            : A.generalizedTimeToDate(value.value as string)
      }
    }
    index++
  }
  const signatureOid = oidOf(children(parts[index]!)[0]!)
  const signature = Buffer.from(parts[index + 1]!.value as string, 'binary')

  return {
    certificates,
    certificateIssuerDer,
    digestOid,
    signatureOid,
    signature,
    signedAttributesDer,
    messageDigest,
    signingTime,
    signerSerialHex: forge.util.bytesToHex(sidParts[1]!.value as string).toUpperCase(),
    signerIssuerDer: derOf(sidParts[0]!),
  }
}

function findSigner(parsed: ParsedSignedData): X509Certificate | undefined {
  const serial = stripLeadingZeros(parsed.signerSerialHex)
  const matches = parsed.certificates
    .map((cert, i) => ({ cert, issuerDer: parsed.certificateIssuerDer[i] }))
    .filter(({ cert }) => stripLeadingZeros(cert.serialNumber.toUpperCase()) === serial)
  return matches.find((m) => m.issuerDer === parsed.signerIssuerDer)?.cert ?? matches[0]?.cert
}

// ── per-signature verification ──────────────────────────────────────────────

interface FieldWithValue {
  name: string
  value: PDFDict
  widget: PDFDict
  widgetRef?: PDFRef
}

/** Every signature field that has a value, depth-first through the AcroForm tree */
export function collectSignatureFields(pdf: PDFDocument): FieldWithValue[] {
  const out: FieldWithValue[] = []
  const acroForm = pdf.catalog.lookupMaybe(PDFName.of('AcroForm'), PDFDict)
  const fields = acroForm?.lookupMaybe(PDFName.of('Fields'), PDFArray)
  if (!fields) return out
  const seen = new Set<PDFDict>()

  const visit = (node: PDFObject, parentName: string, inheritedType: string | undefined) => {
    const dict = pdf.context.lookupMaybe(node, PDFDict)
    if (!dict || seen.has(dict)) return
    seen.add(dict)
    const partial = lookupText(dict, 'T')
    const name = partial ? (parentName ? `${parentName}.${partial}` : partial) : parentName
    const ft = dict.lookup(PDFName.of('FT'))
    const type = ft instanceof PDFName ? ft.decodeText() : inheritedType
    const value = lookupDict(dict, 'V')
    if (type === 'Sig' && value && value.has(PDFName.of('ByteRange'))) {
      out.push({ name, value, widget: dict, widgetRef: node instanceof PDFRef ? node : undefined })
    }
    const kids = dict.lookupMaybe(PDFName.of('Kids'), PDFArray)
    if (kids) for (const kid of kids.asArray()) visit(kid, name, type)
  }
  for (const field of fields.asArray()) visit(field, '', undefined)
  return out
}

const decodeByteRange = (value: PDFObject | undefined): number[] | null => {
  if (!(value instanceof PDFArray) || value.size() !== 4) return null
  const numbers = value
    .asArray()
    .map((entry) => (entry instanceof PDFNumber ? entry.asNumber() : NaN))
  return numbers.every((n) => Number.isInteger(n) && n >= 0) ? numbers : null
}

function certifiedLevelOf(value: PDFDict): 1 | 2 | 3 | undefined {
  const references = value.lookupMaybe(PDFName.of('Reference'), PDFArray)
  if (!references) return undefined
  for (const entry of references.asArray()) {
    const ref = value.context.lookupMaybe(entry, PDFDict)
    const method = ref?.lookup(PDFName.of('TransformMethod'))
    if (!ref || !(method instanceof PDFName) || method.decodeText() !== 'DocMDP') continue
    const params = lookupDict(ref, 'TransformParams')
    const level = params?.lookup(PDFName.of('P'))
    const p = level instanceof PDFNumber ? level.asNumber() : 2
    return p === 1 || p === 3 ? p : 2
  }
  return undefined
}

const pdfDateToIso = (raw: string | undefined): string | undefined => {
  const m = /^D:(\d{4})(\d{2})?(\d{2})?(\d{2})?(\d{2})?(\d{2})?(Z|[+-]\d{2}'?(?:\d{2}'?)?)?/.exec(
    raw ?? '',
  )
  if (!m) return undefined
  const [, y, mo = '01', d = '01', h = '00', mi = '00', s = '00', tz = 'Z'] = m
  let offset = 'Z'
  if (tz !== 'Z' && tz !== '') {
    const digits = tz.replace(/'/g, '')
    offset = `${digits.slice(0, 3)}:${digits.slice(3, 5) || '00'}`
  }
  const date = new Date(`${y}-${mo}-${d}T${h}:${mi}:${s}${offset}`)
  return Number.isNaN(date.getTime()) ? undefined : date.toISOString()
}

const statusOf = (problems: PdfSignatureProblem[]): PdfSignatureInfo['status'] => {
  if (
    problems.some((p) =>
      ['digest-mismatch', 'signature-invalid', 'malformed', 'signer-certificate-missing'].includes(
        p,
      ),
    )
  ) {
    return 'invalid'
  }
  // Who the signer is (self-signed, unknown issuer) is a note, not a failure: the signature is
  // still cryptographically intact, so it counts as valid. Only things that weaken what the
  // signature proves — later changes, an out-of-date certificate, an unchecked algorithm — warn.
  const weakens: PdfSignatureProblem[] = [
    'document-changed-after-signing',
    'certificate-expired',
    'certificate-not-yet-valid',
    'unsupported-algorithm',
  ]
  return problems.some((p) => weakens.includes(p)) ? 'warning' : 'valid'
}

function verifyOne(bytes: Buffer, field: FieldWithValue, pageIndex: number): PdfSignatureInfo {
  const { value } = field
  const subFilterName = value.lookup(PDFName.of('SubFilter'))
  const subFilter = subFilterName instanceof PDFName ? subFilterName.decodeText() : ''
  const typeName = value.lookup(PDFName.of('Type'))
  const kind: PdfSignatureInfo['kind'] =
    subFilter === 'ETSI.RFC3161' ||
    (typeName instanceof PDFName && typeName.decodeText() === 'DocTimeStamp')
      ? 'timestamp'
      : 'signature'

  const claimedTime = pdfDateToIso(lookupText(value, 'M'))
  const info: PdfSignatureInfo = {
    fieldName: field.name,
    pageIndex,
    kind,
    status: 'invalid',
    problems: [],
    subFilter,
    signerName: lookupText(value, 'Name'),
    reason: lookupText(value, 'Reason'),
    location: lookupText(value, 'Location'),
    contactInfo: lookupText(value, 'ContactInfo'),
    signingTime: claimedTime,
    signingTimeSigned: false,
    coversWholeDocument: false,
    certifiedLevel: certifiedLevelOf(value),
    chain: [],
    trusted: false,
  }
  const fail = (problem: PdfSignatureProblem): PdfSignatureInfo => {
    info.problems = [problem]
    info.status = statusOf(info.problems)
    return info
  }

  const byteRange = decodeByteRange(value.lookup(PDFName.of('ByteRange')))
  const contentsObj = value.lookup(PDFName.of('Contents'))
  if (!byteRange || !(contentsObj instanceof PDFHexString || contentsObj instanceof PDFString)) {
    return fail('malformed')
  }
  const [start1, length1, start2, length2] = byteRange as [number, number, number, number]
  const end1 = start1 + length1
  const end2 = start2 + length2
  // The two ranges must hug the /Contents string exactly: any other hole is unsigned content
  // an attacker could swap (signature wrapping), so it is rejected rather than trusted.
  if (start1 !== 0 || end1 > start2 || end2 > bytes.length) return fail('malformed')
  if (bytes[end1] !== 0x3c || bytes[start2 - 1] !== 0x3e) return fail('malformed')
  const trailing = bytes.subarray(end2)
  info.coversWholeDocument = trailing.every((b) => b === 0x0a || b === 0x0d || b === 0x20)

  if (kind === 'timestamp') {
    // Verifying a TSTInfo needs the timestamp-token profile; show it, but never claim it checked out
    info.problems = ['unsupported-algorithm']
    info.status = 'warning'
    return info
  }

  let parsed: ParsedSignedData
  try {
    parsed = parseSignedData(Buffer.from(contentsObj.asBytes()))
  } catch {
    return fail('malformed')
  }
  info.signingTime = parsed.signingTime?.toISOString() ?? claimedTime
  info.signingTimeSigned = parsed.signingTime !== undefined

  const hashName = DIGEST_OIDS[parsed.digestOid]
  info.digestAlgorithm = hashName?.toUpperCase()
  const signer = findSigner(parsed)
  const problems: PdfSignatureProblem[] = []
  if (!signer) return fail('signer-certificate-missing')

  const chain = buildChain(signer, parsed.certificates)
  info.signer = describeCertificate(signer)
  info.chain = chain.map(describeCertificate)
  info.trusted = chainIsTrusted(chain)

  const supported =
    hashName !== undefined &&
    (RSA_OIDS.has(parsed.signatureOid) || ECDSA_OIDS.has(parsed.signatureOid))
  if (!supported) {
    info.problems = ['unsupported-algorithm']
    info.status = 'warning'
    return info
  }

  const range1 = bytes.subarray(start1, end1)
  const range2 = bytes.subarray(start2, end2)
  const digest = createHash(hashName).update(range1).update(range2).digest()
  let intact = true
  let signedBytes: Buffer
  if (parsed.signedAttributesDer) {
    if (!parsed.messageDigest || !digest.equals(parsed.messageDigest)) intact = false
    signedBytes = parsed.signedAttributesDer
  } else {
    signedBytes = Buffer.concat([range1, range2])
  }
  if (!intact) problems.push('digest-mismatch')
  else {
    let signatureOk: boolean
    try {
      signatureOk = cryptoVerify(hashName, signedBytes, signer.publicKey, parsed.signature)
    } catch {
      signatureOk = false
    }
    if (!signatureOk) problems.push('signature-invalid')
  }

  if (problems.length === 0) {
    if (!info.coversWholeDocument) problems.push('document-changed-after-signing')
    const at = parsed.signingTime ?? new Date()
    if (at.getTime() < new Date(signer.validFrom).getTime())
      problems.push('certificate-not-yet-valid')
    else if (at.getTime() > new Date(signer.validTo).getTime()) problems.push('certificate-expired')
    if (!info.trusted)
      problems.push(
        info.chain[info.chain.length - 1]?.selfSigned && chain.length === 1
          ? 'self-signed'
          : 'untrusted-issuer',
      )
  }
  info.problems = problems
  info.status = statusOf(problems)
  return info
}

/** Page index of each widget, by object identity; fields whose widget has no /P fall back to scanning page /Annots */
function pageIndexOf(pdf: PDFDocument, field: FieldWithValue): number {
  const pages = pdf.getPages()
  const direct = field.widget.get(PDFName.of('P'))
  if (direct instanceof PDFRef) {
    const index = pages.findIndex((page) => page.ref === direct)
    if (index >= 0) return index
  }
  if (!field.widgetRef) return -1
  for (let i = 0; i < pages.length; i++) {
    const annots = pages[i]!.node.lookupMaybe(PDFName.of('Annots'), PDFArray)
    if (annots?.asArray().some((entry) => entry === field.widgetRef)) return i
  }
  return -1
}

/**
 * Read and check every certificate signature in a PDF. Nothing the file says about itself is
 * trusted: integrity comes from re-hashing the raw bytes the signature's /ByteRange names.
 * Encrypted or unparsable files report no signatures rather than throwing.
 */
export async function readPdfSignatures(input: Uint8Array): Promise<PdfSignatureInfo[]> {
  const bytes = Buffer.from(input.buffer, input.byteOffset, input.byteLength)
  if (!bytes.includes(BYTE_RANGE_MARKER)) return []
  let pdf: PDFDocument
  try {
    pdf = await PDFDocument.load(bytes, { updateMetadata: false, throwOnInvalidObject: false })
  } catch {
    return []
  }
  return collectSignatureFields(pdf).map((field) =>
    verifyOne(bytes, field, pageIndexOf(pdf, field)),
  )
}
