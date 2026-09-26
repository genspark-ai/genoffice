import { mkdtempSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { runCli, type RunOptions } from '../src/cli'

export interface CapturedRun {
  code: number
  stdout: string
  stderr: string
  json: () => any
}

export async function run(argv: string[], opts: Omit<RunOptions, 'io'> = {}): Promise<CapturedRun> {
  const out: string[] = []
  const err: string[] = []
  const code = await runCli(argv, {
    ...opts,
    io: { stdout: (t) => out.push(t), stderr: (t) => err.push(t) },
  })
  return {
    code,
    stdout: out.join('\n'),
    stderr: err.join('\n'),
    json: () => JSON.parse(out.join('\n')),
  }
}

export function tempDir(): string {
  return mkdtempSync(join(tmpdir(), 'genoffice-test-'))
}

/** A valid PDF with real Helvetica text, one page per entry; enough for page counting, text reading and conversion. */
export function writeMinimalPdf(
  path: string,
  text: string | string[] = 'Hello genoffice',
  info: { title?: string; author?: string } = {},
): string {
  const texts = Array.isArray(text) ? text : [text]
  const kids = texts.map((_, i) => `${4 + 2 * i} 0 R`).join(' ')
  const objects = [
    '<< /Type /Catalog /Pages 2 0 R >>',
    `<< /Type /Pages /Kids [${kids}] /Count ${texts.length} >>`,
    '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>',
  ]
  texts.forEach((t, i) => {
    const content = `BT /F1 24 Tf 72 700 Td (${t}) Tj ET`
    objects.push(
      `<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Contents ${5 + 2 * i} 0 R /Resources << /Font << /F1 3 0 R >> >> >>`,
      `<< /Length ${content.length} >>\nstream\n${content}\nendstream`,
    )
  })
  const entries = Object.entries(info).map(
    ([k, v]) => `/${k[0]!.toUpperCase()}${k.slice(1)} (${v})`,
  )
  if (entries.length) objects.push(`<< ${entries.join(' ')} >>`)
  const infoRef = entries.length ? ` /Info ${objects.length} 0 R` : ''
  let body = '%PDF-1.4\n'
  const offsets: number[] = []
  objects.forEach((obj, i) => {
    offsets.push(body.length)
    body += `${i + 1} 0 obj\n${obj}\nendobj\n`
  })
  const xref = body.length
  body += `xref\n0 ${objects.length + 1}\n0000000000 65535 f \n`
  for (const off of offsets) body += `${String(off).padStart(10, '0')} 00000 n \n`
  body += `trailer\n<< /Size ${objects.length + 1} /Root 1 0 R${infoRef} >>\nstartxref\n${xref}\n%%EOF\n`
  writeFileSync(path, body, 'latin1')
  return path
}
