import { randomBytes } from 'node:crypto'
import { writeFileSync } from 'node:fs'
import { extname } from 'node:path'
import { fileURLToPath } from 'node:url'
import { generateImageTool } from '@genoffice/ai-search'
import { fetchRemoteImage } from '@genoffice/electron-utils/remote-image'
import { flagBool, flagString } from '../args'
import { aiSettingsPath, prepareCloud } from '../cloud'
import { resolveInput, resolveOutput } from '../fs'
import type { CommandDef } from '../registry'
import { CliError, EXIT } from '../result'

const ASPECTS = ['1:1', '4:3', '16:9', '9:16', '3:4', '2:3', '3:2', 'auto']
const SIZES = ['auto', '0.5k', '1k', '2k', '3k', '4k']

const EXTS_BY_MIME: Record<string, readonly string[]> = {
  'image/png': ['png'],
  'image/jpeg': ['jpg', 'jpeg'],
  'image/webp': ['webp'],
  'image/gif': ['gif'],
}

export const imageCommand: CommandDef = {
  name: 'image',
  summary:
    'Generate an image from a prompt with the configured image provider and save it as a file.',
  usage: 'image <prompt> [--out <file>] [--aspect 16:9] [--size 1k] [--ref <image>]... [--force]',
  options: [
    {
      name: 'out',
      value: 'file',
      description: 'where to save (default: generated-image-<time>.<ext> in cwd)',
    },
    {
      name: 'aspect',
      value: 'ratio',
      description: '1:1 | 4:3 | 16:9 | 9:16 | 3:4 | 2:3 | 3:2 | auto',
    },
    { name: 'size', value: 'size', description: 'auto | 0.5k | 1k | 2k | 3k | 4k (Genspark only)' },
    {
      name: 'ref',
      value: 'images',
      description: 'reference or edit-target images (paths or URLs), comma-separated',
    },
    { name: 'model', value: 'name', description: 'Genspark model override (e.g. fal-bria-rmbg)' },
    { name: 'force', description: 'overwrite an existing output file' },
  ],
  async run(args, ctx) {
    const prompt = args.positionals.join(' ').trim()
    if (!prompt) throw new CliError(EXIT.usage, 'missing <prompt>')
    const aspect = flagString(args, 'aspect')
    if (aspect && !ASPECTS.includes(aspect)) {
      throw new CliError(EXIT.usage, `--aspect must be one of ${ASPECTS.join(', ')}`)
    }
    const size = flagString(args, 'size')
    if (size && !SIZES.includes(size)) {
      throw new CliError(EXIT.usage, `--size must be one of ${SIZES.join(', ')}`)
    }
    const refs = (flagString(args, 'ref') ?? '')
      .split(',')
      .map((r) => r.trim())
      .filter(Boolean)
      // only http(s) is remote; file:// and plain paths go through the path policy before upload
      .map((r) =>
        /^https?:\/\//i.test(r)
          ? r
          : resolveInput(r.startsWith('file:') ? fileURLToPath(r) : r, ctx),
      )
    // resolve (policy, --force) the output before spending generation credits; the
    // default name is checked as .png here and re-checked under the real extension
    const out = flagString(args, 'out')
    const force = flagBool(args, 'force')
    const stem = out ?? `generated-image-${stamp()}`
    const outExt = out ? extname(out).slice(1).toLowerCase() : ''
    // an extension-less --out is a stem: the file takes the provider's real extension
    const chosen = resolveOutput(out && outExt ? out : `${stem}.png`, ctx, { force, fresh: true })
    // the provider may answer in another format than --out names; every sibling the
    // file could be saved under is checked now, so nothing fails after credits are spent
    const chosenExt = outExt || 'png'
    const base = chosen.slice(0, -chosenExt.length)
    for (const ext of siblingExtensions(chosenExt)) {
      resolveOutput(`${base}${ext}`, ctx, { force, fresh: true })
    }
    await prepareCloud(ctx.env)
    const r = await generateImageTool(aiSettingsPath(ctx.env), {
      prompt,
      aspectRatio: aspect,
      imageSize: size,
      model: flagString(args, 'model'),
      ...(refs.length ? { referenceImageUrls: refs } : {}),
    })
    if (!r.url) throw new CliError(EXIT.app, r.error ?? 'image generation failed')
    const image = await loadImage(r.url)
    const ext = EXTS_BY_MIME[image.mime]?.[0] ?? 'png'
    // the provider picks the encoding; a .png name holding JPEG bytes would mislead every reader,
    // so the file takes the real extension and output_path says where it went
    const renamed = Boolean(out && outExt && !(EXTS_BY_MIME[image.mime] ?? []).includes(outExt))
    const output = renamed
      ? resolveOutput(`${chosen.slice(0, -outExt.length)}${ext}`, ctx, { force, fresh: true })
      : (out && outExt) || ext === 'png'
        ? chosen
        : resolveOutput(`${stem}.${ext}`, ctx, { force, fresh: true })
    writeFileSync(output, image.bytes)
    return {
      summary: `saved ${image.bytes.byteLength} bytes (${image.mime}) to ${output}`,
      outputPath: output,
      detail: {
        mime: image.mime,
        bytes: image.bytes.byteLength,
        source_url: r.url,
        ...(renamed
          ? { note: `provider returned ${image.mime}; saved with .${ext} instead of .${outExt}` }
          : {}),
      },
    }
  },
}

/** Extensions the saved file could end up with when the provider answers in another format than `outExt` names. */
export function siblingExtensions(outExt: string): string[] {
  const outMime = Object.entries(EXTS_BY_MIME).find(([, exts]) => exts.includes(outExt))?.[0]
  return Object.entries(EXTS_BY_MIME)
    .filter(([mime]) => mime !== outMime)
    .map(([, exts]) => exts[0]!)
}

/** Genspark returns an https URL, BYOK providers a file:// in the app's generated-image store; fetchRemoteImage serves both. */
async function loadImage(url: string): Promise<{ bytes: Uint8Array; mime: string }> {
  const response = await fetchRemoteImage(url)
  if (!response?.ok) throw new CliError(EXIT.app, `could not download the generated image: ${url}`)
  const mime = response.headers.get('content-type')?.split(';')[0]?.trim() || 'image/png'
  return { bytes: new Uint8Array(await response.arrayBuffer()), mime }
}

/** yyyymmdd-hhmmssmmm plus a random tail, so two runs in the same instant do not collide */
function stamp(): string {
  return (
    new Date().toISOString().replace(/[-:.]/g, '').replace('Z', '').replace('T', '-') +
    '-' +
    randomBytes(2).toString('hex')
  )
}
