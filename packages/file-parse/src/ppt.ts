import * as CFB from 'cfb'
import './vendor'

type LegacyPptGlobals = typeof globalThis & {
  CFB?: typeof CFB
  cptable?: {
    utils: {
      decode(codepage: number, bytes: Uint8Array): string
    }
  }
}

// ppt-to-text hides these dependencies behind concatenated CommonJS requires,
// which Rollup cannot bundle. Provide them before loading the parser lazily.
const legacyGlobals = globalThis as LegacyPptGlobals
legacyGlobals.CFB ??= CFB
legacyGlobals.cptable ??= {
  utils: {
    decode(codepage, bytes) {
      if (codepage !== 1200) throw new Error(`Unsupported legacy PPT codepage: ${codepage}`)
      return Buffer.from(bytes).toString('utf16le')
    },
  },
}

function normalizeText(text: string): string {
  return text.replace(/\r\n?/g, '\n').trim()
}

/** Extract one readable text section per slide from a legacy PowerPoint 97-2003 file. */
export async function pptToText(bytes: Uint8Array): Promise<string> {
  const pptParser = (await import('ppt-to-text')).default
  const presentation = pptParser.readBuffer(Buffer.from(bytes))
  const slides = pptParser.utils.to_text(presentation)
  return slides
    .map((text, index) => {
      const content = normalizeText(text)
      return content ? `## Slide ${index + 1}\n${content}` : `## Slide ${index + 1}`
    })
    .join('\n\n')
}
