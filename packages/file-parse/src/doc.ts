import WordExtractor from 'word-extractor'

export async function docToText(buffer: Buffer): Promise<string> {
  const extractor = new WordExtractor()
  const doc = await extractor.extract(buffer)
  const parts: string[][] = []

  const body = doc.getBody()
  if (body) parts.push([body])

  const footnotes = doc.getFootnotes()
  if (footnotes) parts.push(['--- Footnotes ---', footnotes])

  const endnotes = doc.getEndnotes()
  if (endnotes) parts.push(['--- Endnotes ---', endnotes])

  const headers = doc.getHeaders()
  if (headers) parts.push(['--- Headers ---', headers])

  return parts.map((p) => p.join('\n')).join('\n\n')
}
