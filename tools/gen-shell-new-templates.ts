/** Regenerate the checked-in Windows Explorer templates: npx tsx tools/gen-shell-new-templates.ts */
import { mkdir, writeFile } from 'node:fs/promises'
import JSZip from 'jszip'
import { buildBlankDocx } from '../packages/docx-engine/src/blank'
import { createBlankPptx } from '../packages/pptx-engine/src/blank'
import { blankXlsxBuffer } from '../packages/xlsx-gateway/src/gateway/csv-import'

async function main() {
  const output = new URL('../apps/shell/build/shell-new/', import.meta.url)
  await mkdir(output, { recursive: true })
  const templates = {
    docx: await buildBlankDocx(),
    xlsx: await blankXlsxBuffer(),
    pptx: await createBlankPptx(),
  }
  for (const [ext, bytes] of Object.entries(templates)) {
    const zip = await JSZip.loadAsync(bytes)
    // Keep regeneration deterministic when only the ZIP timestamps differ.
    zip.forEach((_, entry) => {
      entry.date = new Date('2000-01-01T00:00:00Z')
    })
    await writeFile(
      new URL(`blank.${ext}`, output),
      await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }),
    )
  }
}

void main()
