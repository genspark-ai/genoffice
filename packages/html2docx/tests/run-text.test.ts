import assert from 'node:assert/strict'
import { TabStopType } from 'docx'
import JSZip from 'jszip'
import { test } from 'vitest'

import { generateDocx } from '../src/generate'
import { tabStopsFor } from '../src/generate/word-utils'

test('coerces a non-string run text instead of throwing', async () => {
  // Tab-stop path: r.text.includes ran straight into a non-string.
  assert.deepEqual(tabStopsFor({ contentDxa: 9000 }, [{ text: 42 }]), [])
  assert.deepEqual(tabStopsFor({ contentDxa: 9000 }, [{ text: 'Total\t12' }]), [
    { type: TabStopType.RIGHT, position: 9000 },
  ])

  // Run path: (r.text || '').split threw on the same value. A para reaches
  // tabStopsFor first, so the whole conversion died before any run was built.
  const buffer = await generateDocx([{ type: 'para', runs: [{ text: 42 }, { text: ' tail' }] }], {})
  const zip = await JSZip.loadAsync(buffer)
  const xml = (await zip.file('word/document.xml')?.async('string')) ?? ''
  assert.match(xml, />42</)
  assert.match(xml, /tail/)
})
