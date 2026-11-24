import JSZip from 'jszip'
import { BUILT_IN_TABLE_NORMAL_XML, BUILT_IN_TABLE_STYLES } from './builtin-table-styles'
import { parseStyles } from './parse-styles'
import type { StyleInfo, ThemeColors } from './types'

/**
 * The built-in table styles read as a document would read them, for previews
 * of styles the document does not define yet (the Table Design gallery).
 */
export async function builtInTableStyleInfos(
  theme?: ThemeColors | null,
): Promise<Map<string, StyleInfo>> {
  const zip = new JSZip()
  zip.file(
    'word/styles.xml',
    '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      BUILT_IN_TABLE_NORMAL_XML +
      BUILT_IN_TABLE_STYLES.map((style) => style.xml).join('') +
      '</w:styles>',
  )
  const { styles } = await parseStyles(zip, theme)
  styles.delete('TableNormal')
  return styles
}
