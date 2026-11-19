/**
 * The current xlsx patcher materializes worksheet XML as JavaScript strings.
 * Rewriting a large entry can therefore require several times its uncompressed
 * size in the Electron main process. Automatic recovery is best-effort and
 * must not make an otherwise responsive workbook crash in the background.
 */
export const MAX_AUTOMATIC_RECOVERY_WORKSHEET_XML_BYTES = 64 * 1024 * 1024

export function allowsAutomaticWorkbookRecovery(
  sheets: readonly { readonly sourceXmlBytes?: number | undefined }[],
): boolean {
  return sheets.every(
    (sheet) =>
      sheet.sourceXmlBytes === undefined ||
      sheet.sourceXmlBytes <= MAX_AUTOMATIC_RECOVERY_WORKSHEET_XML_BYTES,
  )
}
