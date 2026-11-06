/**
 * The ceilings a delimited import (a .csv/.tsv opened as a workbook) may
 * reach.
 *
 * The import streams the grid and the worksheet (see csv-import), so what a
 * file costs is the compression of its archive rather than its size in
 * memory. The ceiling is still there — a text file of any size converts into a
 * workbook of up to a million rows — but it is the user's to set, and it
 * starts where the streaming import can afford it.
 */

/** the steps the setting offers, in MB */
export const CSV_IMPORT_LIMIT_CHOICES_MB = [64, 128, 256, 512, 1024] as const

/** the ceiling a delimited import starts at, in MB */
export const CSV_IMPORT_DEFAULT_LIMIT_MB = 256
